#!/usr/bin/env python3
"""Validate and aggregate a local KA trade workbook without copying it into the repo."""
from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

DIRECTIONS = {"BUY", "SELL", "BUY_BACK", "SELL_SHORT"}
KA_CROWDS = {"赌徒", "稳健老钱", "职业炒家", "高资产沉睡", "中产进取"}
LEGACY_REQUIRED = ["股票代码", "股票名称", "买卖方向", "ka类型", "用户数", "交易订单数", "trade_amt_usd"]
GROUP_FIELDS = ["分区日期", "ka类型", "买卖方向", "证券类型V2"]
METRICS = ["用户数", "交易订单数", "trade_amt_usd"]
MARKETS = ["JP", "US"]
OPTION_RE = re.compile(r"(?:\d{6}[CP]\d+|\s\d{6}\s+\d+(?:\.\d+)?[CP])$", re.I)


def bucket() -> dict[str, float | int]:
    return {"orders": 0, "amount": 0.0, "users": 0}


def number(value: Any, *, row_no: int, field: str) -> float:
    if value in (None, ""):
        return 0.0
    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"第 {row_no} 行字段 {field} 不是数字：{value!r}") from exc
    if result < 0:
        raise ValueError(f"第 {row_no} 行字段 {field} 不得为负数：{value!r}")
    return result


def excel_date(value: Any) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date().isoformat()
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").date().isoformat()
    except ValueError:
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(text))).date().isoformat()
        except ValueError as exc:
            raise ValueError(f"无法识别分区日期：{value!r}") from exc


def asset_class(security_type: str, symbol: str) -> str:
    kind = security_type.upper()
    if kind == "OPTION" or OPTION_RE.search(symbol):
        return "期权"
    if kind in {"FUND", "TRUST"} or "ETF" in symbol.upper():
        return "ETF/基金"
    if kind == "WARRANT":
        return "权证"
    return "股票"


def add(target: dict[str, float | int], users: float, orders: float, amount: float) -> None:
    target["users"] += int(users)
    target["orders"] += int(orders)
    target["amount"] += amount


def empty_result() -> dict[str, Any]:
    return {
        "directions": defaultdict(bucket),
        "crowds": defaultdict(lambda: defaultdict(bucket)),
        "assetClasses": defaultdict(bucket),
        "symbols": defaultdict(lambda: defaultdict(bucket)),
    }


def aggregate_current(ws: Any, report_date: str) -> tuple[dict[str, Any], int]:
    top = [cell.value for cell in ws[1]]
    second = [cell.value for cell in ws[2]]
    missing = [name for name in GROUP_FIELDS + ["股票名称"] if name not in second]
    if missing:
        raise ValueError(f"缺少字段：{', '.join(missing)}")

    columns: dict[tuple[str, str], int] = {}
    metric = None
    for index, market in enumerate(second):
        if index < len(top) and top[index] in METRICS:
            metric = top[index]
        if metric in METRICS and market in MARKETS:
            columns[(str(metric), str(market))] = index
    missing_metrics = [(metric, market) for metric in METRICS for market in MARKETS if (metric, market) not in columns]
    if missing_metrics:
        labels = ", ".join(f"{metric}/{market}" for metric, market in missing_metrics)
        raise ValueError(f"缺少市场指标列：{labels}")

    field_col = {name: second.index(name) for name in GROUP_FIELDS + ["股票名称"]}
    result = empty_result()
    carried: dict[str, Any] = {name: None for name in GROUP_FIELDS}
    matched_rows = 0

    for row_no, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
        values = list(row)
        for field in GROUP_FIELDS:
            value = values[field_col[field]]
            if value not in (None, ""):
                carried[field] = value
        row_date = excel_date(carried["分区日期"])
        if row_date != report_date:
            continue
        crowd = str(carried["ka类型"] or "").strip()
        if crowd not in KA_CROWDS:
            continue
        direction = str(carried["买卖方向"] or "").strip().upper()
        if direction not in DIRECTIONS:
            raise ValueError(f"第 {row_no} 行存在未知买卖方向：{direction!r}")
        security_type = str(carried["证券类型V2"] or "").strip()
        symbol = str(values[field_col["股票名称"]] or "").strip()
        if not symbol:
            raise ValueError(f"第 {row_no} 行缺少股票名称")

        users = sum(number(values[columns[("用户数", market)]], row_no=row_no, field=f"用户数/{market}") for market in MARKETS)
        orders = sum(number(values[columns[("交易订单数", market)]], row_no=row_no, field=f"交易订单数/{market}") for market in MARKETS)
        amount = sum(number(values[columns[("trade_amt_usd", market)]], row_no=row_no, field=f"trade_amt_usd/{market}") for market in MARKETS)
        if not users and not orders and not amount:
            continue

        add(result["directions"][direction], users, orders, amount)
        add(result["crowds"][crowd][direction], users, orders, amount)
        add(result["symbols"][symbol][direction], users, orders, amount)
        add(result["assetClasses"][asset_class(security_type, symbol)], users, orders, amount)
        matched_rows += 1

    return result, matched_rows


def aggregate_legacy(ws: Any, report_date: str) -> tuple[dict[str, Any], int]:
    headers = [cell.value for cell in ws[1]]
    missing = [name for name in LEGACY_REQUIRED if name not in headers]
    if missing:
        raise ValueError(f"缺少字段：{', '.join(missing)}")
    col = {name: headers.index(name) for name in LEGACY_REQUIRED}
    result = empty_result()
    carried = {"股票代码": None, "股票名称": None}
    matched_rows = 0

    for row_no, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        values = list(row)
        for field in carried:
            if values[col[field]] not in (None, ""):
                carried[field] = values[col[field]]
        crowd = str(values[col["ka类型"]] or "").strip()
        if crowd not in KA_CROWDS:
            continue
        direction = str(values[col["买卖方向"]] or "").strip().upper()
        if direction not in DIRECTIONS:
            raise ValueError(f"第 {row_no} 行存在未知买卖方向：{direction!r}")
        symbol = str(carried["股票名称"] or carried["股票代码"] or "").strip()
        if not symbol:
            raise ValueError(f"第 {row_no} 行缺少股票名称或代码")
        users = number(values[col["用户数"]], row_no=row_no, field="用户数")
        orders = number(values[col["交易订单数"]], row_no=row_no, field="交易订单数")
        amount = number(values[col["trade_amt_usd"]], row_no=row_no, field="trade_amt_usd")
        add(result["directions"][direction], users, orders, amount)
        add(result["crowds"][crowd][direction], users, orders, amount)
        add(result["symbols"][symbol][direction], users, orders, amount)
        add(result["assetClasses"][asset_class("", symbol)], users, orders, amount)
        matched_rows += 1

    return result, matched_rows


def aggregate(path: Path, report_date: str) -> dict[str, Any]:
    wb = load_workbook(path, read_only=False, data_only=True)
    if "报表" not in wb.sheetnames:
        raise ValueError("缺少工作表：报表")
    ws = wb["报表"]
    row_one = [cell.value for cell in ws[1]]
    row_two = [cell.value for cell in ws[2]]
    if "分区日期" in row_two and "ka类型" in row_two:
        result, matched_rows = aggregate_current(ws, report_date)
    elif all(name in row_one for name in LEGACY_REQUIRED):
        result, matched_rows = aggregate_legacy(ws, report_date)
    else:
        raise ValueError("无法识别报表表头；需要双层 JP/US 表头或旧版单层表头")
    if not matched_rows:
        raise ValueError(f"未找到 {report_date} 的五类 KA 交易数据")
    result["matchedRows"] = matched_rows
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path, help="仓库外的本地 Excel 路径")
    parser.add_argument("--date", required=True, help="报告交易日 YYYY-MM-DD")
    parser.add_argument("--output", type=Path, help="可选：将聚合 JSON 写入本地路径")
    args = parser.parse_args()
    try:
        report_date = datetime.strptime(args.date, "%Y-%m-%d").date().isoformat()
    except ValueError as exc:
        raise SystemExit("--date 必须是有效的 YYYY-MM-DD 日期") from exc
    if not args.workbook.is_file():
        raise SystemExit(f"找不到 Excel：{args.workbook}")
    result = {
        "reportDate": report_date,
        "scope": "五类KA，JP与US合并；排除非KA",
        "warning": "方向性成交金额差不等于真实净入金或持仓变化；用户数不可跨分组相加去重；源文件不含成交时点和持仓成本。",
        **aggregate(args.workbook, report_date),
    }
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.write_text(payload, encoding="utf-8")
    else:
        print(payload)


if __name__ == "__main__":
    main()
