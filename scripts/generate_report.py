#!/usr/bin/env python3
"""Validate and aggregate a local KA trade workbook without copying it into the repo."""
from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

REQUIRED = ["股票代码", "股票名称", "买卖方向", "ka类型", "用户数", "交易订单数", "trade_amt_usd"]
DIRECTIONS = {"BUY", "SELL", "BUY_BACK", "SELL_SHORT"}
OPTION_RE = re.compile(r"(?:\d{6}[CP]\d+|\s\d{6}\s+\d+(?:\.\d+)?[CP])$", re.I)


def number(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    return float(value)


def asset_class(code: str, name: str) -> str:
    text = f"{code} {name}".strip()
    if OPTION_RE.search(text):
        return "期权"
    if "ETF" in text.upper():
        return "ETF"
    return "其余"


def aggregate(path: Path) -> dict[str, Any]:
    wb = load_workbook(path, read_only=False, data_only=True)
    if "报表" not in wb.sheetnames:
        raise ValueError("缺少工作表：报表")
    ws = wb["报表"]
    headers = [cell.value for cell in ws[1]]
    missing = [name for name in REQUIRED if name not in headers]
    if missing:
        raise ValueError(f"缺少字段：{', '.join(missing)}")
    col = {name: headers.index(name) for name in REQUIRED}

    # Excel merged cells visually carry the first row's symbol/name; inherit them here.
    merged_lookup: dict[tuple[int, int], Any] = {}
    for area in ws.merged_cells.ranges:
        top = ws.cell(area.min_row, area.min_col).value
        for row in range(area.min_row, area.max_row + 1):
            for column in range(area.min_col, area.max_col + 1):
                merged_lookup[(row, column)] = top

    totals = defaultdict(lambda: {"orders": 0, "amount": 0.0, "users": 0})
    crowds = defaultdict(lambda: defaultdict(lambda: {"orders": 0, "amount": 0.0, "users": 0}))
    assets = defaultdict(lambda: {"orders": 0, "amount": 0.0})
    symbols = defaultdict(lambda: defaultdict(lambda: {"orders": 0, "amount": 0.0, "users": 0}))

    for row_no, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        values = list(row)
        code = values[col["股票代码"]]
        name = values[col["股票名称"]]
        if code in (None, ""):
            code = merged_lookup.get((row_no, col["股票代码"] + 1), "")
        if name in (None, ""):
            name = merged_lookup.get((row_no, col["股票名称"] + 1), "")
        direction = str(values[col["买卖方向"]] or "").upper()
        crowd = str(values[col["ka类型"]] or "未分类")
        if direction not in DIRECTIONS:
            raise ValueError(f"第 {row_no} 行存在未知买卖方向：{direction!r}")
        users = int(number(values[col["用户数"]]))
        orders = int(number(values[col["交易订单数"]]))
        amount = number(values[col["trade_amt_usd"]])
        for bucket in (totals[direction], crowds[crowd][direction], symbols[str(code)][direction]):
            bucket["orders"] += orders
            bucket["amount"] += amount
            bucket["users"] += users
        klass = asset_class(str(code), str(name))
        assets[klass]["orders"] += orders
        assets[klass]["amount"] += amount

    return {"directions": totals, "crowds": crowds, "assetClasses": assets, "symbols": symbols}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path, help="仓库外的本地 Excel 路径")
    parser.add_argument("--date", required=True, help="报告交易日 YYYY-MM-DD；源表无日期，必须显式指定")
    parser.add_argument("--output", type=Path, help="可选：将聚合 JSON 写入本地路径")
    args = parser.parse_args()
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", args.date):
        raise SystemExit("--date 必须为 YYYY-MM-DD")
    result = {"reportDate": args.date, "warning": "方向性成交金额差不等于真实净入金；源文件不含成交时点。", **aggregate(args.workbook)}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.write_text(payload, encoding="utf-8")
    else:
        print(payload)


if __name__ == "__main__":
    main()
