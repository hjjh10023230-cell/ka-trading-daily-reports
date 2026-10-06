# KA Trading Daily Reports

KA 交易趋势日报中心，用于按交易日归档聚合、脱敏后的市场观察。

## 在线入口

- 报告中心：`https://hjjh10023230-cell.github.io/ka-trading-daily-reports/`
- 最新一期：`https://hjjh10023230-cell.github.io/ka-trading-daily-reports/latest.html`
- 历史日报：从报告中心进入，或访问对应日期目录

## 目录规则

```text
reports/<年份>/<YYYY-MM-DD>/index.html
```

## 每日更新

1. 原始 Excel 仅保存在仓库外，不得复制进仓库。
2. 生成并复核当日报告，按日期放入独立目录。
3. 更新 `reports.json`、首页卡片和 `latest.html`。
4. 发布前检查聚合口径、低样本披露与敏感信息。
5. 提交并推送到 `main`，GitHub Pages 自动更新。

## 三段式框架

1. 整体交易情绪与方向性成交资金；
2. 人群分化与标的轮动；
3. 交易和当日行情的关系，以及下一交易日的运营动作。

## 口径说明

- `BUY 金额 − SELL 金额` 只是方向性成交金额差，不等于账户净入金或真实资金净流入。
- 未具备成交时间、持仓成本和账户现金变化时，只做日级行情关联，不判断具体盘中买卖点。
- 公开报告仅保留汇总、降精度信息，不上传 UID、账户、姓名、联系方式、内部链接、原始数据或可识别的单客户明细。
- `robots=noindex` 不能替代访问控制；本仓库内容仍属于公开信息。
