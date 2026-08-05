from .loaders import load_csv
from .metrics import current_ratio, days_sales_outstanding, debt_ratio, growth_rate, inventory_turnover, receivables_to_revenue


def calculate_metrics() -> dict:
    inc = load_csv("income_statement.csv").set_index("item")
    bs = load_csv("balance_sheet.csv").set_index("item")
    product = load_csv("revenue_by_product.csv").set_index("product")
    quarter = load_csv("revenue_by_quarter.csv").set_index("quarter")
    purchases = load_csv("raw_material_purchases.csv").set_index("item")
    revenue = {y: float(inc.loc["营业收入", str(y)]) for y in [2020, 2021, 2022]}
    net_profit = {y: float(inc.loc["净利润", str(y)]) for y in [2020, 2021, 2022]}
    ar = {y: float(bs.loc["应收账款", str(y)]) for y in [2020, 2021, 2022]}
    inv = {y: float(bs.loc["存货", str(y)]) for y in [2020, 2021, 2022]}
    cost = float(inc.loc["营业成本", "2022"])
    return {
        "revenue_growth_2021": growth_rate(revenue[2021], revenue[2020]),
        "revenue_growth_2022": growth_rate(revenue[2022], revenue[2021]),
        "profit_growth_2022": growth_rate(net_profit[2022], net_profit[2021]),
        "net_margin_2022": net_profit[2022] / revenue[2022],
        "receivables_to_revenue_2022": receivables_to_revenue(ar[2022], revenue[2022]),
        "debt_ratio_2022": debt_ratio(float(bs.loc["负债合计", "2022"]), float(bs.loc["资产总计", "2022"])),
        "current_ratio_2022": current_ratio(float(bs.loc["流动资产合计", "2022"]), float(bs.loc["流动负债合计", "2022"])),
        "inventory_turnover_2022": inventory_turnover(cost, (inv[2021] + inv[2022]) / 2),
        "dso_2022": days_sales_outstanding((ar[2021] + ar[2022]) / 2, revenue[2022]),
        "prepayment_growth_2022": growth_rate(float(bs.loc["预付款项", "2022"]), float(bs.loc["预付款项", "2021"])),
        "other_current_assets_growth_2022": growth_rate(float(bs.loc["其他流动资产", "2022"]), float(bs.loc["其他流动资产", "2021"])),
        "server_revenue_growth_2022": growth_rate(float(product.loc["服务器滑轨", "2022_amount"]), float(product.loc["服务器滑轨", "2021_amount"])),
        "q4_share_2021": float(quarter.loc["第四季度", "2021_pct"]),
        "q4_share_2022": float(quarter.loc["第四季度", "2022_pct"]),
        "plate_plastic_purchase_share_2022": float(purchases.loc["板材", "2022_pct"]) + float(purchases.loc["塑料组件", "2022_pct"]),
    }


def build_summary(checklist_result: dict | None = None) -> str:
    m = calculate_metrics()
    priority_counts = (checklist_result or {}).get("priority_counts", {"高": 0, "中": 0, "低": 0})
    checklist_total = sum(priority_counts.values()) or 101
    high_count = priority_counts.get("高", 0)
    medium_count = priority_counts.get("中", 0)
    low_count = priority_counts.get("低", 0)
    return f"""# A 公司审计案例分析摘要

## 执行摘要

A 公司主营家具五金配件，家具滑轨为主要产品；服务器滑轨 2022 年收入同比增长 {m['server_revenue_growth_2022']:.2%}。板材与塑料组件占 2022 年原材料采购的 {m['plate_plastic_purchase_share_2022']:.2%}，外销占比为 23.61%。董事长兼总经理，财务团队共 5 人。2022 年营业收入为 26,373.49 万元，同比增长 {m['revenue_growth_2022']:.2%}；净利润 3,013.75 万元，同比增长 {m['profit_growth_2022']:.2%}。主营业务毛利率由 2020 年的 23.40% 降至 2022 年的 21.38%，服务器滑轨毛利率同期由 32.20% 降至 22.87%。这些变化是风险信号，不是已确认错报。

综合金额、波动、估计复杂度、舞弊风险与证据取得难度，最需要投入审计资源的三项为营业收入、应收账款和存货；营业成本与毛利率为中高风险。预付款项、其他流动资产、借款及股本变动需要定向核查。

## 关键指标

| 指标 | 结果 | 审计含义 |
|---|---:|---|
| 2021 年收入增长 | {m['revenue_growth_2021']:.2%} | 高增长年份是趋势对比基准 |
| 2022 年收入增长 | {m['revenue_growth_2022']:.2%} | 增速较 2021 年明显放缓，需结合交易证据判断 |
| 2022 年第四季度收入占比 | {m['q4_share_2022']:.2%} | 低于 2021 年 {m['q4_share_2021']:.2%}；截止风险来自收入重大性、外销条款和期末跨期可能性，不以四季度集中为依据 |
| 2022 年净利润增长 | {m['profit_growth_2022']:.2%} | 利润基本持平，需结合毛利和减值分析 |
| 2022 年净利率 | {m['net_margin_2022']:.2%} | 与收入、成本及减值项目交叉分析 |
| 应收账款/收入 | {m['receivables_to_revenue_2022']:.2%} | 余额重大，函证、期后回款与坏账估计是重点 |
| 流动比率 | {m['current_ratio_2022']:.2f} | 短期偿债缓冲有限，关注资金预测 |
| 资产负债率 | {m['debt_ratio_2022']:.2%} | 负债水平较高，核对借款和流动性披露 |
| 存货周转率 | {m['inventory_turnover_2022']:.2f} 次 | 结合库龄、盘点和跌价准备判断 |
| 应收账款周转天数 | {m['dso_2022']:.1f} 天 | 结合信用政策和期后回款评估可收回性 |
| 预付款项同比 | {m['prepayment_growth_2022']:.2%} | 增幅大，检查交易背景与期后到货 |
| 其他流动资产同比 | {m['other_current_assets_growth_2022']:.2%} | 增幅大，检查构成、分类和结转 |

## 计划阶段拟重点审计事项与证据路径

1. **营业收入**：销售与收款流程、合同/订单/发票/出库/物流/验收证据链、客户函证、期后回款、年末截止、外销报关与提单。
2. **应收账款**：总账与账龄核对、主要及异常余额函证、期后回款、争议款项、坏账准备模型和敏感性。
3. **存货**：年末监盘、账实抽盘、委托加工存货；如存在异地存货，取得第三方确认或安排现场程序；并检查成本归集、库龄、期后售价及入出库截止。
4. **成本与毛利率**：成本核算方法、领料和工时、制造费用分配、单位成本重算、产品/客户/月度毛利率下钻。
5. **其他定向项目**：预付款与其他流动资产的合同、付款、发票、期后结转；借款与银行函证；股本及资本公积变动的批准文件和工商登记。

## 审计计划与交付概况

- 审计程序 30 项；
- 资料需求 {checklist_total} 项，分为 9 个业务模块；其中高优先级 {high_count} 项、中优先级 {medium_count} 项、低优先级 {low_count} 项；
- 项目合伙人与质量复核人（如适用）已分设，重大判断不得由编制人自我复核；
- 重要性、抽样、风险登记和控制了解模板已建立；
- LaTeX 汇报由 XeLaTeX 自动生成 15 页 PDF；
- 当前仍处于计划阶段，不包含真实执行结果。

## 数据质量与边界

- 关键表格已按原 PDF 人工核对并结构化录入；所有金额统一为万元。
- 自动校验覆盖资产负债表平衡、流动项目合计、利润勾稽、收入/成本分类、存货净额、比例和毛利率。
- 2020 年及 2022 年“主营业务+其他业务=营业收入”、个别存货分类存在不超过 0.01 万元的四舍五入差异，低于 0.02 容差。
- 案例未提供交易级明细、重要性水平、内控测试、函证或盘点结果，因此材料仅设计风险应对，不形成真实审计结论或审计意见。

来源：`1-2 Assignment for PwC You Plus.pdf`，第 1—10 页。
"""
