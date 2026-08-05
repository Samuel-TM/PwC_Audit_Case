from dataclasses import asdict, dataclass

import pandas as pd

from .config import TOLERANCE
from .loaders import load_csv
from .metrics import gross_margin


@dataclass
class Check:
    name: str
    status: str
    max_difference: float
    note: str


def _check(name: str, differences: list[float], note: str = "") -> Check:
    maximum = max(abs(value) for value in differences) if differences else 0.0
    return Check(name, "PASS" if maximum <= TOLERANCE + 1e-9 else "FAIL", round(maximum, 4), note)


def validate_all() -> dict:
    checks: list[Check] = []
    years = [2020, 2021, 2022]
    bs = load_csv("balance_sheet.csv")
    inc = load_csv("income_statement.csv")

    def b(item: str, year: int) -> float:
        return float(bs.loc[bs.item == item, str(year)].iloc[0])

    def i(item: str, year: int, missing_zero: bool = False) -> float:
        value = inc.loc[inc.item == item, str(year)].iloc[0]
        if pd.isna(value) and missing_zero:
            return 0.0
        return float(value)

    checks.append(_check("资产总计=负债和所有者权益总计", [b("资产总计", y) - b("负债和所有者权益总计", y) for y in years]))

    current_assets = ["货币资金", "应收票据", "应收账款", "应收款项融资", "预付款项", "其他应收款", "存货", "其他流动资产"]
    current_liabilities = ["短期借款", "交易性金融负债", "应付票据", "应付账款", "合同负债", "应付职工薪酬", "应交税费", "其他应付款", "一年内到期的非流动负债", "其他流动负债"]
    checks.append(_check("流动资产合计", [sum(0.0 if pd.isna(bs.loc[bs.item == item, str(y)].iloc[0]) else float(bs.loc[bs.item == item, str(y)].iloc[0]) for item in current_assets) - b("流动资产合计", y) for y in years]))
    checks.append(_check("流动负债合计", [sum(0.0 if pd.isna(bs.loc[bs.item == item, str(y)].iloc[0]) else float(bs.loc[bs.item == item, str(y)].iloc[0]) for item in current_liabilities) - b("流动负债合计", y) for y in years]))

    operating = []
    total_profit = []
    net_profit = []
    for y in years:
        operating.append(i("营业总收入", y) - i("营业总成本", y) + i("其他收益", y) + i("公允价值变动收益", y, True) + i("信用减值损失", y) + i("资产减值损失", y) - i("营业利润", y))
        total_profit.append(i("营业利润", y) + i("营业外收入", y, True) - i("营业外支出", y) - i("利润总额", y))
        net_profit.append(i("利润总额", y) - i("所得税费用", y) - i("净利润", y))
    checks.append(_check("营业利润勾稽", operating))
    checks.append(_check("利润总额勾稽", total_profit))
    checks.append(_check("净利润勾稽", net_profit))

    product = load_csv("revenue_by_product.csv")
    region = load_csv("revenue_by_region.csv")
    quarter = load_csv("revenue_by_quarter.csv")
    cost = load_csv("cost_structure.csv")
    revenue_class = {2020: (18836.61, 836.64, 19673.24), 2021: (25734.65, 320.24, 26054.89), 2022: (26020.41, 353.09, 26373.49)}
    checks.append(_check("主营业务+其他业务=营业收入", [main + other - total for main, other, total in revenue_class.values()], "2020及2022存在0.01万元四舍五入差异"))
    for table_name, table in [("产品收入", product), ("区域收入", region), ("季度收入", quarter), ("成本构成", cost)]:
        diffs = []
        pct_diffs = []
        detail = table.iloc[:-1]
        total = table.iloc[-1]
        for y in years:
            diffs.append(detail[f"{y}_amount"].sum() - total[f"{y}_amount"])
            pct_diffs.append(detail[f"{y}_pct"].sum() - 1.0)
        checks.append(_check(f"{table_name}金额合计", diffs))
        checks.append(_check(f"{table_name}比例合计", [d * 100 for d in pct_diffs], "单位：百分点"))

    inventory = load_csv("inventory.csv")
    inventory_diffs: list[float] = []
    inventory_total_diffs: list[float] = []
    for y in years:
        subset = inventory[(inventory.year == y) & (inventory.item != "合计")].copy()
        subset["allowance"] = subset["allowance"].fillna(0.0)
        inventory_diffs.extend((subset.gross - subset.allowance - subset.net).tolist())
        total = inventory[(inventory.year == y) & (inventory.item == "合计")].iloc[0]
        inventory_total_diffs.extend([subset.gross.sum() - total.gross, subset.allowance.sum() - total.allowance, subset.net.sum() - total.net])
    checks.append(_check("存货账面余额-跌价准备=账面价值", inventory_diffs, "个别项目存在0.01万元四舍五入差异"))
    checks.append(_check("存货分类合计", inventory_total_diffs))

    margin = load_csv("gross_margin.csv")
    main_margin = margin[margin.item == "主营业务合计"].iloc[0]
    margin_diffs = []
    for y in years:
        revenue = float(product.loc[product["product"] == "合计", f"{y}_amount"].iloc[0])
        main_cost = float(cost.loc[cost.item == "合计", f"{y}_amount"].iloc[0])
        margin_diffs.append((gross_margin(revenue, main_cost) - float(main_margin[f"{y}_margin"])) * 100)
    checks.append(_check("主营业务毛利率重算", margin_diffs, "单位：百分点"))

    zeros_in_source = 0
    for name in ["income_statement.csv", "balance_sheet.csv", "inventory.csv", "receivables_aging.csv"]:
        frame = load_csv(name)
        numeric = frame.select_dtypes(include="number")
        zeros_in_source += int((numeric == 0).sum().sum())
    checks.append(Check("缺失值与0不混用", "PASS" if zeros_in_source == 0 else "FAIL", float(zeros_in_source), "源表中的'-'保留为空值"))

    return {
        "status": "PASS" if all(c.status == "PASS" for c in checks) else "FAIL",
        "tolerance": TOLERANCE,
        "checks": [asdict(c) for c in checks],
    }


def assert_valid() -> dict:
    report = validate_all()
    failures = [c for c in report["checks"] if c["status"] != "PASS"]
    if failures:
        raise AssertionError(f"data validation failed: {failures}")
    return report

