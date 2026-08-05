from src.risk_engine import evaluate_risks, risk_level


def test_risk_level_mapping():
    assert risk_level(4.0) == "高"
    assert risk_level(3.0) == "中高"
    assert risk_level(2.0) == "中"
    assert risk_level(1.99) == "一般"


def test_preset_high_risks():
    risks = {risk.area: risk.level for risk in evaluate_risks()}
    assert risks["营业收入"] == "高"
    assert risks["应收账款"] == "高"
    assert risks["存货"] == "高"
    assert risks["营业成本与毛利率"] == "中高"

