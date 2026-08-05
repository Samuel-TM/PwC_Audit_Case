from pathlib import Path

import fitz
from openpyxl import load_workbook

from src.config import OUTPUT_DIR


def test_charts_exist_and_dimensions():
    from PIL import Image
    charts = sorted((OUTPUT_DIR / "charts").glob("*.png"))
    assert len(charts) == 9
    for chart in charts:
        assert Image.open(chart).size == (1600, 900)


def test_workbook_sheet_names():
    checklist = load_workbook(OUTPUT_DIR / "审计资料清单.xlsx", read_only=False)
    assert checklist.sheetnames == ["总览", "通用资料", "收入", "应收账款", "存货", "采购与成本", "资金与借款", "其他重点项目", "其他报表项目", "费用税务与披露"]
    assert checklist["总览"].freeze_panes == "A5"
    assert checklist["收入"].freeze_panes == "A2"
    plan = load_workbook(OUTPUT_DIR / "审计计划.xlsx", read_only=False)
    assert plan.sheetnames == ["项目管理看板", "总体时间表", "人员分工", "审计程序", "问题跟踪", "函证跟踪", "盘点安排", "重要性与抽样", "风险登记表", "控制了解"]
    assert plan["总体时间表"].freeze_panes == "A5"
    assert plan["审计程序"].freeze_panes == "C2"


def test_latex_and_pdf():
    tex = (OUTPUT_DIR / "A公司审计案例汇报.tex").read_text(encoding="utf-8")
    assert tex.count(r"\begin{frame}") == 15
    assert "关键审计事项 1" in tex
    assert "重要性与抽样" in tex
    pdf = fitz.open(OUTPUT_DIR / "A公司审计案例汇报.pdf")
    assert pdf.page_count == 15
    assert all(len(page.get_text().strip()) > 20 for page in pdf)


def test_plan_governance_and_coverage():
    plan = load_workbook(OUTPUT_DIR / "审计计划.xlsx", data_only=False)
    procedures = plan["审计程序"]
    headers = [cell.value for cell in procedures[1]]
    assert headers[-11:] == ["完成百分比", "状态", "证据索引", "发现的问题", "结论", "依赖资料", "资料状态", "逾期工作日", "阻塞原因", "复核状态", "最终底稿索引"]
    rows = list(procedures.iter_rows(min_row=2, values_only=True))
    assert len(rows) >= 30
    assert all(row[5] != row[6] for row in rows)
    codes = {row[0] for row in rows}
    assert {"AP-01", "FA-01", "LEASE-01", "PAY-01", "TAX-01", "EXP-01", "CF-01", "DISC-01"} <= codes
    timeline = plan["总体时间表"]
    assert all(str(timeline.cell(row, 5).value).startswith("=NETWORKDAYS") for row in range(5, timeline.max_row + 1))
    assert any("重要性" in str(timeline.cell(row, 2).value) for row in range(5, timeline.max_row + 1))


def test_materiality_risk_and_control_templates():
    plan = load_workbook(OUTPUT_DIR / "审计计划.xlsx", data_only=False)
    materiality = plan["重要性与抽样"]
    assert materiality["B5"].value.startswith("=IF(")
    assert "待项目经理" in materiality["B2"].value
    risk = plan["风险登记表"]
    assert risk["H2"].value == "=AVERAGE(B2:G2)"
    assert risk.max_row >= 8
    controls = plan["控制了解"]
    assert controls.max_row == 7
    assert {controls.cell(row, 1).value for row in range(2, 8)} == {"销售与收款", "采购与付款", "生产与存货", "资金", "财务结账", "主数据和系统权限"}
    dashboard = plan["项目管理看板"]
    assert "'审计程序'!$P$2:$P$" in dashboard["B4"].value
    assert "'审计程序'!$P$2:$P$" in dashboard["B5"].value
    assert "'审计程序'!$U$2:$U$" in dashboard["B6"].value


def test_checklist_full_statement_extension():
    checklist = load_workbook(OUTPUT_DIR / "审计资料清单.xlsx", data_only=False)
    assert checklist["其他报表项目"].max_row - 1 >= 13
    assert checklist["费用税务与披露"].max_row - 1 >= 13
    assert len(checklist["总览"].conditional_formatting) >= 1
    new_sheet_rules = sum(len(rules) for rules in checklist["其他报表项目"].conditional_formatting._cf_rules.values())
    assert new_sheet_rules >= 3


def test_plan_validations_and_conditional_formats_follow_new_columns():
    plan = load_workbook(OUTPUT_DIR / "审计计划.xlsx", data_only=False)
    procedures = plan["审计程序"]
    assert len(procedures.data_validations.dataValidation) == 4
    validation_ranges = {str(item.sqref) for item in procedures.data_validations.dataValidation}
    assert any(value.startswith("P2:P") for value in validation_ranges)
    assert any(value.startswith("U2:U") for value in validation_ranges)
    assert any(value.startswith("X2:X") for value in validation_ranges)
    procedure_rules = sum(len(rules) for rules in procedures.conditional_formatting._cf_rules.values())
    assert procedure_rules >= 4


def test_qa_count():
    text = (OUTPUT_DIR / "导师问答手册.md").read_text(encoding="utf-8")
    assert text.count("**30 秒回答：**") >= 25
