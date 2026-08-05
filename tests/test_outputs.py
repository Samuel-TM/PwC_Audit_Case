import json
import re
from datetime import date, datetime

import fitz
from openpyxl import load_workbook

from src.config import (
    CHART_DIR,
    CHECKLIST_OUTPUT_NAME,
    OUTPUT_DIR,
    PDF_VERIFICATION_OUTPUT_NAME,
    PLAN_OUTPUT_NAME,
    QA_OUTPUT_NAME,
    REPORT_OUTPUT_STEM,
    RISK_OUTPUT_NAME,
    SUMMARY_OUTPUT_NAME,
    WORKBOOK_VERIFICATION_OUTPUT_NAME,
)


def test_v3_charts_exist_and_dimensions():
    from PIL import Image

    charts = sorted(CHART_DIR.glob("*.png"))
    assert len(charts) == 9
    assert (CHART_DIR / "06_receivables_net_ratio.png").exists()
    assert (CHART_DIR / "07_inventory_composition_2022.png").exists()
    for chart in charts:
        assert Image.open(chart).size == (1600, 900)


def test_v3_workbook_sheet_names():
    checklist = load_workbook(OUTPUT_DIR / CHECKLIST_OUTPUT_NAME, read_only=False)
    assert checklist.sheetnames == ["总览", "通用资料", "收入", "应收账款", "存货", "采购与成本", "资金与借款", "其他重点项目", "其他报表项目", "费用税务与披露"]
    assert checklist["总览"].freeze_panes == "A6"
    assert checklist["收入"].freeze_panes == "A2"
    plan = load_workbook(OUTPUT_DIR / PLAN_OUTPUT_NAME, read_only=False)
    assert plan.sheetnames == ["项目管理看板", "项目参数", "总体时间表", "人员分工", "审计程序", "问题跟踪", "函证跟踪", "盘点安排", "重要性与抽样", "风险登记表", "控制了解"]
    assert plan["总体时间表"].freeze_panes == "A5"
    assert plan["审计程序"].freeze_panes == "D2"


def test_v3_latex_pdf_page_size_fonts_and_forbidden_text():
    tex = (OUTPUT_DIR / f"{REPORT_OUTPUT_STEM}.tex").read_text(encoding="utf-8")
    assert tex.count(r"\begin{frame}") == 15
    assert r"\fontsize{17.5}{20.5}" in tex
    assert r"\newcommand{\bodyfont}{\fontsize{10.8}{13.2}" in tex
    assert r"\newcommand{\smallbody}{\fontsize{9.8}{12.0}" in tex
    assert r"\newcommand{\tinybody}{\fontsize{8.5}{10.4}" in tex
    assert r"\newcommand{\sourcefont}{\fontsize{7.0}{8.3}" in tex
    assert r"\newcommand{\closingtitlefont}{\fontsize{31}{35}" in tex
    assert "assets/derived/PwC_logo_cropped.png" in tex
    assert "assets/derived/PwC_2025_Logo.pdf" in tex
    pdf = fitz.open(OUTPUT_DIR / f"{REPORT_OUTPUT_STEM}.pdf")
    assert pdf.page_count == 15
    assert all(abs(page.rect.width / page.rect.height - 16 / 9) < 0.01 for page in pdf)
    assert all(len(page.get_text().strip()) > 20 for page in pdf)
    text = "\n".join(page.get_text() for page in pdf)
    for forbidden in ["本轮优化", "Version 2", "source/", "output/", "新增 26 项", "冲突降为 0"]:
        assert forbidden not in text


def test_v3_chart_heights_and_page_specific_content():
    tex = (OUTPUT_DIR / f"{REPORT_OUTPUT_STEM}.tex").read_text(encoding="utf-8")
    expected = {
        "01_revenue_net_profit.png": "4.40cm",
        "02_receivables_inventory.png": "5.00cm",
        "08_audit_risk_heatmap.png": "4.70cm",
        "05_export_share.png": "3.00cm",
        "06_receivables_net_ratio.png": "3.00cm",
        "07_inventory_composition_2022.png": "3.00cm",
        "04_product_gross_margin.png": "3.00cm",
        "09_audit_plan_gantt.png": "5.00cm",
    }
    for chart, height in expected.items():
        assert re.search(rf"height={re.escape(height)}[^\n]+{re.escape(chart)}", tex)
    assert "净利润基本持平，主营业务毛利率继续下降" in tex
    assert "售价下降、单位成本上升是分析方向" in tex


def test_plan_governance_coverage_and_formula_driven_risk_mapping():
    plan = load_workbook(OUTPUT_DIR / PLAN_OUTPUT_NAME, data_only=False)
    procedures = plan["审计程序"]
    headers = [cell.value for cell in procedures[1]]
    assert headers[:10] == ["工作编号", "风险 ID", "审计模块", "风险描述", "审计认定", "审计程序", "负责人", "复核人", "风险等级", "程序执行优先级"]
    rows = list(procedures.iter_rows(min_row=2, values_only=True))
    assert len(rows) == 30
    assert all(row[6] != row[7] for row in rows)
    assert all(str(row[8]).startswith("=IFERROR(INDEX('风险登记表'!") for row in rows)
    assert all(str(row[9]).startswith("=IF(OR(I") for row in rows)
    risk_register = plan["风险登记表"]
    valid_ids = {risk_register.cell(row, 1).value for row in range(2, 10)}
    assert all(row[1] in valid_ids for row in rows)
    mapped = {row[0]: row[1] for row in rows}
    assert mapped["P-01"] == mapped["OCA-01"] == "R-PREOCA"
    assert mapped["B-01"] == mapped["GC-01"] == "R-DEBTGC"


def test_project_mode_overdue_and_dashboard_formulas():
    plan = load_workbook(OUTPUT_DIR / PLAN_OUTPUT_NAME, data_only=False)
    params = plan["项目参数"]
    assert params["B2"].value == "计划"
    assert params["B3"].value in {date(2023, 1, 10), datetime(2023, 1, 10)}
    procedures = plan["审计程序"]
    for row in range(2, procedures.max_row + 1):
        formula = procedures.cell(row, 24).value
        assert "'项目参数'!$B$2=\"计划\"" in formula
        assert "'项目参数'!$B$3" in formula
        assert "TODAY()" not in formula
    dashboard = plan["项目管理看板"]
    assert "'审计程序'!$R$2:$R$" in dashboard["B4"].value
    assert "COUNTIFS" in dashboard["E4"].value and "'审计程序'!$Z$2:$Z$" in dashboard["E4"].value
    assert "'审计程序'!$Y$2:$Y$" in dashboard["B5"].value
    assert "'审计程序'!$X$2:$X$" in dashboard["E5"].value
    assert dashboard["B6"].value == "=IFERROR('[审计资料清单_V3.xlsx]总览'!$G$15,0)"
    assert "/COUNT('函证跟踪'!$F$2:$F$21)" in dashboard["E6"].value
    assert "/COUNTIF('盘点安排'!$B$2:$B$13,\"<>\")" in dashboard["B7"].value


def test_materiality_risk_and_control_templates():
    plan = load_workbook(OUTPUT_DIR / PLAN_OUTPUT_NAME, data_only=False)
    materiality = plan["重要性与抽样"]
    assert materiality["B5"].value.startswith("=IF(")
    assert "待项目经理" in materiality["B2"].value
    risk = plan["风险登记表"]
    assert risk["A1"].value == "风险 ID"
    assert risk["I2"].value == "=AVERAGE(C2:H2)"
    assert risk["M2"].value == '=IF(K2="",J2,K2)'
    assert risk["A9"].value == "R-BASE"
    controls = plan["控制了解"]
    assert controls.max_row == 7
    assert {controls.cell(row, 1).value for row in range(2, 8)} == {"销售与收款", "采购与付款", "生产与存货", "资金", "财务结账", "主数据和系统权限"}


def test_checklist_full_scope_and_historical_date_logic():
    checklist = load_workbook(OUTPUT_DIR / CHECKLIST_OUTPUT_NAME, data_only=False)
    assert checklist["其他报表项目"].max_row - 1 >= 13
    assert checklist["费用税务与披露"].max_row - 1 >= 13
    assert len(checklist["总览"].conditional_formatting) >= 1
    assert checklist["总览"]["K2"].value == "计划"
    overview = checklist["总览"]
    assert overview.tables["ChecklistOverview"].ref == "A5:H15"
    assert overview["A5"].value == "模块"
    assert overview["G6"].value == "=IF(B6=0,0,D6/B6)"
    assert overview["B15"].value == "=SUM(B6:B14)"
    for sheet_name in checklist.sheetnames[1:]:
        rules = checklist[sheet_name].conditional_formatting._cf_rules.values()
        formulas = [formula for entries in rules for entry in entries for formula in entry.formula]
        assert any("'总览'!$K$2=\"执行\"" in formula for formula in formulas)
        assert all("TODAY()" not in formula for formula in formulas)


def test_plan_validations_and_conditional_formats_follow_v3_columns():
    plan = load_workbook(OUTPUT_DIR / PLAN_OUTPUT_NAME, data_only=False)
    procedures = plan["审计程序"]
    assert len(procedures.data_validations.dataValidation) == 3
    validation_ranges = {str(item.sqref) for item in procedures.data_validations.dataValidation}
    assert any(value.startswith("R2:R") for value in validation_ranges)
    assert any(value.startswith("W2:W") for value in validation_ranges)
    assert any(value.startswith("Z2:Z") for value in validation_ranges)
    procedure_rules = sum(len(rules) for rules in procedures.conditional_formatting._cf_rules.values())
    assert procedure_rules >= 4


def test_qa_numbering_and_required_answers():
    text = (OUTPUT_DIR / QA_OUTPUT_NAME).read_text(encoding="utf-8")
    assert text.count("**30 秒回答：**") == 30
    assert "## 17. 本案例的重要性水平是多少？" in text
    assert "## 18. 风险等级如何确定并同步？" in text
    assert "## 23. 审计团队如何分工？" in text
    assert "## 24. 为什么安排四周，如何判断进度？" in text
    assert "## 25. 控制了解覆盖哪些流程？" in text
    assert "谁复核项目经理的重大判断" in text


def test_cross_file_key_numbers_and_verification_outputs():
    pdf = fitz.open(OUTPUT_DIR / f"{REPORT_OUTPUT_STEM}.pdf")
    pdf_text = "\n".join(page.get_text() for page in pdf)
    summary = (OUTPUT_DIR / SUMMARY_OUTPUT_NAME).read_text(encoding="utf-8")
    for value in ["1.22%", "34.35%", "21.38%", "23.61%", "648.33", "4,109.52", "4,504.14", "101", "30"]:
        assert value in pdf_text or value in summary
    workbook_report = json.loads((OUTPUT_DIR / WORKBOOK_VERIFICATION_OUTPUT_NAME).read_text(encoding="utf-8"))
    assert workbook_report["checklist"]["total_requests"] == 101
    assert workbook_report["plan"]["procedure_count"] == 30
    assert workbook_report["plan"]["self_review_conflicts"] == []
    risks = json.loads((OUTPUT_DIR / RISK_OUTPUT_NAME).read_text(encoding="utf-8"))
    assert {item["risk_id"] for item in risks} >= {"R-PREOCA", "R-DEBTGC"}
    pdf_report = json.loads((OUTPUT_DIR / PDF_VERIFICATION_OUTPUT_NAME).read_text(encoding="utf-8"))
    assert pdf_report["status"] == "PASS"
    assert pdf_report["checks"]["all_pages_rendered"] is True


def test_build_source_does_not_report_pptx_as_a_delivery():
    build_source = (OUTPUT_DIR.parent / "scripts" / "build_all.py").read_text(encoding="utf-8")
    assert "未修改 PPTX" not in build_source
    assert 'A公司审计案例汇报.pptx' not in build_source
