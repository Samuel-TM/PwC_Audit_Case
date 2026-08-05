from pathlib import Path

import fitz
from openpyxl import load_workbook
from pptx import Presentation

from src.config import OUTPUT_DIR


def test_charts_exist_and_dimensions():
    from PIL import Image
    charts = sorted((OUTPUT_DIR / "charts").glob("*.png"))
    assert len(charts) == 9
    for chart in charts:
        assert Image.open(chart).size == (1600, 900)


def test_workbook_sheet_names():
    checklist = load_workbook(OUTPUT_DIR / "审计资料清单.xlsx", read_only=False)
    assert checklist.sheetnames == ["总览", "通用资料", "收入", "应收账款", "存货", "采购与成本", "资金与借款", "其他重点项目"]
    assert checklist["总览"].freeze_panes == "A5"
    assert checklist["收入"].freeze_panes == "A2"
    plan = load_workbook(OUTPUT_DIR / "审计计划.xlsx", read_only=False)
    assert plan.sheetnames == ["总体时间表", "人员分工", "审计程序", "问题跟踪", "函证跟踪", "盘点安排"]
    assert plan["总体时间表"].freeze_panes == "A5"
    assert plan["审计程序"].freeze_panes == "C2"


def test_presentation_and_pdf():
    ppt = Presentation(OUTPUT_DIR / "A公司审计案例汇报.pptx")
    assert len(ppt.slides) == 15
    assert all(any(shape.has_text_frame and shape.text.strip() for shape in slide.shapes) for slide in ppt.slides)
    pdf = fitz.open(OUTPUT_DIR / "A公司审计案例汇报.pdf")
    assert pdf.page_count == 15


def test_qa_count():
    text = (OUTPUT_DIR / "导师问答手册.md").read_text(encoding="utf-8")
    assert text.count("**30 秒回答：**") >= 25
