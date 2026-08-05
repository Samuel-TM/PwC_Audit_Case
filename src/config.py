from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
CHART_DIR = OUTPUT_DIR / "charts_V3"
PREVIEW_DIR = OUTPUT_DIR / "previews"
BUILD_DIR = ROOT / ".cache" / "artifact_build"
REPORT_DIR = ROOT / "report"
REPORT_TEX = REPORT_DIR / "A公司审计案例汇报.tex"
REPORT_OUTPUT_STEM = "A公司审计案例汇报_V3"
CHECKLIST_OUTPUT_NAME = "审计资料清单_V3.xlsx"
PLAN_OUTPUT_NAME = "审计计划_V3.xlsx"
QA_OUTPUT_NAME = "导师问答手册_V3.md"
SUMMARY_OUTPUT_NAME = "analysis_summary_V3.md"
VALIDATION_OUTPUT_NAME = "validation_report_V3.json"
RISK_OUTPUT_NAME = "audit_risk_results_V3.json"
WORKBOOK_VERIFICATION_OUTPUT_NAME = "workbook_verification_V3.json"
PDF_VERIFICATION_OUTPUT_NAME = "pdf_verification_V3.json"
BUILD_REPORT_OUTPUT_NAME = "build_report_V3.txt"
SOURCE_PDF = ROOT / "source" / "1-2 Assignment for PwC You Plus.pdf"
TOLERANCE = 0.02


def ensure_directories() -> None:
    for path in (OUTPUT_DIR, CHART_DIR, PREVIEW_DIR, BUILD_DIR, REPORT_DIR):
        path.mkdir(parents=True, exist_ok=True)
