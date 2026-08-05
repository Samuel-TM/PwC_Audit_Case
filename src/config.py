from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
CHART_DIR = OUTPUT_DIR / "charts"
PREVIEW_DIR = OUTPUT_DIR / "previews"
BUILD_DIR = ROOT / ".cache" / "artifact_build"
REPORT_DIR = ROOT / "report"
REPORT_TEX = REPORT_DIR / "A公司审计案例汇报.tex"
SOURCE_PDF = ROOT / "source" / "1-2 Assignment for PwC You Plus.pdf"
TOLERANCE = 0.02


def ensure_directories() -> None:
    for path in (OUTPUT_DIR, CHART_DIR, PREVIEW_DIR, BUILD_DIR, REPORT_DIR):
        path.mkdir(parents=True, exist_ok=True)
