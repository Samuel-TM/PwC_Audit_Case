#!/usr/bin/env python3
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.config import OUTPUT_DIR, REPORT_TEX, ROOT


def export_pdf() -> Path:
    """用 XeLaTeX 直接生成汇报 PDF，不读取或修改 PPTX。"""
    if not REPORT_TEX.exists():
        raise FileNotFoundError(REPORT_TEX)
    xelatex = shutil.which("xelatex")
    if not xelatex:
        raise RuntimeError("xelatex not available")
    env = os.environ.copy()
    env["FONTCONFIG_FILE"] = str(ROOT / "assets" / "fonts" / "fonts.conf")
    latex_dir = ROOT / ".cache" / "latex"
    latex_dir.mkdir(parents=True, exist_ok=True)
    command = [xelatex, "-interaction=nonstopmode", "-halt-on-error", f"-output-directory={latex_dir}", str(REPORT_TEX)]
    for _ in range(2):
        completed = subprocess.run(command, cwd=ROOT, check=False, env=env, capture_output=True, text=True)
        if completed.returncode != 0:
            log = latex_dir / "A公司审计案例汇报.log"
            detail = log.read_text(encoding="utf-8", errors="replace")[-6000:] if log.exists() else completed.stdout[-6000:]
            raise RuntimeError(f"XeLaTeX compilation failed:\n{detail}")
    built_pdf = latex_dir / "A公司审计案例汇报.pdf"
    if not built_pdf.exists():
        raise RuntimeError("XeLaTeX did not create PDF")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(built_pdf, OUTPUT_DIR / "A公司审计案例汇报.pdf")
    shutil.copy2(REPORT_TEX, OUTPUT_DIR / "A公司审计案例汇报.tex")
    pdf = OUTPUT_DIR / "A公司审计案例汇报.pdf"
    return pdf


if __name__ == "__main__":
    print(export_pdf())
