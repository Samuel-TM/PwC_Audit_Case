#!/usr/bin/env python3
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.config import OUTPUT_DIR, ROOT


def export_pdf() -> Path:
    pptx = OUTPUT_DIR / "A公司审计案例汇报.pptx"
    if not pptx.exists():
        raise FileNotFoundError(pptx)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice/soffice not available")
    env = os.environ.copy()
    env["FONTCONFIG_FILE"] = str(ROOT / "assets" / "fonts" / "fonts.conf")
    profile = ROOT / "tmp" / "lo_profile"
    profile.mkdir(parents=True, exist_ok=True)
    subprocess.run([
        soffice, "--headless", f"-env:UserInstallation=file://{profile}",
        "--convert-to", "pdf", "--outdir", str(OUTPUT_DIR), str(pptx)
    ], check=True, env=env, capture_output=True, text=True)
    pdf = OUTPUT_DIR / "A公司审计案例汇报.pdf"
    if not pdf.exists():
        raise RuntimeError("PDF conversion did not create output")
    return pdf


if __name__ == "__main__":
    print(export_pdf())
