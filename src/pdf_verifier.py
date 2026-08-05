from __future__ import annotations

import json
import re
from pathlib import Path

import fitz

from .config import PDF_VERIFICATION_OUTPUT_NAME, ROOT


FORBIDDEN_PDF_TEXT = [
    "本轮优化",
    "Version 2",
    "source/",
    "output/",
    "新增 26 项",
    "冲突降为 0",
]


def verify_pdf(pdf_path: Path, log_path: Path) -> dict[str, object]:
    document = fitz.open(pdf_path)
    render_dir = ROOT / ".cache" / "pdf_v3" / "pages"
    render_dir.mkdir(parents=True, exist_ok=True)
    page_ratios: list[float] = []
    text_lengths: list[int] = []
    image_heights_cm: dict[str, float] = {}
    full_text: list[str] = []
    rendered_pages: list[str] = []

    for page_number, page in enumerate(document, start=1):
        page_ratios.append(round(page.rect.width / page.rect.height, 6))
        text = page.get_text()
        full_text.append(text)
        text_lengths.append(len(text.strip()))
        image_info = page.get_image_info()
        heights = [(item["bbox"][3] - item["bbox"][1]) / 72 * 2.54 for item in image_info]
        image_heights_cm[str(page_number)] = round(max(heights, default=0.0), 2)
        pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
        rendered = render_dir / f"page-{page_number:02d}.png"
        pixmap.save(rendered)
        rendered_pages.append(str(rendered.relative_to(ROOT)))

    combined_text = "\n".join(full_text)
    forbidden_hits = [term for term in FORBIDDEN_PDF_TEXT if term in combined_text]
    log_text = log_path.read_text(encoding="utf-8", errors="replace") if log_path.exists() else ""
    layout_warnings = [
        line.strip()
        for line in log_text.splitlines()
        if re.search(r"Overfull \\hbox|Overfull \\vbox", line)
    ]
    latex_warnings = [line.strip() for line in log_text.splitlines() if "LaTeX Warning:" in line]
    ratio_ok = all(abs(ratio - 16 / 9) < 0.01 for ratio in page_ratios)
    text_ok = all(length > 20 for length in text_lengths)
    checks = {
        "page_count_15": document.page_count == 15,
        "aspect_ratio_16_9": ratio_ok,
        "all_pages_have_text": text_ok,
        "forbidden_text_absent": not forbidden_hits,
        "no_overfull_boxes": not layout_warnings,
        "all_pages_rendered": len(rendered_pages) == 15,
    }
    result: dict[str, object] = {
        "status": "PASS" if all(checks.values()) else "FAIL",
        "file": pdf_path.name,
        "checks": checks,
        "page_count": document.page_count,
        "page_ratios": page_ratios,
        "text_lengths": text_lengths,
        "image_heights_cm": image_heights_cm,
        "forbidden_hits": forbidden_hits,
        "layout_warnings": layout_warnings,
        "latex_warning_count": len(latex_warnings),
        "rendered_pages": rendered_pages,
    }
    output = pdf_path.parent / PDF_VERIFICATION_OUTPUT_NAME
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result
