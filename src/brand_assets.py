from __future__ import annotations

from pathlib import Path

import fitz
from PIL import Image

from .config import ROOT


def prepare_brand_assets() -> list[Path]:
    """生成适合 Beamer 直接引用的紧凑 Logo 资产，不修改用户原图。"""
    source_png = ROOT / "assets" / "PwC.png"
    source_svg = ROOT / "assets" / "PwC_2025_Logo.svg"
    if not source_png.exists() or not source_svg.exists():
        missing = [str(path) for path in (source_png, source_svg) if not path.exists()]
        raise FileNotFoundError(f"missing brand assets: {missing}")

    target_dir = ROOT / "assets" / "derived"
    target_dir.mkdir(parents=True, exist_ok=True)

    cropped_png = target_dir / "PwC_logo_cropped.png"
    with Image.open(source_png) as image:
        rgba = image.convert("RGBA")
        alpha_bbox = rgba.getchannel("A").getbbox()
        if alpha_bbox is None:
            raise ValueError("PwC.png has no visible pixels")
        padding = 20
        left = max(0, alpha_bbox[0] - padding)
        top = max(0, alpha_bbox[1] - padding)
        right = min(rgba.width, alpha_bbox[2] + padding)
        bottom = min(rgba.height, alpha_bbox[3] + padding)
        rgba.crop((left, top, right, bottom)).save(cropped_png)

    vector_pdf = target_dir / "PwC_2025_Logo.pdf"
    svg_text = source_svg.read_text(encoding="utf-8").replace(
        '<svg width="100" height="100" viewBox="0 0 100 100"',
        '<svg width="70" height="34" viewBox="15 33 70 34"',
        1,
    )
    svg_document = fitz.open("svg", svg_text.encode("utf-8"))
    vector_pdf.write_bytes(svg_document.convert_to_pdf())
    return [cropped_png, vector_pdf]
