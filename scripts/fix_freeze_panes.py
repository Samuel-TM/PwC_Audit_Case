#!/usr/bin/env python3
"""Surgical OOXML compatibility fix for freeze panes.

The workbook content and formatting are authored by artifact-tool. Its current
XLSX exporter does not persist freeze panes, so this script adds only the
standard worksheet `<pane>` element without changing cells, formulas, tables,
validation, conditional formatting, or styles.
"""
from __future__ import annotations

import os
import re
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _pane_xml(rows: int, columns: int) -> str:
    col_letter = chr(ord("A") + columns)
    top_left = f"{col_letter}{rows + 1}"
    attrs = []
    if columns:
        attrs.append(f'xSplit="{columns}"')
    if rows:
        attrs.append(f'ySplit="{rows}"')
    active = "bottomRight" if rows and columns else ("topRight" if columns else "bottomLeft")
    attrs.extend([f'topLeftCell="{top_left}"', f'activePane="{active}"', 'state="frozen"'])
    return f'<x:pane {" ".join(attrs)} /><x:selection pane="{active}" activeCell="{top_left}" sqref="{top_left}" />'


def apply_freeze(workbook: Path, settings: list[tuple[int, int]]) -> None:
    with zipfile.ZipFile(workbook, "r") as source:
        members = {info.filename: source.read(info.filename) for info in source.infolist()}
    for index, (rows, columns) in enumerate(settings, start=1):
        name = f"xl/worksheets/sheet{index}.xml"
        text = members[name].decode("utf-8")
        pattern = r"(<x:sheetView\b[^>]*)\s*/>"
        replacement = rf"\1>{_pane_xml(rows, columns)}</x:sheetView>"
        updated, count = re.subn(pattern, replacement, text, count=1)
        if count != 1:
            raise RuntimeError(f"could not locate sheetView in {name}")
        members[name] = updated.encode("utf-8")
    fd, temp_name = tempfile.mkstemp(prefix=workbook.stem + "-", suffix=".xlsx", dir=workbook.parent)
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target:
            for name, data in members.items():
                target.writestr(name, data)
        os.replace(temp_path, workbook)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def main() -> None:
    apply_freeze(ROOT / "output" / "审计资料清单.xlsx", [(4, 0)] + [(1, 0)] * 7)
    apply_freeze(ROOT / "output" / "审计计划.xlsx", [(4, 0), (1, 0), (1, 2), (1, 0), (1, 0), (1, 0)])


if __name__ == "__main__":
    main()

