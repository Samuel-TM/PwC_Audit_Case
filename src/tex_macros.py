from __future__ import annotations

import os
from pathlib import Path

from .config import ROOT, VERSION


def _latex_escape(value: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
    }
    return "".join(replacements.get(char, char) for char in value)


def write_tex_macros(workbook_result: dict[str, object]) -> Path:
    checklist = workbook_result["checklist"]
    priorities = checklist["priority_counts"]
    total = int(checklist["total_requests"])
    business_modules = int(checklist["sheet_count"]) - 1
    team_members = _latex_escape(os.environ.get("PWC_TEAM_MEMBERS", "审计案例项目组"))
    presenter = _latex_escape(os.environ.get("PWC_PRESENTER", "项目组代表"))
    lines = [
        "% Auto-generated build parameters. Do not edit by hand.",
        rf"\renewcommand{{\TeamMembers}}{{{team_members}}}",
        rf"\renewcommand{{\PresenterName}}{{{presenter}}}",
        rf"\renewcommand{{\ChecklistTotal}}{{{total}}}",
        rf"\renewcommand{{\ChecklistHigh}}{{{int(priorities['高'])}}}",
        rf"\renewcommand{{\ChecklistMedium}}{{{int(priorities['中'])}}}",
        rf"\renewcommand{{\ChecklistLow}}{{{int(priorities['低'])}}}",
        rf"\renewcommand{{\BusinessModules}}{{{business_modules}}}",
    ]
    target = ROOT / ".cache" / f"generated_metrics_{VERSION}.tex"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return target

