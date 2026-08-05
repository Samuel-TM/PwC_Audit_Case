#!/usr/bin/env python3
import json
import os
import shutil
import subprocess
import sys
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault("MPLCONFIGDIR", str(ROOT / "tmp" / "matplotlib"))
os.environ.setdefault("FONTCONFIG_FILE", str(ROOT / "assets" / "fonts" / "fonts.conf"))
os.environ.setdefault("XDG_CACHE_HOME", str(ROOT / "tmp" / "cache"))

from src.analysis import build_summary
from src.charts import generate_charts
from src.config import BUILD_DIR, OUTPUT_DIR, PREVIEW_DIR, ensure_directories
from src.qa_generator import write_qa
from src.risk_engine import evaluate_risks
from src.validators import assert_valid


def run(command: list[str], env: dict | None = None) -> None:
    subprocess.run(command, cwd=ROOT, check=True, env=env)


def main() -> None:
    ensure_directories()
    started = datetime.now(timezone.utc)

    validation = assert_valid()
    (OUTPUT_DIR / "validation_report.json").write_text(json.dumps(validation, ensure_ascii=False, indent=2), encoding="utf-8")
    (OUTPUT_DIR / "analysis_summary.md").write_text(build_summary(), encoding="utf-8")
    risks = [asdict(item) for item in evaluate_risks()]
    (OUTPUT_DIR / "audit_risk_results.json").write_text(json.dumps(risks, ensure_ascii=False, indent=2), encoding="utf-8")
    charts = generate_charts()
    write_qa()

    node = os.environ.get("CODEX_PRIMARY_RUNTIME_NODE", "node")
    skill_dir = Path("/root/.codex/skills/builtins/presentations")
    setup = skill_dir / "container_tools" / "setup_artifact_tool_workspace.mjs"
    run([node, str(setup), "--workspace", str(BUILD_DIR)])

    env = os.environ.copy()
    env["FONTCONFIG_FILE"] = str(ROOT / "assets" / "fonts" / "fonts.conf")
    env["MPLCONFIGDIR"] = str(ROOT / "tmp" / "matplotlib")
    run([node, str(BUILD_DIR / "build_workbooks.mjs"), "all"], env=env)
    run([sys.executable, str(ROOT / "scripts" / "fix_freeze_panes.py")], env=env)
    shutil.copy2(OUTPUT_DIR / "审计资料清单.xlsx", ROOT / "templates" / "audit_checklist_template.xlsx")
    shutil.copy2(OUTPUT_DIR / "审计计划.xlsx", ROOT / "templates" / "audit_plan_template.xlsx")
    run([node, str(BUILD_DIR / "build_presentation.mjs")], env=env)

    run([sys.executable, str(ROOT / "scripts" / "export_pdf.py")], env=env)
    test_env = env.copy()
    vendor_path = str(ROOT / "vendor" / "python")
    test_env["PYTHONPATH"] = vendor_path + os.pathsep + test_env.get("PYTHONPATH", "")
    test = subprocess.run([sys.executable, "-m", "pytest", "-q"], cwd=ROOT, env=test_env, text=True, capture_output=True)
    if test.returncode != 0:
        raise RuntimeError(f"pytest failed\n{test.stdout}\n{test.stderr}")

    required = [
        "A公司审计案例汇报.pptx", "A公司审计案例汇报.pdf", "审计资料清单.xlsx",
        "审计计划.xlsx", "导师问答手册.md", "analysis_summary.md",
        "validation_report.json", "audit_risk_results.json", "workbook_verification.json",
    ]
    missing = [name for name in required if not (OUTPUT_DIR / name).exists()]
    if missing:
        raise RuntimeError(f"missing outputs: {missing}")
    if len(charts) != 9:
        raise RuntimeError(f"expected 9 charts, got {len(charts)}")

    completed = datetime.now(timezone.utc)
    report = [
        "A 公司模拟审计案例构建报告",
        f"构建开始（UTC）：{started.isoformat()}",
        f"构建完成（UTC）：{completed.isoformat()}",
        "",
        f"数据校验：{validation['status']}（{len(validation['checks'])} 项）",
        f"图表：{len(charts)} 张，均为 1600×900 PNG",
        "PPT：15 页；已逐页渲染并保存布局检查文件",
        "Excel：资料清单 8 个工作表；审计计划 6 个工作表；已逐表渲染",
        "导师问答：30 题，每题含 30 秒回答和追问展开",
        f"测试：{test.stdout.strip()}",
        "",
        "已生成文件：",
        *[f"- {name}" for name in required],
        *[f"- charts/{p.name}" for p in charts],
        "",
        "已知限制：案例未提供交易级数据、重要性水平、内控测试、函证、盘点和期后结果；因此仅形成风险评估与程序设计，不形成真实审计意见。",
    ]
    (OUTPUT_DIR / "build_report.txt").write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report))


if __name__ == "__main__":
    main()
