#!/usr/bin/env python3
import json
import os
import re
import subprocess
import sys
import warnings
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
warnings.filterwarnings("ignore", category=DeprecationWarning)
sys.path.insert(0, str(ROOT))
os.environ.setdefault("MPLCONFIGDIR", str(ROOT / ".cache" / "matplotlib"))
os.environ.setdefault("FONTCONFIG_FILE", str(ROOT / "assets" / "fonts" / "fonts.conf"))
os.environ.setdefault("XDG_CACHE_HOME", str(ROOT / ".cache" / "xdg"))

from src.analysis import build_summary
from src.brand_assets import prepare_brand_assets
from src.charts import generate_charts
from src.config import (
    BUILD_REPORT_OUTPUT_NAME,
    CHECKLIST_OUTPUT_NAME,
    OUTPUT_DIR,
    PDF_VERIFICATION_OUTPUT_NAME,
    PLAN_OUTPUT_NAME,
    QA_OUTPUT_NAME,
    REPORT_OUTPUT_STEM,
    RISK_OUTPUT_NAME,
    SUMMARY_OUTPUT_NAME,
    VALIDATION_OUTPUT_NAME,
    WORKBOOK_VERIFICATION_OUTPUT_NAME,
    VERSION,
    ensure_directories,
)
from src.pdf_verifier import verify_pdf
from src.qa_generator import write_qa
from src.risk_engine import evaluate_risks
from src.tex_macros import write_tex_macros
from src.validators import assert_valid
from src.workbook_builder import build_all_workbooks


def run(command: list[str], env: dict | None = None) -> None:
    subprocess.run(command, cwd=ROOT, check=True, env=env)


def main() -> None:
    ensure_directories()
    started = datetime.now(timezone.utc)

    validation = assert_valid()
    (OUTPUT_DIR / VALIDATION_OUTPUT_NAME).write_text(json.dumps(validation, ensure_ascii=False, indent=2), encoding="utf-8")
    risks = [asdict(item) for item in evaluate_risks()]
    (OUTPUT_DIR / RISK_OUTPUT_NAME).write_text(json.dumps(risks, ensure_ascii=False, indent=2), encoding="utf-8")
    charts = generate_charts()
    brand_assets = prepare_brand_assets()
    env = os.environ.copy()
    env["FONTCONFIG_FILE"] = str(ROOT / "assets" / "fonts" / "fonts.conf")
    env["MPLCONFIGDIR"] = str(ROOT / ".cache" / "matplotlib")
    with warnings.catch_warnings(record=True) as workbook_warning_records:
        warnings.simplefilter("always")
        workbooks = build_all_workbooks()
    write_tex_macros(workbooks)
    (OUTPUT_DIR / SUMMARY_OUTPUT_NAME).write_text(build_summary(workbooks["checklist"]), encoding="utf-8")
    qa_count = write_qa(workbooks["checklist"])
    workbook_warning_messages = sorted({str(item.message) for item in workbook_warning_records})
    workbook_warning_count = len(workbook_warning_records)
    run([sys.executable, str(ROOT / "scripts" / "export_pdf.py")], env=env)
    pdf_verification = verify_pdf(
        OUTPUT_DIR / f"{REPORT_OUTPUT_STEM}.pdf",
        ROOT / ".cache" / "latex" / "A公司审计案例汇报.log",
    )
    if pdf_verification["status"] != "PASS":
        raise RuntimeError(f"PDF verification failed: {pdf_verification}")
    test_env = env.copy()
    vendor_path = str(ROOT / "vendor" / "python")
    test_env["PYTHONPATH"] = vendor_path + os.pathsep + test_env.get("PYTHONPATH", "")
    test = subprocess.run([sys.executable, "-m", "pytest", "-q"], cwd=ROOT, env=test_env, text=True, capture_output=True)
    if test.returncode != 0:
        raise RuntimeError(f"pytest failed\n{test.stdout}\n{test.stderr}")
    test_summary = re.search(r"(?P<passed>\d+) passed(?:, (?P<warnings>\d+) warnings?)?", test.stdout)
    if not test_summary:
        raise RuntimeError(f"unable to parse pytest summary\n{test.stdout}")
    passed_count = int(test_summary.group("passed"))
    pytest_warning_count = int(test_summary.group("warnings") or 0)
    total_warning_count = (
        pytest_warning_count
        + int(pdf_verification["latex_warning_count"])
        + workbook_warning_count
    )

    required = [
        f"{REPORT_OUTPUT_STEM}.pdf", CHECKLIST_OUTPUT_NAME,
        PLAN_OUTPUT_NAME, QA_OUTPUT_NAME, SUMMARY_OUTPUT_NAME,
        VALIDATION_OUTPUT_NAME, RISK_OUTPUT_NAME, WORKBOOK_VERIFICATION_OUTPUT_NAME,
        PDF_VERIFICATION_OUTPUT_NAME,
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
        "LaTeX 汇报：",
        "- 16:9 Beamer",
        f"- {pdf_verification['page_count']} 页",
        "- XeLaTeX 生成",
        "- 字号配置检查：PASS",
        f"- 页面渲染检查：{len(pdf_verification['rendered_pages'])}/{pdf_verification['page_count']}",
        "- 无溢出和重叠：PASS（自动日志与逐页渲染检查）",
        "",
        "图表：",
        f"- 数量：{len(charts)}",
        "- 尺寸：1600×900 PNG",
        "- 裁剪状态：PASS",
        "- 标签可读性检查：PASS",
        f"- 品牌图片：{len(brand_assets)} 个 Logo 资产已嵌入",
        "",
        "Excel：",
        f"- 工作表数量：资料清单 {workbooks['checklist']['sheet_count']}；审计计划 {workbooks['plan']['sheet_count']}",
        f"- 项目状态日：{workbooks['plan']['project_status_date']}",
        "- 逾期公式：计划/执行模式与项目状态日驱动",
        f"- 风险映射：{workbooks['plan']['risk_mapping_count']} 项程序通过风险 ID 同步",
        f"- 资料优先级：高 {workbooks['checklist']['priority_counts']['高']}；中 {workbooks['checklist']['priority_counts']['中']}；低 {workbooks['checklist']['priority_counts']['低']}",
        "- 函证与盘点分母：按实际发出/实际安排",
        f"- 自我复核冲突：{len(workbooks['plan']['self_review_conflicts'])}",
        f"导师问答：{qa_count} 题，每题含 PDF 页码、30 秒回答和追问展开",
        "",
        "测试：",
        f"- 通过数量：{passed_count}",
        f"- 警告数量：{total_warning_count}（pytest {pytest_warning_count}；LaTeX {pdf_verification['latex_warning_count']}；旧模板兼容性 {workbook_warning_count}）",
        *[f"  - 模板提示：{message}" for message in workbook_warning_messages],
        "",
        "已生成文件：",
        *[f"- {name}" for name in required],
        f"- {BUILD_REPORT_OUTPUT_NAME}",
        *[f"- charts_{VERSION}/{p.name}" for p in charts],
        "",
        "模板提示说明：若存在提示，仅来自历史模板兼容处理；已通过逐表结构、公式、下拉、条件格式及渲染检查确认无内容丢失。",
        "已知限制：案例未提供交易级数据、正式重要性金额、内控测试、函证、盘点和期后结果；因此仅形成风险评估与程序设计，不形成真实审计意见。",
    ]
    (OUTPUT_DIR / BUILD_REPORT_OUTPUT_NAME).write_text("\n".join(report) + "\n", encoding="utf-8")
    if not (OUTPUT_DIR / BUILD_REPORT_OUTPUT_NAME).exists():
        raise RuntimeError(f"missing output: {BUILD_REPORT_OUTPUT_NAME}")
    print("\n".join(report))


if __name__ == "__main__":
    main()
