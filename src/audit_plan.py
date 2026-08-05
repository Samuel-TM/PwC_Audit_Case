import os
import subprocess

from .config import BUILD_DIR, ROOT


def build_audit_plan() -> None:
    node = os.environ.get("CODEX_PRIMARY_RUNTIME_NODE", "node")
    subprocess.run([node, str(BUILD_DIR / "build_workbooks.mjs"), "plan"], cwd=ROOT, check=True)

