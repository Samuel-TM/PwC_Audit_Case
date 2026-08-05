import os
import subprocess

from .config import BUILD_DIR, ROOT


def build_checklist() -> None:
    node = os.environ.get("CODEX_PRIMARY_RUNTIME_NODE", "node")
    subprocess.run([node, str(BUILD_DIR / "build_workbooks.mjs"), "checklist"], cwd=ROOT, check=True)

