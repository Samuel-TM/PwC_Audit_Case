import os
import subprocess

from .config import BUILD_DIR, ROOT


def build_presentation() -> None:
    node = os.environ.get("CODEX_PRIMARY_RUNTIME_NODE", "node")
    subprocess.run([node, str(BUILD_DIR / "build_presentation.mjs")], cwd=ROOT, check=True)

