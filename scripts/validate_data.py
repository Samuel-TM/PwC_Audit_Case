#!/usr/bin/env python3
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.validators import assert_valid


if __name__ == "__main__":
    print(json.dumps(assert_valid(), ensure_ascii=False, indent=2))
