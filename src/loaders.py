from pathlib import Path
from typing import Any

import pandas as pd
import yaml

from .config import DATA_DIR


def load_csv(name: str) -> pd.DataFrame:
    return pd.read_csv(DATA_DIR / name)


def load_yaml(name: str) -> dict[str, Any]:
    with (DATA_DIR / name).open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def numeric_value(df: pd.DataFrame, label_col: str, label: str, year: int) -> float:
    value = df.loc[df[label_col] == label, str(year)].iloc[0]
    return float(value)


def data_files() -> list[Path]:
    return sorted(DATA_DIR.glob("*.csv")) + sorted(DATA_DIR.glob("*.yaml"))

