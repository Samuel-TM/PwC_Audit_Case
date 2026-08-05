from dataclasses import dataclass
from typing import Any

from .loaders import load_yaml


@dataclass(frozen=True)
class RiskResult:
    area: str
    average_score: float
    level: str
    assertions: list[str]
    reasons: list[str]
    scores: dict[str, int]


def risk_level(score: float) -> str:
    if score >= 4.0:
        return "高"
    if score >= 3.0:
        return "中高"
    if score >= 2.0:
        return "中"
    return "一般"


def evaluate_risks() -> list[RiskResult]:
    items: list[dict[str, Any]] = load_yaml("audit_risks.yaml")["risks"]
    results: list[RiskResult] = []
    for item in items:
        average = sum(item["scores"].values()) / len(item["scores"])
        results.append(
            RiskResult(
                area=item["area"],
                average_score=round(average, 2),
                level=risk_level(average),
                assertions=item["assertions"],
                reasons=item["reasons"],
                scores=item["scores"],
            )
        )
    return results

