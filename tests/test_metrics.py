import pytest

from src.analysis import calculate_metrics
from src.metrics import current_ratio, debt_ratio, gross_margin, growth_rate, ratio


def test_growth_rate():
    assert growth_rate(26373.49, 26054.89) == pytest.approx(0.012228, abs=1e-6)


def test_gross_margin():
    assert gross_margin(26020.41, 20456.71) == pytest.approx(0.213821, abs=1e-6)


def test_ratios():
    assert current_ratio(19477.47, 14210.27) == pytest.approx(1.37066, abs=1e-5)
    assert debt_ratio(14855.04, 25437.63) == pytest.approx(0.583979, abs=1e-6)
    assert ratio(9059.00, 26373.49) == pytest.approx(0.343489, abs=1e-6)


def test_analysis_metrics():
    metrics = calculate_metrics()
    assert metrics["inventory_turnover_2022"] == pytest.approx(4.3059, abs=1e-4)
    assert metrics["dso_2022"] == pytest.approx(127.30, abs=0.02)
