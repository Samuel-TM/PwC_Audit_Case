def growth_rate(current: float, previous: float) -> float:
    if previous == 0:
        raise ZeroDivisionError("previous cannot be zero")
    return current / previous - 1


def ratio(numerator: float, denominator: float) -> float:
    if denominator == 0:
        raise ZeroDivisionError("denominator cannot be zero")
    return numerator / denominator


def gross_margin(revenue: float, cost: float) -> float:
    return ratio(revenue - cost, revenue)


def current_ratio(current_assets: float, current_liabilities: float) -> float:
    return ratio(current_assets, current_liabilities)


def debt_ratio(total_liabilities: float, total_assets: float) -> float:
    return ratio(total_liabilities, total_assets)


def receivables_to_revenue(receivables: float, revenue: float) -> float:
    return ratio(receivables, revenue)


def inventory_turnover(cost: float, avg_inventory: float) -> float:
    return ratio(cost, avg_inventory)


def days_sales_outstanding(avg_receivables: float, revenue: float) -> float:
    return ratio(avg_receivables, revenue) * 365

