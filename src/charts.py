from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager

from .config import CHART_DIR, ROOT
from .loaders import load_csv
from .risk_engine import evaluate_risks

BLUE = "#3D8DFF"
LIGHT_BLUE = "#6DCBF4"
PALE = "#D0EDFA"
INK = "#171717"
GRAY = "#A8AFB8"
RED = "#E45D5D"


def _setup() -> None:
    font_path = ROOT / "assets" / "fonts" / "NotoSansSC-Regular.ttf"
    if font_path.exists():
        font_manager.fontManager.addfont(font_path)
        plt.rcParams["font.family"] = ["Noto Sans SC Thin", "DejaVu Sans"]
    plt.rcParams.update({
        "axes.unicode_minus": False,
        "font.size": 17,
        "axes.titlesize": 26,
        "axes.titleweight": "bold",
        "axes.labelcolor": INK,
        "text.color": INK,
        "xtick.color": "#444444",
        "ytick.color": "#444444",
        "axes.edgecolor": "#D9DDE3",
        "figure.facecolor": "white",
        "axes.facecolor": "white",
    })


def _save(fig: plt.Figure, name: str) -> None:
    fig.set_size_inches(16, 9)
    fig.tight_layout(pad=2.2)
    fig.savefig(CHART_DIR / name, dpi=100, facecolor="white")
    plt.close(fig)


def _clean(ax: plt.Axes) -> None:
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.grid(axis="y", color="#E8EAED", linewidth=1)
    ax.set_axisbelow(True)


def generate_charts() -> list[Path]:
    _setup()
    CHART_DIR.mkdir(parents=True, exist_ok=True)
    years = np.array([2020, 2021, 2022])
    inc = load_csv("income_statement.csv").set_index("item")
    bs = load_csv("balance_sheet.csv").set_index("item")
    margin = load_csv("gross_margin.csv").set_index("item")

    fig, ax = plt.subplots()
    x = np.arange(3)
    revenue = [float(inc.loc["营业收入", str(y)]) for y in years]
    profit = [float(inc.loc["净利润", str(y)]) for y in years]
    ax.bar(x - 0.18, revenue, 0.36, label="营业收入", color=BLUE)
    ax.bar(x + 0.18, profit, 0.36, label="净利润", color=LIGHT_BLUE)
    ax.set_title("营业收入增长在 2022 年明显放缓")
    ax.set_xticks(x, years)
    ax.set_ylabel("万元")
    ax.legend(frameon=False, ncols=2, loc="upper left")
    for idx, value in enumerate(revenue): ax.text(idx - 0.18, value + 500, f"{value:,.0f}", ha="center", fontsize=14)
    for idx, value in enumerate(profit): ax.text(idx + 0.18, value + 500, f"{value:,.0f}", ha="center", fontsize=14)
    _clean(ax)
    _save(fig, "01_revenue_net_profit.png")

    fig, ax = plt.subplots()
    ar = [float(bs.loc["应收账款", str(y)]) for y in years]
    inv = [float(bs.loc["存货", str(y)]) for y in years]
    ax.plot(years, ar, marker="o", markersize=10, linewidth=4, color=BLUE, label="应收账款净额")
    ax.plot(years, inv, marker="o", markersize=10, linewidth=4, color=LIGHT_BLUE, label="存货净额")
    ax.set_title("应收账款持续处于高位，存货仍高于 2020 年")
    ax.set_ylabel("万元")
    ax.set_xticks(years)
    ax.legend(frameon=False, ncols=2, loc="upper left")
    for series in [ar, inv]:
        for y, value in zip(years, series): ax.text(y, value + 280, f"{value:,.0f}", ha="center", fontsize=14)
    _clean(ax)
    _save(fig, "02_receivables_inventory.png")

    fig, ax = plt.subplots()
    vals = [float(margin.loc["主营业务合计", f"{y}_margin"]) * 100 for y in years]
    ax.plot(years, vals, marker="o", linewidth=5, markersize=11, color=BLUE)
    ax.fill_between(years, vals, min(vals) - 1.5, color=PALE, alpha=0.75)
    ax.set_ylim(19, 25)
    ax.set_title("主营业务毛利率连续下降至 21.38%")
    ax.set_ylabel("毛利率（%）")
    ax.set_xticks(years)
    for y, value in zip(years, vals): ax.text(y, value + 0.25, f"{value:.2f}%", ha="center", fontsize=16, fontweight="bold")
    _clean(ax)
    _save(fig, "03_main_gross_margin.png")

    fig, ax = plt.subplots()
    appliance = [float(margin.loc["家电滑轨", f"{y}_margin"]) * 100 for y in years]
    server = [float(margin.loc["服务器滑轨", f"{y}_margin"]) * 100 for y in years]
    ax.plot(years, appliance, marker="o", linewidth=4, color=LIGHT_BLUE, label="家电滑轨")
    ax.plot(years, server, marker="o", linewidth=4, color=BLUE, label="服务器滑轨")
    ax.set_ylim(18, 35)
    ax.set_title("服务器滑轨毛利率三年下降 9.33 个百分点")
    ax.set_ylabel("毛利率（%）")
    ax.set_xticks(years)
    ax.legend(frameon=False, ncols=2, loc="upper right")
    for series in [appliance, server]:
        for y, value in zip(years, series): ax.text(y, value + 0.5, f"{value:.2f}%", ha="center", fontsize=14)
    _clean(ax)
    _save(fig, "04_product_gross_margin.png")

    region = load_csv("revenue_by_region.csv").set_index("region")
    fig, ax = plt.subplots()
    export = [float(region.loc["外销", f"{y}_pct"]) * 100 for y in years]
    bars = ax.bar(years, export, color=[PALE, LIGHT_BLUE, BLUE], width=0.58)
    ax.set_ylim(0, 28)
    ax.set_title("外销占比三年提升至 23.61%")
    ax.set_ylabel("外销占比（%）")
    ax.set_xticks(years)
    for bar, value in zip(bars, export): ax.text(bar.get_x() + bar.get_width()/2, value + 0.7, f"{value:.2f}%", ha="center", fontsize=16, fontweight="bold")
    _clean(ax)
    _save(fig, "05_export_share.png")

    quarter = load_csv("revenue_by_quarter.csv")
    quarter = quarter[quarter.quarter != "合计"].set_index("quarter")
    fig, ax = plt.subplots()
    bottom = np.zeros(3)
    colors = ["#D0EDFA", "#9DDAF5", "#6DCBF4", "#3D8DFF"]
    for q, color in zip(quarter.index, colors):
        vals = np.array([float(quarter.loc[q, f"{y}_pct"]) * 100 for y in years])
        ax.bar(years, vals, bottom=bottom, label=q, color=color, width=0.6)
        for xval, btm, val in zip(years, bottom, vals):
            if val >= 8: ax.text(xval, btm + val/2, f"{val:.1f}%", ha="center", va="center", fontsize=13)
        bottom += vals
    ax.set_title("季度收入结构变化，2022 年第四季度占比回落")
    ax.set_ylabel("主营业务收入占比（%）")
    ax.set_xticks(years)
    ax.legend(frameon=False, ncols=4, loc="upper center", bbox_to_anchor=(0.5, 1.01))
    ax.spines[["top", "right", "left"]].set_visible(False)
    _save(fig, "06_quarterly_revenue_structure.png")

    cost = load_csv("cost_structure.csv")
    cost = cost[cost.item != "合计"]
    fig, ax = plt.subplots()
    values = cost["2022_pct"].astype(float).values * 100
    labels = cost.item.tolist()
    ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=90, colors=[BLUE, LIGHT_BLUE, PALE], textprops={"fontsize": 17}, wedgeprops={"linewidth": 2, "edgecolor": "white"})
    ax.set_title("2022 年主营业务成本中直接材料占 58.50%")
    _save(fig, "07_cost_structure_2022.png")

    risks = evaluate_risks()
    dims = ["金额重大性", "同比波动", "管理层判断程度", "舞弊可能性", "交易复杂度", "证据获取难度"]
    matrix = np.array([[risk.scores[d] for d in dims] for risk in risks])
    fig, ax = plt.subplots()
    image = ax.imshow(matrix, cmap="Blues", vmin=1, vmax=5, aspect="auto")
    ax.set_title("收入、应收和存货的综合风险评分最高")
    ax.set_xticks(range(len(dims)), dims, rotation=15, ha="right")
    ax.set_yticks(range(len(risks)), [r.area for r in risks])
    for row in range(matrix.shape[0]):
        for col in range(matrix.shape[1]): ax.text(col, row, str(matrix[row, col]), ha="center", va="center", color="white" if matrix[row, col] >= 4 else INK, fontweight="bold")
    fig.colorbar(image, ax=ax, shrink=0.75, label="风险评分（1-5）")
    _save(fig, "08_audit_risk_heatmap.png")

    fig, ax = plt.subplots()
    phases = [("计划与风险评估", 0, 5, BLUE), ("控制了解与初测", 5, 5, LIGHT_BLUE), ("实质性审计程序", 10, 5, BLUE), ("收尾与汇报", 15, 5, LIGHT_BLUE)]
    for idx, (name, start, duration, color) in enumerate(phases):
        ax.barh(idx, duration, left=start, color=color, height=0.55)
        ax.text(start + duration/2, idx, name, ha="center", va="center", color="white" if color == BLUE else INK, fontweight="bold")
    ax.set_yticks(range(4), ["第1周", "第2周", "第3周", "第4周"])
    ax.set_xticks([0, 5, 10, 15, 20], ["启动", "W1结束", "W2结束", "W3结束", "项目结束"])
    ax.set_title("四周审计计划：先锁定证据，再集中实施实质性程序")
    ax.invert_yaxis()
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)
    ax.grid(axis="x", color="#E8EAED")
    _save(fig, "09_audit_plan_gantt.png")

    return sorted(CHART_DIR.glob("*.png"))
