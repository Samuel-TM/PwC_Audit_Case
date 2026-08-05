from pathlib import Path
from datetime import date

import matplotlib.dates as mdates
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
        "font.size": 25,
        "axes.titlesize": 34,
        "axes.titleweight": "bold",
        "axes.labelsize": 25,
        "xtick.labelsize": 23,
        "ytick.labelsize": 23,
        "legend.fontsize": 23,
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
    fig.tight_layout(pad=1.0)
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
    for idx, value in enumerate(revenue): ax.text(idx - 0.18, value + 500, f"{value:,.0f}", ha="center", fontsize=23)
    for idx, value in enumerate(profit): ax.text(idx + 0.18, value + 500, f"{value:,.0f}", ha="center", fontsize=23)
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
        for y, value in zip(years, series): ax.text(y, value + 280, f"{value:,.0f}", ha="center", fontsize=23)
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
    for y, value in zip(years, vals): ax.text(y, value + 0.25, f"{value:.2f}%", ha="center", fontsize=25, fontweight="bold")
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
        for y, value in zip(years, series): ax.text(y, value + 0.5, f"{value:.2f}%", ha="center", fontsize=23)
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
    for bar, value in zip(bars, export): ax.text(bar.get_x() + bar.get_width()/2, value + 0.7, f"{value:.2f}%", ha="center", fontsize=25, fontweight="bold")
    _clean(ax)
    _save(fig, "05_export_share.png")

    # 应收页面只保留净额及其占收入比例，避免重复展示存货趋势。
    fig, ax = plt.subplots()
    ar_ratio = np.array(ar) / np.array(revenue) * 100
    bars = ax.bar(years, ar, color=[PALE, LIGHT_BLUE, BLUE], width=0.58)
    ax.set_ylim(0, max(ar) * 1.22)
    ax.set_title("应收账款净额维持高位，2022 年占收入 34.35%")
    ax.set_ylabel("应收账款净额（万元）")
    ax.set_xticks(years)
    for bar, value, share in zip(bars, ar, ar_ratio):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            value + max(ar) * 0.035,
            f"{value:,.0f}\n{share:.2f}% / 收入",
            ha="center",
            fontsize=23,
            fontweight="bold",
        )
    _clean(ax)
    _save(fig, "06_receivables_net_ratio.png")

    # 存货页面改用 2022 年存货构成及跌价准备，不再使用成本结构饼图。
    inventory = load_csv("inventory.csv")
    inventory = inventory[(inventory.year == 2022) & (inventory.item != "合计")].copy()
    inventory["allowance"] = inventory["allowance"].fillna(0.0)
    major = inventory[inventory.item.isin(["原材料", "在产品", "库存商品"])].copy()
    other = inventory[~inventory.item.isin(["原材料", "在产品", "库存商品"])][["gross", "allowance", "net"]].sum()
    labels = major.item.tolist() + ["其他"]
    net_values = major.net.astype(float).tolist() + [float(other["net"])]
    allowance_values = major.allowance.astype(float).tolist() + [float(other["allowance"])]
    fig, ax = plt.subplots()
    positions = np.arange(len(labels))
    ax.barh(positions, net_values, color=BLUE, label="账面价值")
    ax.barh(positions, allowance_values, left=net_values, color=LIGHT_BLUE, label="跌价准备")
    ax.set_yticks(positions, labels)
    ax.set_xlabel("万元")
    ax.set_title("2022 年存货构成：库存商品与原材料占主体")
    ax.legend(frameon=False, ncols=2, loc="lower right")
    for position, value in zip(positions, net_values):
        ax.text(value + 35, position, f"{value:,.0f}", va="center", fontsize=23, fontweight="bold")
    ax.invert_yaxis()
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.grid(axis="x", color="#E8EAED", linewidth=1)
    ax.set_axisbelow(True)
    _save(fig, "07_inventory_composition_2022.png")

    # 投影版改为横向排序，避免六维热力图在整页缩放后文字过小。
    risks = sorted(evaluate_risks(), key=lambda item: item.average_score)
    level_colors = {"高": RED, "中高": "#EB8C00", "中": GRAY, "一般": "#D9DDE3"}
    fig, ax = plt.subplots()
    positions = np.arange(len(risks))
    bars = ax.barh(
        positions,
        [risk.average_score for risk in risks],
        color=[level_colors[risk.level] for risk in risks],
        height=0.62,
    )
    ax.set_yticks(positions, [risk.area for risk in risks])
    ax.set_xlim(0, 5.25)
    ax.set_xlabel("六维平均分（1-5）")
    ax.set_title("计划阶段风险排序：收入、应收与存货投入最多资源")
    for bar, risk in zip(bars, risks):
        ax.text(
            risk.average_score + 0.08,
            bar.get_y() + bar.get_height() / 2,
            f"{risk.average_score:.2f}｜{risk.level}",
            va="center",
            fontsize=23,
            fontweight="bold",
        )
    ax.spines[["top", "right", "left"]].set_visible(False)
    ax.grid(axis="x", color="#E8EAED", linewidth=1)
    ax.set_axisbelow(True)
    _save(fig, "08_audit_risk_ranking.png")

    fig, ax = plt.subplots()
    tasks = [
        ("盘点安排确认", date(2022, 12, 27), date(2022, 12, 30), BLUE),
        ("函证核验与首次发出", date(2023, 1, 3), date(2023, 1, 6), LIGHT_BLUE),
        ("流程穿行与分析", date(2023, 1, 9), date(2023, 1, 13), "#89CFF0"),
        ("高风险实质性程序", date(2023, 1, 16), date(2023, 1, 24), BLUE),
        ("未决事项清理", date(2023, 1, 23), date(2023, 1, 26), "#EB8C00"),
        ("项目经理及合伙人复核", date(2023, 1, 25), date(2023, 1, 27), RED),
    ]
    for idx, (name, start, finish, color) in enumerate(tasks):
        start_num = mdates.date2num(start)
        finish_num = mdates.date2num(finish)
        ax.barh(idx, finish_num - start_num + 1, left=start_num, color=color, height=0.56)
        ax.text(
            start_num + (finish_num - start_num + 1) / 2,
            idx,
            name,
            ha="center",
            va="center",
            color="white" if color in {BLUE, RED} else INK,
            fontsize=21,
            fontweight="bold",
        )
        ax.plot(finish_num, idx, marker="D", markersize=7, color=INK)
    ax.set_yticks(range(len(tasks)), ["前置", "第1周", "第2周", "第3周", "第4周", "第4周"])
    tick_dates = [date(2022, 12, 27), date(2023, 1, 3), date(2023, 1, 9), date(2023, 1, 16), date(2023, 1, 23), date(2023, 1, 27)]
    ax.set_xticks([mdates.date2num(item) for item in tick_dates])
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))
    ax.set_xlim(mdates.date2num(date(2022, 12, 26)), mdates.date2num(date(2023, 1, 29)))
    ax.set_title("四周关键路径：外部程序前置，分析与实质性程序并行")
    ax.invert_yaxis()
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)
    ax.grid(axis="x", color="#E8EAED")
    _save(fig, "09_audit_plan_gantt.png")

    return sorted(CHART_DIR.glob("*.png"))
