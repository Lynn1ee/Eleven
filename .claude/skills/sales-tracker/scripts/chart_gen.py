"""图表生成：折线图、饼图、柱状图、数据表格，输出为PNG图片用于飞书卡片内嵌"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.ticker as mticker
import numpy as np
import os
from datetime import datetime

# ── 全局样式 ──────────────────────────────────

# 中文字体
_font_candidates = [
    "SimHei", "Microsoft YaHei", "PingFang SC",
    "Noto Sans CJK SC", "WenQuanYi Micro Hei", "sans-serif",
]
for _font in _font_candidates:
    try:
        fm.findfont(_font, fallback_to_default=False)
        plt.rcParams["font.sans-serif"] = [_font, "DejaVu Sans"]
        break
    except Exception:
        continue

plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["figure.dpi"] = 200
plt.rcParams["savefig.dpi"] = 200
plt.rcParams["axes.grid"] = True
plt.rcParams["grid.alpha"] = 0.15
plt.rcParams["grid.linestyle"] = "-"
plt.rcParams["grid.color"] = "#E8E8E8"
plt.rcParams["axes.spines.top"] = False
plt.rcParams["axes.spines.right"] = False
plt.rcParams["axes.spines.left"] = False
plt.rcParams["axes.spines.bottom"] = False
plt.rcParams["font.size"] = 11
plt.rcParams["axes.titlesize"] = 14
plt.rcParams["axes.titleweight"] = "600"
plt.rcParams["axes.labelcolor"] = "#666666"
plt.rcParams["xtick.color"] = "#888888"
plt.rcParams["ytick.color"] = "#888888"

# 现代配色
C_PALETTE = ["#5470C6", "#91CC75", "#FAC858", "#EE6666", "#73C0DE",
             "#3BA272", "#FC8452", "#9A60B4", "#EA7CCC", "#B6A2DE"]
C_BLUE = "#5470C6"
C_GREEN = "#91CC75"
C_ORANGE = "#FAC858"
C_RED = "#EE6666"
C_BG = "#F5F6FA"

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def _save_chart(fig, filename: str) -> str:
    path = os.path.join(OUTPUT_DIR, filename)
    fig.savefig(path, dpi=200, bbox_inches="tight", facecolor="white", edgecolor="none")
    plt.close(fig)
    return path


def _auto_filename(prefix: str) -> str:
    return f"{prefix}_{datetime.now().strftime('%Y%m%d%H%M%S')}.png"


# ── 折线图 ────────────────────────────────────

def trend_line(dates: list, values: list, title: str = "趋势图",
               ylabel: str = "", filename: str = None) -> str:
    """折线图：日期 vs 数值，展示时间序列趋势。"""
    if filename is None:
        filename = _auto_filename("trend")

    fig, ax = plt.subplots(figsize=(10, 4))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    ax.plot(dates, values, marker="o", linewidth=2.5, markersize=7,
            color=C_BLUE, markeredgecolor="white", markeredgewidth=1.5,
            zorder=3)
    ax.fill_between(range(len(dates)), values, alpha=0.1, color=C_BLUE)
    ax.set_title(title, pad=14, color="#333333")
    ax.set_ylabel(ylabel, fontsize=10)
    ax.set_xlabel("日期", fontsize=10)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.tick_params(axis="both", labelsize=9)
    fig.autofmt_xdate()

    if values:
        max_idx = values.index(max(values))
        min_idx = values.index(min(values))
        ax.annotate(f"最高 {values[max_idx]}", xy=(max_idx, values[max_idx]),
                    xytext=(0, 14), textcoords="offset points", fontsize=9,
                    color=C_RED, ha="center", fontweight="bold")
        ax.annotate(f"最低 {values[min_idx]}", xy=(min_idx, values[min_idx]),
                    xytext=(0, -18), textcoords="offset points", fontsize=9,
                    color="#888888", ha="center", fontweight="bold")

    return _save_chart(fig, filename)


# ── 饼图 ──────────────────────────────────────

def pie_chart(labels: list, values: list, title: str = "占比分析",
              filename: str = None) -> str:
    """饼图：各平台占比，带百分比标签。"""
    if filename is None:
        filename = _auto_filename("pie")

    colors = [C_BLUE, C_GREEN, C_ORANGE, C_RED, "#9A60B4", "#73C0DE"]
    explode = [0.02] * len(labels)

    fig, ax = plt.subplots(figsize=(6.5, 6.5))
    fig.patch.set_facecolor("white")
    wedges, texts, autotexts = ax.pie(
        values, explode=explode, labels=labels,
        colors=colors[:len(labels)],
        autopct="%1.1f%%", startangle=140, pctdistance=0.58,
        wedgeprops={"linewidth": 2.5, "edgecolor": "white"},
    )
    for t in autotexts:
        t.set_fontsize(12)
        t.set_fontweight("bold")
        t.set_color("white")
    for t in texts:
        t.set_fontsize(11)
        t.set_color("#444444")
    ax.set_title(title, pad=16, color="#333333")

    return _save_chart(fig, filename)


# ── 柱状图 ────────────────────────────────────

def bar_chart(labels: list, values: list, title: str = "排名对比",
              ylabel: str = "", filename: str = None) -> str:
    """柱状图：人员排名对比，前三名金银铜高亮。"""
    if filename is None:
        filename = _auto_filename("bar")

    n = len(labels)
    colors = [C_BLUE] * n
    if n >= 1:
        colors[0] = "#F3C846"  # 金
    if n >= 2:
        colors[1] = "#B0BCC9"  # 银
    if n >= 3:
        colors[2] = "#CD9A6C"  # 铜

    fig, ax = plt.subplots(figsize=(max(7, n * 0.7), 5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    bars = ax.bar(labels, values, color=colors, edgecolor="white",
                  linewidth=1.2, width=0.65, zorder=3)

    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.015,
                str(val), ha="center", fontsize=10, fontweight="bold", color="#444444")

    ax.set_title(title, pad=14, color="#333333")
    ax.set_ylabel(ylabel, fontsize=10)
    ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax.tick_params(axis="both", labelsize=9)
    ax.set_ylim(0, max(values) * 1.18)
    ax.tick_params(axis="y", length=0)
    ax.tick_params(axis="x", length=0)

    return _save_chart(fig, filename)


# ── 数据表格图 ────────────────────────────────

def table_image(headers: list, rows: list, title: str = "",
                col_widths: list = None, filename: str = None) -> str:
    """生成格式化的数据表格 PNG，适合嵌入飞书卡片。

    headers: 列标题列表
    rows: 二维列表，每行是数据值列表
    title: 表格标题（可选）
    col_widths: 列宽比例列表（可选，默认均分）
    """
    if filename is None:
        filename = _auto_filename("table")

    n_rows = len(rows)
    n_cols = len(headers)
    fig_height = max(2.2, n_rows * 0.42 + 1.6)
    fig_width = max(6, sum(col_widths) * 1.1 if col_widths else n_cols * 1.3)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.axis("off")

    if title:
        ax.set_title(title, fontsize=13, fontweight="bold", pad=10, loc="left", color="#222")

    # 构建表格数据
    cell_text = [[str(c) for c in row] for row in rows]
    cell_colors = []
    for i in range(n_rows):
        if i % 2 == 0:
            cell_colors.append(["#F7F8FA"] * n_cols)
        else:
            cell_colors.append(["white"] * n_cols)

    table = ax.table(
        cellText=cell_text,
        colLabels=headers,
        cellColours=cell_colors,
        colColours=["#4C78A8"] * n_cols,
        cellLoc="center",
        loc="upper center",
    )

    # 样式
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.35)

    for key, cell in table.get_celld().items():
        cell.set_edgecolor("#E0E0E0")
        cell.set_linewidth(0.5)
        if key[0] == 0:  # header row
            cell.set_text_props(color="white", fontweight="bold", fontsize=9.5)
            cell.set_edgecolor("#3A6A96")
            cell.set_linewidth(1)
            cell.get_text().set_color("white")

    # 设置列宽
    if col_widths:
        for i, w in enumerate(col_widths):
            for key, cell in table.get_celld().items():
                if key[1] == i:
                    cell.set_width(w)

    table.set_fontsize(9)

    return _save_chart(fig, filename)


# ── 双折线图（预留）────────────────────────────

def dual_line(dates: list, values1: list, values2: list,
              label1: str = "会话量", label2: str = "订单数",
              title: str = "双指标趋势", filename: str = None) -> str:
    """双折线图：两组数据在同一图上展示。"""
    if filename is None:
        filename = _auto_filename("dual")

    fig, ax1 = plt.subplots(figsize=(10, 4))
    ax1.plot(dates, values1, marker="o", linewidth=2.5, markersize=6,
             color=C_BLUE, label=label1, markeredgecolor="white", markeredgewidth=1.2)
    ax1.set_ylabel(label1, color=C_BLUE, fontsize=10)
    ax1.fill_between(range(len(dates)), values1, alpha=0.06, color=C_BLUE)

    ax2 = ax1.twinx()
    ax2.plot(dates, values2, marker="s", linewidth=2.5, markersize=6,
             color=C_ORANGE, label=label2, linestyle="--")
    ax2.set_ylabel(label2, color=C_ORANGE, fontsize=10)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left", framealpha=0.9)

    ax1.set_title(title, pad=12)
    ax1.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    ax2.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
    fig.autofmt_xdate()

    return _save_chart(fig, filename)
