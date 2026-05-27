"""图表生成：折线图、饼图、柱状图，输出为PNG图片用于飞书消息附件"""

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import os
from datetime import datetime

# 尝试使用中文字体
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

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)


def _save_chart(fig, filename: str) -> str:
    """保存图表为PNG，返回文件路径。"""
    path = os.path.join(OUTPUT_DIR, filename)
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


def trend_line(dates: list, values: list, title: str = "趋势图",
               ylabel: str = "", filename: str = None) -> str:
    """折线图：日期 vs 数值，展示时间序列趋势。"""
    if filename is None:
        filename = f"trend_{datetime.now().strftime('%Y%m%d%H%M%S')}.png"

    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(dates, values, marker="o", linewidth=2, markersize=6, color="#3370FF")
    ax.fill_between(range(len(dates)), values, alpha=0.1, color="#3370FF")
    ax.set_title(title, fontsize=14, fontweight="bold")
    ax.set_ylabel(ylabel)
    ax.set_xlabel("日期")
    ax.grid(axis="y", alpha=0.3)
    fig.autofmt_xdate()

    # 标注最高点和最低点
    if values:
        max_idx = values.index(max(values))
        min_idx = values.index(min(values))
        ax.annotate(f"最高 {values[max_idx]}", xy=(max_idx, values[max_idx]),
                    xytext=(0, 10), textcoords="offset points", fontsize=9,
                    color="#D32F2F", ha="center")
        ax.annotate(f"最低 {values[min_idx]}", xy=(min_idx, values[min_idx]),
                    xytext=(0, -15), textcoords="offset points", fontsize=9,
                    color="#388E3C", ha="center")

    return _save_chart(fig, filename)


def pie_chart(labels: list, values: list, title: str = "占比分析",
              filename: str = None) -> str:
    """饼图：各平台占比。"""
    if filename is None:
        filename = f"pie_{datetime.now().strftime('%Y%m%d%H%M%S')}.png"

    colors = ["#3370FF", "#34C759", "#F5A623", "#FF6B6B"]
    explode = [0.02] * len(labels)

    fig, ax = plt.subplots(figsize=(6, 6))
    wedges, texts, autotexts = ax.pie(
        values, explode=explode, labels=labels, colors=colors[:len(labels)],
        autopct="%1.1f%%", startangle=140, pctdistance=0.6,
    )
    for t in autotexts:
        t.set_fontsize(11)
        t.set_fontweight("bold")
    ax.set_title(title, fontsize=14, fontweight="bold")

    return _save_chart(fig, filename)


def bar_chart(labels: list, values: list, title: str = "排名对比",
              ylabel: str = "", filename: str = None) -> str:
    """柱状图：人员排名对比。"""
    if filename is None:
        filename = f"bar_{datetime.now().strftime('%Y%m%d%H%M%S')}.png"

    colors = ["#3370FF"] * len(labels)
    # 前三名高亮
    if len(colors) >= 1:
        colors[0] = "#FFB800"  # 金
    if len(colors) >= 2:
        colors[1] = "#A0A0A0"  # 银
    if len(colors) >= 3:
        colors[2] = "#CD7F32"  # 铜

    fig, ax = plt.subplots(figsize=(8, 5))
    bars = ax.bar(labels, values, color=colors, edgecolor="white", linewidth=0.5)

    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.01,
                str(val), ha="center", fontsize=10, fontweight="bold")

    ax.set_title(title, fontsize=14, fontweight="bold")
    ax.set_ylabel(ylabel)
    ax.grid(axis="y", alpha=0.3)

    return _save_chart(fig, filename)


def dual_line(dates: list, values1: list, values2: list,
              label1: str = "会话量", label2: str = "订单数",
              title: str = "双指标趋势", filename: str = None) -> str:
    """双折线图：两组数据在同一图上展示。"""
    if filename is None:
        filename = f"dual_{datetime.now().strftime('%Y%m%d%H%M%S')}.png"

    fig, ax1 = plt.subplots(figsize=(10, 4))

    ax1.plot(dates, values1, marker="o", linewidth=2, markersize=6,
             color="#3370FF", label=label1)
    ax1.set_ylabel(label1, color="#3370FF")
    ax1.fill_between(range(len(dates)), values1, alpha=0.1, color="#3370FF")

    ax2 = ax1.twinx()
    ax2.plot(dates, values2, marker="s", linewidth=2, markersize=6,
             color="#F5A623", label=label2, linestyle="--")
    ax2.set_ylabel(label2, color="#F5A623")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left")

    ax1.set_title(title, fontsize=14, fontweight="bold")
    ax1.grid(axis="y", alpha=0.3)
    fig.autofmt_xdate()

    return _save_chart(fig, filename)
