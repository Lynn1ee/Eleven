#!/usr/bin/env python
r"""
客服数据周报 HTML 生成器

用法:
    python gen_weekly_report.py <Excel路径> [输出路径]

示例:
    python gen_weekly_report.py data.xlsx report.html

Excel 格式要求:
  Row 1:  指标 | 时间 | 9个客服名字
  Row 2-4:  会话量（上周/本周/环比）
  Row 5-7:  平均处理时长
  Row 8-10: 评价率
  Row 11-13: 差评率(Excel中标为"满意度")
  Row 14-16: 无声占比
  Row 17-19: 无声会话平均处理时长
  Row 20:  空行
  Row 21-23: 会话总量
  Row 24-25: 无声会话量
"""

import json
import sys
import os
import re
from pathlib import Path

import openpyxl


def read_excel(excel_path):
    """从 Excel 读取所有数据，返回结构化的 dict"""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    def get_row(row_num):
        row = list(ws.iter_rows(min_row=row_num, max_row=row_num, values_only=True))[0]
        return row

    # Row 1: 表头
    header = get_row(1)
    agents = list(header[2:])  # 跳过"指标"和"时间"列

    # Row 2-4: 会话量
    week1_label = str(get_row(2)[1])
    week2_label = str(get_row(3)[1])

    def parse_data(row_num):
        """提取某行的9个客服数据"""
        row = get_row(row_num)
        return [float(v) if v is not None and v != '' else 0.0 for v in row[2:]]

    data = {
        'agents': agents,
        'week1': week1_label,  # 如 "06.08-06.14"
        'week2': week2_label,  # 如 "06.15-06.21"
        # 日均会话量
        'sa1': parse_data(2), 'sa2': parse_data(3), 'saQ': parse_data(4),
        # 平均处理时长
        'du1': parse_data(5), 'du2': parse_data(6), 'duQ': parse_data(7),
        # 评价率（原始为小数，HTML中×100显示）
        'ra1': parse_data(8), 'ra2': parse_data(9), 'raQ': parse_data(10),
        # 差评率（Excel中标为"满意度"）
        'st1': parse_data(11), 'st2': parse_data(12), 'stQ': parse_data(13),
        # 无声占比
        'si1': parse_data(14), 'si2': parse_data(15), 'siQ': parse_data(16),
        # 无声处理时长
        'sd1': parse_data(17), 'sd2': parse_data(18), 'sdQ': parse_data(19),
        # Row 21-23: 会话总量
        'ts1': [int(v) for v in parse_data(22)],
        'ts2': [int(v) for v in parse_data(23)],
        # Row 24-25: 无声会话量
        'sv1': [int(v) for v in parse_data(24)],
        'sv2': [int(v) for v in parse_data(25)],
    }
    return data


# ============================================================
# 自动分析文字生成
# ============================================================

def find_top_changes(values, qoq, n=3, reverse=False):
    """找出环比变化最大的 n 个人，返回 [(name, old, new, qoq), ...]"""
    indexed = list(enumerate(qoq))
    indexed.sort(key=lambda x: x[1], reverse=not reverse)  # 默认降序
    result = []
    for i, q in indexed[:n]:
        result.append((i, values[0][i], values[1][i], q))
    return result


def pct(v, decimals=2):
    return f"{v*100:+.{decimals}f}%"


def pct_v(v, decimals=2):
    """将小数转为百分比字符串"""
    return f"{v*100:.{decimals}f}%"


def value_fmt(v, is_pct=False):
    if is_pct:
        return pct_v(v)
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return f"{v:.2f}" if isinstance(v, float) else str(v)


def generate_insight_text(d):
    """基于数据自动生成三部分分析文字"""
    agents = d['agents']
    w1, w2 = d['week1'], d['week2']

    # ---- 团队整体趋势 ----
    w2_total = sum(d['ts2'])
    w1_total = sum(d['ts1'])
    total_chg = (w2_total - w1_total) / w1_total

    avg_du2 = sum(d['du2']) / len(agents)
    avg_du1 = sum(d['du1']) / len(agents)
    du_chg = (avg_du2 - avg_du1) / avg_du1

    avg_ra2 = sum(d['ra2']) / len(agents)
    avg_ra1 = sum(d['ra1']) / len(agents)
    ra_chg = (avg_ra2 - avg_ra1) / avg_ra1

    avg_st2 = sum(d['st2']) / len(agents)
    avg_st1 = sum(d['st1']) / len(agents)
    st_chg = (avg_st2 - avg_st1) / avg_st1

    avg_si2 = sum(d['si2']) / len(agents)
    avg_si1 = sum(d['si1']) / len(agents)
    si_chg = (avg_si2 - avg_si1) / avg_si1

    # 计数上升/下降人数
    sa_up = sum(1 for q in d['saQ'] if q > 0)
    sa_dn = sum(1 for q in d['saQ'] if q < 0)

    st_up = sum(1 for q in d['stQ'] if q > 0.01)
    st_dn = sum(1 for q in d['stQ'] if q < -0.01)

    si_up = sum(1 for q in d['siQ'] if q > 0.01)
    si_dn = sum(1 for q in d['siQ'] if q < -0.01)

    ra_up = sum(1 for q in d['raQ'] if q > 0.01)
    ra_dn = sum(1 for q in d['raQ'] if q < -0.01)

    du_up = sum(1 for q in d['duQ'] if q > 0.01)
    du_dn = sum(1 for q in d['duQ'] if q < -0.01)

    # 找出极端变化的人
    sa_top = find_top_changes((d['sa1'], d['sa2']), d['saQ'], 2)
    st_worst = find_top_changes((d['st1'], d['st2']), d['stQ'], 2)  # 差评率上升最多的
    st_best = find_top_changes((d['st1'], d['st2']), d['stQ'], 1, reverse=True)  # 差评率下降的
    si_worst = find_top_changes((d['si1'], d['si2']), d['siQ'], 2)
    du_best = find_top_changes((d['du1'], d['du2']), d['duQ'], 1, reverse=True)
    du_worst = find_top_changes((d['du1'], d['du2']), d['duQ'], 1)

    def name(i):
        return agents[i]

    # === Part 1: 团队整体趋势 ===
    trend_lines = []
    trend_lines.append(
        f'<li>总会话量 {w2_total:,} 较上周 {w1_total:,} '
        f'{"上升" if total_chg>=0 else "下降"} <b>{pct(total_chg, 1)}</b>，'
        f'{sa_up}人上升{sa_dn}人下降。</li>'
    )

    du_desc = "改善" if du_chg < 0 else ("上升" if du_chg > 0.005 else "基本持平")
    du_color = ""
    trend_lines.append(
        f'<li>处理时长团队均值 {avg_du2:.2f}s，较上周 {avg_du1:.2f}s '
        f'{du_desc} <b>{pct(du_chg, 1)}</b>。</li>'
    )

    trend_lines.append(
        f'<li>评价率团队均值从 {pct_v(avg_ra1)} → {pct_v(avg_ra2)}（<b>{pct(ra_chg)}</b>），'
        f'{ra_up}人上升{ra_dn}人下降。</li>'
    )

    st_alarm = " &#x26A0;&#xFE0F;" if st_chg > 0.2 else ""
    trend_lines.append(
        f'<li>差评率团队均值从 {pct_v(avg_st1)} → {pct_v(avg_st2)}（<b>{pct(st_chg)}</b>），'
        f'{st_up}人上升{st_dn}人下降。{st_alarm}</li>'
    )

    si_alarm = " &#x26A0;&#xFE0F; 连续恶化！" if si_chg > 0.5 else ""
    trend_lines.append(
        f'<li><b>无声占比：</b>团队均值从 {pct_v(avg_si1)} → {pct_v(avg_si2)}，'
        f'<b>{pct(si_chg)}</b>，{si_up}人上升{si_dn}人下降。{si_alarm}</li>'
    )

    # 建议
    alerts = []
    if st_chg > 0.15:
        alerts.append(f'差评率管控（{name(st_worst[0][0])}{pct(st_worst[0][3])}、{name(st_worst[1][0])}{pct(st_worst[1][3])}）')
    if si_chg > 0.2:
        alerts.append(f'无声占比管控（{name(si_worst[0][0])}{pct(si_worst[0][3])}、{name(si_worst[1][0])}{pct(si_worst[1][3])}）')
    if alerts:
        trend_lines.append(f'<li>建议下周重点关注：{"、".join(alerts)}。</li>')

    part1 = '\n'.join(trend_lines)

    # === Part 2: 本周核心发现 ===
    findings = []

    # 差评率
    if st_chg > 0.1:
        st_items = '、'.join([f'{name(i)}{pct(q)}（{pct_v(o)}→{pct_v(n)}）'
                              for i, o, n, q in st_worst[:3]])
        findings.append(f'<li><b>差评率恶化：</b>{st_up}人上升，{st_items}增幅最大</li>')
    else:
        st_items_best = '、'.join([f'{name(i)}{pct(q)}（{pct_v(o)}→{pct_v(n)}）'
                                   for i, o, n, q in st_best[:2]])
        findings.append(f'<li><b>差评率改善：</b>{st_dn}人下降，{st_items_best}改善最显著</li>')

    # 无声占比
    if si_up > len(agents) // 2:
        si_items = '、'.join([f'{name(i)}{pct(q)}（{pct_v(o)}→{pct_v(n)}）'
                              for i, o, n, q in si_worst[:3]])
        findings.append(f'<li><b>无声占比大面积恶化：</b>{si_up}人上升，{si_items}增幅最大</li>')
    else:
        findings.append(f'<li><b>无声占比：</b>{si_up}人上升{si_dn}人下降，团队均值{pct_v(avg_si2)}</li>')

    # 处理时长
    du_best_name = name(du_best[0][0])
    findings.append(
        f'<li><b>处理时长：</b>{du_best_name}改善最显著{pct(du_best[0][3])}'
        f'（{du_best[0][1]:.1f}s→{du_best[0][2]:.1f}s），'
        f'{name(du_worst[0][0]) if du_worst else ""}'
        f'{"上升" if du_worst and du_worst[0][3] > 0 else ""}最多'
        f'{pct(du_worst[0][3]) if du_worst else ""}'
        f'（{du_worst[0][1]:.1f}s→{du_worst[0][2]:.1f}s）' if du_worst else ''
        f'</li>'
    )

    # 评价率
    ra_top = find_top_changes((d['ra1'], d['ra2']), d['raQ'], 1)
    ra_bottom = find_top_changes((d['ra1'], d['ra2']), d['raQ'], 1, reverse=True)
    findings.append(
        f'<li><b>评价率：</b>{name(ra_top[0][0])}上升最多{pct(ra_top[0][3])}'
        f'（{pct_v(ra_top[0][1])}→{pct_v(ra_top[0][2])}），'
        f'{name(ra_bottom[0][0])}下降最多{pct(ra_bottom[0][3])}'
        f'（{pct_v(ra_bottom[0][1])}→{pct_v(ra_bottom[0][2])}）</li>'
    )

    # 找出各指标的最值
    max_sa_idx = max(range(len(agents)), key=lambda i: d['sa2'][i])
    max_st_idx = max(range(len(agents)), key=lambda i: d['st2'][i])
    min_st_idx = min(range(len(agents)), key=lambda i: d['st2'][i])
    max_si_idx = max(range(len(agents)), key=lambda i: d['si2'][i])
    min_si_idx = min(range(len(agents)), key=lambda i: d['si2'][i])
    max_du_idx = max(range(len(agents)), key=lambda i: d['du2'][i])
    max_ts_idx = max(range(len(agents)), key=lambda i: d['ts2'][i])

    findings.append(
        f'<li><b>本周之最：</b>日均会话量最高 {name(max_sa_idx)}（{d["sa2"][max_sa_idx]:.0f}），'
        f'差评率最高 {name(max_st_idx)}（{pct_v(d["st2"][max_st_idx])}），'
        f'无声占比最高 {name(max_si_idx)}（{pct_v(d["si2"][max_si_idx])}），'
        f'处理时长最高 {name(max_du_idx)}（{d["du2"][max_du_idx]:.1f}s）</li>'
    )

    part2 = '\n'.join(findings)

    # === Part 3: 综合排名 + 亮点/警示 ===
    # 五维评分
    def normalize(arr, asc=True):
        mn, mx = min(arr), max(arr)
        if mx == mn:
            return [50] * len(arr)
        return [(v - mn) / (mx - mn) * 100 if asc else (mx - v) / (mx - mn) * 100 for v in arr]

    ss = normalize(d['sa2'], True)
    sd_n = normalize(d['du2'], False)
    sr = normalize(d['ra2'], True)
    sst_n = normalize(d['st2'], False)
    ssi_n = normalize(d['si2'], False)
    scores = [ss[i]*0.2 + sd_n[i]*0.2 + sr[i]*0.2 + sst_n[i]*0.2 + ssi_n[i]*0.2 for i in range(len(agents))]
    ranked = sorted(zip(agents, scores), key=lambda x: x[1], reverse=True)

    # 亮点：前3名中在某项指标上表现突出的
    highlights = []
    warnings = []

    # 倒数3名
    bottom3 = [ranked[-1], ranked[-2], ranked[-3]]

    # 各项指标找亮点
    for metric_name, arr, asc, good_label in [
        ('会话量', d['sa2'], True, '日均会话量'),
        ('处理时长', d['du2'], False, '处理时长'),
        ('评价率', d['ra2'], True, '评价率'),
        ('差评率', d['st2'], False, '差评率'),
        ('无声占比', d['si2'], False, '无声占比'),
    ]:
        if asc:
            best_idx = max(range(len(agents)), key=lambda i: arr[i])
        else:
            best_idx = min(range(len(agents)), key=lambda i: arr[i])
        val = arr[best_idx]
        val_str = pct_v(val) if metric_name in ('评价率', '差评率', '无声占比') else (
            f'{val:.0f}' if metric_name == '会话量' else f'{val:.1f}s')
        highlights.append(f'{name(best_idx)}（{good_label}{val_str}）')

    # 警示：各项指标最差的
    for metric_name, arr, asc, bad_label in [
        ('差评率', d['st2'], True, '差评率'),
        ('无声占比', d['si2'], True, '无声占比'),
        ('处理时长', d['du2'], True, '处理时长'),
    ]:
        worst_idx = max(range(len(agents)), key=lambda i: arr[i])
        q = None
        if metric_name == '差评率':
            q = d['stQ'][worst_idx]
        elif metric_name == '无声占比':
            q = d['siQ'][worst_idx]
        elif metric_name == '处理时长':
            q = d['duQ'][worst_idx]

        val_str = pct_v(arr[worst_idx]) if metric_name in ('差评率', '无声占比') else f'{arr[worst_idx]:.1f}s'
        q_str = f'环比{pct(q)}' if q and abs(q) > 0.05 else ''
        warnings.append(f'{name(worst_idx)}（{bad_label}{val_str}全队最高{q_str}）')

    # 综合排名倒数3名
    bottom_str = '、'.join([f'{n}（综合倒数第{len(ranked)-i}）' for i, (n, s) in enumerate(ranked[-3:])])

    highlights_str = '、'.join(highlights[:5])
    warnings_str = '、'.join(warnings[:5])

    part3 = (
        f'<li><b>亮点：</b>{highlights_str}</li>\n'
        f'<li><b>警示：</b>{warnings_str}</li>'
    )

    return part1, part2, part3, ranked


# ============================================================
# HTML 生成
# ============================================================

COLORS = ['#667eea', '#f093fb', '#4facfe', '#43e97b', '#fa709a', '#fee140', '#a18cd1', '#fbc2eb', '#ff9a76']

CSS = r'''<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','PingFang SC','Microsoft YaHei',sans-serif;background:#f0f2f5;color:#333;padding:20px}
.header{text-align:center;padding:30px 20px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:#fff;border-radius:16px;margin-bottom:24px;box-shadow:0 4px 15px rgba(102,126,234,0.4)}
.header h1{font-size:28px;font-weight:700;letter-spacing:2px}.header .subtitle{font-size:15px;opacity:0.9;margin-top:8px}
.header .date-badge{display:inline-block;background:rgba(255,255,255,0.2);padding:6px 20px;border-radius:20px;font-size:14px;margin-top:12px}
.summary-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;margin-bottom:24px}
.card{background:#fff;border-radius:12px;padding:20px;box-shadow:0 2px 8px rgba(0,0,0,0.06);transition:transform 0.2s}.card:hover{transform:translateY(-2px);box-shadow:0 4px 16px rgba(0,0,0,0.1)}
.card .label{font-size:13px;color:#888;margin-bottom:6px}.card .value{font-size:26px;font-weight:700;color:#333}
.card .change{font-size:13px;margin-top:4px}.card .change.good{color:#27ae60}.card .change.bad{color:#e74c3c}
.section{background:#fff;border-radius:12px;padding:24px;margin-bottom:24px;box-shadow:0 2px 8px rgba(0,0,0,0.06)}
.section-title{font-size:18px;font-weight:700;color:#333;margin-bottom:20px;padding-left:12px;border-left:4px solid #667eea}
.chart-row{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}
.chart-container{position:relative;width:100%;height:320px}.chart-full{position:relative;width:100%;height:380px}
table{width:100%;border-collapse:collapse;font-size:13px}th{background:#f7f8fc;padding:10px 8px;text-align:center;font-weight:600;color:#555;border-bottom:2px solid #eee}
td{padding:8px;text-align:center;border-bottom:1px solid #f0f0f0}tr:hover td{background:#fafbff}
.name-col{font-weight:600;color:#444;text-align:left!important}.trend-up{color:#e74c3c;font-weight:600}.trend-down{color:#27ae60;font-weight:600}
.analysis-box{background:#f8f9ff;border-radius:10px;padding:18px;margin:16px 0;border-left:4px solid #667eea;line-height:1.8;font-size:14px}
.analysis-box h4{color:#667eea;margin-bottom:8px;font-size:15px}.analysis-box ul{padding-left:18px}.analysis-box li{margin:4px 0}
.tag{display:inline-block;padding:2px 10px;border-radius:10px;font-size:12px;font-weight:600}
.tag-red{background:#fee;color:#e74c3c}.tag-green{background:#efe;color:#27ae60}.tag-yellow{background:#fef9e7;color:#f39c12}
.footer{text-align:center;padding:20px;color:#aaa;font-size:12px}
@media(max-width:900px){.chart-row{grid-template-columns:1fr}}
</style>'''


def build_html(d, part1, part2, part3, ranked):
    agents = d['agents']
    w1_label = d['week1']
    w2_label = d['week2']
    # 格式化日期用于显示
    def fmt_date(label):
        parts = label.replace('.', '-').split('-')
        if len(parts) == 2:
            m1, d1 = parts
            return f"6月{d1}日"
        elif len(parts) == 4:
            return f"{parts[0]}月{parts[1]}日 ~ {parts[2]}月{parts[3]}日"
        return label

    # 解析月份和日期
    w2_parts = w2_label.replace('.', '-').split('-')
    if len(w2_parts) == 4:
        date_display = f"2026年{w2_parts[0]}月{w2_parts[1]}日 ~ {w2_parts[2]}月{w2_parts[3]}日"
    else:
        date_display = w2_label

    w1_parts = w1_label.replace('.', '-').split('-')
    if len(w1_parts) == 4:
        w1_date_display = f"{w1_parts[0]}.{w1_parts[1]}-{w1_parts[2]}.{w1_parts[3]}"
    else:
        w1_date_display = w1_label

    title_date = w2_label.replace('.', '-')

    html = f'''<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>客服数据周报 - {title_date}</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
{CSS}
</head><body>
<div class="header"><h1>&#x1F4CA; 客服团队数据周报</h1><div class="subtitle">Customer Service Weekly Analytics Report</div>
<div class="date-badge">&#x1F4C5; 统计周期：{date_display}（对比上周 {w1_date_display}）</div></div>
<div class="summary-cards" id="summaryCards"></div>
<div class="section"><div class="section-title">&#x1F4CB; 核心指标总览（{w2_label}）</div><div style="overflow-x:auto"><table id="mainTable"></table></div></div>
<div class="section"><div class="section-title">&#x1F4AC; 会话量分析</div><div class="chart-row"><div class="chart-container"><canvas id="chartSessionVol"></canvas></div><div class="chart-container"><canvas id="chartSessionQoQ"></canvas></div></div><div id="sessionAnalysis" class="analysis-box"></div></div>
<div class="section"><div class="section-title">&#x23F1;&#xFE0F; 平均处理时长分析</div><div class="chart-row"><div class="chart-container"><canvas id="chartDuration"></canvas></div><div class="chart-container"><canvas id="chartDurationQoQ"></canvas></div></div><div id="durationAnalysis" class="analysis-box"></div></div>
<div class="section"><div class="section-title">&#x2B50; 评价率 & 差评率分析</div><div class="chart-row"><div class="chart-container"><canvas id="chartRating"></canvas></div><div class="chart-container"><canvas id="chartSatisfaction"></canvas></div></div><div class="chart-row"><div class="chart-container"><canvas id="chartRadar"></canvas></div><div class="chart-container"><canvas id="chartRatingQoQ"></canvas></div></div><div id="ratingAnalysis" class="analysis-box"></div></div>
<div class="section"><div class="section-title">&#x1F507; 无声会话专项分析</div><div class="chart-row"><div class="chart-container"><canvas id="chartSilentRatio"></canvas></div><div class="chart-container"><canvas id="chartSilentVol"></canvas></div></div><div class="chart-full"><canvas id="chartSilentDuration"></canvas></div><div id="silentAnalysis" class="analysis-box"></div></div>
<div class="section"><div class="section-title">&#x1F3C6; 本周综合表现排名</div><div class="chart-full"><canvas id="chartRanking"></canvas></div><div id="rankingAnalysis" class="analysis-box"></div></div>
<div class="section"><div class="section-title">&#x1F4A1; 数据洞察</div>
<div class="analysis-box"><h4>&#x1F4CA; 团队整体趋势</h4><ul>
{part1}
</ul></div>
<div class="analysis-box"><h4>&#x1F4C8; 本周核心发现</h4><ul>
{part2}
</ul></div>
<div class="analysis-box"><h4>&#x1F3C6; 值得关注的个人表现</h4><ul>
{part3}
</ul></div></div>
<div class="footer">客服数据周报 · 自动生成 · 数据来源：{os.path.basename(sys.argv[1]) if len(sys.argv) > 1 else 'Excel'}</div>
'''

    # JS 数据部分
    js_data = f'''
<script>
const dataLabelsPlugin = {{
  id: 'dataLabels',
  afterDatasetsDraw(chart, args, options) {{
    if (chart.config.type === 'radar') return;
    const {{ ctx }} = chart;
    ctx.save();
    ctx.font = 'bold 10px "PingFang SC","Microsoft YaHei",sans-serif';
    ctx.fillStyle = '#444';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'bottom';
    chart.data.datasets.forEach((dataset, i) => {{
      const meta = chart.getDatasetMeta(i);
      if (!meta.hidden) {{
        meta.data.forEach((element, index) => {{
          const value = dataset.data[index];
          if (value !== null && value !== undefined) {{
            const x = element.x;
            const y = element.y - 3;
            let text = typeof value === 'number' ? value.toFixed(2) : String(value);
            ctx.fillText(text, x, y);
          }}
        }});
      }}
    }});
    ctx.restore();
  }}
}};
Chart.register(dataLabelsPlugin);
Chart.defaults.font.family='-apple-system,"PingFang SC","Microsoft YaHei",sans-serif';
Chart.defaults.color='#666';

const agents={json.dumps(agents)};
const colors={json.dumps(COLORS)};

const sa1={json.dumps(d['sa1'])};
const sa2={json.dumps(d['sa2'])};
const saQ={json.dumps(d['saQ'])};
const du1={json.dumps(d['du1'])};
const du2={json.dumps(d['du2'])};
const duQ={json.dumps(d['duQ'])};
const ra1={json.dumps(d['ra1'])};
const ra2={json.dumps(d['ra2'])};
const raQ={json.dumps(d['raQ'])};
const st1={json.dumps(d['st1'])};
const st2={json.dumps(d['st2'])};
const stQ={json.dumps(d['stQ'])};
const si1={json.dumps(d['si1'])};
const si2={json.dumps(d['si2'])};
const siQ={json.dumps(d['siQ'])};
const sd1={json.dumps(d['sd1'])};
const sd2={json.dumps(d['sd2'])};
const sdQ={json.dumps(d['sdQ'])};
const sv1={json.dumps(d['sv1'])};
const sv2={json.dumps(d['sv2'])};
const ts1={json.dumps(d['ts1'])};
const ts2={json.dumps(d['ts2'])};
const week1Label='{w1_label}';
const week2Label='{w2_label}';

function fmt(n,d){{d=d||2;return Number(n).toFixed(d)}}
function fmtPct(n){{return(n>=0?'+':'')+(n*100).toFixed(2)+'%'}}
function tag(c,t){{return'<span class="tag tag-'+c+'">'+t+'</span>'}}

(function(){{
const w2T=ts2.reduce((a,b)=>a+b,0),w1T=ts1.reduce((a,b)=>a+b,0);
const ad2=du2.reduce((a,b)=>a+b,0)/9,ad1=du1.reduce((a,b)=>a+b,0)/9;
const as2=st2.reduce((a,b)=>a+b,0)/9,as1=st1.reduce((a,b)=>a+b,0)/9;
const ar2=ra2.reduce((a,b)=>a+b,0)/9,ar1=ra1.reduce((a,b)=>a+b,0)/9;
const ai2=si2.reduce((a,b)=>a+b,0)/9,ai1=si1.reduce((a,b)=>a+b,0)/9;
const cards=[
{{l:'&#x1F4CC; 本周总会话量',v:w2T.toLocaleString(),c:((w2T-w1T)/w1T*100).toFixed(1)+'%',d:w2T>=w1T?'上升':'下降',g:w2T>=w1T}},
{{l:'&#x23F1;&#xFE0F; 团队均处理时长',v:fmt(ad2)+'s',c:((ad2-ad1)/ad1*100).toFixed(1)+'%',d:ad2>=ad1?'上升':'下降',g:ad2<=ad1}},
{{l:'&#x2B50; 团队平均差评率',v:fmt(as2*100)+'%',c:((as2-as1)/as1*100).toFixed(1)+'%',d:as2>=as1?'上升':'下降',g:as2<=as1}},
{{l:'&#x1F4DD; 团队平均评价率',v:fmt(ar2*100)+'%',c:((ar2-ar1)/ar1*100).toFixed(1)+'%',d:ar2>=ar1?'上升':'下降',g:ar2>=ar1}},
{{l:'&#x1F507; 团队无声占比',v:fmt(ai2*100)+'%',c:((ai2-ai1)/ai1*100).toFixed(1)+'%',d:ai2>=ai1?'上升':'下降',g:ai2<=ai1}}
];
document.getElementById('summaryCards').innerHTML=cards.map(c=>'<div class="card"><div class="label">'+c.l+'</div><div class="value">'+c.v+'</div><div class="change '+(c.g?'good':'bad')+'">'+(c.c?'较上周 '+c.c+'（'+c.d+'）':'-')+'</div></div>').join('');
}})();

(function(){{
const rows=[
{{n:'日均会话量',w1:sa1,w2:sa2,q:saQ,fv:v=>Math.round(v),fq:v=>fmtPct(v),inv:false}},
{{n:'平均处理时长(s)',w1:du1,w2:du2,q:duQ,fv:v=>fmt(v,2),fq:v=>fmtPct(v),inv:true}},
{{n:'评价率(%)',w1:ra1.map(v=>v*100),w2:ra2.map(v=>v*100),q:raQ,fv:v=>fmt(v,2),fq:v=>fmtPct(v),inv:false}},
{{n:'差评率(%)',w1:st1.map(v=>v*100),w2:st2.map(v=>v*100),q:stQ,fv:v=>fmt(v,2),fq:v=>fmtPct(v),inv:true}},
{{n:'无声占比(%)',w1:si1.map(v=>v*100),w2:si2.map(v=>v*100),q:siQ,fv:v=>fmt(v,2),fq:v=>fmtPct(v),inv:true}},
{{n:'无声处理时长(s)',w1:sd1,w2:sd2,q:sdQ,fv:v=>fmt(v,2),fq:v=>fmtPct(v),inv:true}},
{{n:'会话总量',w1:ts1,w2:ts2,q:null,fv:v=>Math.round(v).toLocaleString(),fq:null,inv:false}},
{{n:'无声会话量',w1:sv1,w2:sv2,q:null,fv:v=>Math.round(v),fq:null,inv:false}}
];
let h='<tr><th>指标</th><th>周期</th>';agents.forEach(a=>h+='<th>'+a+'</th>');h+='</tr>';
rows.forEach(r=>{{
h+='<tr><td class="name-col" rowspan="'+(r.q?3:2)+'">'+r.n+'</td><td>上周</td>';r.w1.forEach(v=>h+='<td>'+r.fv(v)+'</td>');h+='</tr><tr><td>本周</td>';r.w2.forEach(v=>h+='<td>'+r.fv(v)+'</td>');h+='</tr>';
if(r.q){{h+='<tr><td>环比</td>';r.q.forEach(v=>{{const g=r.inv?v<0:v>0;h+='<td class="'+(g?'trend-down':'trend-up')+'">'+r.fq(v)+'</td>';}});h+='</tr>';}}
}});document.getElementById('mainTable').innerHTML=h;
}})();

// Charts
new Chart(document.getElementById('chartSessionVol'),{{type:'bar',data:{{labels:agents,datasets:[
{{label:'上周('+week1Label+')',data:sa1,backgroundColor:'rgba(102,126,234,0.6)',borderRadius:4}},
{{label:'本周('+week2Label+')',data:sa2,backgroundColor:'rgba(118,75,162,0.7)',borderRadius:4}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'top'}},title:{{display:true,text:'日均会话量对比',font:{{size:14}}}}}},scales:{{y:{{beginAtZero:true,grid:{{color:'#f0f0f0'}}}},x:{{grid:{{display:false}}}}}}}}}});

new Chart(document.getElementById('chartSessionQoQ'),{{type:'bar',data:{{labels:agents,datasets:[{{label:'环比变化(%)',data:saQ.map(v=>(v*100).toFixed(2)),backgroundColor:saQ.map(v=>v>=0?'rgba(39,174,96,0.7)':'rgba(231,76,60,0.7)'),borderRadius:4}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},title:{{display:true,text:'会话量环比变化 %',font:{{size:14}}}}}},scales:{{y:{{grid:{{color:'#f0f0f0'}}}},x:{{grid:{{display:false}}}}}}}}}});

new Chart(document.getElementById('chartDuration'),{{type:'bar',data:{{labels:agents,datasets:[
{{label:'上周(s)',data:du1,backgroundColor:'rgba(102,126,234,0.6)',borderRadius:4}},
{{label:'本周(s)',data:du2,backgroundColor:'rgba(231,76,60,0.6)',borderRadius:4}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'top'}},title:{{display:true,text:'平均处理时长对比(秒)',font:{{size:14}}}}}},scales:{{y:{{beginAtZero:false,grid:{{color:'#f0f0f0'}}}},x:{{grid:{{display:false}}}}}}}}}});

new Chart(document.getElementById('chartDurationQoQ'),{{type:'bar',data:{{labels:agents,datasets:[{{label:'环比变化(%)',data:duQ.map(v=>(v*100).toFixed(2)),backgroundColor:duQ.map(v=>v<=0?'rgba(39,174,96,0.7)':'rgba(231,76,60,0.7)'),borderRadius:4}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},title:{{display:true,text:'处理时长环比变化 %',font:{{size:14}}}}}},scales:{{y:{{grid:{{color:'#f0f0f0'}}}},x:{{grid:{{display:false}}}}}}}}}});

new Chart(document.getElementById('chartRating'),{{type:'bar',data:{{labels:agents,datasets:[
{{label:'上周(%)',data:ra1.map(v=>v*100),backgroundColor:'rgba(102,126,234,0.6)',borderRadius:4}},
{{label:'本周(%)',data:ra2.map(v=>v*100),backgroundColor:'rgba(240,147,251,0.7)',borderRadius:4}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'top'}},title:{{display:true,text:'评价率对比 (%)',font:{{size:14}}}}}},scales:{{y:{{beginAtZero:true,max:50,grid:{{color:'#f0f0f0'}}}},x:{{grid:{{display:false}}}}}}}}}});

new Chart(document.getElementById('chartSatisfaction'),{{type:'bar',data:{{labels:agents,datasets:[
{{label:'上周(%)',data:st1.map(v=>v*100),backgroundColor:'rgba(102,126,234,0.6)',borderRadius:4}},
{{label:'本周(%)',data:st2.map(v=>v*100),backgroundColor:'rgba(79,172,254,0.7)',borderRadius:4}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'top'}},title:{{display:true,text:'差评率对比 (%)',font:{{size:14}}}}}},scales:{{y:{{beginAtZero:true,max:Math.ceil(Math.max(...st2)*105),grid:{{color:'#f0f0f0'}}}},x:{{grid:{{display:false}}}}}}}}}});

new Chart(document.getElementById('chartRadar'),{{type:'radar',data:{{labels:agents,datasets:[
{{label:'评价率(本周%)',data:ra2.map(v=>v*100),borderColor:'#667eea',backgroundColor:'rgba(102,126,234,0.15)',pointBackgroundColor:'#667eea'}},
{{label:'差评率(本周%×3)',data:st2.map(v=>v*300),borderColor:'#f093fb',backgroundColor:'rgba(240,147,251,0.15)',pointBackgroundColor:'#f093fb'}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'top'}},title:{{display:true,text:'服务质量雷达图（本周）',font:{{size:14}}}},dataLabels:{{display:false}}}},scales:{{r:{{beginAtZero:true,grid:{{color:'#e0e0e0'}},angleLines:{{color:'#e0e0e0'}},pointLabels:{{font:{{size:11}}}}}}}}}}}});

new Chart(document.getElementById('chartRatingQoQ'),{{type:'bar',data:{{labels:agents,datasets:[
{{label:'评价率环比%',data:raQ.map(v=>(v*100).toFixed(2)),backgroundColor:raQ.map(v=>v>=0?'rgba(39,174,96,0.7)':'rgba(231,76,60,0.7)'),borderRadius:4}},
{{label:'差评率环比%',data:stQ.map(v=>(v*100).toFixed(2)),backgroundColor:stQ.map(v=>v<=0?'rgba(39,174,96,0.7)':'rgba(231,76,60,0.7)'),borderRadius:4}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'top'}},title:{{display:true,text:'评价率 & 差评率环比变化 %',font:{{size:14}}}}}},scales:{{y:{{grid:{{color:'#f0f0f0'}}}},x:{{grid:{{display:false}}}}}}}}}});

new Chart(document.getElementById('chartSilentRatio'),{{type:'bar',data:{{labels:agents,datasets:[
{{label:'上周(%)',data:si1.map(v=>v*100),backgroundColor:'rgba(102,126,234,0.6)',borderRadius:4}},
{{label:'本周(%)',data:si2.map(v=>v*100),backgroundColor:'rgba(231,76,60,0.7)',borderRadius:4}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'top'}},title:{{display:true,text:'无声占比对比 (%)',font:{{size:14}}}}}},scales:{{y:{{beginAtZero:true,max:Math.ceil(Math.max(...si2)*110),grid:{{color:'#f0f0f0'}}}},x:{{grid:{{display:false}}}}}}}}}});

new Chart(document.getElementById('chartSilentVol'),{{type:'bar',data:{{labels:agents,datasets:[
{{label:'上周',data:sv1,backgroundColor:'rgba(102,126,234,0.6)',borderRadius:4}},
{{label:'本周',data:sv2,backgroundColor:'rgba(231,76,60,0.7)',borderRadius:4}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'top'}},title:{{display:true,text:'无声会话量（绝对数）',font:{{size:14}}}}}},scales:{{y:{{beginAtZero:true,grid:{{color:'#f0f0f0'}}}},x:{{grid:{{display:false}}}}}}}}}});

new Chart(document.getElementById('chartSilentDuration'),{{type:'bar',data:{{labels:agents,datasets:[
{{label:'上周(s)',data:sd1,backgroundColor:'rgba(102,126,234,0.6)',borderRadius:4}},
{{label:'本周(s)',data:sd2,backgroundColor:'rgba(231,76,60,0.7)',borderRadius:4}}
]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{position:'top'}},title:{{display:true,text:'无声会话平均处理时长对比（秒）',font:{{size:14}}}}}},scales:{{y:{{beginAtZero:true,grid:{{color:'#f0f0f0'}}}},x:{{grid:{{display:false}}}}}}}}}});

// Auto-generated analysis text
(function(){{
'''

    # 自动生成分析文字
    agents_js = json.dumps(agents)
    sa1_js = json.dumps(d['sa1'])
    sa2_js = json.dumps(d['sa2'])
    saQ_js = json.dumps(d['saQ'])
    du1_js = json.dumps(d['du1'])
    du2_js = json.dumps(d['du2'])
    duQ_js = json.dumps(d['duQ'])
    ra1_js = json.dumps(d['ra1'])
    ra2_js = json.dumps(d['ra2'])
    raQ_js = json.dumps(d['raQ'])
    st1_js = json.dumps(d['st1'])
    st2_js = json.dumps(d['st2'])
    stQ_js = json.dumps(d['stQ'])
    si1_js = json.dumps(d['si1'])
    si2_js = json.dumps(d['si2'])
    siQ_js = json.dumps(d['siQ'])
    du1_vals = [f'{v:.1f}' for v in d['du1']]
    du2_vals = [f'{v:.1f}' for v in d['du2']]
    ts1_js = json.dumps(d['ts1'])
    ts2_js = json.dumps(d['ts2'])
    sv1_js = json.dumps(d['sv1'])
    sv2_js = json.dumps(d['sv2'])

    # 构建各section的分析文字（简洁版：各指标top变化 + 团队汇总）
    def top_n_html(arr1, arr2, qoq, label, n=3, pct_fmt=True, is_reverse=False, bad_news=False):
        """生成变化最大N人的HTML片段。bad_news=True 时用红色标记恶化项"""
        indexed = list(enumerate(qoq))
        indexed.sort(key=lambda x: x[1], reverse=True)
        parts = []
        for idx, q in indexed[:n]:
            if abs(q) < 0.01:
                continue
            if bad_news:
                cls = 'red' if abs(q) > 0.05 else 'yellow'
            else:
                cls = 'green' if ((q < 0) if is_reverse else (q > 0)) else 'red'
                if cls == 'red' and abs(q) < 0.05:
                    cls = 'yellow'
            old_val = arr1[idx]
            new_val = arr2[idx]
            if pct_fmt:
                old_s, new_s = pct_v(old_val), pct_v(new_val)
            else:
                old_s = f'{old_val:.0f}' if old_val == int(old_val) else f'{old_val:.1f}'
                new_s = f'{new_val:.0f}' if new_val == int(new_val) else f'{new_val:.1f}'
            parts.append(
                f"'<li>'+tag('{cls}','{agents[idx]}')+' {label} <b>{pct(q)}</b>"
                f"（{old_s}→{new_s}）</li>'"
            )
        return '+'.join(parts) if parts else "''"

    # 各区块分析JS
    session_top = top_n_html(d['sa1'], d['sa2'], d['saQ'], '会话量上升', 4, False)
    session_bot = top_n_html(d['sa1'], d['sa2'], d['saQ'], '会话量下降', 2, False, True, True)

    dur_best = top_n_html(d['du1'], d['du2'], d['duQ'], '处理时长改善', 3, False, True)
    dur_worst = top_n_html(d['du1'], d['du2'], d['duQ'], '处理时长上升', 2, False, False, True)

    rating_top = top_n_html(d['ra1'], d['ra2'], d['raQ'], '评价率上升', 2)
    rating_bot = top_n_html(d['ra1'], d['ra2'], d['raQ'], '评价率下降', 2, True, True, True)
    st_worst_html = top_n_html(d['st1'], d['st2'], d['stQ'], '差评率暴涨', 3, True, False, True)
    st_best_html = top_n_html(d['st1'], d['st2'], d['stQ'], '差评率改善', 2, True, True)

    si_worst_html = top_n_html(d['si1'], d['si2'], d['siQ'], '无声占比上升', 3, True, False, True)
    si_best_html = top_n_html(d['si1'], d['si2'], d['siQ'], '无声占比改善', 1, True, True)

    # 计算团队汇总
    w2_total = sum(d['ts2'])
    w1_total = sum(d['ts1'])
    total_chg = (w2_total - w1_total) / w1_total

    # 生成完整的分析JS
    js_analysis = f'''
const agentsArr={agents_js};

// 会话量分析
document.getElementById('sessionAnalysis').innerHTML='<h4>&#x1F4CA; 分析要点</h4><ul>'+
{session_top}+
{session_bot}+
'<li><b>团队总会话量：</b>本周 {w2_total:,} vs 上周 {w1_total:,}，整体{"上升" if total_chg>=0 else "下降"} <b>{pct(total_chg, 1)}</b></li></ul>';

// 处理时长分析
document.getElementById('durationAnalysis').innerHTML='<h4>&#x23F1;&#xFE0F; 分析要点</h4><ul>'+
{dur_best}+
{dur_worst}+
'<li><b>团队均值：</b>本周 {sum(d["du2"])/len(agents):.2f}s vs 上周 {sum(d["du1"])/len(agents):.2f}s，{"基本持平" if abs(du_chg:= (sum(d["du2"])/len(agents)-sum(d["du1"])/len(agents))/(sum(d["du1"])/len(agents))) < 0.01 else ("改善" if du_chg<0 else "上升")} <b>{pct((sum(d["du2"])/len(agents)-sum(d["du1"])/len(agents))/(sum(d["du1"])/len(agents)), 1)}</b></li></ul>';

// 评价率 & 差评率分析
document.getElementById('ratingAnalysis').innerHTML='<h4>&#x2B50; 分析要点</h4><ul>'+
'<li>'+tag('yellow','评价率变化')+' — 团队均值从 {pct_v(sum(d["ra1"])/len(agents))} → {pct_v(sum(d["ra2"])/len(agents))}（{pct((sum(d["ra2"])-sum(d["ra1"]))/sum(d["ra1"]))}）</li>'+
{rating_top}+
{rating_bot}+
'<li><b>差评率变化（越低越好）：</b></li>'+
{st_worst_html}+
{st_best_html}+
'</ul>';

// 无声分析
document.getElementById('silentAnalysis').innerHTML='<h4>&#x1F507; 分析要点</h4><ul>'+
'<li>'+tag('red','&#x26A0;&#xFE0F; 无声占比')+' — 团队均值从 {pct_v(sum(d["si1"])/len(agents))} → {pct_v(sum(d["si2"])/len(agents))}，<b>{pct((sum(d["si2"])-sum(d["si1"]))/sum(d["si1"]))}</b></li>'+
{si_worst_html}+
{si_best_html}+
'</ul>';

// 综合排名
(function(){{
function normalize(arr,asc){{const mn=Math.min(...arr),mx=Math.max(...arr);if(mx===mn)return arr.map(()=>50);return arr.map(v=>asc?((v-mn)/(mx-mn)*100):((mx-v)/(mx-mn)*100));}}
const ss=normalize(sa2,true),sd_n=normalize(du2,false),sr=normalize(ra2,true),sst_n=normalize(st2,false),ssi_n=normalize(si2,false);
const scores=agents.map((_,i)=>(ss[i]*0.2+sd_n[i]*0.2+sr[i]*0.2+sst_n[i]*0.2+ssi_n[i]*0.2));
new Chart(document.getElementById('chartRanking'),{{type:'bar',data:{{labels:agents,datasets:[{{label:'综合得分',data:scores.map(v=>v.toFixed(1)),backgroundColor:colors.map(c=>c+'cc'),borderColor:colors,borderWidth:2,borderRadius:6}}]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},title:{{display:true,text:'本周综合表现排名（越高越好）',font:{{size:14}}}}}},scales:{{x:{{beginAtZero:true,max:100,grid:{{color:'#f0f0f0'}}}},y:{{grid:{{display:false}},reverse:true}}}}}}}});
const ranked=agents.map((a,i)=>({{n:a,s:scores[i]}})).sort((a,b)=>b.s-a.s);
document.getElementById('rankingAnalysis').innerHTML='<h4>&#x1F3C6; 综合排名（五维等权评分）</h4><ul>'+
ranked.map((r,i)=>{{const m=i===0?'&#x1F947;':i===1?'&#x1F948;':i===2?'&#x1F949;':(i+1)+'.';return '<li><b>'+m+' '+r.n+'</b> — 综合得分 <b>'+r.s.toFixed(1)+'</b> 分</li>';}}).join('')+'</ul>';
}})();
}})();
</script>'''

    html += js_data + js_analysis + '\n</body></html>'
    return html


# ============================================================
# 主入口
# ============================================================

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    excel_path = sys.argv[1]

    # 输出路径：默认同目录同名 .html
    if len(sys.argv) >= 3:
        out_path = sys.argv[2]
    else:
        excel_name = Path(excel_path).stem
        out_path = str(Path(excel_path).parent / f'{excel_name}.html')

    d = read_excel(excel_path)
    part1, part2, part3, ranked = generate_insight_text(d)
    html = build_html(d, part1, part2, part3, ranked)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f'[OK] HTML 周报已生成: {out_path}')
    print(f'     周期: {d["week1"]} → {d["week2"]}')
    print(f'     客服: {len(d["agents"])} 人')
    print(f'     总会话量: {sum(d["ts2"]):,}')


if __name__ == '__main__':
    main()
