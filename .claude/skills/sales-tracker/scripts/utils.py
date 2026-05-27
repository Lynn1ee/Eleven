"""工具函数：消息解析、数据提取、合计计算"""

import re
from datetime import datetime


def extract_table1_data(text: str) -> dict:
    """从自然语言中提取表1（平台会话量）数据。

    返回: {"拼多多火车票": int, "拼多多机票": int, "千牛": int, "抖音": int, "备注": str}
    未提及的平台默认为 0。
    """
    result = {"拼多多火车票": 0, "拼多多机票": 0, "千牛": 0, "抖音": 0, "备注": ""}

    patterns = {
        "拼多多火车票": r"(?:拼多多火车票|火车票拼多多)\s*(\d+)",
        "拼多多机票": r"(?:拼多多机票|机票拼多多)\s*(\d+)",
        "千牛": r"千牛\s*(\d+)",
        "抖音": r"抖音\s*(\d+)",
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            result[key] = int(match.group(1))

    # 提取备注（平台数字后面的文字说明）
    remark_match = re.search(r"备注[：:]\s*(.+)", text)
    if remark_match:
        result["备注"] = remark_match.group(1).strip()

    return result


def extract_table2_data(text: str) -> list:
    """从文本中提取订单号列表。

    支持逗号、空格、换行、顿号等分隔符。
    返回订单号字符串列表。
    """
    # 先用常见分隔符拆分
    parts = re.split(r"[,，、\s\n]+", text.strip())
    # 过滤：空字符串、@提及、纯字母（非订单号）
    orders = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if p.startswith("@"):
            continue
        orders.append(p)
    return orders


def classify_message(text: str) -> str:
    """判断消息类型。

    返回:
        "table1" — 包含平台关键词+数字
        "table2" — 包含订单号格式
        "query" — 查询统计类
        "modify" — 修改/删除类
        "report" — 报表类
        "remind" — 提醒类
        "unknown" — 无法识别
    """
    text_lower = text.lower()

    # 报表类优先判断
    if any(kw in text for kw in ["日报", "周报", "月报", "报表"]):
        return "report"

    # 提醒类
    if any(kw in text for kw in ["提醒", "催一下", "催一下数据", "上报"]):
        return "remind"

    # 修改删除类
    if any(kw in text for kw in ["修改", "改一下", "删除", "删掉", "去掉"]):
        return "modify"

    # 查询类
    if any(kw in text for kw in ["查询", "统计", "汇总", "多少", "排名", "帮我查", "帮我看看"]):
        return "query"

    # 表1：平台+数字（支持拼多多火车票/火车票拼多多 两种顺序）
    if re.search(r"((拼多多(火车票|机票)|(火车票|机票)\s*拼多多)|千牛|抖音)\s*\d+", text):
        return "table1"

    # 表2：看起来像订单号（字母+数字组合，或纯数字编号）
    # 支持格式: HT001, HT-001, 20240527001 等
    has_order_pattern = bool(
        re.search(r"[A-Za-z]{2,}[-\s]?\d{3,}", text)  # HT001, HT-001
        or re.search(r"\d{6,}", text)  # 纯数字6位以上
    )
    if has_order_pattern:
        return "table2"

    return "unknown"


def calc_total(pdd_train: int = 0, pdd_flight: int = 0, qianniu: int = 0, douyin: int = 0) -> int:
    """计算四平台合计。"""
    return pdd_train + pdd_flight + qianniu + douyin


def get_current_month() -> str:
    """返回当前月份，格式 YYYY-MM。"""
    return datetime.now().strftime("%Y-%m")


def get_base_name(month: str = None) -> str:
    """生成当月 Base 名称。"""
    if month is None:
        month = get_current_month()
    return f"业务数据登记_{month}"


def calc_change_rate(current: float, previous: float) -> str:
    """计算环比变化率，返回格式化字符串如 '↑12%'。"""
    if previous == 0:
        return "新增" if current > 0 else "—"
    rate = (current - previous) / previous * 100
    if rate > 5:
        return f"↑{rate:.0f}%"
    elif rate < -5:
        return f"↓{abs(rate):.0f}%"
    else:
        return "→持平"
