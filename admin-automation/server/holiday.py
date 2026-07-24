"""节假日工具：从 API 获取国务院法定节假日 + 调休，计算月度标准工时"""
import json
import calendar
import urllib.request
from datetime import date
from urllib.parse import urlparse, parse_qs
from .db import get_db

API = "https://timor.tech/api/holiday/year"


def _fetch_year(year):
    """获取指定年份的节假日数据，缓存到 app_data"""
    db = get_db()
    row = db.execute(
        "SELECT data_value FROM app_data WHERE data_key = ? AND user_id IS NULL",
        (f"holiday_{year}",)).fetchone()
    if row:
        return json.loads(row["data_value"])

    try:
        req = urllib.request.Request(
            f"{API}/{year}",
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        )
        resp = urllib.request.urlopen(req, timeout=10)
        data = json.loads(resp.read())
        if data.get("code") == 0 and data.get("holiday"):
            db.execute(
                "INSERT OR REPLACE INTO app_data (user_id, data_key, data_value) VALUES (NULL, ?, ?)",
                (f"holiday_{year}", json.dumps(data["holiday"], ensure_ascii=False)))
            db.commit()
            return data["holiday"]
    except Exception:
        pass
    return None


def calc_work_hours(year, month):
    """计算指定年月的标准工时（工作日 × 8）"""
    holidays = _fetch_year(year) or {}
    days_in_month = calendar.monthrange(year, month)[1]
    work_days = 0

    for d in range(1, days_in_month + 1):
        key = f"{month:02d}-{d:02d}"
        dt = date(year, month, d)
        weekday = dt.weekday()  # 0=周一, 6=周日
        is_weekend = weekday >= 5

        if key in holidays:
            info = holidays[key]
            if info.get("holiday"):
                continue  # 法定假日，休息
            else:
                work_days += 1  # 调休补班，工作日
        else:
            if not is_weekend:
                work_days += 1

    return work_days * 8


def handle_work_hours(handler):
    """GET /api/work-hours?year=2026&month=6 返回标准工时"""
    parsed = urlparse(handler.path)
    params = parse_qs(parsed.query)
    y = int(params.get("year", [date.today().year])[0])
    m = int(params.get("month", [date.today().month])[0])
    hours = calc_work_hours(y, m)
    return 200, {"success": True, "year": y, "month": m, "work_hours": hours}
