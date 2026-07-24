# 在线客服排班系统

一个 **Python 自动化排班引擎**，用历史进线量数据自动生成月度排班表。

## 快速开始

```bash
pip install openpyxl
python generate_schedule.py
```

## 适配新月份

打开 `generate_schedule.py`，改顶部 5 个参数：

```python
YEAR = 2026          # 年份
MONTH = 8            # 目标月份
NUM_DAYS = 31        # 天数
START_WEEKDAY = 6    # 当月1日是周几（0=周一, 6=周日）
DAILY_TARGETS = [...] # 每日目标人数（按公式重算）
```

## 文档

| 文件 | 说明 |
|------|------|
| `交接说明.md` | 完整使用手册（必读） |
| `7月排班逻辑说明.md` | 算法原理、公式推导 |
| `6月份原始班表.xlsx` | 6月原始数据参考 |
| `旧版存档/` | 历史班表 |

## 依赖

- Python 3.9+
- openpyxl
