# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

飞书客服业务任务量登记机器人。通过 WebSocket 长连接接收飞书群聊消息，自动识别平台会话量（拼多多火车票/机票、千牛、抖音）或订单号，写入飞书多维表格（Bitable），支持跨月查询、删除、定时提醒，以及含内嵌图表的日报/周报/月报（飞书互动卡片消息）。

## 启动方式

```bash
cd sales-tracker-bot
cp .env.example .env   # 填入飞书应用凭证
pip install -r requirements.txt
python server.py
```

Python 3.12+, 依赖: `lark-oapi>=1.6`, `requests`, `python-dotenv`, `matplotlib`.

## 架构

**`sales-tracker-bot/server.py`**（~1400 行）— 整个 bot 的单文件实现：

- `FeishuClient` — 飞书 API 封装：token 管理、消息发送/回复（文本、图片、交互卡片）、Bitable CRUD、图片上传。直接用 `requests`，不走 SDK 的 HTTP 客户端。
- `MessageHandler` — 业务引擎。入口 `handle()` 调用 `classify_message()` 判断消息类型后分发：登记 → 表1 或 表2、查询 → 个人/排名、报表 → 日报/周报/月报、删除、提醒。
- SDK 的 `WSClient` + `EventDispatcherHandler` 维持 WebSocket 长连接，回调 `on_message_receive` 做文本提取、姓名解析、去重检查后交给 `MessageHandler.handle()`。

**`.claude/skills/sales-tracker/scripts/utils.py`** — 纯函数，供 bot 和 Claude Code skill 共用：
- `classify_message(text)` → `"table1" | "table2" | "query" | "report" | "remind" | "modify" | "unknown"`
- `extract_table1_data(text)` → `{"拼多多火车票": int, ...}` — 正则提取，支持品牌缩写（拼多多/多多/PDD）和产品别名（火车/火车票、机票/飞机）
- `extract_table2_data(text)` → `list[str]` — 分隔符拆分订单号
- `calc_change_rate(current, previous)` → `"↑12%" | "↓8%" | "→持平"`

**`.claude/skills/sales-tracker/scripts/chart_gen.py`** — matplotlib 图表生成，输出 PNG 到 `output/`：
- `trend_line()` — 折线图（标注最高/最低点）
- `pie_chart()` — 饼图
- `bar_chart()` — 柱状图（前三名金银铜高亮）
- `table_image()` — 数据表格图
- `dual_line()` — 双 Y 轴折线图

## 关键设计

- **UTF-8 编码**：`add_records()` 用 `data=json.dumps(body, ensure_ascii=False).encode("utf-8")` 防止中文姓名在 HTTP 传输中乱码。
- **日期处理**：Bitable 日期字段存毫秒时间戳。`_parse_date()` 从消息文本提取日期，支持 `M-D`、`M/D`、`M月D日`、`M月D号`、`YYYY年M月D日`、混合格式如 `5-29日`。`_ts_to_date()` 转回 `YYYY-MM-DD` 字符串比对。
- **消息去重**：已处理消息 ID 持久化到 `.processed_ids.json`（上限 5000），重启不丢失。
- **跨月聚合**：`_load_all_states()` 扫描所有 `.state_YYYY-MM.json` 文件实现跨月查询和报表。
- **姓名解析**：三级优先级 — `name_map.json`（热加载）→ 飞书通讯录 API → `外部用户_xxxx` / `用户_xxxx`。
- **卡片降级**：报表优先发交互卡片（图表内嵌），`reply_card()` 失败则降级为文本 + 单独图片。
- **月度归档**：每月独立 Base（两张表），状态文件 `.state_YYYY-MM.json` 跟踪 ID。Schema 见 `.claude/skills/sales-tracker/references/table-schemas.md`。

## 修改指引

- **消息解析 / 模糊匹配**：`utils.py` → `extract_table1_data()` 和 `classify_message()` 的正则
- **日期格式**：`server.py` → `MessageHandler._parse_date()`
- **报表内容 / 图表**：逻辑在 `server.py` 的 `_daily_report()` / `_weekly_report()` / `_monthly_report()`，样式在 `chart_gen.py`
- **API 交互**：`server.py` → `FeishuClient`
- **用户名映射**：`sales-tracker-bot/name_map.json`（`"open_id": "姓名"`）
- **定时提醒时间**：`server.py` → `run_scheduler()` → `reminder_times`
