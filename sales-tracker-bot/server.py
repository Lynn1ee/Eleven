"""飞书 Bot 长连接模式 — 使用 lark-oapi SDK，无需公网 URL、无需 ngrok"""

import json
import os
import sys
import time
import threading
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from collections import defaultdict

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import re
import requests

# lark-oapi SDK
from lark_oapi.ws import Client as WSClient
from lark_oapi.event.dispatcher_handler import EventDispatcherHandler

# utils
SKILL_DIR = Path(os.getenv("SKILL_DIR", Path(__file__).parent.parent /
                           ".claude" / "skills" / "sales-tracker"))
sys.path.insert(0, str(SKILL_DIR / "scripts"))
from utils import extract_table1_data, extract_table2_data, classify_message

# ══════════════════════════════════════════════════════
# 配置
# ══════════════════════════════════════════════════════

FEISHU_APP_ID = os.getenv("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")

# ══════════════════════════════════════════════════════
# 飞书 API 客户端（发送消息、操作表格）
# ══════════════════════════════════════════════════════

class FeishuClient:
    def __init__(self, app_id: str, app_secret: str):
        self.app_id = app_id
        self.app_secret = app_secret
        self._token = None
        self._token_expire = 0

    def _get_tenant_token(self) -> str:
        now = time.time()
        if self._token and now < self._token_expire - 60:
            return self._token
        resp = requests.post(
            "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
            json={"app_id": self.app_id, "app_secret": self.app_secret},
            timeout=10,
        )
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"获取 token 失败: {data}")
        self._token = data["tenant_access_token"]
        self._token_expire = now + data.get("expire", 7200)
        return self._token

    def _load_name_map(self) -> dict:
        f = Path(__file__).parent / "name_map.json"
        sys.stderr.write(f"[DEBUG] _load_name_map path: {f.resolve()}\n")
        sys.stderr.write(f"[DEBUG] _load_name_map exists: {f.exists()}\n")
        sys.stderr.flush()
        if f.exists():
            try:
                raw_bytes = f.read_bytes()
                sys.stderr.write(f"[DEBUG] _load_name_map raw hex: {raw_bytes.hex()}\n")
                sys.stderr.flush()
                text = f.read_text(encoding="utf-8")
                sys.stderr.write(f"[DEBUG] _load_name_map text repr: {repr(text[:200])}\n")
                sys.stderr.flush()
                result = json.loads(text)
                for k, v in result.items():
                    sys.stderr.write(f"[DEBUG] _load_name_map key={k} value={repr(v)} v_hex={v.encode('utf-8').hex()}\n")
                sys.stderr.flush()
                return result
            except (json.JSONDecodeError, KeyError) as e:
                sys.stderr.write(f"[DEBUG] _load_name_map error: {e}\n")
                sys.stderr.flush()
                pass
        return {}

    def get_user_name(self, open_id: str) -> str:
        # 先查自定义映射
        name_map = self._load_name_map()
        if open_id in name_map:
            resolved = name_map[open_id]
            sys.stderr.write(f"[DEBUG] name_map hit: {open_id} -> {repr(resolved)} (UTF-8: {resolved.encode('utf-8').hex()})\n")
            sys.stderr.flush()
            return resolved

        resp = requests.get(
            f"https://open.feishu.cn/open-apis/contact/v3/users/{open_id}",
            headers={"Authorization": f"Bearer {self._get_tenant_token()}"},
            timeout=10,
        )
        data = resp.json()
        code = data.get("code", -1)
        if code == 0:
            user = data.get("data", {}).get("user", {})
            name = user.get("name") or user.get("en_name") or user.get("nickname")
            if name:
                return name

        # 41050 = 跨租户外部用户，Contact API 无权查姓名
        if code == 41050:
            return f"外部用户_{open_id[-4:]}"

        return f"用户_{open_id[-4:]}"

    def send_message(self, chat_id: str, content: str):
        body = {
            "receive_id": chat_id,
            "msg_type": "text",
            "content": json.dumps({"text": content}),
        }
        requests.post(
            "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            json=body, timeout=10,
        )

    def reply_message(self, msg_id: str, content: str):
        """引用回复某条消息。"""
        body = {
            "msg_type": "text",
            "content": json.dumps({"text": content}),
        }
        resp = requests.post(
            f"https://open.feishu.cn/open-apis/im/v1/messages/{msg_id}/reply",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            json=body, timeout=10,
        )
        return resp.json()

    def add_records(self, app_token: str, table_id: str, records: list):
        # 手动序列化以捕获实际字节并确保编码正确
        body = {"records": records}
        body_json = json.dumps(body, ensure_ascii=False)
        body_bytes = body_json.encode("utf-8")

        sys.stderr.write(f"[DEBUG] add_records body JSON repr: {repr(body_json[:500])}\n")
        sys.stderr.write(f"[DEBUG] add_records body hex (first 200): {body_bytes[:200].hex()}\n")
        sys.stderr.flush()

        resp = requests.post(
            f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_create",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            data=body_bytes, timeout=15,
        )
        result = resp.json()
        sys.stderr.write(f"[DEBUG] add_records response: code={result.get('code')} msg={result.get('msg','')}\n")
        sys.stderr.flush()
        return result

    def search_records(self, app_token: str, table_id: str) -> list:
        all_records = []
        page_token = ""
        while True:
            params = {"page_size": "100"}
            if page_token:
                params["page_token"] = page_token
            resp = requests.get(
                f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records",
                headers={"Authorization": f"Bearer {self._get_tenant_token()}"},
                params=params, timeout=15,
            )
            data = resp.json()
            if data.get("code") != 0:
                break
            items = data.get("data", {}).get("items", [])
            all_records.extend(items)
            if not data.get("data", {}).get("has_more"):
                break
            page_token = data["data"].get("page_token", "")
        return all_records


# ══════════════════════════════════════════════════════
# 消息处理引擎
# ══════════════════════════════════════════════════════

class MessageHandler:
    def __init__(self, feishu_client):
        self.fs = feishu_client
        self._state = self._load_state()
        self._chat_ids: set = set()
        self._remind_lock = threading.Lock()

    def _load_state(self) -> dict:
        month = datetime.now().strftime("%Y-%m")
        state_file = Path(__file__).parent / f".state_{month}.json"
        if state_file.exists():
            return json.loads(state_file.read_text())
        return {"app_token": "", "table1_id": "", "table2_id": ""}

    # ── 消息去重 ──────────────────────────────

    @property
    def _dedup_file(self) -> Path:
        return Path(__file__).parent / ".processed_ids.json"

    def _load_processed_ids(self) -> set:
        f = self._dedup_file
        if f.exists():
            try:
                data = json.loads(f.read_text())
                return set(data.get("ids", []))
            except (json.JSONDecodeError, KeyError):
                pass
        return set()

    def _save_processed_ids(self, ids: set):
        # 只保留最近 5000 条
        data = {"ids": list(ids)[-5000:]}
        self._dedup_file.write_text(json.dumps(data))

    def is_duplicate(self, msg_id: str) -> bool:
        if not msg_id:
            return False
        processed = self._load_processed_ids()
        return msg_id in processed

    def mark_processed(self, msg_id: str):
        if not msg_id:
            return
        processed = self._load_processed_ids()
        processed.add(msg_id)
        self._save_processed_ids(processed)

    # ── 登记处理 ──────────────────────────────

    def handle_register(self, msg_text: str, user_name: str, chat_id: str) -> str:
        msg_type = classify_message(msg_text)

        if msg_type == "table1":
            data = extract_table1_data(msg_text)
            total = data["拼多多火车票"] + data["拼多多机票"] + data["千牛"] + data["抖音"]
            if total == 0:
                return "未识别到有效数据。格式：拼多多火车票5 拼多多机票3 千牛1"

            record = {"fields": {
                "日期": int(datetime.now().timestamp() * 1000),
                "姓名": user_name,
                "拼多多火车票": data["拼多多火车票"],
                "拼多多机票": data["拼多多机票"],
                "千牛": data["千牛"],
                "抖音": data["抖音"],
                "合计": total,
                "备注": data.get("备注", ""),
            }}
            result = self.fs.add_records(self._state["app_token"],
                                         self._state["table1_id"], [record])
            sys.stderr.write(f"[DEBUG] table1 add_records: {result.get('code')} {result.get('msg','')}\n")
            sys.stderr.flush()
            if result.get("code") == 0:
                return (f"✅ 已登记 IM会话量\n"
                        f"拼多多火车票={data['拼多多火车票']} 拼多多机票={data['拼多多机票']} "
                        f"千牛={data['千牛']} 抖音={data['抖音']} 合计={total}")
            return f"❌ 登记失败: {result.get('msg', '')}"

        elif msg_type == "table2":
            orders = extract_table2_data(msg_text)
            if not orders:
                return "未识别到订单号。格式：HT001, HT002, HT003"

            records = []
            for order_no in orders:
                records.append({"fields": {
                    "日期": int(datetime.now().timestamp() * 1000),
                    "姓名": user_name,
                    "订单号": order_no,
                    "备注": "",
                }})
            result = self.fs.add_records(self._state["app_token"],
                                         self._state["table2_id"], records)
            sys.stderr.write(f"[DEBUG] table2 add_records: {result.get('code')} {result.get('msg','')}\n")
            sys.stderr.flush()
            if result.get("code") == 0:
                return f"✅ 已登记 {len(orders)} 条火车票订单: {', '.join(orders[:10])}"
            return f"❌ 登记失败: {result.get('msg', '')}"

        return None

    # ── 查询处理 ──────────────────────────────

    def _load_all_states(self) -> list:
        """查找所有历史 state 文件，返回 [(月份, app_token, table1_id, table2_id), ...]。
        按月份倒序排列。
        """
        state_files = sorted(
            Path(__file__).parent.glob(".state_????-??.json"),
            reverse=True,
        )
        result = []
        for f in state_files:
            try:
                data = json.loads(f.read_text())
                if data.get("app_token") and data.get("table1_id"):
                    month = f.stem.replace(".state_", "")
                    result.append((month, data["app_token"], data["table1_id"], data.get("table2_id", "")))
            except (json.JSONDecodeError, KeyError):
                pass
        return result

    def handle_query(self, msg_text: str, user_name: str, chat_id: str) -> str:
        if "排名" in msg_text:
            return self._query_ranking()
        if "平台" in msg_text or "占比" in msg_text:
            return self._query_platform_share()
        return self._query_user(user_name)

    def _query_user(self, user_name: str) -> str:
        all_states = self._load_all_states()
        grand_ttl = grand_train = grand_flight = grand_qn = grand_dy = grand_days = 0
        grand_orders = 0
        monthly_lines = []

        for month, app_token, t1_id, t2_id in all_states:
            records1 = self.fs.search_records(app_token, t1_id)
            m_train = m_flight = m_qn = m_dy = m_days = 0
            for r in records1:
                f = r.get("fields", {})
                if _get_field(f, "姓名") == user_name:
                    m_train += int(_get_field(f, "拼多多火车票") or 0)
                    m_flight += int(_get_field(f, "拼多多机票") or 0)
                    m_qn += int(_get_field(f, "千牛") or 0)
                    m_dy += int(_get_field(f, "抖音") or 0)
                    m_days += 1

            m_total = m_train + m_flight + m_qn + m_dy
            if m_total > 0 or m_days > 0:
                monthly_lines.append(
                    f"  {month}: 火车票{m_train} 机票{m_flight} 千牛{m_qn} 抖音{m_dy} 合计{m_total} ({m_days}天)"
                )
                grand_train += m_train
                grand_flight += m_flight
                grand_qn += m_qn
                grand_dy += m_dy
                grand_days += m_days

            if t2_id:
                records2 = self.fs.search_records(app_token, t2_id)
                m_orders = sum(1 for r in records2
                               if _get_field(r.get("fields", {}), "姓名") == user_name)
                grand_orders += m_orders

        grand_ttl = grand_train + grand_flight + grand_qn + grand_dy
        lines = [f"📊 {user_name} 全部历史汇总"]
        if monthly_lines:
            lines.append("各月明细:")
            lines.extend(monthly_lines)
            lines.append("")
        lines.append(
            f"总计({grand_days}天): 火车票{grand_train} 机票{grand_flight} "
            f"千牛{grand_qn} 抖音{grand_dy} 合计{grand_ttl}\n"
            f"订单: {grand_orders}条"
        )
        return "\n".join(lines)

    def _query_ranking(self) -> str:
        all_states = self._load_all_states()
        user_totals = defaultdict(int)
        for month, app_token, t1_id, _ in all_states:
            records = self.fs.search_records(app_token, t1_id)
            for r in records:
                name = _get_field(r.get("fields", {}), "姓名")
                total = int(_get_field(r.get("fields", {}), "合计") or 0)
                user_totals[name] += total

        sorted_users = sorted(user_totals.items(), key=lambda x: x[1], reverse=True)
        lines = [f"🏆 全部历史排名（{len(all_states)} 个月）"]
        medals = ["🥇", "🥈", "🥉"]
        for i, (name, total) in enumerate(sorted_users[:10]):
            prefix = medals[i] if i < 3 else f"{i+1}."
            lines.append(f"  {prefix} {name}: {total}")
        if not sorted_users:
            lines.append("  暂无数据")
        return "\n".join(lines)

    def _query_platform_share(self) -> str:
        all_states = self._load_all_states()
        pdd_train, pdd_flight, qn, dy = 0, 0, 0, 0
        for month, app_token, t1_id, _ in all_states:
            records = self.fs.search_records(app_token, t1_id)
            for r in records:
                f = r.get("fields", {})
                pdd_train += int(_get_field(f, "拼多多火车票") or 0)
                pdd_flight += int(_get_field(f, "拼多多机票") or 0)
                qn += int(_get_field(f, "千牛") or 0)
                dy += int(_get_field(f, "抖音") or 0)
        total = pdd_train + pdd_flight + qn + dy
        if total == 0:
            return "暂无数据。"
        months = len(all_states)
        return (f"📊 平台占比（{months} 个月汇总）\n"
                f"火车票 {pdd_train} ({pdd_train/total*100:.1f}%)\n"
                f"机票 {pdd_flight} ({pdd_flight/total*100:.1f}%)\n"
                f"千牛 {qn} ({qn/total*100:.1f}%)\n"
                f"抖音 {dy} ({dy/total*100:.1f}%)\n"
                f"合计 {total}")

    # ── 报表 ──────────────────────────────────

    def handle_report(self, msg_text: str, user_name: str, chat_id: str) -> str:
        if "日报" in msg_text:
            return self._daily_report()
        elif "周报" in msg_text:
            return self._weekly_report()
        elif "月报" in msg_text:
            return self._monthly_report()
        else:
            return self._daily_report()

    def _daily_report(self) -> str:
        yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        records1 = self.fs.search_records(self._state["app_token"], self._state["table1_id"])
        records2 = self.fs.search_records(self._state["app_token"], self._state["table2_id"])

        user_data = defaultdict(lambda: {"total": 0, "pdd_train": 0, "pdd_flight": 0, "qn": 0, "dy": 0})
        for r in records1:
            f = r.get("fields", {})
            date_val = _get_field(f, "日期")
            if isinstance(date_val, str) and date_val[:10] == yesterday:
                name = _get_field(f, "姓名")
                user_data[name]["total"] += int(_get_field(f, "合计") or 0)
                user_data[name]["pdd_train"] += int(_get_field(f, "拼多多火车票") or 0)
                user_data[name]["pdd_flight"] += int(_get_field(f, "拼多多机票") or 0)
                user_data[name]["qn"] += int(_get_field(f, "千牛") or 0)
                user_data[name]["dy"] += int(_get_field(f, "抖音") or 0)

        order_count = sum(1 for r in records2
                          if _get_field(r.get("fields", {}), "日期")[:10] == yesterday)

        lines = [f"📊 日报 — {yesterday}"]
        grand_total = 0
        for name, data in user_data.items():
            lines.append(f"{name}: 火车票{data['pdd_train']} 机票{data['pdd_flight']} "
                         f"千牛{data['qn']} 抖音{data['dy']} 合计{data['total']}")
            grand_total += data["total"]
        lines.append(f"IM总会话量: {grand_total} | 订单数: {order_count}")
        if not user_data:
            lines.append("暂无昨日数据")
        return "\n".join(lines)

    def _weekly_report(self) -> str:
        today = datetime.now()
        monday = today - timedelta(days=today.weekday())
        week_start = (monday - timedelta(days=7)).strftime("%Y-%m-%d")
        week_end = (monday - timedelta(days=1)).strftime("%Y-%m-%d")

        records1 = self.fs.search_records(self._state["app_token"], self._state["table1_id"])
        user_totals = defaultdict(int)
        pdd_train, pdd_flight, qn, dy = 0, 0, 0, 0

        for r in records1:
            f = r.get("fields", {})
            date_str = _get_field(f, "日期")[:10]
            if week_start <= date_str <= week_end:
                name = _get_field(f, "姓名")
                t = int(_get_field(f, "合计") or 0)
                user_totals[name] += t
                pdd_train += int(_get_field(f, "拼多多火车票") or 0)
                pdd_flight += int(_get_field(f, "拼多多机票") or 0)
                qn += int(_get_field(f, "千牛") or 0)
                dy += int(_get_field(f, "抖音") or 0)

        records2 = self.fs.search_records(self._state["app_token"], self._state["table2_id"])
        order_count = sum(1 for r in records2
                          if week_start <= _get_field(r.get("fields", {}), "日期")[:10] <= week_end)

        total = pdd_train + pdd_flight + qn + dy
        lines = [f"📈 周报 — {week_start} ~ {week_end}",
                 f"IM总会话量: {total} | 订单数: {order_count}",
                 f"平台: 火车票{pdd_train} 机票{pdd_flight} 千牛{qn} 抖音{dy}"]
        sorted_users = sorted(user_totals.items(), key=lambda x: x[1], reverse=True)
        for i, (name, t) in enumerate(sorted_users[:5]):
            prefix = "🥇" if i == 0 else "🥈" if i == 1 else "🥉" if i == 2 else f"  {i+1}."
            lines.append(f"{prefix} {name}: {t}")
        return "\n".join(lines)

    def _monthly_report(self) -> str:
        last_month_dt = datetime.now().replace(day=1) - timedelta(days=1)
        last_month = last_month_dt.strftime("%Y-%m")
        records1 = self.fs.search_records(self._state["app_token"], self._state["table1_id"])
        user_totals = defaultdict(int)
        pdd_train, pdd_flight, qn, dy = 0, 0, 0, 0

        for r in records1:
            f = r.get("fields", {})
            date_str = _get_field(f, "日期")[:7]
            if date_str == last_month:
                name = _get_field(f, "姓名")
                t = int(_get_field(f, "合计") or 0)
                user_totals[name] += t
                pdd_train += int(_get_field(f, "拼多多火车票") or 0)
                pdd_flight += int(_get_field(f, "拼多多机票") or 0)
                qn += int(_get_field(f, "千牛") or 0)
                dy += int(_get_field(f, "抖音") or 0)

        records2 = self.fs.search_records(self._state["app_token"], self._state["table2_id"])
        order_count = sum(1 for r in records2
                          if _get_field(r.get("fields", {}), "日期")[:7] == last_month)

        total = pdd_train + pdd_flight + qn + dy
        lines = [f"📋 月报 — {last_month}",
                 f"IM总会话量: {total} | 订单数: {order_count}",
                 f"平台: 火车票{pdd_train} 机票{pdd_flight} 千牛{qn} 抖音{dy}"]
        sorted_users = sorted(user_totals.items(), key=lambda x: x[1], reverse=True)
        for i, (name, t) in enumerate(sorted_users[:10]):
            prefix = "🥇" if i == 0 else "🥈" if i == 1 else "🥉" if i == 2 else f"  {i+1}."
            lines.append(f"{prefix} {name}: {t}")
        return "\n".join(lines)

    # ── 提醒 ──────────────────────────────────

    def handle_remind(self, msg_text: str, user_name: str, chat_id: str) -> str:
        today = datetime.now().strftime("%Y-%m-%d")
        records1 = self.fs.search_records(self._state["app_token"], self._state["table1_id"])
        reported = set()
        for r in records1:
            f = r.get("fields", {})
            if _get_field(f, "日期")[:10] == today:
                reported.add(_get_field(f, "姓名"))
        if reported:
            return f"📢 今日已上报: {', '.join(sorted(reported))}\n还没上报的同事请尽快提交！"
        return "📢 今日暂无上报，请大家及时提交数据！"

    # ── 定时提醒 ──────────────────────────────

    def scheduled_remind(self):
        """定时任务：向所有已知群聊发送提醒。"""
        msg = self.handle_remind("提醒", "", "")
        chat_ids = list(self._chat_ids)
        if not chat_ids:
            sys.stderr.write("[SCHEDULE] 无已知群聊，跳过提醒\n")
            sys.stderr.flush()
            return
        for cid in chat_ids:
            try:
                self.fs.send_message(cid, msg)
                sys.stderr.write(f"[SCHEDULE] 已发送提醒到 {cid}\n")
            except Exception:
                traceback.print_exc()
        sys.stderr.flush()

    # ── 入口 ──────────────────────────────────

    def handle(self, msg_text: str, user_name: str, chat_id: str) -> str:
        sys.stderr.write(f"[DEBUG] handle msg: {msg_text} from {user_name}\n")
        sys.stderr.flush()

        if chat_id:
            self._chat_ids.add(chat_id)

        msg_type = classify_message(msg_text)

        if msg_type == "table1" or msg_type == "table2":
            return self.handle_register(msg_text, user_name, chat_id)

        if msg_type == "query":
            return self.handle_query(msg_text, user_name, chat_id)

        if msg_type == "report":
            return self.handle_report(msg_text, user_name, chat_id)

        if msg_type == "remind":
            return self.handle_remind(msg_text, user_name, chat_id)

        return (f"未识别操作类型。支持：\n"
                f"• 登记: 拼多多火车票5 拼多多机票3 千牛1\n"
                f"• 登记: HT001, HT002\n"
                f"• 查询: 查询/排名/占比\n"
                f"• 报表: 日报/周报/月报\n"
                f"• 提醒: 提醒上报")


def _get_field(fields: dict, name: str) -> str:
    val = fields.get(name, "")
    if isinstance(val, list) and val:
        val = val[0].get("text", "") if isinstance(val[0], dict) else str(val[0])
    return str(val) if val else ""


# ══════════════════════════════════════════════════════
# 定时提醒调度器
# ══════════════════════════════════════════════════════

def run_scheduler(handler: MessageHandler):
    """后台线程：每分钟检查是否到提醒时间（14:55 和 23:55）。"""
    reminder_times = {"14:55", "23:55"}
    last_fired = set()

    while True:
        try:
            now = datetime.now()
            current_time = now.strftime("%H:%M")

            if current_time in reminder_times and current_time not in last_fired:
                sys.stderr.write(f"[SCHEDULE] 触发定时提醒: {current_time}\n")
                sys.stderr.flush()
                handler.scheduled_remind()
                last_fired.add(current_time)

            # 过了提醒分钟就清除标记，为下一天准备
            if current_time not in reminder_times:
                last_fired.discard(current_time)

            time.sleep(30)
        except Exception:
            traceback.print_exc()
            time.sleep(30)


# ══════════════════════════════════════════════════════
# SDK 事件回调 — 收到消息后转交 MessageHandler
# ══════════════════════════════════════════════════════

def create_event_handler(feishu: FeishuClient, handler: MessageHandler):
    """创建 SDK 事件处理器。"""

    def on_message_receive(event):
        try:
            msg = event.event.message
            msg_id = msg.message_id or ""

            # 检查 EventContext 中是否有额外字段
            sys.stderr.write(f"[DEBUG] EventContext: schema={event.schema}, type={event.type}, header={event.header}\n")
            sys.stderr.flush()

            # 持久化去重：同一条消息绝不处理两次
            if handler.is_duplicate(msg_id):
                sys.stderr.write(f"[DEBUG] skip duplicate: {msg_id}\n")
                sys.stderr.flush()
                return

            # 先标记已处理，再执行业务逻辑（防止并发重复）
            handler.mark_processed(msg_id)
            content_str = msg.content or "{}"
            try:
                msg_text = json.loads(content_str).get("text", "")
            except (json.JSONDecodeError, AttributeError):
                msg_text = content_str if isinstance(content_str, str) else ""

            # 提取发送者
            open_id = ""
            sender = event.event.sender
            if sender and sender.sender_id:
                open_id = sender.sender_id.open_id or ""

            chat_id = msg.chat_id or ""
            user_name = feishu.get_user_name(open_id) if open_id else "未知用户"

            sys.stderr.write(f"[DEBUG] WS msg: '{msg_text}' from {user_name}\n")
            sys.stderr.flush()

            # 查ID 命令（去掉 @提及后匹配）
            clean_text = re.sub(r'@\S+', '', msg_text).strip()
            if clean_text.lower() in ("查id", "我的id", "myid", "id"):
                reply = f"你的 open_id: {open_id}\n当前姓名: {user_name}"
            else:
                reply = handler.handle(msg_text, user_name, chat_id)
            sys.stderr.write(f"[DEBUG] handle result: {repr(reply)[:200]}\n")
            sys.stderr.flush()
            if reply:
                if msg_id:
                    result = feishu.reply_message(msg_id, reply)
                    sys.stderr.write(f"[DEBUG] reply_message result: {result}\n")
                else:
                    feishu.send_message(chat_id, reply)
                    sys.stderr.write(f"[DEBUG] send_message to {chat_id}\n")
                sys.stderr.flush()
        except Exception:
            traceback.print_exc()

    event_handler = EventDispatcherHandler.builder("", "")\
        .register_p2_im_message_receive_v1(on_message_receive)\
        .build()

    return event_handler


# ══════════════════════════════════════════════════════
# 启动
# ══════════════════════════════════════════════════════

if __name__ == "__main__":
    print("=" * 50)
    print("飞书 Bot 长连接模式 (lark-oapi SDK)")
    print("=" * 50)

    missing = []
    for key, val in [("FEISHU_APP_ID", FEISHU_APP_ID),
                     ("FEISHU_APP_SECRET", FEISHU_APP_SECRET)]:
        status = "OK" if val else "MISSING"
        print(f"  [{status}] {key}")
        if not val:
            missing.append(key)

    if missing:
        print(f"\n缺少环境变量: {', '.join(missing)}")
        sys.exit(1)

    feishu = FeishuClient(FEISHU_APP_ID, FEISHU_APP_SECRET)
    handler = MessageHandler(feishu)

    if not handler._state.get("app_token"):
        print("❌ 未找到当月 Base 状态文件，请先创建表格")
        sys.exit(1)

    print(f"  Base: {handler._state['app_token']}")
    print(f"  表1: {handler._state['table1_id']}")
    print(f"  表2: {handler._state['table2_id']}")
    print("\n启动定时提醒 (14:55 / 23:55)...")
    threading.Thread(target=run_scheduler, args=(handler,), daemon=True).start()
    print("启动长连接...")
    print("(请在飞书群里发消息测试)")

    event_handler = create_event_handler(feishu, handler)
    client = WSClient(FEISHU_APP_ID, FEISHU_APP_SECRET, event_handler=event_handler)
    client.start()
