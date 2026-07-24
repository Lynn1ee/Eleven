"""飞书 Bot — lark-oapi SDK WebSocket 长连接 + Flask/ngrok Webhook 双通道"""

import json
import os
import sys
import io

# Windows 控制台/管道默认 GBK 编码无法输出 emoji，强制切换为 UTF-8
for stream_name in ('stdout', 'stderr'):
    try:
        stream = getattr(sys, stream_name)
        # detach 旧 TextIOWrapper(GDK)，换新的 UTF-8 包装
        raw = stream.detach()
        setattr(sys, stream_name, io.TextIOWrapper(raw, encoding='utf-8'))
    except (AttributeError, OSError, ValueError):
        pass
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

# 禁用系统代理 — 防止 Clash/V2Ray 等代理软件关闭后导致飞书 API 连接失败
for _pv in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy", "ALL_PROXY", "all_proxy"):
    os.environ.pop(_pv, None)
os.environ["NO_PROXY"] = "*"

import re
import requests

# lark-oapi SDK（导入较慢，约需 15-20 秒）
sys.stderr.write("[INIT] 正在加载飞书 SDK...\n")
sys.stderr.flush()
from lark_oapi.event.dispatcher_handler import EventDispatcherHandler
from lark_oapi.adapter.flask.parser import parse_req, parse_resp
from flask import Flask
try:
    from pyngrok import ngrok
except ImportError:
    ngrok = None
sys.stderr.write("[INIT] SDK 加载完成\n")
sys.stderr.flush()

# utils
SKILL_DIR = Path(os.getenv("SKILL_DIR", Path(__file__).parent.parent /
                           ".claude" / "skills" / "sales-tracker"))
sys.path.insert(0, str(SKILL_DIR / "scripts"))
from utils import extract_table1_data, extract_table2_data, classify_message, calc_change_rate
from chart_gen import trend_line, pie_chart, bar_chart, table_image

# ══════════════════════════════════════════════════════
# 配置
# ══════════════════════════════════════════════════════

FEISHU_APP_ID = os.getenv("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")
FEISHU_ENCRYPT_KEY = os.getenv("FEISHU_ENCRYPT_KEY", "")
FEISHU_VERIFICATION_TOKEN = os.getenv("FEISHU_VERIFICATION_TOKEN", "")

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
        if f.exists():
            try:
                return json.loads(f.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, KeyError):
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

    def upload_image(self, file_path: str) -> str:
        """上传图片到飞书，返回 image_key。"""
        with open(file_path, "rb") as f:
            resp = requests.post(
                "https://open.feishu.cn/open-apis/im/v1/images",
                headers={"Authorization": f"Bearer {self._get_tenant_token()}"},
                files={"image": f},
                data={"image_type": "message"},
                timeout=30,
            )
        data = resp.json()
        code = data.get("code", -1)
        if code != 0:
            sys.stderr.write(f"[DEBUG] upload_image failed: {data}\n")
            sys.stderr.flush()
            return ""
        return data["data"]["image_key"]

    def reply_image(self, msg_id: str, image_key: str):
        """以图片回复某条消息。"""
        body = {"msg_type": "image", "content": json.dumps({"image_key": image_key})}
        return requests.post(
            f"https://open.feishu.cn/open-apis/im/v1/messages/{msg_id}/reply",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            json=body, timeout=10,
        ).json()

    def send_image(self, chat_id: str, image_key: str):
        """发送图片到群聊。"""
        body = {
            "receive_id": chat_id,
            "msg_type": "image",
            "content": json.dumps({"image_key": image_key}),
        }
        requests.post(
            "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            json=body, timeout=10,
        )

    def reply_card(self, msg_id: str, card: dict):
        """以卡片消息引用回复。"""
        body = {"msg_type": "interactive", "content": json.dumps(card, ensure_ascii=False)}
        return requests.post(
            f"https://open.feishu.cn/open-apis/im/v1/messages/{msg_id}/reply",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            data=json.dumps(body, ensure_ascii=False).encode("utf-8"), timeout=15,
        ).json()

    def send_card(self, chat_id: str, card: dict):
        """发送卡片消息到群聊，返回 API 结果 dict。"""
        body = {
            "receive_id": chat_id,
            "msg_type": "interactive",
            "content": json.dumps(card, ensure_ascii=False),
        }
        resp = requests.post(
            "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            data=json.dumps(body, ensure_ascii=False).encode("utf-8"), timeout=15,
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

    def update_record(self, app_token: str, table_id: str, record_id: str, fields: dict):
        """更新单条记录的部分字段。"""
        resp = requests.put(
            f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/{record_id}",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            json={"fields": fields}, timeout=15,
        )
        return resp.json()

    def delete_records(self, app_token: str, table_id: str, record_ids: list):
        """批量删除记录（每次最多 500 条）。"""
        resp = requests.post(
            f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_delete",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            json={"records": record_ids}, timeout=15,
        )
        return resp.json()

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

    def add_collaborator(self, app_token: str, open_id: str, perm: str = "full_access"):
        """将成员加入多维表格协作列表。
        perm: "full_access"（可编辑）, "manage"（管理员）
        """
        resp = requests.post(
            f"https://open.feishu.cn/open-apis/drive/v1/permissions/{app_token}/members?type=bitable",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            json={
                "member_type": "openid",
                "member_id": open_id,
                "perm": perm,
            },
            timeout=10,
        )
        return resp.json()

    def create_bitable(self, name: str) -> dict:
        """创建多维表格 Base，返回 API 响应 JSON。"""
        resp = requests.post(
            "https://open.feishu.cn/open-apis/bitable/v1/apps",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            json={"name": name},
            timeout=15,
        )
        return resp.json()

    def create_table(self, app_token: str, name: str, fields: list) -> dict:
        """在 Base 中创建数据表，fields = [{"field_name": str, "type": int}, ...]。"""
        resp = requests.post(
            f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables",
            headers={
                "Authorization": f"Bearer {self._get_tenant_token()}",
                "Content-Type": "application/json; charset=utf-8",
            },
            json={"table": {"name": name, "fields": fields}},
            timeout=15,
        )
        return resp.json()

    def list_messages(self, chat_id: str, start_time: str = None,
                      end_time: str = None, page_size: int = 50) -> list:
        """拉取群聊历史消息（分页），返回原始消息 dict 列表。"""
        all_msgs = []
        page_token = ""
        while True:
            params = {
                "container_id_type": "chat",
                "container_id": chat_id,
                "sort_type": "ByCreateTimeAsc",
                "page_size": str(page_size),
            }
            if start_time:
                params["start_time"] = start_time
            if end_time:
                params["end_time"] = end_time
            if page_token:
                params["page_token"] = page_token
            resp = requests.get(
                "https://open.feishu.cn/open-apis/im/v1/messages",
                headers={"Authorization": f"Bearer {self._get_tenant_token()}"},
                params=params, timeout=15,
            )
            data = resp.json()
            if data.get("code") != 0:
                sys.stderr.write(f"[BACKFILL] list_messages 失败: "
                                 f"code={data.get('code')} msg={data.get('msg','')}\n")
                break
            items = data.get("data", {}).get("items", [])
            all_msgs.extend(items)
            if not data.get("data", {}).get("has_more"):
                break
            page_token = data["data"].get("page_token", "")
            if not page_token:
                break
            time.sleep(0.3)
        return all_msgs


# ══════════════════════════════════════════════════════
# 消息处理引擎
# ══════════════════════════════════════════════════════

class MessageHandler:
    def __init__(self, feishu_client):
        self.fs = feishu_client
        self._state = self._load_state()
        self._chat_ids: set = self._load_chat_ids()
        self._remind_lock = threading.Lock()
        self._dedup_lock = threading.Lock()
        self._pending_charts: list = []  # 待发送的图表文件路径
        self._pending_card: dict | None = None  # 待发送的卡片 JSON
        self._ensure_monthly_base()

    def _load_state(self) -> dict:
        month = datetime.now().strftime("%Y-%m")
        state_file = Path(__file__).parent / f".state_{month}.json"
        if state_file.exists():
            return json.loads(state_file.read_text())
        return {"app_token": "", "table1_id": "", "table2_id": ""}

    # ── 群聊 ID 持久化 ─────────────────────────

    @property
    def _chat_ids_file(self) -> Path:
        return Path(__file__).parent / ".chat_ids.json"

    def _load_chat_ids(self) -> set:
        f = self._chat_ids_file
        if f.exists():
            try:
                data = json.loads(f.read_text())
                return set(data.get("chat_ids", []))
            except (json.JSONDecodeError, KeyError):
                pass
        return set()

    def _save_chat_ids(self):
        data = {"chat_ids": sorted(self._chat_ids),
                "updated_at": datetime.now().isoformat()}
        self._chat_ids_file.write_text(json.dumps(data, ensure_ascii=False))

    # ── 补录状态持久化 ────────────────────────

    @property
    def _backfill_state_file(self) -> Path:
        return Path(__file__).parent / ".backfill_state.json"

    def _load_backfill_state(self) -> dict:
        f = self._backfill_state_file
        if f.exists():
            try:
                return json.loads(f.read_text())
            except (json.JSONDecodeError, KeyError):
                pass
        return {}

    def _save_backfill_state(self, state: dict):
        self._backfill_state_file.write_text(json.dumps(state, ensure_ascii=False))

    def _ensure_monthly_base(self) -> bool:
        """确保当月有独立的 Base：不存在则自动创建，并从上月迁移当日数据。"""
        now = datetime.now()
        month_str = now.strftime("%Y-%m")
        state_file = Path(__file__).parent / f".state_{month_str}.json"
        current_app_token = self._state.get("app_token", "")

        # 上月的 app_token，用于检测 stale state
        prev_month = (now.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
        prev_state_file = Path(__file__).parent / f".state_{prev_month}.json"
        prev_app_token = ""
        if prev_state_file.exists():
            try:
                prev_app_token = json.loads(prev_state_file.read_text()).get("app_token", "")
            except (json.JSONDecodeError, KeyError):
                pass

        # 已有独立 Base — 无需操作
        if current_app_token and current_app_token != prev_app_token:
            return True

        print(f"[AUTO] 自动创建 {month_str} Base...")

        # 1. 创建 Base
        base_name = f"{month_str} 业务任务量"
        resp = self.fs.create_bitable(base_name)
        if resp.get("code") != 0:
            print(f"[AUTO] 创建 Base 失败: {resp}")
            return False
        app_token = resp["data"]["app"]["app_token"]

        # 2. 创建表1 — IM会话量
        table1_fields = [
            {"field_name": "日期", "type": 5},
            {"field_name": "姓名", "type": 1},
            {"field_name": "拼多多火车票", "type": 2},
            {"field_name": "拼多多机票", "type": 2},
            {"field_name": "千牛", "type": 2},
            {"field_name": "抖音", "type": 2},
            {"field_name": "合计", "type": 2},
            {"field_name": "备注", "type": 1},
        ]
        resp = self.fs.create_table(app_token, "IM会话量", table1_fields)
        if resp.get("code") != 0:
            print(f"[AUTO] 创建表1失败: {resp}")
            return False
        table1_id = resp["data"]["table_id"]

        # 3. 创建表2 — 火车票订单
        table2_fields = [
            {"field_name": "日期", "type": 5},
            {"field_name": "姓名", "type": 1},
            {"field_name": "订单号", "type": 1},
            {"field_name": "备注", "type": 1},
        ]
        resp = self.fs.create_table(app_token, "火车票订单", table2_fields)
        if resp.get("code") != 0:
            print(f"[AUTO] 创建表2失败: {resp}")
            return False
        table2_id = resp["data"]["table_id"]

        # 4. 保存状态（先保存再迁移，确保迁移写入目标正确）
        self._state = {
            "app_token": app_token,
            "table1_id": table1_id,
            "table2_id": table2_id,
        }
        state_file.write_text(json.dumps(self._state, ensure_ascii=False), encoding="utf-8")

        # 5. 从上月 Base 迁移当日数据
        if prev_app_token and prev_state_file.exists():
            self._migrate_current_month_data(prev_app_token, prev_state_file)

        print(f"[AUTO] {month_str} Base 创建完成: {app_token}")
        return True

    def _migrate_current_month_data(self, prev_app_token: str, prev_state_file: Path):
        """将上月 Base 中当月 1 日的数据迁移到当前 Base。"""
        now = datetime.now()
        day1_start = int(datetime(now.year, now.month, 1).timestamp() * 1000)
        day1_end = int(datetime(now.year, now.month, 2).timestamp() * 1000) - 1

        prev_state = json.loads(prev_state_file.read_text())
        # table1 的数值字段（旧 Base 可能存为字符串，需转数字）
        table1_number_fields = {"拼多多火车票", "拼多多机票", "千牛", "抖音", "合计"}
        targets = [
            (prev_state.get("table1_id", ""), "IM会话量", self._state["table1_id"], table1_number_fields),
            (prev_state.get("table2_id", ""), "火车票订单", self._state["table2_id"], set()),
        ]

        for prev_table_id, label, new_table_id, number_fields in targets:
            if not prev_table_id:
                continue

            records = self.fs.search_records(prev_app_token, prev_table_id)
            to_migrate = []
            for r in records:
                date_val = r.get("fields", {}).get("日期")
                if date_val and day1_start <= int(date_val) <= day1_end:
                    to_migrate.append(r)

            if not to_migrate:
                continue

            # 清洗字段：去掉自动编号 ID、字符串数字转为 int
            new_records = []
            for r in to_migrate:
                fields = {}
                for k, v in r["fields"].items():
                    if k == "ID":
                        continue
                    if k in number_fields and isinstance(v, str):
                        try:
                            fields[k] = int(v)
                        except ValueError:
                            fields[k] = v
                    else:
                        fields[k] = v
                new_records.append({"fields": fields})

            # 写入新 Base
            result = self.fs.add_records(self._state["app_token"], new_table_id, new_records)
            if result.get("code") != 0:
                print(f"[MIGRATE] 迁移 {label} 写入失败: {result}")
                continue

            # 从旧 Base 删除（分批 ≤500）
            record_ids = [r["record_id"] for r in to_migrate]
            for i in range(0, len(record_ids), 500):
                self.fs.delete_records(prev_app_token, prev_table_id, record_ids[i:i + 500])

            print(f"[MIGRATE] 已迁移 {len(to_migrate)} 条 {label} 记录")

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
        with self._dedup_lock:
            processed = self._load_processed_ids()
            processed.add(msg_id)
            self._save_processed_ids(processed)

    def _check_and_mark_processed(self, msg_id: str) -> bool:
        """原子操作：检查并标记已处理。返回 True 表示重复（跳过）。"""
        if not msg_id:
            return False
        with self._dedup_lock:
            processed = self._load_processed_ids()
            if msg_id in processed:
                return True
            processed.add(msg_id)
            self._save_processed_ids(processed)
            return False

    # ── 姓名解析 ──────────────────────────────

    def _resolve_name(self, msg_text: str, fallback: str) -> str:
        """从消息文本中匹配已知姓名，匹配到则返回该姓名，否则返回 fallback。"""
        name_map = self.fs._load_name_map()
        known_names = sorted(set(name_map.values()), key=len, reverse=True)
        for name in known_names:
            if name in msg_text:
                return name
        return fallback

    @staticmethod
    def _parse_date(msg_text: str):
        """从消息文本开头提取日期，返回毫秒时间戳，无日期则返回 None。

        支持格式（M=月 D=日）:
          分隔符:   M-D, M/D, M.D, M-D日, M/D号 等
          中文:     M月D, M月D日, M月D号
          完整年份: YYYY-M-D, YYYY年M月D日, YYYY/M/D 等
        """
        m = re.match(
            r'(?:(\d{4})\s*[年/\-.]\s*)?'
            r'(\d{1,2})\s*(?:[/\-.]|月)\s*'
            r'(\d{1,2})\s*[日号]?',
            msg_text,
        )
        if m:
            year = int(m.group(1)) if m.group(1) else datetime.now().year
            month, day = int(m.group(2)), int(m.group(3))
            return int(datetime(year, month, day).timestamp() * 1000)
        return None

    # ── 登记处理 ──────────────────────────────

    def handle_register(self, msg_text: str, user_name: str, chat_id: str) -> str:
        resolved_name = self._resolve_name(msg_text, user_name)
        record_date = self._parse_date(msg_text) or int(datetime.now().timestamp() * 1000)
        msg_type = classify_message(msg_text)

        if msg_type == "table1":
            data = extract_table1_data(msg_text)
            total = data["拼多多火车票"] + data["拼多多机票"] + data["千牛"] + data["抖音"]
            if total == 0:
                return "未识别到有效数据。格式：拼多多火车票5 拼多多机票3 千牛1"

            record = {"fields": {
                "日期": record_date,
                "姓名": resolved_name,
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
                    "日期": record_date,
                    "姓名": resolved_name,
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
        # 从消息文本中提取查询目标姓名
        import re
        target = msg_text
        for kw in ["查询", "统计", "汇总", "多少", "帮我查", "帮我看看", "排名"]:
            target = target.replace(kw, "")
        target = target.replace("的", "")
        target = re.sub(r"\s+", "", target).strip()
        has_explicit_name = bool(target)

        if "排名" in msg_text:
            if has_explicit_name:
                return self._query_user_rank(target)
            return self._query_ranking()
        return self._query_user(target if has_explicit_name else user_name)

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

    def _query_user_rank(self, user_name: str) -> str:
        """查询指定人的排名 + 个人汇总。"""
        all_states = self._load_all_states()
        user_totals = defaultdict(int)
        for _, app_token, t1_id, _ in all_states:
            records = self.fs.search_records(app_token, t1_id)
            for r in records:
                name = _get_field(r.get("fields", {}), "姓名")
                total = int(_get_field(r.get("fields", {}), "合计") or 0)
                user_totals[name] += total

        sorted_users = sorted(user_totals.items(), key=lambda x: x[1], reverse=True)
        rank = None
        for i, (name, _) in enumerate(sorted_users):
            if name == user_name:
                rank = i + 1
                break

        # 个人汇总
        personal = self._query_user(user_name)
        total_people = len(sorted_users)
        rank_str = f"第{rank}名 / 共{total_people}人" if rank else "暂无排名数据"
        return f"{personal}\n\n🏆 排名: {rank_str}"

    # ── 跨月搜索 ────────────────────────────

    def _search_all_states(self, table_type: str) -> list:
        """搜索所有月份的 Base，返回合并后的记录列表。
        table_type: "table1" 或 "table2"
        """
        all_states = self._load_all_states()
        all_records = []
        for _, app_token, t1_id, t2_id in all_states:
            table_id = t1_id if table_type == "table1" else t2_id
            if table_id:
                all_records.extend(self.fs.search_records(app_token, table_id))
        return all_records

    # ── 卡片构建 ────────────────────────────

    def _upload_chart_for_card(self, file_path: str) -> str:
        """上传单个图表到飞书，返回 image_key，失败返回空字符串。"""
        try:
            return self.fs.upload_image(file_path)
        except Exception:
            traceback.print_exc()
            return ""

    def _make_md_table(self, headers: list, rows: list) -> str:
        """构建 Markdown 表格文本，用于飞书卡片内嵌展示。"""
        cols = len(headers)
        sep = "|" + "|".join(["---"] * cols) + "|"
        header_line = "|" + "|".join(str(h) for h in headers) + "|"
        data_lines = ["|" + "|".join(str(c) for c in row) + "|" for row in rows]
        return "\n".join([header_line, sep] + data_lines)

    def _make_table_img(self, headers: list, rows: list, title: str = "") -> str:
        """生成数据表格 PNG，上传飞书，返回 image_key。"""
        try:
            path = table_image(headers, rows, title=title)
            self._pending_charts.append(path)
            return self._upload_chart_for_card(path)
        except Exception:
            traceback.print_exc()
            return ""

    def _make_card(self, title: str, color: str, elements: list) -> dict:
        """构建飞书卡片 JSON 2.0。"""
        return {
            "schema": "2.0",
            "config": {"enable_forward": True, "width_mode": "fill"},
            "header": {
                "template": color,
                "title": {"tag": "plain_text", "content": title},
            },
            "body": {
                "direction": "vertical",
                "elements": elements,
            },
        }

    @staticmethod
    def _md(content: str) -> dict:
        return {"tag": "markdown", "content": content}

    @staticmethod
    def _hr() -> dict:
        return {"tag": "hr"}

    @staticmethod
    def _note(content: str) -> dict:
        return {"tag": "markdown", "content": f"---\n{content}"}

    @staticmethod
    def _img(image_key: str, alt: str = "") -> dict:
        return {
            "tag": "img",
            "img_key": image_key,
            "alt": {"tag": "plain_text", "content": alt},
            "scale_type": "fit_horizontal",
        }

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
        weekday = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"][
            (datetime.now() - timedelta(days=1)).weekday()
        ]
        records1 = self._search_all_states("table1")
        records2 = self._search_all_states("table2")

        user_data = defaultdict(lambda: {"total": 0, "pdd_train": 0, "pdd_flight": 0, "qn": 0, "dy": 0})
        for r in records1:
            f = r.get("fields", {})
            if _ts_to_date(_get_field(f, "日期")) == yesterday:
                name = _get_field(f, "姓名")
                user_data[name]["total"] += int(_get_field(f, "合计") or 0)
                user_data[name]["pdd_train"] += int(_get_field(f, "拼多多火车票") or 0)
                user_data[name]["pdd_flight"] += int(_get_field(f, "拼多多机票") or 0)
                user_data[name]["qn"] += int(_get_field(f, "千牛") or 0)
                user_data[name]["dy"] += int(_get_field(f, "抖音") or 0)

        order_by_user = defaultdict(int)
        for r in records2:
            f = r.get("fields", {})
            if _ts_to_date(_get_field(f, "日期")) == yesterday:
                name = _get_field(f, "姓名")
                order_by_user[name] += 1

        if not user_data and not order_by_user:
            return f"📊 日报 — {yesterday} {weekday}\n\n暂无昨日数据"

        all_names = sorted(set(list(user_data.keys()) + list(order_by_user.keys())))

        g_train = sum(d["pdd_train"] for d in user_data.values())
        g_flight = sum(d["pdd_flight"] for d in user_data.values())
        g_qn = sum(d["qn"] for d in user_data.values())
        g_dy = sum(d["dy"] for d in user_data.values())
        g_total = g_train + g_flight + g_qn + g_dy
        g_orders = sum(order_by_user.values())
        participant_count = len([n for n in all_names
                                 if user_data.get(n, {}).get('total', 0) or order_by_user.get(n, 0)])

        # ── 文本 fallback ──
        lines = [
            f"📊 日报 — {yesterday} {weekday}",
            "━━━━━━━━━━━━━━━━━━━━",
            "",
            "【IM 会话量】",
        ]
        for name in all_names:
            d = user_data.get(name, {})
            t = d.get("total", 0)
            if t == 0:
                continue
            lines.append(
                f"  {name}：火车票 {d.get('pdd_train',0):>5} ｜机票 {d.get('pdd_flight',0):>5}"
                f" ｜千牛 {d.get('qn',0):>5} ｜抖音 {d.get('dy',0):>5} ｜合计 {t:>5}"
            )
        lines.append("  " + "─" * 36)
        lines.append(
            f"  平台合计：火车票 {g_train:>5} ｜机票 {g_flight:>5}"
            f" ｜千牛 {g_qn:>5} ｜抖音 {g_dy:>5} ｜总计 {g_total:>5}"
        )
        if order_by_user:
            lines.append("")
            lines.append("【火车票订单】")
            for name in all_names:
                cnt = order_by_user.get(name, 0)
                if cnt == 0:
                    continue
                lines.append(f"  {name}：{cnt} 条")
            lines.append(f"  ─────────────")
            lines.append(f"  订单合计：{g_orders} 条")
        lines.append("")
        lines.append(f"📌 参与人数：{participant_count}")

        # ── 7天趋势图 ──
        trend_img_key = ""
        try:
            days = []
            day_vals = []
            for i in range(6, -1, -1):
                d = (datetime.now() - timedelta(days=i + 1)).strftime("%Y-%m-%d")
                days.append(d[-5:])
                day_sum = sum(
                    int(r.get("fields", {}).get("合计") or 0)
                    for r in records1
                    if _ts_to_date(_get_field(r.get("fields", {}), "日期")) == d
                )
                day_vals.append(day_sum)
            if sum(day_vals) > 0:
                chart_path = trend_line(days, day_vals, title=f"IM 会话量近 7 天趋势", ylabel="会话量")
                trend_img_key = self._upload_chart_for_card(chart_path)
                self._pending_charts.append(chart_path)
        except Exception:
            traceback.print_exc()

        # ── 卡片 ──
        card_els = []

        # IM 会话量数据表格
        im_headers = ["姓名", "火车票", "机票", "千牛", "抖音", "合计"]
        im_rows = []
        for name in all_names:
            d = user_data.get(name, {})
            t = d.get("total", 0)
            if t == 0:
                continue
            im_rows.append([name, str(d.get("pdd_train", 0)), str(d.get("pdd_flight", 0)),
                            str(d.get("qn", 0)), str(d.get("dy", 0)), str(t)])
        im_rows.append(["平台合计", str(g_train), str(g_flight), str(g_qn), str(g_dy), str(g_total)])
        card_els.append(self._md(f"**📋 IM 会话量**\n\n{self._make_md_table(im_headers, im_rows)}"))
        card_els.append(self._hr())

        # 火车票订单数据表格
        if order_by_user:
            order_headers = ["姓名", "订单数"]
            order_rows = [[name, str(cnt)] for name, cnt in
                          sorted(order_by_user.items(), key=lambda x: x[1], reverse=True)]
            order_rows.append(["合计", str(g_orders)])
            card_els.append(self._md(f"**📋 火车票订单**\n\n{self._make_md_table(order_headers, order_rows)}"))
            card_els.append(self._hr())

        if trend_img_key:
            card_els.append(self._img(trend_img_key, "7天趋势图"))

        card_els.append(self._note(f"📌 参与人数：{participant_count}"))
        self._pending_card = self._make_card(f"📊 日报 — {yesterday} {weekday}", "blue", card_els)

        return "\n".join(lines)

    def _weekly_report(self) -> str:
        today = datetime.now()
        monday = today - timedelta(days=today.weekday())
        week_start = (monday - timedelta(days=7)).strftime("%Y-%m-%d")
        week_end = (monday - timedelta(days=1)).strftime("%Y-%m-%d")

        # 上上周（环比用）
        prev_week_start = (monday - timedelta(days=14)).strftime("%Y-%m-%d")
        prev_week_end = (monday - timedelta(days=8)).strftime("%Y-%m-%d")

        records1 = self._search_all_states("table1")
        user_totals = defaultdict(lambda: {"total": 0, "pdd_train": 0, "pdd_flight": 0, "qn": 0, "dy": 0, "days": set()})
        prev_total = 0
        for r in records1:
            f = r.get("fields", {})
            date_str = _ts_to_date(_get_field(f, "日期"))
            t = int(_get_field(f, "合计") or 0)
            if week_start <= date_str <= week_end:
                name = _get_field(f, "姓名")
                user_totals[name]["total"] += t
                user_totals[name]["pdd_train"] += int(_get_field(f, "拼多多火车票") or 0)
                user_totals[name]["pdd_flight"] += int(_get_field(f, "拼多多机票") or 0)
                user_totals[name]["qn"] += int(_get_field(f, "千牛") or 0)
                user_totals[name]["dy"] += int(_get_field(f, "抖音") or 0)
                user_totals[name]["days"].add(date_str)
            if prev_week_start <= date_str <= prev_week_end:
                prev_total += t

        records2 = self._search_all_states("table2")
        order_by_user = defaultdict(int)
        for r in records2:
            f = r.get("fields", {})
            if week_start <= _ts_to_date(_get_field(f, "日期")) <= week_end:
                order_by_user[_get_field(f, "姓名")] += 1

        all_names = sorted(set(list(user_totals.keys()) + list(order_by_user.keys())))
        g_train = sum(d["pdd_train"] for d in user_totals.values())
        g_flight = sum(d["pdd_flight"] for d in user_totals.values())
        g_qn = sum(d["qn"] for d in user_totals.values())
        g_dy = sum(d["dy"] for d in user_totals.values())
        g_total = g_train + g_flight + g_qn + g_dy
        g_orders = sum(order_by_user.values())
        change_str = calc_change_rate(g_total, prev_total)

        # ── 文本 fallback ──
        lines = [
            f"📈 周报 — {week_start} ~ {week_end}",
            "━━━━━━━━━━━━━━━━━━━━",
            "",
            "【IM 会话量】",
        ]
        sorted_users = sorted(user_totals.items(), key=lambda x: x[1]["total"], reverse=True)
        for i, (name, d) in enumerate(sorted_users):
            days = len(d["days"])
            avg = d["total"] // days if days > 0 else 0
            medal = ["🥇", "🥈", "🥉"][i] if i < 3 else "  "
            lines.append(
                f"  {medal} {name}：火车票 {d['pdd_train']:>5} ｜机票 {d['pdd_flight']:>5}"
                f" ｜千牛 {d['qn']:>5} ｜抖音 {d['dy']:>5} ｜合计 {d['total']:>5} ｜日均 {avg:>4}"
            )
        lines.append("  " + "─" * 44)
        lines.append(
            f"  平台合计：火车票 {g_train:>5} ｜机票 {g_flight:>5}"
            f" ｜千牛 {g_qn:>5} ｜抖音 {g_dy:>5} ｜总计 {g_total:>5}"
        )
        if order_by_user:
            lines.append("")
            lines.append("【火车票订单】")
            sorted_orders = sorted(order_by_user.items(), key=lambda x: x[1], reverse=True)
            for name, cnt in sorted_orders:
                lines.append(f"  {name}：{cnt} 条")
            lines.append(f"  ─────────────")
            lines.append(f"  订单合计：{g_orders} 条")
        lines.append("")
        lines.append(f"📌 参与人数：{len(all_names)} ｜工作日：{7} 天")

        # ── 图表生成并上传 ──
        bar_img_key = ""
        pie_img_key = ""
        try:
            top_users = sorted(user_totals.items(), key=lambda x: x[1]["total"], reverse=True)[:10]
            if top_users:
                names = [u[0] for u in top_users]
                vals = [u[1]["total"] for u in top_users]
                chart_path = bar_chart(names, vals, title=f"周排名对比 ({week_start}~{week_end})", ylabel="会话量")
                bar_img_key = self._upload_chart_for_card(chart_path)
                self._pending_charts.append(chart_path)
        except Exception:
            traceback.print_exc()
        try:
            if g_total > 0:
                pie_path = pie_chart(
                    ["火车票", "机票", "千牛", "抖音"],
                    [g_train, g_flight, g_qn, g_dy],
                    title=f"平台占比 ({week_start}~{week_end})",
                )
                pie_img_key = self._upload_chart_for_card(pie_path)
                self._pending_charts.append(pie_path)
        except Exception:
            traceback.print_exc()

        # ── 卡片 ──
        card_els = []

        # IM 会话量数据表格（含排名）
        im_headers = ["排名", "姓名", "火车票", "机票", "千牛", "抖音", "合计", "日均"]
        im_rows = []
        for i, (name, d) in enumerate(sorted_users):
            days = len(d["days"])
            avg = d["total"] // days if days > 0 else 0
            medal = ["🥇", "🥈", "🥉"][i] if i < 3 else str(i + 1)
            im_rows.append([medal, name, str(d["pdd_train"]), str(d["pdd_flight"]),
                            str(d["qn"]), str(d["dy"]), str(d["total"]), str(avg)])
        im_rows.append(["—", "平台合计", str(g_train), str(g_flight), str(g_qn), str(g_dy), str(g_total), "—"])
        card_els.append(self._md(f"**环比：{change_str}**") if change_str else None)
        card_els = [e for e in card_els if e is not None]
        card_els.append(self._md(f"**📋 IM 会话量**\n\n{self._make_md_table(im_headers, im_rows)}"))
        card_els.append(self._hr())

        # 火车票订单数据表格
        if order_by_user:
            order_headers = ["姓名", "订单数"]
            order_rows = [[name, str(cnt)] for name, cnt in
                          sorted(order_by_user.items(), key=lambda x: x[1], reverse=True)]
            order_rows.append(["合计", str(g_orders)])
            card_els.append(self._md(f"**📋 火车票订单**\n\n{self._make_md_table(order_headers, order_rows)}"))
            card_els.append(self._hr())

        if bar_img_key:
            card_els.append(self._img(bar_img_key, "排名对比"))
        if pie_img_key:
            card_els.append(self._img(pie_img_key, "平台占比"))

        card_els.append(self._note(f"📌 参与人数：{len(all_names)} ｜工作日：7 天"))
        self._pending_card = self._make_card(f"📈 周报 — {week_start} ~ {week_end}", "turquoise", card_els)

        return "\n".join(lines)

    def _monthly_report(self) -> str:
        last_month_dt = datetime.now().replace(day=1) - timedelta(days=1)
        last_month = last_month_dt.strftime("%Y-%m")

        # 上上月（环比用）
        prev_month = (last_month_dt.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")

        records1 = self._search_all_states("table1")
        user_totals = defaultdict(lambda: {"total": 0, "pdd_train": 0, "pdd_flight": 0, "qn": 0, "dy": 0, "days": set()})
        prev_total = 0
        daily_totals = defaultdict(int)
        for r in records1:
            f = r.get("fields", {})
            date_str = _ts_to_date(_get_field(f, "日期"))
            t = int(_get_field(f, "合计") or 0)
            if date_str[:7] == last_month:
                name = _get_field(f, "姓名")
                user_totals[name]["total"] += t
                user_totals[name]["pdd_train"] += int(_get_field(f, "拼多多火车票") or 0)
                user_totals[name]["pdd_flight"] += int(_get_field(f, "拼多多机票") or 0)
                user_totals[name]["qn"] += int(_get_field(f, "千牛") or 0)
                user_totals[name]["dy"] += int(_get_field(f, "抖音") or 0)
                user_totals[name]["days"].add(date_str)
                daily_totals[date_str] += t
            if date_str[:7] == prev_month:
                prev_total += t

        records2 = self._search_all_states("table2")
        order_by_user = defaultdict(int)
        for r in records2:
            f = r.get("fields", {})
            if _ts_to_date(_get_field(f, "日期"))[:7] == last_month:
                order_by_user[_get_field(f, "姓名")] += 1

        all_names = sorted(set(list(user_totals.keys()) + list(order_by_user.keys())))
        g_train = sum(d["pdd_train"] for d in user_totals.values())
        g_flight = sum(d["pdd_flight"] for d in user_totals.values())
        g_qn = sum(d["qn"] for d in user_totals.values())
        g_dy = sum(d["dy"] for d in user_totals.values())
        g_total = g_train + g_flight + g_qn + g_dy
        g_orders = sum(order_by_user.values())
        change_str = calc_change_rate(g_total, prev_total)

        # ── 文本 fallback ──
        lines = [
            f"📋 月报 — {last_month}",
            "━━━━━━━━━━━━━━━━━━━━",
            "",
            "【IM 会话量】",
        ]
        sorted_users = sorted(user_totals.items(), key=lambda x: x[1]["total"], reverse=True)
        for i, (name, d) in enumerate(sorted_users):
            days = len(d["days"])
            avg = d["total"] // days if days > 0 else 0
            medal = ["🥇", "🥈", "🥉"][i] if i < 3 else "  "
            lines.append(
                f"  {medal} {name}：火车票 {d['pdd_train']:>5} ｜机票 {d['pdd_flight']:>5}"
                f" ｜千牛 {d['qn']:>5} ｜抖音 {d['dy']:>5} ｜合计 {d['total']:>5} ｜日均 {avg:>4}"
            )
        lines.append("  " + "─" * 44)
        lines.append(
            f"  平台合计：火车票 {g_train:>5} ｜机票 {g_flight:>5}"
            f" ｜千牛 {g_qn:>5} ｜抖音 {g_dy:>5} ｜总计 {g_total:>5}"
        )
        if g_total > 0:
            lines.append("")
            lines.append("【平台占比】")
            lines.append(f"  火车票  {g_train:>5}  ({g_train/g_total*100:5.1f}%)")
            lines.append(f"  机票    {g_flight:>5}  ({g_flight/g_total*100:5.1f}%)")
            lines.append(f"  千牛    {g_qn:>5}  ({g_qn/g_total*100:5.1f}%)")
            lines.append(f"  抖音    {g_dy:>5}  ({g_dy/g_total*100:5.1f}%)")
        if order_by_user:
            lines.append("")
            lines.append("【火车票订单】")
            sorted_orders = sorted(order_by_user.items(), key=lambda x: x[1], reverse=True)
            for name, cnt in sorted_orders:
                lines.append(f"  {name}：{cnt} 条")
            lines.append(f"  ─────────────")
            lines.append(f"  订单合计：{g_orders} 条")
        lines.append("")
        lines.append(f"📌 参与人数：{len(all_names)}")

        # ── 图表生成并上传 ──
        bar_img_key = ""
        pie_img_key = ""
        trend_img_key = ""
        try:
            top_users = sorted(user_totals.items(), key=lambda x: x[1]["total"], reverse=True)[:10]
            if top_users:
                names = [u[0] for u in top_users]
                vals = [u[1]["total"] for u in top_users]
                chart_path = bar_chart(names, vals, title=f"月排名对比 ({last_month})", ylabel="会话量")
                bar_img_key = self._upload_chart_for_card(chart_path)
                self._pending_charts.append(chart_path)
        except Exception:
            traceback.print_exc()
        try:
            if g_total > 0:
                pie_path = pie_chart(
                    ["火车票", "机票", "千牛", "抖音"],
                    [g_train, g_flight, g_qn, g_dy],
                    title=f"平台占比 ({last_month})",
                )
                pie_img_key = self._upload_chart_for_card(pie_path)
                self._pending_charts.append(pie_path)
        except Exception:
            traceback.print_exc()
        # 日均趋势折线图
        try:
            if daily_totals:
                sorted_dates = sorted(daily_totals.keys())
                date_labels = [d[-5:] for d in sorted_dates]
                vals = [daily_totals[d] for d in sorted_dates]
                if sum(vals) > 0 and len(vals) >= 2:
                    trend_path = trend_line(date_labels, vals, title=f"每日会话量趋势 ({last_month})", ylabel="会话量")
                    trend_img_key = self._upload_chart_for_card(trend_path)
                    self._pending_charts.append(trend_path)
        except Exception:
            traceback.print_exc()

        # ── 卡片 ──
        card_els = []

        # IM 会话量数据表格（含排名）
        im_headers = ["排名", "姓名", "火车票", "机票", "千牛", "抖音", "合计", "日均"]
        im_rows = []
        for i, (name, d) in enumerate(sorted_users):
            days = len(d["days"])
            avg = d["total"] // days if days > 0 else 0
            medal = ["🥇", "🥈", "🥉"][i] if i < 3 else str(i + 1)
            im_rows.append([medal, name, str(d["pdd_train"]), str(d["pdd_flight"]),
                            str(d["qn"]), str(d["dy"]), str(d["total"]), str(avg)])
        im_rows.append(["—", "平台合计", str(g_train), str(g_flight), str(g_qn), str(g_dy), str(g_total), "—"])
        header_lines = []
        if change_str:
            header_lines.append(f"环比：{change_str}")
        if g_total > 0:
            header_lines.append(
                f"占比：火车票 {g_train/g_total*100:.1f}% ｜机票 {g_flight/g_total*100:.1f}%"
                f" ｜千牛 {g_qn/g_total*100:.1f}% ｜抖音 {g_dy/g_total*100:.1f}%"
            )
        if header_lines:
            card_els.append(self._md("\n".join(header_lines)))
        card_els.append(self._md(f"**📋 IM 会话量**\n\n{self._make_md_table(im_headers, im_rows)}"))
        card_els.append(self._hr())

        # 火车票订单数据表格
        if order_by_user:
            order_headers = ["姓名", "订单数"]
            order_rows = [[name, str(cnt)] for name, cnt in
                          sorted(order_by_user.items(), key=lambda x: x[1], reverse=True)]
            order_rows.append(["合计", str(g_orders)])
            card_els.append(self._md(f"**📋 火车票订单**\n\n{self._make_md_table(order_headers, order_rows)}"))
            card_els.append(self._hr())

        if bar_img_key:
            card_els.append(self._img(bar_img_key, "排名对比"))
        if pie_img_key:
            card_els.append(self._img(pie_img_key, "平台占比"))
        if trend_img_key:
            card_els.append(self._img(trend_img_key, "每日趋势"))

        card_els.append(self._note(f"📌 参与人数：{len(all_names)}"))
        self._pending_card = self._make_card(f"📋 月报 — {last_month}", "purple", card_els)

        return "\n".join(lines)

    # ── 提醒 ──────────────────────────────────

    def handle_remind(self, msg_text: str, user_name: str, chat_id: str) -> str:
        today = datetime.now().strftime("%Y-%m-%d")
        reported = set()

        # 查 table1（IM 会话量）
        records1 = self.fs.search_records(self._state["app_token"], self._state["table1_id"])
        for r in records1:
            f = r.get("fields", {})
            if _ts_to_date(_get_field(f, "日期")) == today:
                reported.add(_get_field(f, "姓名"))

        # 查 table2（火车票订单）
        if self._state.get("table2_id"):
            records2 = self.fs.search_records(self._state["app_token"], self._state["table2_id"])
            for r in records2:
                f = r.get("fields", {})
                if _ts_to_date(_get_field(f, "日期")) == today:
                    reported.add(_get_field(f, "姓名"))

        if reported:
            return f"📢 今日已上报: {', '.join(sorted(reported))}\n还没上报的同事请尽快提交！"
        return "📢 今日暂无上报，请大家及时提交数据！"

    # ── 修改与删除 ────────────────────────────

    def handle_delete(self, msg_text: str, user_name: str, chat_id: str) -> str:
        """解析删除消息，匹配记录并删除。

        格式：删除 <关键词>   （按关键词搜索删除）
             删除             （删除自己今天的记录）
        """
        # 去掉 "删除"/"删掉"/"去掉" 前缀
        keyword = re.sub(r'^(删除|删掉|去掉)\s*', '', msg_text).strip()

        # 在两张表中搜索
        today = datetime.now().strftime("%Y-%m-%d")
        to_delete = []
        for table_id, table_name in [(self._state["table1_id"], "表1"), (self._state["table2_id"], "表2")]:
            records = self.fs.search_records(self._state["app_token"], table_id)
            for r in records:
                f = r.get("fields", {})
                if keyword:
                    r_text = " ".join(str(v) for v in f.values())
                    if keyword in r_text:
                        to_delete.append((r["record_id"], table_id, table_name, f))
                else:
                    # 无关键词：删除自己今天的记录
                    r_name = _get_field(f, "姓名")
                    r_date = _ts_to_date(_get_field(f, "日期"))
                    if r_name == user_name and r_date == today:
                        to_delete.append((r["record_id"], table_id, table_name, f))

        if not to_delete:
            if keyword:
                return f"未找到匹配「{keyword}」的记录。"
            return f"今天没有 {user_name} 的登记记录。"

        # 删除按表分组，同一批记录可能跨表
        grouped = defaultdict(list)
        for rid, tid, _, _ in to_delete[:20]:
            grouped[tid].append(rid)
        deleted_count = 0
        errors = []
        for tid, rids in grouped.items():
            result = self.fs.delete_records(self._state["app_token"], tid, rids)
            if result.get("code") != 0:
                errors.append(result.get("msg", ""))
            else:
                deleted_count += len(rids)

        if errors:
            return f"❌ 删除失败: {'; '.join(errors)}"

        details = "\n".join(
            f"  {tname}: {_get_field(f, '订单号') or _get_field(f, '姓名') or keyword}"
            for _, _, tname, f in to_delete[:5]
        )
        return f"✅ 已删除 {deleted_count} 条记录:\n{details}"

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

    # ── 断连补录 ──────────────────────────────

    @staticmethod
    def _extract_text_from_raw_msg(msg: dict) -> str:
        """从原始 API 消息 dict 提取文本（兼容 text/post 类型）。"""
        msg_type = msg.get("msg_type", "text") or "text"
        body = msg.get("body", {}) or {}
        content_str = body.get("content", "{}") or "{}"
        try:
            content_json = json.loads(content_str)
        except (json.JSONDecodeError, AttributeError):
            return content_str if isinstance(content_str, str) else ""
        if msg_type == "post":
            parts = []
            for paragraph in content_json.get("content", []):
                for element in paragraph:
                    if isinstance(element, dict) and element.get("tag") == "text":
                        parts.append(element.get("text", ""))
            return "".join(parts)
        return content_json.get("text", "")

    def backfill_messages(self):
        """后台线程：拉取各群聊历史消息并补录遗漏的注册消息。"""
        try:
            sys.stderr.write("[BACKFILL] 开始自动补录...\n")
            sys.stderr.flush()

            window_minutes = int(os.getenv("BACKFILL_WINDOW_MINUTES", "360"))
            max_window_minutes = int(os.getenv("BACKFILL_MAX_WINDOW_MINUTES", "1440"))
            window_minutes = min(window_minutes, max_window_minutes)

            chat_ids = list(self._chat_ids)
            if not chat_ids:
                sys.stderr.write("[BACKFILL] 无已知群聊，跳过补录\n")
                sys.stderr.flush()
                return

            backfill_state = self._load_backfill_state()
            now_ts = time.time()
            global_t1 = global_t2 = 0

            for chat_id in chat_ids:
                prev_ts = backfill_state.get(chat_id, 0)
                if prev_ts > 0:
                    start_ts = prev_ts  # 从上次补录时间开始，覆盖整个断连期
                else:
                    start_ts = now_ts - window_minutes * 60
                start_ts = max(start_ts, now_ts - max_window_minutes * 60)
                start_time_str = str(int(start_ts))

                sys.stderr.write(f"[BACKFILL] 拉取群聊 {chat_id} "
                                 f"起点={datetime.fromtimestamp(start_ts).strftime('%m-%d %H:%M')}\n")
                sys.stderr.flush()

                messages = self.fs.list_messages(chat_id, start_time=start_time_str)
                sys.stderr.write(f"[BACKFILL] 群聊 {chat_id} 拉取到 {len(messages)} 条\n")
                sys.stderr.flush()

                t1 = t2 = 0
                for msg in messages:
                    msg_id = msg.get("message_id", "")
                    if self._check_and_mark_processed(msg_id):
                        continue

                    sender = msg.get("sender", {}) or {}
                    if sender.get("sender_type") == "app":
                        continue

                    msg_type = msg.get("msg_type", "") or ""
                    if msg_type not in ("text", "post"):
                        continue

                    text = self._extract_text_from_raw_msg(msg)
                    if not text:
                        continue

                    clean = re.sub(r'@\S+', '', text).strip()
                    clean = re.sub(r'^回复\s*\S+[:：]\s*\n?', '', clean).strip()
                    if not clean:
                        continue

                    if re.match(r'^(✅|已记录映射|查询|报表|提醒|📊|📈|📋|📢|🏆)', clean):
                        continue

                    msg_type_class = classify_message(clean)
                    if msg_type_class not in ("table1", "table2"):
                        continue

                    open_id = sender.get("id", "") or ""
                    user_name = self._resolve_name(clean, "")
                    if not user_name and open_id:
                        user_name = self.fs.get_user_name(open_id)
                    if not user_name:
                        user_name = f"用户_{open_id[-4:]}" if open_id else "未知用户"

                    record_date = self._parse_date(clean)
                    if not record_date:
                        try:
                            create_ts = int(msg.get("create_time", "0") or "0") // 1000
                            d = datetime.fromtimestamp(create_ts)
                            record_date = int(datetime(d.year, d.month, d.day).timestamp() * 1000)
                        except (ValueError, OSError):
                            record_date = int(datetime.now().timestamp() * 1000)

                    if msg_type_class == "table1":
                        data = extract_table1_data(clean)
                        total = data["拼多多火车票"] + data["拼多多机票"] + data["千牛"] + data["抖音"]
                        if total > 0:
                            record = {"fields": {
                                "日期": record_date,
                                "姓名": user_name,
                                "拼多多火车票": data["拼多多火车票"],
                                "拼多多机票": data["拼多多机票"],
                                "千牛": data["千牛"],
                                "抖音": data["抖音"],
                                "合计": total,
                                "备注": data.get("备注", ""),
                            }}
                            result = self.fs.add_records(
                                self._state["app_token"], self._state["table1_id"], [record])
                            if result.get("code") == 0:
                                t1 += 1
                            else:
                                sys.stderr.write(f"[BACKFILL] table1 失败: {result.get('msg')}\n")
                    else:
                        orders = extract_table2_data(clean)
                        if orders:
                            records = [{"fields": {
                                "日期": record_date, "姓名": user_name,
                                "订单号": o, "备注": "",
                            }} for o in orders]
                            result = self.fs.add_records(
                                self._state["app_token"], self._state["table2_id"], records)
                            if result.get("code") == 0:
                                t2 += len(orders)
                            else:
                                sys.stderr.write(f"[BACKFILL] table2 失败: {result.get('msg')}\n")

                    sys.stderr.write(f"[BACKFILL] 补录: {user_name} {clean[:60]}\n")
                    sys.stderr.flush()
                    # 更新最后收消息时间（供 watchdog 检测存活）
                    try:
                        (Path(__file__).parent / ".last_msg_ts").write_text(str(time.time()))
                    except Exception:
                        pass

                backfill_state[chat_id] = now_ts
                self._save_backfill_state(backfill_state)
                global_t1 += t1
                global_t2 += t2

                if t1 > 0 or t2 > 0:
                    parts = []
                    if t1 > 0:
                        parts.append(f"IM会话量 {t1} 条")
                    if t2 > 0:
                        parts.append(f"火车票订单 {t2} 条")
                    summary = f"🔄 断连补录完成：{', '.join(parts)}"
                    try:
                        self.fs.send_message(chat_id, summary)
                    except Exception:
                        pass

                sys.stderr.write(f"[BACKFILL] 群聊 {chat_id} 补录: table1={t1} table2={t2}\n")

            sys.stderr.write(f"[BACKFILL] 补录结束: table1={global_t1} table2={global_t2}\n")
            sys.stderr.flush()

        except Exception:
            sys.stderr.write(f"[BACKFILL] 异常:\n{traceback.format_exc()}\n")
            sys.stderr.flush()

    # ── 入口 ──────────────────────────────────

    def handle(self, msg_text: str, user_name: str, chat_id: str) -> str:
        sys.stderr.write(f"[DEBUG] handle msg: {msg_text} from {user_name}\n")
        sys.stderr.flush()

        if chat_id:
            if chat_id not in self._chat_ids:
                self._chat_ids.add(chat_id)
                self._save_chat_ids()

        msg_type = classify_message(msg_text)

        if msg_type == "table1" or msg_type == "table2":
            return self.handle_register(msg_text, user_name, chat_id)

        if msg_type == "query":
            return self.handle_query(msg_text, user_name, chat_id)

        if msg_type == "report":
            return self.handle_report(msg_text, user_name, chat_id)

        if msg_type == "remind":
            return self.handle_remind(msg_text, user_name, chat_id)

        if msg_type == "modify":
            # 只有删除走这里，修改功能已移除
            if any(kw in msg_text for kw in ["删除", "删掉", "去掉"]):
                return self.handle_delete(msg_text, user_name, chat_id)
            return "修改功能已移除。请使用删除后重新登记的方式修改数据。"

        return (f"未识别操作类型。支持：\n"
                f"• 登记: 拼多多火车票5 拼多多机票3 千牛1\n"
                f"• 登记: HT001, HT002\n"
                f"• 查询: 查询/排名\n"
                f"• 报表: 日报/周报/月报\n"
                f"• 删除: 删除 HT001\n"
                f"• 提醒: 提醒上报")


def _get_field(fields: dict, name: str) -> str:
    val = fields.get(name, "")
    if isinstance(val, list) and val:
        val = val[0].get("text", "") if isinstance(val[0], dict) else str(val[0])
    return str(val) if val else ""


def _ts_to_date(val) -> str:
    """将飞书日期字段值（毫秒时间戳）转为日期字符串 YYYY-MM-DD。"""
    try:
        ts = int(float(str(val))) / 1000
        return datetime.fromtimestamp(ts).strftime("%Y-%m-%d")
    except (ValueError, OSError, TypeError):
        s = str(val)
        return s[:10] if len(s) >= 10 else s


# ══════════════════════════════════════════════════════
# 定时提醒调度器
# ══════════════════════════════════════════════════════

def run_scheduler(handler: MessageHandler):
    """后台线程：每分钟检查是否到提醒时间（14:55 和 23:55）。"""
    reminder_times = {"14:55", "23:55"}
    last_fired_date = ""
    fired_today = set()

    while True:
        try:
            now = datetime.now()
            current_date = now.strftime("%Y-%m-%d")
            current_time = now.strftime("%H:%M")

            # 新的一天，重置触发记录
            if current_date != last_fired_date:
                last_fired_date = current_date
                fired_today.clear()

            if current_time in reminder_times and current_time not in fired_today:
                sys.stderr.write(f"[SCHEDULE] 触发定时提醒: {current_date} {current_time}\n")
                sys.stderr.flush()
                handler.scheduled_remind()
                fired_today.add(current_time)

            time.sleep(30)
        except Exception:
            traceback.print_exc()
            time.sleep(30)


# ══════════════════════════════════════════════════════
# SDK 事件回调 — 收到消息后转交 MessageHandler
# ══════════════════════════════════════════════════════

def _extract_text(msg) -> str:
    """从飞书消息中提取文本，兼容 text 和 post（富文本含图片）两种类型。"""
    msg_type = getattr(msg, "message_type", "text") or "text"
    content_str = msg.content or "{}"
    try:
        content_json = json.loads(content_str)
    except (json.JSONDecodeError, AttributeError):
        return content_str if isinstance(content_str, str) else ""

    if msg_type == "post":
        parts = []
        for paragraph in content_json.get("content", []):
            for element in paragraph:
                if isinstance(element, dict) and element.get("tag") == "text":
                    parts.append(element.get("text", ""))
        return "".join(parts)

    return content_json.get("text", "")


def create_event_handler(feishu: FeishuClient, handler: MessageHandler):
    """创建 SDK 事件处理器。"""

    def on_message_receive(event):
        try:
            msg = event.event.message
            msg_id = msg.message_id or ""

            # 检查 EventContext 中是否有额外字段
            sys.stderr.write(f"[DEBUG] EventContext: schema={event.schema}, type={event.type}, header={event.header}\n")
            sys.stderr.flush()

            # 原子去重：检查+标记，防止 WS 线程和 backfill 线程竞态
            if handler._check_and_mark_processed(msg_id):
                sys.stderr.write(f"[DEBUG] skip duplicate: {msg_id}\n")
                sys.stderr.flush()
                return

            msg_text = _extract_text(msg)

            # 提取发送者
            open_id = ""
            sender = event.event.sender
            if sender and sender.sender_id:
                open_id = sender.sender_id.open_id or ""

            chat_id = msg.chat_id or ""
            user_name = feishu.get_user_name(open_id) if open_id else "未知用户"

            sys.stderr.write(f"[DEBUG] WS msg type={getattr(msg, 'message_type', '?')} '{msg_text}' from {user_name}\n")
            sys.stderr.flush()

            # 清洗消息文本：去掉 @提及、引用回复前缀 "回复 XXX:"
            clean_text = re.sub(r'@\S+', '', msg_text).strip()
            clean_text = re.sub(r'^回复\s*\S+[:：]\s*\n?', '', clean_text).strip()

            clean_lower_no_space = re.sub(r'\s+', '', clean_text).lower()
            id_keywords = ("查id", "查询id", "我的id", "myid", "id", "查看id", "查看openid", "openid")
            if clean_lower_no_space in id_keywords:
                reply = f"你的 open_id: {open_id}\n当前姓名: {user_name}"
            else:
                # 检查是否 "查询ID 张三" 格式（带姓名参数）
                name_match = re.match(
                    r'(?:查询id|查id|查看id|查询openid|查看openid|我的id)\s+(.+)',
                    clean_text.strip(),
                    re.IGNORECASE,
                )
                if not name_match:
                    name_match = re.match(r'id\s+(.+)', clean_text.strip(), re.IGNORECASE)
                if name_match:
                    name = name_match.group(1).strip()
                    if open_id:
                        name_map = handler.fs._load_name_map()
                        name_map[open_id] = name
                        nm_file = Path(__file__).parent / "name_map.json"
                        nm_file.write_text(
                            json.dumps(name_map, ensure_ascii=False, indent=2),
                            encoding="utf-8",
                        )
                        reply = f"已记录映射: {name}\nopen_id: {open_id}"
                    else:
                        reply = "未能获取 open_id，请重试"
                elif clean_lower_no_space in ("授权", "授权我", "加权限", "给我权限"):
                    chat_type = getattr(msg, "chat_type", "") if msg else ""
                    if chat_type != "p2p":
                        reply = "权限管理请在私聊中发送此命令（DM bot）。"
                    elif open_id and handler._state.get("app_token"):
                        results = []
                        seen = set()
                        for app_token in [handler._state["app_token"]] + [
                            s[1] for s in handler._load_all_states()
                        ]:
                            if app_token in seen:
                                continue
                            seen.add(app_token)
                            r = feishu.add_collaborator(app_token, open_id)
                            code = r.get("code")
                            results.append(f"  Base {app_token[-6:]}: {'OK' if code == 0 else r.get('msg', code)}")
                        reply = f"已授权 {user_name} 为以下 Base 的编辑者:\n" + "\n".join(results)
                    else:
                        reply = "未找到 Base 或 open_id，请先确保有当月状态文件"
                else:
                    handler._pending_charts = []  # 清空上次的图表
                    handler._pending_card = None
                    reply = handler.handle(clean_text, user_name, chat_id)
            sys.stderr.write(f"[DEBUG] handle result: {repr(reply)[:200]}\n")
            sys.stderr.flush()
            if handler._pending_card:
                # 发送卡片消息（图表已内嵌在卡片中）
                if msg_id:
                    result = feishu.reply_card(msg_id, handler._pending_card)
                else:
                    feishu.send_card(chat_id, handler._pending_card)
                    result = {"code": 0}
                sys.stderr.write(f"[DEBUG] reply_card result: {result}\n")
                if result.get("code") != 0:
                    # 卡片失败，降级为文本 + 单独图片
                    sys.stderr.write(f"[DEBUG] card failed, fallback to text\n")
                    if reply:
                        if msg_id:
                            feishu.reply_message(msg_id, reply)
                        else:
                            feishu.send_message(chat_id, reply)
                    for chart_path in handler._pending_charts:
                        try:
                            image_key = feishu.upload_image(chart_path)
                            if image_key:
                                if msg_id:
                                    feishu.reply_image(msg_id, image_key)
                                else:
                                    feishu.send_image(chat_id, image_key)
                        except Exception:
                            traceback.print_exc()
                handler._pending_card = None
                sys.stderr.flush()
            elif reply:
                if msg_id:
                    result = feishu.reply_message(msg_id, reply)
                    sys.stderr.write(f"[DEBUG] reply_message result: {result}\n")
                else:
                    feishu.send_message(chat_id, reply)
                    sys.stderr.write(f"[DEBUG] send_message to {chat_id}\n")
                # 发送附带图表
                for chart_path in handler._pending_charts:
                    try:
                        image_key = feishu.upload_image(chart_path)
                        if image_key:
                            if msg_id:
                                feishu.reply_image(msg_id, image_key)
                            else:
                                feishu.send_image(chat_id, image_key)
                    except Exception:
                        traceback.print_exc()
                sys.stderr.flush()
            # 更新最后收消息时间（供 watchdog 检测 WebSocket 存活）
            try:
                (Path(__file__).parent / ".last_msg_ts").write_text(str(time.time()))
            except Exception:
                pass
        except Exception:
            traceback.print_exc()

    event_handler = EventDispatcherHandler.builder(FEISHU_ENCRYPT_KEY, FEISHU_VERIFICATION_TOKEN)\
        .register_p2_im_message_receive_v1(on_message_receive)\
        .build()

    return event_handler


# ══════════════════════════════════════════════════════
# Flask 应用
# ══════════════════════════════════════════════════════

app = Flask(__name__)


@app.route("/webhook", methods=["POST"])
def webhook():
    raw_request = parse_req()
    event_handler = app.config["event_handler"]
    raw_response = event_handler.do(raw_request)
    return parse_resp(raw_response)


# ══════════════════════════════════════════════════════
# 启动
# ══════════════════════════════════════════════════════

if __name__ == "__main__":
    print("=" * 50)
    print("飞书 Bot HTTP Webhook 模式")
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
        print("❌ 无法创建当月 Base，请检查飞书 API 凭证和权限")
        sys.exit(1)

    print(f"  Base: {handler._state['app_token']}")
    print(f"  表1: {handler._state['table1_id']}")
    print(f"  表2: {handler._state['table2_id']}")
    print("\n启动定时提醒 (14:55 / 23:55)...")
    threading.Thread(target=run_scheduler, args=(handler,), daemon=True).start()

    print("启动断连消息补录 (后台)...")
    threading.Thread(target=handler.backfill_messages, name="backfill", daemon=True).start()

    event_handler = create_event_handler(feishu, handler)
    app.config["event_handler"] = event_handler

    PORT = int(os.getenv("PORT", "8080"))

    # 尝试获取已有的 ngrok 隧道（避免 ERR_NGROK_334）
    # 加重试：ngrok API 可能短暂不可达（进程启动时序）
    ngrok_url = None
    import time as _time
    for attempt in range(3):
        try:
            for api_port in (4040, 4041):
                try:
                    r = requests.get(f"http://127.0.0.1:{api_port}/api/tunnels", timeout=5)
                    if r.status_code == 200:
                        tunnels = r.json().get("tunnels", [])
                        for t in tunnels:
                            if f"localhost:{PORT}" in t.get("config", {}).get("addr", ""):
                                ngrok_url = t["public_url"]
                                break
                    if ngrok_url:
                        break
                except Exception:
                    continue
        except Exception:
            pass
        if ngrok_url:
            break
        if attempt < 2:
            sys.stderr.write(f"[NGROK] 第{attempt+1}次检测未发现隧道，2秒后重试...\n")
            sys.stderr.flush()
            _time.sleep(2)

    if ngrok_url:
        print(f"\n✅ 复用已有 ngrok 隧道: {ngrok_url}")
        print("=" * 50)
        print("请确保飞书开发者后台「事件订阅」配置为：")
        print(f"  {ngrok_url}/webhook")
        print("=" * 50)
    elif ngrok:
        auth_token = os.getenv("NGROK_AUTHTOKEN", "")
        if auth_token:
            ngrok.set_auth_token(auth_token)
        # 先确认没有残留 ngrok 进程，避免 ERR_NGROK_334
        import subprocess as _sp
        has_existing_ngrok = False
        try:
            result = _sp.run(
                ["tasklist", "/FI", "IMAGENAME eq ngrok.exe", "/FO", "CSV", "/NH"],
                capture_output=True, text=True, timeout=8,
            )
            if "ngrok.exe" in result.stdout:
                has_existing_ngrok = True
        except Exception:
            pass

        if has_existing_ngrok:
            # 已有 ngrok 进程，最后尝试一次 API 检测
            sys.stderr.write("[NGROK] 检测到已有 ngrok 进程，最后一次尝试 API 检测...\n")
            sys.stderr.flush()
            for api_port in (4040, 4041):
                try:
                    r = requests.get(f"http://127.0.0.1:{api_port}/api/tunnels", timeout=5)
                    if r.status_code == 200:
                        tunnels = r.json().get("tunnels", [])
                        for t in tunnels:
                            if f"localhost:{PORT}" in t.get("config", {}).get("addr", ""):
                                ngrok_url = t["public_url"]
                                break
                    if ngrok_url:
                        break
                except Exception:
                    continue
            if not ngrok_url:
                sys.stderr.write("[NGROK] API 检测仍未发现隧道，跳过 ngrok 启动\n")
                sys.stderr.flush()
        else:
            try:
                tunnel = ngrok.connect(PORT, "http")
                ngrok_url = tunnel.public_url
                print(f"\n✅ ngrok 公网 URL: {ngrok_url}")
                print("=" * 50)
                print("请将以下 URL 配置到飞书开发者后台「事件订阅」：")
                print(f"  {ngrok_url}/webhook")
                print("=" * 50)
            except Exception as e:
                print(f"\n⚠ ngrok 启动失败: {e}")
                print(f"请手动运行 ngrok http {PORT}，将生成的 URL 配置到飞书后台")
    else:
        print("\n⚠ pyngrok 未安装，如需公网 URL 请手动启动 ngrok")
        print("  pip install pyngrok")
        print(f"  ngrok http {PORT}")

    from waitress import serve
    print(f"\n启动 Waitress 生产服务器 (0.0.0.0:{PORT})...")
    serve(app, host="0.0.0.0", port=PORT)
