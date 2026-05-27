"""飞书 OpenAPI 客户端：Token、消息、多维表格、用户信息"""

import os
import time
import requests


class FeishuClient:
    def __init__(self, app_id: str, app_secret: str):
        self.app_id = app_id
        self.app_secret = app_secret
        self._token = None
        self._token_expire = 0

    # ── Token ────────────────────────────────────────────

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
            raise Exception(f"获取 tenant token 失败: {data}")
        self._token = data["tenant_access_token"]
        self._token_expire = now + data.get("expire", 7200)
        return self._token

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self._get_tenant_token()}",
            "Content-Type": "application/json; charset=utf-8",
        }

    # ── 用户信息 ─────────────────────────────────────────

    def get_user_name(self, open_id: str) -> str:
        """根据 open_id 获取用户姓名。"""
        resp = requests.get(
            f"https://open.feishu.cn/open-apis/contact/v3/users/{open_id}",
            headers=self._headers(),
            timeout=10,
        )
        data = resp.json()
        if data.get("code") == 0:
            user = data.get("data", {}).get("user", {})
            # 飞书 API 可能返回 name / en_name / nickname 等不同字段
            return user.get("name") or user.get("en_name") or user.get("nickname") or user.get("email") or "未知用户"
        return "未知用户"

    # ── 消息发送 ─────────────────────────────────────────

    def send_message(self, receive_id_type: str, receive_id: str, content: str,
                     msg_type: str = "interactive") -> dict:
        """
        发送消息到飞书会话。
        receive_id_type: "chat_id" / "open_id" / "user_id"
        receive_id: 对应的 ID
        """
        if msg_type == "text":
            body = {"receive_id": receive_id, "msg_type": "text",
                    "content": f'{{"text":"{content}"}}'}
        elif msg_type == "interactive":
            body = {"receive_id": receive_id, "msg_type": "interactive",
                    "content": content}
        else:
            body = {"receive_id": receive_id, "msg_type": msg_type,
                    "content": content}

        resp = requests.post(
            f"https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type={receive_id_type}",
            headers=self._headers(), json=body, timeout=10,
        )
        return resp.json()

    def send_image(self, receive_id_type: str, receive_id: str,
                   image_key: str) -> dict:
        """发送图片消息。"""
        body = {
            "receive_id": receive_id,
            "msg_type": "image",
            "content": f'{{"image_key":"{image_key}"}}',
        }
        resp = requests.post(
            f"https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type={receive_id_type}",
            headers=self._headers(), json=body, timeout=10,
        )
        return resp.json()

    def upload_image(self, file_path: str) -> str:
        """上传图片，返回 image_key。"""
        import io
        resp = requests.post(
            "https://open.feishu.cn/open-apis/im/v1/images",
            headers={"Authorization": f"Bearer {self._get_tenant_token()}"},
            files={"image_type": "message",
                   "image": (os.path.basename(file_path),
                             open(file_path, "rb"), "image/png")},
            timeout=30,
        )
        data = resp.json()
        if data.get("code") == 0:
            return data["data"]["image_key"]
        raise Exception(f"上传图片失败: {data}")

    # ── 多维表格 (Base) 操作 ─────────────────────────────

    def create_base(self, name: str, folder_token: str = "") -> dict:
        """创建多维表格 Base。"""
        body = {"name": name}
        if folder_token:
            body["folder_token"] = folder_token
        resp = requests.post(
            "https://open.feishu.cn/open-apis/bitable/v1/apps",
            headers=self._headers(), json=body, timeout=10,
        )
        return resp.json()

    def list_tables(self, app_token: str) -> list:
        """列出 Base 中所有表。"""
        resp = requests.get(
            f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables",
            headers=self._headers(), timeout=10,
        )
        data = resp.json()
        if data.get("code") == 0:
            return data["data"].get("items", [])
        return []

    def create_table(self, app_token: str, table_name: str, fields: list) -> dict:
        """
        在 Base 中创建表。
        fields: [{"field_name": "...", "type": 1}, ...]
        类型: 1=文本, 2=数字, 5=日期
        """
        body = {"table": {"name": table_name},
                "fields": fields}
        resp = requests.post(
            f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables",
            headers=self._headers(), json=body, timeout=10,
        )
        return resp.json()

    def add_records(self, app_token: str, table_id: str, records: list) -> dict:
        """
        批量添加记录。
        records: [{"fields": {"字段名": 值, ...}}, ...]
        """
        body = {"records": records}
        resp = requests.post(
            f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_create",
            headers=self._headers(), json=body, timeout=15,
        )
        return resp.json()

    def search_records(self, app_token: str, table_id: str, query: dict = None) -> list:
        """搜索记录。"""
        all_records = []
        page_token = ""
        while True:
            params = {"page_size": "100"}
            if page_token:
                params["page_token"] = page_token
            resp = requests.get(
                f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records",
                headers=self._headers(), params=params, timeout=15,
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

    def update_record(self, app_token: str, table_id: str, record_id: str,
                      fields: dict) -> dict:
        """更新单条记录。"""
        resp = requests.put(
            f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/{record_id}",
            headers=self._headers(), json={"fields": fields}, timeout=10,
        )
        return resp.json()

    def delete_records(self, app_token: str, table_id: str, record_ids: list) -> dict:
        """批量删除记录。"""
        body = {"records": record_ids}
        resp = requests.post(
            f"https://open.feishu.cn/open-apis/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_delete",
            headers=self._headers(), json=body, timeout=15,
        )
        return resp.json()
