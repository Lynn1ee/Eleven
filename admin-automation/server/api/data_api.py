"""数据同步 API：通用键值存储 CRUD，支持 shared/user scope"""
import json
from ..auth import validate_token
from ..db import get_db


def _get_user_id(handler):
    auth_header = handler.headers.get("Authorization", "")
    token = auth_header.replace("Bearer ", "")
    user = validate_token(token)
    return user["id"] if user else None


def handle_get_data(handler, key):
    user_id = _get_user_id(handler)
    if not user_id:
        return 401, {"success": False, "error": "未登录"}

    db = get_db()
    # 先查个人数据
    row = db.execute(
        "SELECT data_value FROM app_data WHERE user_id = ? AND data_key = ?",
        (user_id, key)).fetchone()
    # 再查共享数据
    if not row:
        row = db.execute(
            "SELECT data_value FROM app_data WHERE user_id IS NULL AND data_key = ?",
            (key,)).fetchone()

    if row:
        return 200, {"success": True, "data": json.loads(row["data_value"])}
    else:
        return 200, {"success": True, "data": None}


def handle_set_data(handler, key, data):
    user_id = _get_user_id(handler)
    if not user_id:
        return 401, {"success": False, "error": "未登录"}

    value = data.get("value")
    scope = data.get("scope", "shared")

    db = get_db()
    db_user_id = None if scope == "shared" else user_id

    existing = db.execute(
        "SELECT id FROM app_data WHERE user_id IS ? AND data_key = ?",
        (db_user_id, key)).fetchone()

    json_value = json.dumps(value, ensure_ascii=False)
    if existing:
        db.execute(
            "UPDATE app_data SET data_value = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            (json_value, existing["id"]))
    else:
        db.execute(
            "INSERT INTO app_data (user_id, data_key, data_value) VALUES (?, ?, ?)",
            (db_user_id, key, json_value))
    db.commit()
    return 200, {"success": True}


def handle_delete_data(handler, key):
    user_id = _get_user_id(handler)
    if not user_id:
        return 401, {"success": False, "error": "未登录"}

    db = get_db()
    db.execute(
        "DELETE FROM app_data WHERE user_id = ? AND data_key = ?",
        (user_id, key))
    db.commit()
    return 200, {"success": True}
