"""用户 API：个人信息、SMTP 配置管理"""
import json
import ssl
import smtplib
from ..auth import validate_token
from ..db import get_db
from ..encryption import encrypt, decrypt


def handle_get_smtp_config(handler, _data):
    auth_header = handler.headers.get("Authorization", "")
    token = auth_header.replace("Bearer ", "")
    user = validate_token(token)
    if not user:
        return 401, {"success": False, "error": "未登录"}

    db = get_db()
    row = db.execute("SELECT * FROM user_smtp WHERE user_id = ?", (user["id"],)).fetchone()
    if not row:
        return 200, {"success": True, "config": None}

    return 200, {"success": True, "config": {
        "smtp_host": row["smtp_host"],
        "smtp_port": row["smtp_port"],
        "smtp_user": row["smtp_user"],
        # 密码脱敏
        "smtp_pass_masked": "****" + row["smtp_pass_encrypted"][-4:] if len(row["smtp_pass_encrypted"]) > 4 else "****",
        "from_name": row["from_name"],
        "to_addr": row["to_addr"],
        "cc_addr": row["cc_addr"],
    }}


def handle_set_smtp_config(handler, data):
    auth_header = handler.headers.get("Authorization", "")
    token = auth_header.replace("Bearer ", "")
    user = validate_token(token)
    if not user:
        return 401, {"success": False, "error": "未登录"}

    db = get_db()
    smtp_host = data.get("smtp_host", "smtp.qq.com")
    smtp_port = data.get("smtp_port", 465)
    smtp_user = data.get("smtp_user", user["email"])
    smtp_pass = data.get("smtp_pass", "")
    from_name = data.get("from_name", "元气工作站")
    to_addr = data.get("to_addr", "")
    cc_addr = data.get("cc_addr", "")

    existing = db.execute("SELECT user_id FROM user_smtp WHERE user_id = ?", (user["id"],)).fetchone()

    if smtp_pass and smtp_pass != "****":
        encrypted_pass = encrypt(smtp_pass)
    elif existing:
        # 保留旧密码
        encrypted_pass = db.execute(
            "SELECT smtp_pass_encrypted FROM user_smtp WHERE user_id = ?", (user["id"],)
        ).fetchone()["smtp_pass_encrypted"]
    else:
        return 400, {"success": False, "error": "请提供 SMTP 授权码"}

    if existing:
        db.execute(
            "UPDATE user_smtp SET smtp_host=?, smtp_port=?, smtp_user=?, smtp_pass_encrypted=?, "
            "from_name=?, to_addr=?, cc_addr=?, updated_at=CURRENT_TIMESTAMP WHERE user_id=?",
            (smtp_host, smtp_port, smtp_user, encrypted_pass, from_name, to_addr, cc_addr, user["id"]))
    else:
        db.execute(
            "INSERT INTO user_smtp (user_id, smtp_host, smtp_port, smtp_user, smtp_pass_encrypted, "
            "from_name, to_addr, cc_addr) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (user["id"], smtp_host, smtp_port, smtp_user, encrypted_pass, from_name, to_addr, cc_addr))
    db.commit()

    return 200, {"success": True, "message": "SMTP 配置已更新"}


def handle_test_smtp(handler, _data):
    auth_header = handler.headers.get("Authorization", "")
    token = auth_header.replace("Bearer ", "")
    user = validate_token(token)
    if not user:
        return 401, {"success": False, "error": "未登录"}

    db = get_db()
    row = db.execute("SELECT * FROM user_smtp WHERE user_id = ?", (user["id"],)).fetchone()
    if not row:
        return 400, {"success": False, "error": "未配置 SMTP"}

    try:
        smtp_pass = decrypt(row["smtp_pass_encrypted"])
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(row["smtp_host"], row["smtp_port"], context=context, timeout=10) as server:
            server.login(row["smtp_user"], smtp_pass)
        return 200, {"success": True, "message": "SMTP 连接测试成功"}
    except Exception as e:
        return 500, {"success": False, "error": f"SMTP 连接失败：{e}"}
