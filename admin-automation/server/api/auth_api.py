"""认证 API 端点：登录、登出、用户信息"""
import json
from urllib.parse import urlparse, parse_qs
from ..auth import login, logout, validate_token


def handle_login(handler, data):
    email = (data.get("email") or "").strip()
    smtp_pass = (data.get("smtp_pass") or "").strip()
    smtp_host = data.get("smtp_host", "smtp.qq.com")
    smtp_port = data.get("smtp_port", 465)

    if not email or not smtp_pass:
        return 400, {"success": False, "error": "请输入邮箱和 SMTP 授权码"}

    token = login(email, smtp_pass, smtp_host, smtp_port)
    if token:
        return 200, {"success": True, "token": token, "email": email}
    else:
        return 401, {"success": False, "error": "SMTP 验证失败，请检查邮箱和授权码"}


def handle_logout(handler, data):
    auth_header = handler.headers.get("Authorization", "")
    token = auth_header.replace("Bearer ", "")
    if token:
        logout(token)
    return 200, {"success": True}


def handle_profile(handler, _data):
    auth_header = handler.headers.get("Authorization", "")
    token = auth_header.replace("Bearer ", "")
    user = validate_token(token)
    if not user:
        return 401, {"success": False, "error": "未登录或会话已过期"}
    return 200, {"success": True, "email": user["email"]}
