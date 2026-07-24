"""登录验证、token 管理"""
import secrets
import ssl
import smtplib
from datetime import datetime, timedelta
from .db import get_db
from .encryption import encrypt, decrypt


def verify_smtp(email, smtp_pass, smtp_host="smtp.qq.com", smtp_port=465):
    """通过尝试 SMTP 连接验证邮箱和授权码"""
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context, timeout=10) as server:
            server.login(email, smtp_pass)
        return True
    except smtplib.SMTPAuthenticationError:
        return False
    except Exception:
        return False


def login(email, smtp_pass, smtp_host="smtp.qq.com", smtp_port=465):
    """登录：验证 SMTP → 创建/更新用户 → 存储加密 SMTP → 创建会话 → 返回 token"""
    if not verify_smtp(email, smtp_pass, smtp_host, smtp_port):
        return None

    db = get_db()
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")

    # 创建或更新用户
    user = db.execute("SELECT id FROM users WHERE email = ?", (email,)).fetchone()
    if user:
        db.execute("UPDATE users SET last_login_at = ? WHERE id = ?", (now, user["id"]))
    else:
        db.execute("INSERT INTO users (email, last_login_at) VALUES (?, ?)", (email, now))
        user = db.execute("SELECT id FROM users WHERE email = ?", (email,)).fetchone()

    user_id = user["id"]

    # 存储加密的 SMTP 配置
    encrypted_pass = encrypt(smtp_pass)
    existing = db.execute("SELECT user_id FROM user_smtp WHERE user_id = ?", (user_id,)).fetchone()
    if existing:
        db.execute(
            "UPDATE user_smtp SET smtp_host=?, smtp_port=?, smtp_user=?, smtp_pass_encrypted=?, updated_at=? WHERE user_id=?",
            (smtp_host, smtp_port, email, encrypted_pass, now, user_id))
    else:
        db.execute(
            "INSERT INTO user_smtp (user_id, smtp_host, smtp_port, smtp_user, smtp_pass_encrypted) VALUES (?, ?, ?, ?, ?)",
            (user_id, smtp_host, smtp_port, email, encrypted_pass))

    # 创建会话（30 天过期）
    token = secrets.token_hex(32)
    expires_at = (datetime.utcnow() + timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
    db.execute(
        "INSERT INTO sessions (user_id, token, expires_at) VALUES (?, ?, ?)",
        (user_id, token, expires_at))
    db.commit()

    return token


def validate_token(token):
    """验证 token，返回 user dict 或 None"""
    if not token:
        return None
    db = get_db()
    row = db.execute(
        "SELECT u.id, u.email FROM users u "
        "JOIN sessions s ON s.user_id = u.id "
        "WHERE s.token = ? AND s.expires_at > datetime('now')",
        (token,)
    ).fetchone()
    return dict(row) if row else None


def logout(token):
    """删除会话"""
    db = get_db()
    db.execute("DELETE FROM sessions WHERE token = ?", (token,))
    db.commit()
