"""邮件发送：使用用户 SMTP 配置发送邮件"""
import re
import ssl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from email.utils import formataddr
from .db import get_db
from .encryption import decrypt


def _parse_addrs(addr_str):
    """解析 Name <email> 或纯 email，返回 [(name, email), ...]"""
    if not addr_str:
        return []
    result = []
    for part in addr_str.split(','):
        part = part.strip()
        if not part:
            continue
        m = re.match(r'^(.*?)\s*<([^>]+)>\s*$', part)
        if m:
            result.append((m.group(1).strip(), m.group(2).strip()))
        else:
            result.append(('', part))
    return result


def get_user_smtp_config(user_id):
    """获取用户 SMTP 配置（解密密码）"""
    db = get_db()
    row = db.execute("SELECT * FROM user_smtp WHERE user_id = ?", (user_id,)).fetchone()
    if not row:
        return None
    return {
        "smtp_host": row["smtp_host"],
        "smtp_port": row["smtp_port"],
        "smtp_user": row["smtp_user"],
        "smtp_pass": decrypt(row["smtp_pass_encrypted"]),
        "from_name": row["from_name"],
        "to_addr": row["to_addr"],
        "cc_addr": row["cc_addr"],
    }


def send_email(user_id, to_addr, cc_addr, subject, body, attachment_data, attachment_filename, from_name=''):
    """使用用户 SMTP 配置发送邮件"""
    config = get_user_smtp_config(user_id)
    if not config:
        raise Exception("未配置 SMTP，请重新登录")

    # 解析收件人/抄送的 Name <email> 格式
    to_entries = _parse_addrs(to_addr)
    cc_entries = _parse_addrs(cc_addr)

    sender_name = from_name or config['from_name']
    msg = MIMEMultipart()
    msg["From"] = formataddr((sender_name, config['smtp_user']))
    msg["To"] = ", ".join(formataddr((n, e)) if n else e for n, e in to_entries) if to_entries else to_addr
    if cc_entries:
        msg["Cc"] = ", ".join(formataddr((n, e)) if n else e for n, e in cc_entries)
    msg["Subject"] = Header(subject, "utf-8")
    import html as _html
    safe_body = _html.escape(body)
    safe_body = safe_body.replace('\r\n', '<br>').replace('\n', '<br>')
    html_body = f"""<html><body style="font-family: '微软雅黑', 'Microsoft YaHei', sans-serif;">{safe_body}</body></html>"""
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    if attachment_data is not None:
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(attachment_data.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{Header(attachment_filename, "utf-8").encode()}"')
        msg.attach(part)

    all_recipients = [e for _, e in to_entries]
    if cc_entries:
        all_recipients += [e for _, e in cc_entries]

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(config["smtp_host"], config["smtp_port"], context=context, timeout=15) as server:
        server.login(config["smtp_user"], config["smtp_pass"])
        server.sendmail(config["smtp_user"], all_recipients, msg.as_string())

    # 记录发送历史
    db = get_db()
    db.execute(
        "INSERT INTO email_history (user_id, subject, to_addr, cc_addr, from_addr, status) VALUES (?, ?, ?, ?, ?, 'sent')",
        (0, subject, to_addr, cc_addr or "", config["smtp_user"]))
    db.commit()
