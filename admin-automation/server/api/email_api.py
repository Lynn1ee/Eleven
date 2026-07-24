"""邮件 API：发送邮件、发送历史、模板管理"""
import json
import smtplib
from datetime import datetime
from urllib.parse import urlparse, parse_qs
from ..auth import validate_token
from ..email_sender import send_email, get_user_smtp_config
from ..excel_generator import generate_excel, generate_exam_excel
from ..db import get_db


def _get_user(handler):
    auth_header = handler.headers.get("Authorization", "")
    token = auth_header.replace("Bearer ", "")
    return validate_token(token)


def _get_app_data(user_id, key):
    """从 app_data 表读取 JSON 数据"""
    db = get_db()
    row = db.execute(
        "SELECT data_value FROM app_data WHERE data_key = ? AND (user_id = ? OR user_id IS NULL) ORDER BY user_id DESC LIMIT 1",
        (key, user_id)).fetchone()
    if row:
        return json.loads(row["data_value"])
    return None


def _monthly_history_to_groups(history, month_str):
    """将 monthly_history 中的月份数据转为 generate_excel 需要的 groups 格式"""
    if not history or month_str not in history:
        return None
    groups_data = history[month_str]  # 三个 group 的数组：专职分销、专职大夜、在线客服
    group_keys = ["fenxiao", "daye", "kefu"]
    support_fields = ["waibu", "waibu", "jiaodai"]
    result = {}
    for i, g in enumerate(groups_data):
        key = group_keys[i]
        sup_field = support_fields[i]
        result[key] = []
        for row in (g.get("rows") or []):
            result[key].append({
                "id": row.get("id", ""),
                "name": row.get("name", ""),
                "online": row.get("online", 0),
                "flight": row.get("flight", 0),
                "train": row.get("train", 0),
                "im": row.get("im", 0),
                "support": row.get(sup_field, 0),
            })
    return result


def _auto_detect_month(user_id, data_source):
    """自动识别数据源的最新月份"""
    if data_source == "monthly":
        history = _get_app_data(user_id, "monthly_history")
        if history:
            keys = sorted(history.keys(), reverse=True)
            return keys[0] if keys else None
    elif data_source == "exam":
        return _get_app_data(user_id, "exam_month")
    return None


def handle_send_email(handler, data):
    user = _get_user(handler)
    if not user:
        return 401, {"success": False, "error": "未登录"}

    to_addr = data.get("to", "").strip()
    cc_addr = data.get("cc", "").strip()
    from_name = data.get("from_name", "").strip()
    subject = data.get("subject", "")
    body = data.get("body", "")
    month_str = data.get("month", "")
    data_source = data.get("data_source", "")
    groups = data.get("groups", {})
    records = data.get("records", [])

    if not to_addr:
        return 400, {"success": False, "error": "缺少收件人地址"}

    # 自动识别月份（data_source 有值时，优先用传入的 month，其次自动检测）
    if data_source and not month_str:
        month_str = _auto_detect_month(user["id"], data_source) or ""

    # 如果主题/正文包含 {month} 变量但月份仍为空，尝试从所有数据源自动检测
    if not month_str and ('{month}' in subject or '{month}' in body):
        for src in ("monthly", "exam"):
            month_str = _auto_detect_month(user["id"], src)
            if month_str:
                break
        month_str = month_str or ""

    # {month} 变量替换
    if month_str:
        subject = subject.replace("{month}", month_str)
        body = body.replace("{month}", month_str)

    # 附件生成：优先用直接传入的 groups/records，其次根据 data_source 查数据库
    attachment_data = None
    filename = None
    if groups:
        filename = f"在线客服合计任务量统计表（{month_str}）.xlsx"
        attachment_data = generate_excel(groups, month_str)
    elif records:
        filename = f"{month_str}在线客服月考成绩.xlsx"
        attachment_data = generate_exam_excel(records, month_str)
    elif data_source == "monthly":
        history = _get_app_data(user["id"], "monthly_history")
        groups = _monthly_history_to_groups(history, month_str) if history else None
        if groups:
            filename = f"在线客服合计任务量统计表（{month_str}）.xlsx"
            attachment_data = generate_excel(groups, month_str)
    elif data_source == "exam":
        saved = _get_app_data(user["id"], "exam_data")
        if saved:
            filename = f"{month_str}在线客服月考成绩.xlsx"
            attachment_data = generate_exam_excel(saved, month_str)

    # 默认主题/正文
    if not subject and month_str:
        subject = f"在线客服合计任务量统计表（{month_str}）"
    if not body and month_str:
        body = f"您好，\n\n附件为 {month_str} 在线客服合计任务量统计表，请查收。\n\n—— 元气工作站自动发送"

    try:
        send_email(user["id"], to_addr, cc_addr, subject, body, attachment_data, filename, from_name)
        return 200, {"success": True, "message": "邮件已发送"}
    except smtplib.SMTPAuthenticationError:
        return 500, {"success": False, "error": "SMTP 认证失败，请检查邮箱和授权码"}
    except smtplib.SMTPException as e:
        return 500, {"success": False, "error": f"SMTP 错误：{e}"}
    except Exception as e:
        return 500, {"success": False, "error": str(e)}


def handle_email_history(handler, _data):
    user = _get_user(handler)
    if not user:
        return 401, {"success": False, "error": "未登录"}

    db = get_db()
    rows = db.execute(
        "SELECT subject, to_addr, cc_addr, from_addr, status, error_message, sent_at "
        "FROM email_history ORDER BY sent_at DESC LIMIT 50").fetchall()
    return 200, {"success": True, "history": [dict(r) for r in rows]}


def handle_list_templates(handler, _data):
    user = _get_user(handler)
    if not user:
        return 401, {"success": False, "error": "未登录"}
    db = get_db()
    rows = db.execute(
        "SELECT id, name, data_source, to_addrs, cc_addrs, send_type, subject, body, created_at, updated_at "
        "FROM email_templates ORDER BY updated_at DESC").fetchall()
    return 200, {"success": True, "templates": [dict(r) for r in rows]}


def handle_create_template(handler, data):
    user = _get_user(handler)
    if not user:
        return 401, {"success": False, "error": "未登录"}
    name = data.get("name", "").strip()
    if not name:
        return 400, {"success": False, "error": "模板名称不能为空"}
    db = get_db()
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    db.execute(
        "INSERT INTO email_templates (user_id, name, data_source, to_addrs, cc_addrs, send_type, subject, body, created_at, updated_at) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (0, name, data.get("data_source", ""), data.get("to_addrs", ""),
         data.get("cc_addrs", ""), data.get("send_type", "manual"),
         data.get("subject", ""), data.get("body", ""), now, now))
    db.commit()
    return 200, {"success": True, "message": "模板已创建"}


def handle_update_template(handler, data):
    user = _get_user(handler)
    if not user:
        return 401, {"success": False, "error": "未登录"}
    parsed = urlparse(handler.path)
    template_id = parse_qs(parsed.query).get("id", [None])[0]
    if not template_id:
        return 400, {"success": False, "error": "缺少模板 ID"}
    db = get_db()
    row = db.execute("SELECT id FROM email_templates WHERE id = ?",
                     (template_id,)).fetchone()
    if not row:
        return 404, {"success": False, "error": "模板不存在"}
    now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    db.execute(
        "UPDATE email_templates SET name=?, data_source=?, to_addrs=?, cc_addrs=?, send_type=?, subject=?, body=?, updated_at=? "
        "WHERE id=?",
        (data.get("name", ""), data.get("data_source", ""), data.get("to_addrs", ""),
         data.get("cc_addrs", ""), data.get("send_type", "manual"),
         data.get("subject", ""), data.get("body", ""), now,
         template_id))
    db.commit()
    return 200, {"success": True, "message": "模板已更新"}


def handle_delete_template(handler, _data):
    user = _get_user(handler)
    if not user:
        return 401, {"success": False, "error": "未登录"}
    parsed = urlparse(handler.path)
    template_id = parse_qs(parsed.query).get("id", [None])[0]
    if not template_id:
        return 400, {"success": False, "error": "缺少模板 ID"}
    db = get_db()
    db.execute("DELETE FROM email_templates WHERE id = ?",
               (template_id,))
    db.commit()
    return 200, {"success": True, "message": "模板已删除"}


def handle_confirm_and_send(handler, data):
    """数据完成后自动按模板发送：查找 auto 模板 → 抓取数据 → 逐一发送"""
    user = _get_user(handler)
    if not user:
        return 401, {"success": False, "error": "未登录"}

    data_source = data.get("data_source", "").strip()
    if not data_source:
        return 400, {"success": False, "error": "缺少 data_source 参数"}

    # 检查 SMTP 配置
    config = get_user_smtp_config(user["id"])
    if not config:
        return 400, {"success": False, "error": "未配置 SMTP，请在邮件管理中设置"}

    # 查找匹配的 auto 模板
    db = get_db()
    templates = db.execute(
        "SELECT * FROM email_templates WHERE send_type = 'auto' AND data_source = ?",
        (data_source,)).fetchall()
    if not templates:
        source_label = {"monthly": "月度任务量统计", "exam": "月考成绩"}.get(data_source, data_source)
        return 400, {"success": False, "error": f"没有匹配的自动发送模板，请先在邮件管理中为「{source_label}」创建 auto 类型模板"}

    # 优先使用前端传入的月份，未传则自动识别最新月份
    month_str = data.get("month", "").strip()
    if not month_str:
        month_str = _auto_detect_month(user["id"], data_source)
    if not month_str:
        return 400, {"success": False, "error": "未找到已保存的数据，请先在数据统计中生成报表"}

    # 从数据库抓取数据生成附件
    attachment_data = None
    filename = None
    if data_source == "monthly":
        history = _get_app_data(user["id"], "monthly_history")
        groups = _monthly_history_to_groups(history, month_str) if history else None
        if not groups:
            return 400, {"success": False, "error": f"未找到 {month_str} 的月度任务量数据"}
        filename = f"在线客服合计任务量统计表（{month_str}）.xlsx"
        attachment_data = generate_excel(groups, month_str)
    elif data_source == "exam":
        records = _get_app_data(user["id"], "exam_data")
        if not records:
            return 400, {"success": False, "error": f"未找到 {month_str} 的月考成绩数据"}
        filename = f"{month_str}在线客服月考成绩.xlsx"
        attachment_data = generate_exam_excel(records, month_str)
    else:
        return 400, {"success": False, "error": f"不支持的数据来源：{data_source}"}

    # 逐一按照模板发送
    results = []
    for t in templates:
        t = dict(t)
        to_addr = t.get("to_addrs", "")
        # 发件人姓名和默认抄送
        global_email = _get_app_data(user["id"], "email_config") or {}
        from_name = global_email.get("from_name", "") or config.get("from_name", "")
        # 合并模板抄送 + 全局默认抄送
        t_cc = t.get("cc_addrs", "")
        g_cc = global_email.get("default_cc", "")
        cc_parts = [p.strip() for p in (t_cc + "," + g_cc).split(",") if p.strip()]
        cc_addr = ",".join(dict.fromkeys(cc_parts))  # 去重保留顺序
        subject = t.get("subject", "")
        body = t.get("body", "")

        if not to_addr:
            results.append({"template_id": t["id"], "template_name": t["name"], "status": "skipped", "reason": "模板未设置收件人"})
            continue

        # {month} 变量替换
        subject = subject.replace("{month}", month_str)
        body = body.replace("{month}", month_str)
        if not subject:
            subject = filename
        if not body:
            body = f"您好，\n\n附件为 {month_str} 数据报表，请查收。\n\n—— 元气工作站自动发送"

        try:
            send_email(user["id"], to_addr, cc_addr, subject, body, attachment_data, filename, from_name)
            results.append({"template_id": t["id"], "template_name": t["name"], "status": "sent", "to": to_addr})
        except smtplib.SMTPAuthenticationError:
            results.append({"template_id": t["id"], "template_name": t["name"], "status": "failed", "reason": "SMTP 认证失败"})
        except Exception as e:
            results.append({"template_id": t["id"], "template_name": t["name"], "status": "failed", "reason": str(e)})

    sent_count = sum(1 for r in results if r["status"] == "sent")
    return 200, {"success": True, "sent": sent_count, "total": len(results), "results": results, "month": month_str}
