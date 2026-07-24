"""导出 API：月度任务量 Excel、月考成绩 Excel、月考打字截图 Word"""
import json
from email.header import Header
from ..auth import validate_token
from ..excel_generator import generate_excel, generate_exam_excel
from ..word_generator import generate_exam_word


def _get_user(handler):
    auth_header = handler.headers.get("Authorization", "")
    token = auth_header.replace("Bearer ", "")
    return validate_token(token)


def handle_export_excel(handler, data):
    user = _get_user(handler)
    if not user:
        return 401, {"success": False, "error": "未登录"}

    groups = data.get("groups", {})
    month_str = data.get("month", "")
    filename = data.get("filename", f"在线客服合计任务量统计表（{month_str}）.xlsx")

    excel_data = generate_excel(groups, month_str)

    handler.send_response(200)
    handler.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    handler.send_header("Content-Disposition", f'attachment; filename="{Header(filename, "utf-8").encode()}"')
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.send_header("Access-Control-Expose-Headers", "Content-Disposition")
    handler.end_headers()
    handler.wfile.write(excel_data.getvalue())
    return None  # 信号：响应已直接写入，不需要 JSON 响应


def handle_export_exam_excel(handler, data):
    user = _get_user(handler)
    if not user:
        return 401, {"success": False, "error": "未登录"}

    records = data.get("records", [])
    month_str = data.get("month", "")
    filename = data.get("filename", f"{month_str}在线客服月考成绩.xlsx")

    excel_data = generate_exam_excel(records, month_str)

    handler.send_response(200)
    handler.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    handler.send_header("Content-Disposition", f'attachment; filename="{Header(filename, "utf-8").encode()}"')
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.send_header("Access-Control-Expose-Headers", "Content-Disposition")
    handler.end_headers()
    handler.wfile.write(excel_data.getvalue())
    return None


def handle_export_exam_word(handler, data):
    user = _get_user(handler)
    if not user:
        return 401, {"success": False, "error": "未登录"}

    records = data.get("records", [])
    month_str = data.get("month", "")
    images = data.get("images", {})
    filename = data.get("filename", f"{month_str}在线客服月考打字截图.docx")

    word_data = generate_exam_word(records, month_str, images)

    handler.send_response(200)
    handler.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    handler.send_header("Content-Disposition", f'attachment; filename="{Header(filename, "utf-8").encode()}"')
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.send_header("Access-Control-Expose-Headers", "Content-Disposition")
    handler.end_headers()
    handler.wfile.write(word_data.getvalue())
    return None
