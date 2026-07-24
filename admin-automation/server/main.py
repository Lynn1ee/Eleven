"""元气工作站 — 多用户 HTTPS 服务器"""
import http.server
import json
import os
import ssl
from socketserver import ThreadingMixIn
from urllib.parse import urlparse, parse_qs

from .api.auth_api import handle_login, handle_logout, handle_profile
from .api.data_api import handle_get_data, handle_set_data, handle_delete_data
from .api.export_api import handle_export_excel, handle_export_exam_excel, handle_export_exam_word
from .api.email_api import handle_send_email, handle_email_history, handle_list_templates, handle_create_template, handle_update_template, handle_delete_template, handle_confirm_and_send
from .holiday import handle_work_hours
from .api.user_api import handle_get_smtp_config, handle_set_smtp_config, handle_test_smtp
from .api.invoice_api import handle_list_cycles, handle_create_cycle, handle_delete_cycle
from .api.invoice_api import handle_upload_invoice, handle_list_records, handle_update_record, handle_delete_record
from .api.invoice_api import handle_export_summary, handle_verify_record, handle_verify_upload
from .api.performance_api import (
    handle_import, handle_history, handle_ranking, handle_confirm_rank,
    handle_get_staff_ranks, handle_update_staff_ranks, handle_yearly_history, handle_notice,
    handle_import_file, handle_import_history, handle_set_status_note, handle_get_status_reasons,
    handle_get_exclusions, handle_add_exclusion, handle_remove_exclusion
)

PORT = 8899
HTTP_PORT = 8898
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INVOICE_DIR = os.path.join(BASE_DIR, "data", "invoices")


def _get_lan_ip():
    try:
        import socket
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'


class RequestHandler(http.server.BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        # 首页 — 直接 serve HTML
        if path == "/" or path == "/index.html":
            self._serve_html()
            return
        if path == "/login.html":
            self._serve_login_html()
            return
        if path == "/invoice-upload.html":
            self._serve_invoice_upload_html()
            return
        if path == "/print-invoice.html":
            self._serve_print_invoice_html()
            return
        if path == "/xlsx.full.min.js":
            self._serve_xlsx_js()
            return

        try:
            if path == "/api/server-info":
                result = 200, {"success": True, "lan_ip": _get_lan_ip(), "http_port": HTTP_PORT}
            elif path == "/api/user/profile":
                result = handle_profile(self, {})
            elif path == "/api/user/smtp-config":
                result = handle_get_smtp_config(self, {})
            elif path == "/api/email-history":
                result = handle_email_history(self, {})
            elif path == "/api/email-templates":
                result = handle_list_templates(self, {})
            elif path == "/api/work-hours":
                result = handle_work_hours(self)
            elif path == "/api/invoice/cycles":
                result = handle_list_cycles(self, {})
            elif path == "/api/invoice/records":
                qs = {k: v[0] for k, v in parse_qs(parsed.query).items()}
                result = handle_list_records(self, qs)
            elif path == "/api/invoice-file":
                result = self._serve_invoice_file(parsed.query)
            elif path == "/api/invoice-png":
                result = self._serve_invoice_png(parsed.query)
            elif path == "/api/performance/history":
                result = handle_history(self, {})
            elif path == "/api/performance/ranking":
                qs = {k: v[0] for k, v in parse_qs(parsed.query).items()}
                result = handle_ranking(self, qs)
            elif path == "/api/performance/yearly":
                qs = {k: v[0] for k, v in parse_qs(parsed.query).items()}
                result = handle_yearly_history(self, qs)
            elif path == "/api/performance/notice":
                qs = {k: v[0] for k, v in parse_qs(parsed.query).items()}
                result = handle_notice(self, qs)
            elif path == "/api/staff/ranks":
                result = handle_get_staff_ranks(self, {})
            elif path == "/api/performance/status-reasons":
                result = handle_get_status_reasons(self, {})
            elif path == "/api/performance/exclusions":
                qs = {k: v[0] for k, v in parse_qs(parsed.query).items()}
                result = handle_get_exclusions(self, qs)
            elif path.startswith("/api/data/"):
                key = path[len("/api/data/"):]
                result = handle_get_data(self, key)
            else:
                result = (404, {"success": False, "error": "未知路径"})
        except Exception as e:
            result = (500, {"success": False, "error": str(e)})

        if result is not None:
            self._respond(*result)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        try:
            # multipart/form-data 路由（需在 JSON 解析前处理）
            if path == "/api/invoice/upload":
                result = self._handle_invoice_upload()
                if result is not None:
                    self._respond(*result)
                return
            if path == "/api/invoice/verify-upload":
                result = self._handle_verify_upload()
                if result is not None:
                    self._respond(*result)
                return
            if path == "/api/performance/import-file":
                result = self._handle_promo_upload()
                if result is not None:
                    self._respond(*result)
                return
            if path == "/api/performance/import-history":
                result = self._handle_promo_history_upload()
                if result is not None:
                    self._respond(*result)
                return

            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length > 0 else b"{}"
            data = json.loads(raw.decode("utf-8")) if raw else {}

            if path == "/api/login":
                result = handle_login(self, data)
            elif path == "/api/logout":
                result = handle_logout(self, data)
            elif path == "/api/send-email":
                result = handle_send_email(self, data)
            elif path == "/api/confirm-and-send":
                result = handle_confirm_and_send(self, data)
            elif path == "/api/email-templates":
                result = handle_create_template(self, data)
            elif path == "/api/export-excel":
                result = handle_export_excel(self, data)
            elif path == "/api/export-exam-excel":
                result = handle_export_exam_excel(self, data)
            elif path == "/api/export-exam-word":
                result = handle_export_exam_word(self, data)
            elif path == "/api/user/smtp-config":
                result = handle_set_smtp_config(self, data)
            elif path == "/api/user/smtp-test":
                result = handle_test_smtp(self, data)
            elif path == "/api/invoice/cycles":
                result = handle_create_cycle(self, data)
            elif path == "/api/invoice/records":
                result = handle_update_record(self, data)
            elif path == "/api/invoice/verify":
                result = handle_verify_record(self, data)
            elif path == "/api/invoice/export":
                result = handle_export_summary(self, data)
            elif path == "/api/performance/import":
                result = handle_import(self, data)
            elif path == "/api/performance/confirm-rank":
                result = handle_confirm_rank(self, data)
            elif path == "/api/performance/status-note":
                result = handle_set_status_note(self, data)
            elif path == "/api/performance/exclusions/add":
                result = handle_add_exclusion(self, data)
            elif path == "/api/performance/exclusions/remove":
                result = handle_remove_exclusion(self, data)
            elif path == "/api/staff/ranks":
                result = handle_update_staff_ranks(self, data)
            elif path.startswith("/api/data/"):
                key = path[len("/api/data/"):]
                result = handle_set_data(self, key, data)
            elif path == "/export-excel":
                result = self._handle_legacy_export(data)
            elif path == "/send-email":
                result = self._handle_legacy_send(data)
            elif path == "/export-exam-excel":
                result = self._handle_legacy_exam_export(data)
            elif path == "/export-exam-word":
                result = self._handle_legacy_exam_word(data)
            else:
                result = (404, {"success": False, "error": "未知路径"})
        except Exception as e:
            result = (500, {"success": False, "error": str(e)})

        if result is not None:
            self._respond(*result)

    def do_DELETE(self):
        parsed = urlparse(self.path)
        path = parsed.path

        try:
            # 读取 DELETE 请求体（如果有）
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length > 0 else b"{}"
            data = json.loads(raw.decode("utf-8")) if raw else {}

            if path.startswith("/api/data/"):
                key = path[len("/api/data/"):]
                result = handle_delete_data(self, key)
            elif path == "/api/email-templates":
                result = handle_delete_template(self, {})
            elif path == "/api/invoice/cycles":
                result = handle_delete_cycle(self, data)
            elif path == "/api/invoice/records":
                result = handle_delete_record(self, data)
            else:
                result = (404, {"success": False, "error": "未知路径"})
        except Exception as e:
            result = (500, {"success": False, "error": str(e)})

        if result is not None:
            self._respond(*result)

    def do_PUT(self):
        parsed = urlparse(self.path)
        path = parsed.path

        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length > 0 else b"{}"
            data = json.loads(raw.decode("utf-8")) if raw else {}

            if path == "/api/email-templates":
                result = handle_update_template(self, data)
            else:
                result = (404, {"success": False, "error": "未知路径"})
        except Exception as e:
            result = (500, {"success": False, "error": str(e)})

        if result is not None:
            self._respond(*result)

    # ── 静态文件服务 ──

    def _serve_html(self):
        html_path = os.path.join(BASE_DIR, "style-preview.html")
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self._set_no_cache()
            self.end_headers()
            self.wfile.write(html.encode("utf-8"))
        except FileNotFoundError:
            self._respond(404, {"success": False, "error": "HTML file not found"})

    def _serve_login_html(self):
        html_path = os.path.join(BASE_DIR, "login.html")
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self._set_no_cache()
            self.end_headers()
            self.wfile.write(html.encode("utf-8"))
        except FileNotFoundError:
            self._respond(404, {"success": False, "error": "Login page not found"})

    def _serve_invoice_upload_html(self):
        html_path = os.path.join(BASE_DIR, "invoice-upload.html")
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self._set_no_cache()
            self.end_headers()
            self.wfile.write(html.encode("utf-8"))
        except FileNotFoundError:
            self._respond(404, {"success": False, "error": "上传页面未找到"})

    def _serve_print_invoice_html(self):
        html_path = os.path.join(BASE_DIR, "print-invoice.html")
        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self._set_no_cache()
            self.end_headers()
            self.wfile.write(html.encode("utf-8"))
        except FileNotFoundError:
            self._respond(404, {"success": False, "error": "打印模版页面未找到"})

    def _serve_xlsx_js(self):
        js_path = os.path.join(BASE_DIR, "xlsx.full.min.js")
        try:
            with open(js_path, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/javascript")
            self._set_no_cache()
            self.end_headers()
            self.wfile.write(data)
        except FileNotFoundError:
            self._respond(404, {"success": False, "error": "JS file not found"})

    def _serve_invoice_file(self, query_string):
        from urllib.parse import parse_qs
        qs = {k: v[0] for k, v in parse_qs(query_string).items()}
        cycle_id = qs.get("cycle_id", "")
        file_path = qs.get("file", "")
        if not cycle_id or not file_path:
            self._respond(400, {"success": False, "error": "缺少参数"})
            return

        from .db import get_db
        db = get_db()
        cycle = db.execute("SELECT * FROM invoice_cycles WHERE id = ?", (cycle_id,)).fetchone()
        if not cycle:
            self._respond(404, {"success": False, "error": "周期不存在"})
            return

        full_path = os.path.join(INVOICE_DIR, cycle["cycle_name"], file_path)
        # 安全检查：确保路径在 INVOICE_DIR 内
        if not os.path.normpath(full_path).startswith(os.path.normpath(INVOICE_DIR)):
            self._respond(403, {"success": False, "error": "禁止访问"})
            return
        if not os.path.exists(full_path):
            self._respond(404, {"success": False, "error": "文件不存在"})
            return

        # 根据扩展名设置 Content-Type
        ext = os.path.splitext(file_path)[1].lower()
        content_types = {
            ".pdf": "application/pdf",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".gif": "image/gif",
        }
        ct = content_types.get(ext, "application/octet-stream")

        try:
            with open(full_path, "rb") as f:
                data = f.read()
            self.send_response(200)
            self.send_header("Content-Type", ct)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self._set_no_cache()
            self.end_headers()
            self.wfile.write(data)
        except Exception as e:
            self._respond(500, {"success": False, "error": str(e)})

    def _serve_invoice_png(self, query_string):
        """将 PDF 发票转为 PNG 图片"""
        from urllib.parse import parse_qs
        qs = {k: v[0] for k, v in parse_qs(query_string).items()}
        cycle_id = qs.get("cycle_id", "")
        file_path = qs.get("file", "")
        if not cycle_id or not file_path:
            self._respond(400, {"success": False, "error": "缺少参数"})
            return

        from .db import get_db
        db = get_db()
        cycle = db.execute("SELECT * FROM invoice_cycles WHERE id = ?", (cycle_id,)).fetchone()
        if not cycle:
            self._respond(404, {"success": False, "error": "周期不存在"})
            return

        full_path = os.path.join(INVOICE_DIR, cycle["cycle_name"], file_path)
        if not os.path.normpath(full_path).startswith(os.path.normpath(INVOICE_DIR)):
            self._respond(403, {"success": False, "error": "禁止访问"})
            return
        if not os.path.exists(full_path):
            self._respond(404, {"success": False, "error": "文件不存在"})
            return

        try:
            import fitz
            doc = fitz.open(full_path)
            page = doc[0]
            pix = page.get_pixmap(dpi=200)
            img_bytes = pix.tobytes("png")
            doc.close()

            self.send_response(200)
            self.send_header("Content-Type", "image/png")
            self.send_header("Content-Length", str(len(img_bytes)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self._set_no_cache()
            self.end_headers()
            self.wfile.write(img_bytes)
        except Exception as e:
            self._respond(500, {"success": False, "error": "PDF 转图片失败: " + str(e)})

    # ── 兼容旧端点 ──

    def _handle_legacy_export(self, data):
        from email.header import Header
        groups = data.get("groups", {})
        month_str = data.get("month", "")
        filename = data.get("filename", f"在线客服合计任务量统计表（{month_str}）.xlsx")
        from .excel_generator import generate_excel
        excel_data = generate_excel(groups, month_str)
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{Header(filename, "utf-8").encode()}"')
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Expose-Headers", "Content-Disposition")
        self.end_headers()
        self.wfile.write(excel_data.getvalue())

    def _handle_legacy_send(self, data):
        import smtplib
        import ssl as ssl_mod
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        from email.mime.base import MIMEBase
        from email import encoders
        from email.header import Header
        from .excel_generator import generate_excel
        from .auth import validate_token
        from .db import get_db

        SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.qq.com")
        SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))
        SMTP_USER = os.environ.get("SMTP_USER", "")
        SMTP_PASS = os.environ.get("SMTP_PASS", "")
        FROM_NAME = os.environ.get("FROM_NAME", "元气工作站")

        to_addr = data.get("to", "").strip()
        cc_addr = data.get("cc", "").strip()

        # 合并全局默认抄送
        auth_header = self.headers.get("Authorization", "")
        token = auth_header.replace("Bearer ", "")
        user = validate_token(token)
        if user:
            db = get_db()
            row = db.execute(
                "SELECT data_value FROM app_data WHERE data_key = 'email_config' AND (user_id = ? OR user_id IS NULL) ORDER BY user_id DESC LIMIT 1",
                (user["id"],)).fetchone()
            if row:
                global_email = json.loads(row["data_value"])
                g_cc = global_email.get("default_cc", "")
                if g_cc:
                    cc_parts = [p.strip() for p in (cc_addr + "," + g_cc).split(",") if p.strip()]
                    cc_addr = ",".join(dict.fromkeys(cc_parts))
        month_str = data.get("month", "")
        groups = data.get("groups", {})

        if not to_addr:
            return (400, {"success": False, "error": "缺少收件人地址"})
        if not SMTP_USER or not SMTP_PASS:
            return (500, {"success": False, "error": "旧端点需要设置环境变量 SMTP_USER/SMTP_PASS"})

        filename = f"在线客服合计任务量统计表（{month_str}）.xlsx"
        excel_data = generate_excel(groups, month_str)

        subject = f"在线客服合计任务量统计表（{month_str}）"
        body = f"您好，\n\n附件为 {month_str} 在线客服合计任务量统计表，请查收。\n\n—— 元气工作站自动发送"

        msg = MIMEMultipart()
        msg["From"] = f"{FROM_NAME} <{SMTP_USER}>"
        msg["To"] = to_addr
        if cc_addr:
            msg["Cc"] = cc_addr
        msg["Subject"] = Header(subject, "utf-8")
        msg.attach(MIMEText(body, "plain", "utf-8"))

        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(excel_data.getvalue())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{Header(filename, "utf-8").encode()}"')
        msg.attach(part)

        all_recipients = [a.strip() for a in to_addr.split(",") if a.strip()]
        if cc_addr:
            all_recipients += [a.strip() for a in cc_addr.split(",") if a.strip()]

        try:
            context = ssl_mod.create_default_context()
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context, timeout=15) as server:
                server.login(SMTP_USER, SMTP_PASS)
                server.sendmail(SMTP_USER, all_recipients, msg.as_string())
            return (200, {"success": True, "message": "邮件已发送"})
        except smtplib.SMTPAuthenticationError:
            return (500, {"success": False, "error": "SMTP 认证失败"})
        except smtplib.SMTPException as e:
            return (500, {"success": False, "error": f"SMTP 错误：{e}"})

    def _handle_legacy_exam_export(self, data):
        from email.header import Header
        records = data.get("records", [])
        month_str = data.get("month", "")
        filename = data.get("filename", f"{month_str}在线客服月考成绩.xlsx")
        from .excel_generator import generate_exam_excel
        excel_data = generate_exam_excel(records, month_str)
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{Header(filename, "utf-8").encode()}"')
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Expose-Headers", "Content-Disposition")
        self.end_headers()
        self.wfile.write(excel_data.getvalue())

    def _handle_legacy_exam_word(self, data):
        from email.header import Header
        records = data.get("records", [])
        month_str = data.get("month", "")
        images = data.get("images", {})
        filename = data.get("filename", f"{month_str}在线客服月考打字截图.docx")
        from .word_generator import generate_exam_word
        word_data = generate_exam_word(records, month_str, images)
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        self.send_header("Content-Disposition", f'attachment; filename="{Header(filename, "utf-8").encode()}"')
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Expose-Headers", "Content-Disposition")
        self.end_headers()
        self.wfile.write(word_data.getvalue())

    def _handle_invoice_upload(self):
        """解析 multipart/form-data 并调用发票上传处理"""
        content_type = self.headers.get('Content-Type', '')
        length = int(self.headers.get('Content-Length', 0))
        raw = self.rfile.read(length)

        boundary = self._extract_boundary(content_type)
        if not boundary:
            return 400, {"success": False, "error": "无法解析上传数据"}

        data = self._parse_multipart(raw, boundary)
        return handle_upload_invoice(self, data)

    def _handle_verify_upload(self):
        """解析 multipart/form-data 并调用验真截图上传"""
        content_type = self.headers.get('Content-Type', '')
        length = int(self.headers.get('Content-Length', 0))
        raw = self.rfile.read(length)

        boundary = self._extract_boundary(content_type)
        if not boundary:
            return 400, {"success": False, "error": "无法解析上传数据"}

        data = self._parse_multipart(raw, boundary)
        return handle_verify_upload(self, data)

    def _extract_boundary(self, content_type):
        """从 Content-Type 提取 boundary"""
        for part in content_type.split(';'):
            part = part.strip()
            if part.startswith('boundary='):
                return part[len('boundary='):].strip('"')
        return None

    def _parse_multipart(self, raw, boundary):
        """简易 multipart/form-data 解析，避免 cgi.FieldStorage 兼容问题"""
        import io

        data = {}
        boundary_bytes = ('--' + boundary).encode('utf-8')
        end_boundary = ('--' + boundary + '--').encode('utf-8')

        # 按 boundary 分割
        parts = raw.split(boundary_bytes)
        for part in parts:
            if not part or part == b'--' or part == b'--\r\n':
                continue
            if part.startswith(b'--'):
                break  # 结束边界

            # 去掉开头的 \r\n
            if part.startswith(b'\r\n'):
                part = part[2:]
            elif part.startswith(b'\n'):
                part = part[1:]

            # 分割头部和内容
            header_end = part.find(b'\r\n\r\n')
            if header_end == -1:
                header_end = part.find(b'\n\n')
            if header_end == -1:
                continue

            headers_raw = part[:header_end].decode('utf-8', errors='replace')
            body = part[header_end + 4:]  # skip \r\n\r\n

            # 去掉尾部 \r\n
            if body.endswith(b'\r\n'):
                body = body[:-2]
            elif body.endswith(b'\n'):
                body = body[:-1]

            # 解析 Content-Disposition
            name = None
            filename = None
            for line in headers_raw.split('\r\n'):
                line = line.strip()
                if not line:
                    continue
                key_val = line.split(':', 1)
                if len(key_val) != 2:
                    continue
                header_name = key_val[0].strip().lower()
                header_val = key_val[1].strip()

                if header_name == 'content-disposition':
                    for param in header_val.split(';'):
                        param = param.strip()
                        if param.startswith('name='):
                            name = param[5:].strip('"')
                        elif param.startswith('filename='):
                            filename = param[9:].strip('"')

            if name is None:
                continue

            if filename:
                # 文件字段：创建一个简单对象模拟 cgi.FieldStorage 的行为
                class FileField:
                    pass
                field = FileField()
                field.filename = filename
                field.file = io.BytesIO(body)
                field.value = body
                data[name] = field
            else:
                data[name] = body.decode('utf-8', errors='replace')

        return data

    def _handle_promo_upload(self):
        """解析 multipart/form-data 并调用绩效导入"""
        content_type = self.headers.get('Content-Type', '')
        length = int(self.headers.get('Content-Length', 0))
        raw = self.rfile.read(length)

        boundary = self._extract_boundary(content_type)
        if not boundary:
            return 400, {"success": False, "error": "无法解析上传数据"}

        data = self._parse_multipart(raw, boundary)
        return handle_import_file(self, data)

    def _handle_promo_history_upload(self):
        """解析 multipart/form-data 并调用历史绩效导入"""
        content_type = self.headers.get('Content-Type', '')
        length = int(self.headers.get('Content-Length', 0))
        raw = self.rfile.read(length)

        boundary = self._extract_boundary(content_type)
        if not boundary:
            return 400, {"success": False, "error": "无法解析上传数据"}

        data = self._parse_multipart(raw, boundary)
        return handle_import_history(self, data)

    def _set_no_cache(self):
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")

    def _respond(self, status, data):
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self._set_no_cache()
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))


def main():
    import threading

    class ThreadingHTTPServer(ThreadingMixIn, http.server.HTTPServer):
        daemon_threads = True

    # ── HTTPS 服务器（8899）──
    server = ThreadingHTTPServer(("0.0.0.0", PORT), RequestHandler)

    cert_file = os.path.join(BASE_DIR, "data", "server.crt")
    key_file = os.path.join(BASE_DIR, "data", "server.key")
    if os.path.exists(cert_file) and os.path.exists(key_file):
        ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ctx.load_cert_chain(cert_file, key_file)
        server.socket = ctx.wrap_socket(server.socket, server_side=True)
        print(f"[OK] HTTPS: https://localhost:{PORT}")
    else:
        print(f"[OK] HTTP: http://localhost:{PORT}")

    # ── HTTP 服务器（8898，供局域网访问，无需证书）──
    http_server = ThreadingHTTPServer(("0.0.0.0", HTTP_PORT), RequestHandler)
    http_thread = threading.Thread(target=http_server.serve_forever, daemon=True)
    http_thread.start()
    print(f"[OK] HTTP（局域网）: http://{_get_lan_ip()}:{HTTP_PORT}")

    print("  打开浏览器访问即可使用")
    print("按 Ctrl+C 停止")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[OK] 服务已停止")
        server.server_close()
        http_server.server_close()


if __name__ == "__main__":
    main()
