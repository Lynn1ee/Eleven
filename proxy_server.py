"""Local proxy server for ticket-refund-explainer.html.
Serves static files AND proxies /api/* to CoffeeClaw with session cookie.
This eliminates CORS issues by making everything same-origin (localhost:3456).
"""
import http.server
import urllib.request
import urllib.error
import json
import os
import sys
import ssl

COFFEECLAW_BASE = "http://59.110.155.133:7301"
STATIC_DIR = os.path.dirname(os.path.abspath(__file__))
PORT = 3457


class ProxyHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=STATIC_DIR, **kwargs)

    def do_GET(self):
        if self.path.startswith('/api/'):
            self.proxy_request('GET')
        else:
            super().do_GET()

    def do_POST(self):
        if self.path.startswith('/api/'):
            self.proxy_request('POST')
        else:
            self.send_error(404)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, X-Session-Cookie')
        self.end_headers()

    def proxy_request(self, method):
        target_url = COFFEECLAW_BASE + self.path
        if self.path.startswith('/api/'):
            target_url = COFFEECLAW_BASE + '/' + self.path.lstrip('/')

        # Read request body
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length) if content_length > 0 else None

        # Build headers for upstream request
        upstream_headers = {
            'Content-Type': self.headers.get('Content-Type', 'application/json'),
        }

        # Forward session cookie from custom header
        session_cookie = self.headers.get('X-Session-Cookie', '')
        if session_cookie:
            upstream_headers['Cookie'] = session_cookie

        try:
            req = urllib.request.Request(
                target_url,
                data=body,
                headers=upstream_headers,
                method=method
            )
            # Skip SSL verification for internal service
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE

            resp = urllib.request.urlopen(req, timeout=120, context=ctx)

            # Forward response
            self.send_response(resp.status)
            for key, val in resp.getheaders():
                if key.lower() in ('transfer-encoding', 'connection', 'keep-alive'):
                    continue
                self.send_header(key, val)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()

            # Stream response (supports SSE)
            while True:
                chunk = resp.read(8192)
                if not chunk:
                    break
                self.wfile.write(chunk)
                self.wfile.flush()

        except urllib.error.HTTPError as e:
            self.send_response(e.code)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            err_body = e.read()
            self.wfile.write(err_body)
        except Exception as e:
            self.send_response(502)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({
                'ok': False,
                'error': {'type': 'proxy_error', 'message': str(e)}
            }).encode())

    def log_message(self, format, *args):
        # Suppress log noise for static files
        if '/api/' in str(args[0]):
            print(f"[proxy] {args[0]}")
        else:
            pass  # silent for static files


if __name__ == '__main__':
    print(f"Serving {STATIC_DIR} at http://localhost:{PORT}")
    print(f"Proxying /api/* to {COFFEECLAW_BASE}")
    print(f"Use X-Session-Cookie header to pass coffeeclaw_session cookie")
    print(f"\nOpen: http://localhost:{PORT}/ticket-refund-explainer.html")
    httpd = http.server.HTTPServer(('0.0.0.0', PORT), ProxyHandler)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down...")
        httpd.shutdown()
