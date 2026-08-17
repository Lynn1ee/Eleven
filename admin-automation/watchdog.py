"""守护进程（无窗口后台运行）：每 30 秒检查服务器，挂了就拉起"""
import socket
import ssl
import subprocess
import sys
import os
import time
import logging
import atexit
import traceback
from logging.handlers import RotatingFileHandler

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(SCRIPT_DIR, 'watchdog.log')
WATCHDOG_PID_FILE = os.path.join(SCRIPT_DIR, '.watchdog.pid')
LAUNCHER_PID_FILE = os.path.join(SCRIPT_DIR, '.launcher.pid')

_handler = RotatingFileHandler(
    LOG_FILE,
    maxBytes=10 * 1024 * 1024, backupCount=3, encoding='utf-8'
)
_handler.setFormatter(logging.Formatter('%(asctime)s %(message)s'))
logging.basicConfig(level=logging.INFO, handlers=[_handler])

PORT = 8899
NO_WINDOW = subprocess.CREATE_NO_WINDOW

# 重启防抖：记录最近重启时间，防止死循环
_restart_times = []
_RESTART_LIMIT = 3       # 5分钟内最多重启3次
_RESTART_WINDOW = 300    # 5分钟时间窗口


def is_process_alive(pid):
    """检查指定 PID 的进程是否存在（用 tasklist /FI 精确匹配）。"""
    try:
        result = subprocess.run(
            ['tasklist', '/FI', f'PID eq {pid}', '/FO', 'CSV', '/NH'],
            capture_output=True, text=True, timeout=5,
            creationflags=NO_WINDOW,
        )
        return str(pid) in result.stdout
    except Exception:
        return False


def get_port_owner():
    """用 netstat 查出占用 PORT 的进程 PID，无占用返回 None。"""
    try:
        result = subprocess.run(
            ['netstat', '-ano', '-p', 'TCP'],
            capture_output=True, text=True, timeout=5,
            creationflags=NO_WINDOW,
        )
        for line in result.stdout.splitlines():
            if f':{PORT}' in line and 'LISTENING' in line:
                parts = line.strip().split()
                return int(parts[-1])
    except Exception:
        pass
    return None


def kill_process(pid, label=''):
    """强制杀掉指定 PID 的进程。"""
    try:
        logging.info(f'Killing {label} (PID {pid})')
        subprocess.run(['taskkill', '/F', '/PID', str(pid)],
                       capture_output=True, timeout=5,
                       creationflags=NO_WINDOW)
        return True
    except Exception as e:
        logging.error(f'Failed to kill {label} (PID {pid}): {e}')
        return False


def is_server_running():
    """HTTPS 健康检查：验证服务器能真正返回页面，而非仅端口通。"""
    try:
        sock = socket.create_connection(('127.0.0.1', PORT), timeout=5)
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        tls = ctx.wrap_socket(sock, server_hostname='localhost')
        tls.settimeout(5)
        tls.sendall(b'GET / HTTP/1.1\r\nHost: localhost\r\nConnection: close\r\n\r\n')
        # 循环读取直到连接关闭（Python http.server 分块发送，单次 recv 只拿到头）
        chunks = []
        while True:
            chunk = tls.recv(4096)
            if not chunk:
                break
            chunks.append(chunk)
        # 发送 TLS close_notify 再关闭，避免服务器端产生 ssl.SSLEOFError
        try:
            tls = tls.unwrap()
        except Exception:
            pass
        finally:
            sock.close()
        resp = b''.join(chunks)
        if b'DOCTYPE' not in resp:
            logging.warning(f'Health check: response missing DOCTYPE (got {len(resp)} bytes)')
        return b'DOCTYPE' in resp
    except Exception as e:
        logging.warning(f'Health check exception: {type(e).__name__}: {e}')
        return False


def ensure_firewall():
    try:
        result = subprocess.run(
            ['netsh', 'advfirewall', 'firewall', 'show', 'rule', 'name=yuanqi-8899'],
            capture_output=True, text=True, timeout=5,
            creationflags=NO_WINDOW,
        )
        if result.returncode != 0 or 'yuanqi-8899' not in result.stdout:
            logging.info('Firewall rule missing, re-adding...')
            subprocess.run(
                ['netsh', 'advfirewall', 'firewall', 'add', 'rule',
                 'name=yuanqi-8899', 'dir=in', 'action=allow',
                 'protocol=TCP', 'localport=8899'],
                capture_output=True, timeout=5,
                creationflags=NO_WINDOW,
            )
    except Exception as e:
        logging.error(f'Firewall check failed: {e}')


def start_server():
    global _restart_times

    # 防抖检查
    now = time.time()
    _restart_times = [t for t in _restart_times if now - t < _RESTART_WINDOW]
    if len(_restart_times) >= _RESTART_LIMIT:
        logging.warning(f'{_RESTART_LIMIT} restarts in {_RESTART_WINDOW}s, pausing 120s to break loop')
        time.sleep(120)

    _restart_times.append(now)

    # 1. 检查端口是否被其他进程占用
    port_owner = get_port_owner()
    tracked_pid = None
    if os.path.exists(LAUNCHER_PID_FILE):
        try:
            tracked_pid = int(open(LAUNCHER_PID_FILE).read().strip())
        except Exception:
            pass

    if port_owner:
        if tracked_pid and port_owner == tracked_pid:
            # 端口由 tracked launcher 占用——进程活着但服务不响应，杀
            logging.info(f'Port owned by tracked launcher (PID {port_owner}) but not responding')
            kill_process(port_owner, 'stale launcher')
            time.sleep(2)
        elif tracked_pid and port_owner != tracked_pid:
            # PID 错配：端口被其他进程占用，两个都杀
            logging.info(f'PID mismatch: port owned by {port_owner}, tracked PID is {tracked_pid}')
            kill_process(port_owner, 'port owner')
            kill_process(tracked_pid, 'tracked launcher')
            time.sleep(2)
        elif not tracked_pid:
            # 端口被未知进程占用（可能是手动启动的残留），杀
            logging.info(f'Port occupied by unknown process (PID {port_owner}), killing')
            kill_process(port_owner, 'unknown port owner')
            time.sleep(2)

    # 2. 如果 tracked PID 还在（但端口不属于它），也要清理
    if tracked_pid and is_process_alive(tracked_pid) and tracked_pid != port_owner:
        kill_process(tracked_pid, 'orphaned launcher')
        time.sleep(2)

    # 3. 清理 PID 文件
    try:
        if os.path.exists(LAUNCHER_PID_FILE):
            os.remove(LAUNCHER_PID_FILE)
    except Exception:
        pass

    # 4. 启动新 launcher
    launcher = os.path.join(SCRIPT_DIR, 'launcher.py')
    logging.info('Starting new launcher...')
    try:
        subprocess.Popen(
            [sys.executable, launcher],
            cwd=SCRIPT_DIR,
            creationflags=NO_WINDOW,
        )
    except Exception as e:
        logging.error(f'Failed to start launcher: {e}')
        return

    # 5. 等待并验证启动成功
    for _ in range(10):  # 最多等 10 秒
        time.sleep(1)
        if is_server_running():
            logging.info('Launcher started successfully, server is responding')
            return

    logging.error('Launcher started but server not responding after 10s')


def cleanup_pid():
    try:
        if os.path.exists(WATCHDOG_PID_FILE):
            os.remove(WATCHDOG_PID_FILE)
    except Exception:
        pass


def main():
    # PID 锁 — 防止重复启动 watchdog 自身
    if os.path.exists(WATCHDOG_PID_FILE):
        try:
            old_pid = int(open(WATCHDOG_PID_FILE).read().strip())
            if is_process_alive(old_pid):
                logging.info(f'Watchdog already running (PID {old_pid}), exiting')
                return
        except Exception:
            pass
        try:
            os.remove(WATCHDOG_PID_FILE)
        except Exception:
            pass

    with open(WATCHDOG_PID_FILE, 'w') as f:
        f.write(str(os.getpid()))
    atexit.register(cleanup_pid)

    logging.info('Watchdog started (pythonw, silent mode)')
    ensure_firewall()

    if not is_server_running():
        start_server()

    fail_count = 0
    while True:
        time.sleep(30)
        try:
            if not is_server_running():
                fail_count += 1
                logging.info(f'Health check failed ({fail_count}/2)')
                if fail_count >= 2:
                    logging.warning('2 consecutive failures, restarting server')
                    start_server()
                    fail_count = 0
            else:
                if fail_count > 0:
                    logging.info(f'Health check recovered after {fail_count} failure(s)')
                fail_count = 0
        except Exception as e:
            logging.error(f'Loop error: {e}')
            time.sleep(5)


if __name__ == '__main__':
    main()
