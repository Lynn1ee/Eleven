"""排班系统守护进程 — 每分钟检测端口 8897，异常时自动重启。"""
import os
import sys
import io
import time
import subprocess
import traceback
from datetime import datetime
from pathlib import Path

# Windows 控制台 UTF-8
for stream_name in ('stdout', 'stderr'):
    try:
        stream = getattr(sys, stream_name)
        raw = stream.detach()
        setattr(sys, stream_name, io.TextIOWrapper(raw, encoding='utf-8'))
    except (AttributeError, OSError, ValueError):
        pass

import requests

SCRIPT_DIR = Path(__file__).parent
SERVER_SCRIPT = SCRIPT_DIR / "server.py"
PID_FILE = SCRIPT_DIR / ".watchdog_pid"
LOG_FILE = SCRIPT_DIR / "watchdog.log"
HEARTBEAT_FILE = SCRIPT_DIR / ".watchdog_heartbeat"

CHECK_INTERVAL = 60       # 检查间隔 60 秒
COOLDOWN_SECONDS = 30     # 重启后冷却 30 秒
STARTUP_WAIT = 5          # server.py 启动等待
PORT = 8897


def log(msg: str):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    try:
        print(line)
    except Exception:
        pass
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def check_port() -> bool:
    """检测本地端口是否响应"""
    try:
        r = requests.get(f"http://127.0.0.1:{PORT}/", timeout=5)
        return r.status_code == 200
    except Exception:
        return False


def get_port_owner(port):
    """用 netstat 查出占用指定端口的进程 PID，无占用返回 None。"""
    try:
        result = subprocess.run(
            ["netstat", "-ano", "-p", "TCP"],
            capture_output=True, text=True, timeout=5,
        )
        for line in result.stdout.splitlines():
            if f":{port}" in line and "LISTENING" in line:
                parts = line.strip().split()
                return int(parts[-1])
    except Exception:
        pass
    return None


def kill_server():
    """精准杀掉占用排班系统端口的进程，不误伤其他服务。"""
    port_owner = get_port_owner(PORT)
    if port_owner is None:
        log(f"端口 {PORT} 无占用进程，无需清理")
        return
    if port_owner == os.getpid():
        return
    try:
        subprocess.run(
            ["taskkill", "/F", "/PID", str(port_owner)],
            capture_output=True, timeout=5,
        )
        log(f"已终止占用端口 {PORT} 的进程 (PID {port_owner})")
    except Exception as e:
        log(f"终止进程 {port_owner} 失败: {e}")
    time.sleep(2)


def start_server() -> bool:
    """后台启动 server.py"""
    try:
        subprocess.Popen(
            [sys.executable, "-u", str(SERVER_SCRIPT)],
            cwd=str(SCRIPT_DIR),
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return True
    except Exception:
        log(f"启动失败:\n{traceback.format_exc()}")
        return False


def main():
    if PID_FILE.exists():
        try:
            old_pid = PID_FILE.read_text().strip()
            subprocess.run(
                ["tasklist", "/FI", f"PID eq {old_pid}"],
                capture_output=True, timeout=5,
            )
        except Exception:
            pass
    PID_FILE.write_text(str(os.getpid()))

    log("=== 排班系统守护进程启动 ===")
    log(f"检测间隔: {CHECK_INTERVAL}s | 端口: {PORT}")

    last_restart = 0
    restart_count = 0

    while True:
        try:
            now = time.time()

            # 心跳
            try:
                HEARTBEAT_FILE.write_text(str(int(now)))
            except Exception:
                pass

            # 冷却期跳过
            if now - last_restart < COOLDOWN_SECONDS:
                time.sleep(CHECK_INTERVAL)
                continue

            # 健康检查
            if not check_port():
                restart_count += 1
                log(f"端口 {PORT} 无响应，开始重启（第 {restart_count} 次）")

                kill_server()
                if start_server():
                    time.sleep(STARTUP_WAIT)
                    if check_port():
                        log("重启成功")
                    else:
                        log("重启后端口仍未响应")
                else:
                    log("启动失败")

                last_restart = now

            time.sleep(CHECK_INTERVAL)

        except Exception:
            log(f"主循环异常:\n{traceback.format_exc()}")
            time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()
