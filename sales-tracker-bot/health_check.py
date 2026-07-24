"""Watchdog 存活检查 — 无窗口运行，不闪屏。"""
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
HEARTBEAT_FILE = SCRIPT_DIR / ".watchdog_heartbeat"
WATCHDOG_LOG = SCRIPT_DIR / "watchdog.log"
PYTHON = Path(sys.executable).parent / "python.exe"


def is_alive() -> bool:
    if not HEARTBEAT_FILE.exists():
        return False
    try:
        ts = int(HEARTBEAT_FILE.read_text().strip())
        dt = datetime.fromtimestamp(ts, tz=timezone.utc).astimezone()
        age_seconds = (datetime.now().astimezone() - dt).total_seconds()
        return age_seconds < 600  # 10 分钟内有心跳就算存活
    except (ValueError, OSError):
        return False


def start_watchdog():
    with open(WATCHDOG_LOG, "a", encoding="utf-8") as log:
        log.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] health_check: 拉起 watchdog\n")
    subprocess.Popen(
        [str(PYTHON), "-u", str(SCRIPT_DIR / "watchdog.py")],
        cwd=str(SCRIPT_DIR),
        stdout=open(WATCHDOG_LOG, "a"),
        stderr=subprocess.STDOUT,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )


if __name__ == "__main__":
    if not is_alive():
        start_watchdog()
