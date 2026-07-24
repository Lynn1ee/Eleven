"""Bot 守护进程 — 每 3 分钟检测健康状态，异常时自动重启并通知飞书群。"""

import json
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

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import requests

# ── 配置 ──
SCRIPT_DIR = Path(__file__).parent
BOT_SCRIPT = SCRIPT_DIR / "server.py"
STATE_FILE = SCRIPT_DIR / ".watchdog_state.json"
PID_FILE = SCRIPT_DIR / ".watchdog_pid"
LOG_FILE = SCRIPT_DIR / "watchdog.log"
BOT_LOG_FILE = SCRIPT_DIR / "bot.log"

CHECK_INTERVAL = 1800   # 30 分钟，正常检查间隔
COOLDOWN_SECONDS = 120    # 重启后冷却 2 分钟
# 重启失败退避：连续失败时依次使用（秒）
RETRY_BACKOFF = [120, 300, 900, 1800]  # 2min → 5min → 15min → 30min
STARTUP_WAIT = 30       # server.py 启动等待（SDK 加载需要 20s+）
WS_STALE_SECONDS = 14400  # WebSocket 无消息超过 4 小时视为异常（仅 08:00-23:00 时段）

FEISHU_APP_ID = os.getenv("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.getenv("FEISHU_APP_SECRET", "")
CHAT_ID = os.getenv("CHAT_ID", "oc_8fe150c12cf172df19a52ba745dab9e6")
PORT = int(os.getenv("PORT", "8080"))


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


def load_state() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text())
        except (json.JSONDecodeError, KeyError):
            pass
    return {}


def save_state(state: dict):
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False))


def check_local_port() -> bool:
    """检测本地 8080 端口是否响应（重试 3 次，间隔 5 秒）。"""
    for attempt in range(3):
        try:
            r = requests.post(
                f"http://127.0.0.1:{PORT}/webhook",
                data=b"{}",
                timeout=8,
            )
            return True
        except requests.ConnectionError:
            if attempt < 2:
                time.sleep(5)
        except Exception:
            return True
    return False


def get_ngrok_url() -> str | None:
    """通过 ngrok 本地 API 动态获取当前隧道公网 URL。"""
    for api_port in (4040, 4041):
        try:
            r = requests.get(f"http://127.0.0.1:{api_port}/api/tunnels", timeout=5)
            if r.status_code == 200:
                tunnels = r.json().get("tunnels", [])
                for t in tunnels:
                    if f"localhost:{PORT}" in t.get("config", {}).get("addr", ""):
                        return t["public_url"]
        except Exception:
            continue
    return None


def check_ngrok() -> bool:
    """检测 ngrok 公网 URL 是否可达（重试 3 次，间隔 5 秒）。"""
    for attempt in range(3):
        url = get_ngrok_url()
        if not url:
            if attempt < 2:
                time.sleep(5)
            continue
        try:
            r = requests.get(
                f"{url}/",
                timeout=10,
                headers={"User-Agent": "watchdog"},
            )
            return True
        except Exception:
            if attempt < 2:
                time.sleep(5)
    return False


def kill_processes():
    """杀死 server.py 对应的 python.exe 进程，排除 watchdog 自身。ngrok 不杀以保持隧道存活。"""
    my_pid = str(os.getpid())
    # 逐个杀 python.exe，跳过自己
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq python.exe", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10,
        )
        for line in result.stdout.strip().split("\n"):
            parts = line.replace('"', '').split(",")
            if len(parts) >= 2:
                pid = parts[1].strip()
                if pid and pid != my_pid:
                    try:
                        subprocess.run(
                            ["taskkill", "/F", "/PID", pid],
                            capture_output=True, timeout=5,
                        )
                    except Exception:
                        pass
    except Exception:
        pass
    time.sleep(2)


def start_bot():
    """后台启动 server.py。"""
    # 清除旧的最后收消息时间（新 bot 实例还没收到消息）
    last_msg_file = SCRIPT_DIR / ".last_msg_ts"
    try:
        last_msg_file.unlink()
    except Exception:
        pass
    try:
        subprocess.Popen(
            [sys.executable, "-u", str(BOT_SCRIPT)],
            cwd=str(SCRIPT_DIR),
            stdout=open(BOT_LOG_FILE, "a"),
            stderr=subprocess.STDOUT,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return True
    except Exception:
        log(f"启动 bot 失败:\n{traceback.format_exc()}")
        return False


def wait_for_startup() -> bool:
    """等待 bot 启动完成，返回是否启动成功。"""
    log(f"等待 bot 启动（最长 {STARTUP_WAIT}s）...")
    time.sleep(STARTUP_WAIT)
    healthy = check_local_port()
    if healthy:
        log("bot 启动成功")
    return healthy


def get_feishu_token() -> str:
    """获取飞书 tenant access token。"""
    resp = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET},
        timeout=10,
    )
    data = resp.json()
    if data.get("code") != 0:
        raise Exception(f"获取 token 失败: {data}")
    return data["tenant_access_token"]


def send_feishu_notification(msg: str):
    """发送消息到飞书群聊。"""
    try:
        token = get_feishu_token()
        body = {
            "receive_id": CHAT_ID,
            "msg_type": "text",
            "content": json.dumps({"text": msg}),
        }
        resp = requests.post(
            "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json; charset=utf-8",
            },
            json=body,
            timeout=10,
        )
        data = resp.json()
        if data.get("code") != 0:
            log(f"飞书通知失败: {data}")
        else:
            log("飞书通知已发送")
    except Exception:
        log(f"飞书通知异常:\n{traceback.format_exc()}")


def check_ws_alive() -> bool:
    """检查 WebSocket 是否仍在接收消息（通过 bot 写入的 .last_msg_ts 判断）。"""
    last_msg_file = SCRIPT_DIR / ".last_msg_ts"
    if not last_msg_file.exists():
        return True  # 刚启动还没收到消息，不误报
    try:
        last_ts = float(last_msg_file.read_text().strip())
        elapsed = time.time() - last_ts
        if elapsed > WS_STALE_SECONDS:
            hour = datetime.now().hour
            if 8 <= hour < 23:
                return False
        return True
    except Exception:
        return True


def get_bot_status() -> str:
    """返回 bot 状态描述。"""
    local_ok = check_local_port()
    ngrok_ok = check_ngrok() if local_ok else False
    if not local_ok:
        return "dead"
    if not ngrok_ok:
        return "ngrok_down"
    if not check_ws_alive():
        return "ws_dead"
    return "healthy"


def run_check() -> int:
    """执行一次健康检查，返回下次检查前的等待秒数。"""
    state = load_state()
    last_restart = state.get("last_restart_ts", 0)
    now = time.time()

    # 冷却期检查
    if now - last_restart < COOLDOWN_SECONDS:
        remaining = int(COOLDOWN_SECONDS - (now - last_restart))
        log(f"冷却期中，剩余 {remaining}s，跳过检查")
        return CHECK_INTERVAL

    status = get_bot_status()
    log(f"状态: {status}")

    if status == "healthy":
        state["last_healthy_ts"] = now
        state["consecutive_failures"] = 0
        save_state(state)
        return CHECK_INTERVAL

    # ngrok 断线不杀 bot：网络抖动时 ngrok 会自行重连，bot 本身正常
    if status == "ngrok_down":
        log("ngrok 隧道断开，等待自动恢复，不重启 bot")
        msg = "⚠️ ngrok 隧道断开，请检查网络。bot 仍在运行，网络恢复后自动重连。"
        send_feishu_notification(msg)
        return CHECK_INTERVAL

    # ── 异常：执行自愈 ──
    consecutive = state.get("consecutive_failures", 0)
    restart_count = state.get("restart_count", 0) + 1
    log(f"检测到异常 ({status})，开始自愈流程（第 {restart_count} 次重启）...")

    kill_processes()
    time.sleep(2)

    if not start_bot():
        log("启动 bot 失败，下次检查重试")
        consecutive += 1
        state["consecutive_failures"] = consecutive
        state["last_restart_ts"] = now
        save_state(state)
        backoff = RETRY_BACKOFF[min(consecutive - 1, len(RETRY_BACKOFF) - 1)]
        log(f"连续失败 {consecutive} 次，{backoff}s 后重试")
        return backoff

    ok = wait_for_startup()

    state["last_restart_ts"] = time.time()
    state["restart_count"] = restart_count
    if ok:
        state["last_healthy_ts"] = time.time()
        state["consecutive_failures"] = 0
    else:
        consecutive += 1
        state["consecutive_failures"] = consecutive
    save_state(state)

    if ok:
        msg = (
            f"🔄 机器人自动恢复\n"
            f"检测到异常（{status}），已自动重启完成。\n"
            f"重启次数：第 {restart_count} 次\n"
            f"断线期间的消息已通过补录机制自动追回。"
        )
        send_feishu_notification(msg)
        return CHECK_INTERVAL
    else:
        backoff = RETRY_BACKOFF[min(consecutive - 1, len(RETRY_BACKOFF) - 1)]
        msg = (
            f"⚠️ 机器人自动恢复失败\n"
            f"检测到异常（{status}），重启后仍未恢复。\n"
            f"连续失败 {consecutive} 次，{backoff}s 后自动重试。"
        )
        send_feishu_notification(msg)
        return backoff


def main():
    # PID 锁 — 防止重复启动
    if PID_FILE.exists():
        old_pid = PID_FILE.read_text().strip()
        try:
            subprocess.run(
                ["tasklist", "/FI", f"PID eq {old_pid}"],
                capture_output=True, timeout=5,
            )
        except Exception:
            pass
    PID_FILE.write_text(str(os.getpid()))

    log("=== Watchdog 守护进程启动 ===")
    log(f"检测间隔: {CHECK_INTERVAL}s | 冷却期: {COOLDOWN_SECONDS}s | WS超时: {WS_STALE_SECONDS}s")

    if not FEISHU_APP_ID or not FEISHU_APP_SECRET:
        log("缺少飞书凭证，通知功能将不可用")

    # 短轮询模式：每 60 秒醒来一次，避免长时间 time.sleep 导致假死
    next_check_time = 0  # 0 = 立即执行首次检查
    last_heartbeat = 0

    while True:
        try:
            now = time.time()

            # 心跳：每分钟写一次，证明 watchdog 自身存活
            if now - last_heartbeat >= 60:
                try:
                    (SCRIPT_DIR / ".watchdog_heartbeat").write_text(str(int(now)))
                except Exception:
                    pass
                last_heartbeat = now

            # 到达检查时间，执行健康检查
            if now >= next_check_time:
                try:
                    wait = run_check()
                except Exception:
                    log(f"检测异常:\n{traceback.format_exc()}")
                    wait = CHECK_INTERVAL
                next_check_time = time.time() + wait
                log(f"下次检查: {wait}s 后")

            time.sleep(60)

        except Exception:
            log(f"主循环异常:\n{traceback.format_exc()}")
            time.sleep(60)


if __name__ == "__main__":
    main()
