"""启动器：供 pythonw.exe 调用，导入并运行服务器"""
import sys
import os
import logging
from logging.handlers import RotatingFileHandler

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

# 日志写入文件，避免 pythonw 无控制台时静默崩溃
_handler = RotatingFileHandler(
    os.path.join(SCRIPT_DIR, 'launcher.log'),
    maxBytes=10 * 1024 * 1024, backupCount=3, encoding='utf-8'
)
_handler.setFormatter(logging.Formatter('%(asctime)s %(message)s'))
logging.basicConfig(level=logging.INFO, handlers=[_handler])
logging.info(f'Launcher started (PID {os.getpid()})')

# 写 PID 文件
with open(os.path.join(SCRIPT_DIR, '.launcher.pid'), 'w') as f:
    f.write(str(os.getpid()))

# 重定向 stderr 到日志，stdout 丢弃（http.server 的访问日志太吵）
sys.stdout = open(os.devnull, 'w')
sys.stderr = open(os.path.join(SCRIPT_DIR, 'launcher.log'), 'a')

try:
    from server.main import main
    main()
except Exception as e:
    logging.error(f'Launcher crashed: {e}', exc_info=True)
    # 清理 PID 文件，防止 watchdog 误判
    try:
        os.remove(os.path.join(SCRIPT_DIR, '.launcher.pid'))
    except Exception:
        pass
    raise
