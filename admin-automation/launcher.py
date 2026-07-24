"""启动器：供 pythonw.exe 调用，导入并运行服务器"""
import sys
import os
import logging

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

# 日志写入文件，避免 pythonw 无控制台时静默崩溃
logging.basicConfig(
    filename=os.path.join(SCRIPT_DIR, 'launcher.log'),
    level=logging.INFO,
    format='%(asctime)s %(message)s'
)
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
