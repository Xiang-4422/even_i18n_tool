"""
launcher.py — PyInstaller 打包入口
负责启动 Streamlit 服务并自动打开浏览器。
"""

import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path


def find_free_port(start: int = 8501) -> int:
    for port in range(start, start + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("localhost", port)) != 0:
                return port
    return start


def wait_until_ready(port: int, timeout: int = 30) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("localhost", port)) == 0:
                return True
        time.sleep(0.3)
    return False


def main() -> None:
    # PyInstaller 打包后资源在 sys._MEIPASS，开发时在脚本同目录
    if getattr(sys, "frozen", False):
        base_dir = Path(sys._MEIPASS)
    else:
        base_dir = Path(__file__).parent

    app_script = str(base_dir / "sanitize_i18n_app.py")
    port = find_free_port()

    # 配置 Streamlit 启动参数
    # global.developmentMode=false：PyInstaller 环境下必须显式关闭，否则冲突报错
    sys.argv = [
        "streamlit", "run", app_script,
        f"--server.port={port}",
        "--server.headless=true",
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
        "--logger.level=error",
    ]

    # 后台线程等待服务就绪后打开浏览器
    def open_browser():
        if wait_until_ready(port):
            webbrowser.open(f"http://localhost:{port}")

    threading.Thread(target=open_browser, daemon=True).start()

    # 启动 Streamlit（阻塞直到退出）
    from streamlit.web import cli as stcli
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
