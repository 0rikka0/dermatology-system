# -*- coding: utf-8 -*-
"""
皮肤病理报告系统 (Dermatology Pathology Reporting System)
Main Application Entry Point
"""
import os
import sys
import webbrowser
import threading
import time

# Ensure project root is in sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from app.config import load_config, save_config
from app.path_manager import ensure_workspace_structure, get_database_path
from app.database import init_db
from app.server import run_server

def main():
    print("=" * 60)
    print("      皮肤病理报告系统 (Dermatology Pathology System)")
    print("=" * 60)
    
    cfg = load_config()
    port = cfg.get("server_port", 8765)
    
    # If already initialized, make sure directories and DB are ready
    if cfg.get("initialized") and cfg.get("work_dir"):
        work_dir = cfg.get("work_dir")
        try:
            ensure_workspace_structure(work_dir)
            init_db(get_database_path(work_dir))
            print(f"[System] 工作目录就绪: {work_dir}")
        except Exception as e:
            print(f"[System] 初始化目录提示: {e}")
            
    # Bind server first with automatic port fallback
    actual_port = port
    httpd = None
    for p in range(port, port + 10):
        try:
            httpd = run_server(port=p)
            actual_port = p
            break
        except OSError as e:
            print(f"[System] 端口 {p} 无法绑定 ({e})，正在尝试备用端口 {p+1}...")

    if not httpd:
        print(f"[System] 错误：无法在端口 {port}~{port+9} 范围内启动服务，请检查网络权限。")
        return

    url = f"http://127.0.0.1:{actual_port}"
    print(f"[System] WebUI 访问地址: {url}")
    print("[System] 正在启动后台服务，按 Ctrl+C 可停止系统...")

    # Auto open browser in a background thread
    if cfg.get("auto_open_browser", True):
        def open_browser():
            time.sleep(0.8)
            try:
                webbrowser.open(url)
            except Exception as e:
                print(f"[System] 自动打开浏览器失败: {e}")
        threading.Thread(target=open_browser, daemon=True).start()

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[System] 系统已安全关闭。")
    except Exception as e:
        print(f"\n[System] 服务运行异常: {e}")

if __name__ == "__main__":
    main()
