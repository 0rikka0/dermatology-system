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

from app.config import load_config
from app.instance_lock import (
    acquire_instance_mutex,
    release_instance_mutex,
    discover_running_instance,
    probe_instance,
    write_runtime_info,
    remove_runtime_info,
)


def _open_browser_url(url):
    """Open URL in default browser unless suppressed for automated tests."""
    if "--no-browser" in sys.argv or os.environ.get("DERM_NO_BROWSER") == "1":
        return
    try:
        webbrowser.open(url)
    except Exception as e:
        print(f"[System] 自动打开浏览器失败: {e}")


def main():
    print("=" * 60)
    print("      皮肤病理报告系统 (Dermatology Pathology System)")
    print("=" * 60)

    cfg = load_config()
    port = int(cfg.get("server_port", 8765) or 8765)

    # Acquire OS-level named mutex before heavy imports and socket binding
    acquired, mutex_handle = acquire_instance_mutex()
    if not acquired:
        found_port, _ = discover_running_instance(base_port=port, max_wait=8.0)
        if found_port:
            url = f"http://127.0.0.1:{found_port}"
            print(f"[System] 检测到系统已在后台运行 ({url})，正在为您打开浏览器...")
            _open_browser_url(url)
            return
        # Retry once in case primary instance exited/crashed during startup
        acquired, mutex_handle = acquire_instance_mutex()
        if not acquired:
            print("[System] 检测到系统后台实例正在启动或忙碌，请稍后重试。")
            return

    httpd = None
    try:
        # Remove stale runtime state left by any previous unclean exit
        remove_runtime_info()

        from app.path_manager import ensure_workspace_structure, get_database_path
        from app.database import init_db
        from app.server import run_server

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
        for p in range(port, port + 10):
            try:
                httpd = run_server(port=p)
                actual_port = p
                break
            except OSError as e:
                if probe_instance(p, timeout=0.6):
                    url = f"http://127.0.0.1:{p}"
                    print(f"[System] 检测到端口 {p} 已有本系统实例运行，正在打开浏览器 ({url})...")
                    _open_browser_url(url)
                    return
                print(f"[System] 端口 {p} 无法绑定 ({e})，正在尝试备用端口 {p+1}...")

        if not httpd:
            print(f"[System] 错误：无法在端口 {port}~{port+9} 范围内启动服务，请检查网络权限。")
            return

        write_runtime_info(actual_port, pid=os.getpid())

        url = f"http://127.0.0.1:{actual_port}"
        print(f"[System] WebUI 访问地址: {url}")
        print("[System] 正在启动后台服务，按 Ctrl+C 可停止系统...")

        # Auto open browser in a background thread
        if cfg.get("auto_open_browser", True):
            def open_browser():
                time.sleep(0.8)
                _open_browser_url(url)
            threading.Thread(target=open_browser, daemon=True).start()

        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n[System] 系统已安全关闭。")
        except Exception as e:
            print(f"\n[System] 服务运行异常: {e}")
    finally:
        if httpd is not None:
            try:
                httpd.server_close()
            except Exception:
                pass
        remove_runtime_info(expected_pid=os.getpid())
        release_instance_mutex(mutex_handle)


if __name__ == "__main__":
    main()
