# -*- coding: utf-8 -*-
"""
Single-instance locking, runtime state tracking, HTTP identity probing,
and graceful shutdown helpers for Dermatology Pathology Reporting System.

Uses only Python standard library modules so main.py can check existing
instances immediately before importing heavy document/database modules.
"""
import os
import sys
import json
import time
import ctypes
import subprocess
import urllib.request
import urllib.error
from datetime import datetime

APP_ID = "dermatology-system"
DEFAULT_MUTEX_NAME = os.environ.get("DERM_INSTANCE_MUTEX") or "Local\\DermatologySystem_Instance_Lock"
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_RUNTIME_FILE = os.environ.get("DERM_RUNTIME_FILE") or os.path.join(BASE_DIR, ".server_runtime.json")

_ERROR_ACCESS_DENIED = 5
_ERROR_ALREADY_EXISTS = 183
_SYNCHRONIZE = 0x00100000
_PROCESS_TERMINATE = 0x0001
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_STILL_ACTIVE = 259


def _get_kernel32():
    if os.name != "nt":
        return None
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    from ctypes import wintypes
    k32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    k32.CreateMutexW.restype = wintypes.HANDLE
    k32.CloseHandle.argtypes = [wintypes.HANDLE]
    k32.CloseHandle.restype = wintypes.BOOL
    k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    k32.OpenProcess.restype = wintypes.HANDLE
    k32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    k32.WaitForSingleObject.restype = wintypes.DWORD
    k32.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
    k32.GetExitCodeProcess.restype = wintypes.BOOL
    k32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    k32.TerminateProcess.restype = wintypes.BOOL
    k32.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)
    ]
    k32.QueryFullProcessImageNameW.restype = wintypes.BOOL
    return k32


def acquire_instance_mutex(mutex_name=DEFAULT_MUTEX_NAME):
    """
    Acquire a Windows kernel named mutex for single-instance protection.
    Returns (acquired: bool, handle: int | None).
    """
    if os.name != "nt":
        return True, None
    k32 = _get_kernel32()
    ctypes.set_last_error(0)
    handle = k32.CreateMutexW(None, False, str(mutex_name))
    last_err = ctypes.get_last_error()
    if not handle or last_err in (_ERROR_ALREADY_EXISTS, _ERROR_ACCESS_DENIED):
        if handle:
            try:
                k32.CloseHandle(handle)
            except Exception:
                pass
        return False, None
    return True, handle


def release_instance_mutex(handle):
    """Release an acquired Windows named mutex handle."""
    if not handle or os.name != "nt":
        return
    try:
        k32 = _get_kernel32()
        k32.CloseHandle(handle)
    except Exception:
        pass


def read_runtime_info(runtime_file=DEFAULT_RUNTIME_FILE):
    """Read and validate .server_runtime.json; returns dict or None."""
    if not runtime_file or not os.path.exists(runtime_file):
        return None
    try:
        with open(runtime_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return None
        if data.get("app_id") != APP_ID:
            return None
        port = int(data.get("port", 0))
        if port <= 0 or port > 65535:
            return None
        return data
    except Exception:
        return None


def write_runtime_info(port, pid=None, runtime_file=DEFAULT_RUNTIME_FILE):
    """Atomically write current server runtime metadata (pid, port, app_id)."""
    target_pid = int(pid or os.getpid())
    target_port = int(port)
    payload = {
        "app_id": APP_ID,
        "pid": target_pid,
        "port": target_port,
        "url": f"http://127.0.0.1:{target_port}",
        "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    parent = os.path.dirname(os.path.abspath(runtime_file))
    if parent:
        os.makedirs(parent, exist_ok=True)
    tmp_path = f"{runtime_file}.tmp.{os.getpid()}"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, runtime_file)
    return payload


def remove_runtime_info(runtime_file=DEFAULT_RUNTIME_FILE, expected_pid=None):
    """Remove runtime file, optionally guarding on owning PID."""
    if not runtime_file or not os.path.exists(runtime_file):
        return
    if expected_pid is not None:
        try:
            with open(runtime_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and data.get("pid") not in (None, int(expected_pid)):
                return
        except Exception:
            pass
    try:
        os.remove(runtime_file)
    except Exception:
        pass


def probe_instance(port, timeout=1.5):
    """
    Probe http://127.0.0.1:{port}/api/status bypassing system proxies.
    Returns parsed status dict if it belongs to this application, else None.
    """
    try:
        p = int(port)
    except (TypeError, ValueError):
        return None
    if p <= 0 or p > 65535:
        return None
    url = f"http://127.0.0.1:{p}/api/status"
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    req = urllib.request.Request(url, method="GET")
    try:
        with opener.open(req, timeout=timeout) as resp:
            if getattr(resp, "status", 200) != 200:
                return None
            raw = resp.read().decode("utf-8", errors="replace")
        data = json.loads(raw)
        if isinstance(data, dict) and data.get("success") is True and data.get("app_id") == APP_ID:
            return data
    except Exception:
        return None
    return None


def discover_running_instance(
    base_port=8765,
    runtime_file=DEFAULT_RUNTIME_FILE,
    max_wait=8.0,
    poll_interval=0.2,
    probe_timeout=1.5,
    port_span=10,
):
    """
    Wait briefly and locate an active verified instance of this application.
    Reads runtime_file first, then scans [base_port .. base_port + port_span - 1].
    Returns (port: int | None, status_data: dict | None).
    """
    deadline = time.monotonic() + max(0.2, float(max_wait))
    attempt = 0
    while True:
        checked_ports = set()
        rt = read_runtime_info(runtime_file=runtime_file)
        if rt:
            rt_port = int(rt["port"])
            checked_ports.add(rt_port)
            info = probe_instance(rt_port, timeout=probe_timeout)
            if info:
                return rt_port, info

        # Probe base_port on every attempt; scan full fallback span after first 1.0s or every 3 attempts
        if attempt == 0 or attempt % 3 == 0 or not rt:
            scan_ports = [int(base_port) + i for i in range(port_span)] if (attempt >= 2 or not rt) else [int(base_port)]
            for cand in scan_ports:
                if cand in checked_ports:
                    continue
                checked_ports.add(cand)
                info = probe_instance(cand, timeout=min(probe_timeout, 0.8))
                if info:
                    return cand, info

        if time.monotonic() >= deadline:
            break
        attempt += 1
        time.sleep(poll_interval)

    return None, None


def _is_pid_alive(pid):
    if not pid or int(pid) <= 0:
        return False
    if os.name != "nt":
        try:
            os.kill(int(pid), 0)
            return True
        except OSError:
            return False
    k32 = _get_kernel32()
    h = k32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION | _SYNCHRONIZE, False, int(pid))
    if not h:
        return False
    try:
        from ctypes import wintypes
        code = wintypes.DWORD(0)
        if k32.GetExitCodeProcess(h, ctypes.byref(code)):
            return code.value == _STILL_ACTIVE
        return True
    finally:
        k32.CloseHandle(h)


def _wait_for_pid_exit(pid, timeout=2.5):
    if not pid or int(pid) <= 0:
        return True
    if os.name != "nt":
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if not _is_pid_alive(pid):
                return True
            time.sleep(0.1)
        return not _is_pid_alive(pid)
    k32 = _get_kernel32()
    h = k32.OpenProcess(_SYNCHRONIZE, False, int(pid))
    if not h:
        return True
    try:
        res = k32.WaitForSingleObject(h, int(max(0.1, timeout) * 1000))
        return res == 0
    finally:
        k32.CloseHandle(h)


def _get_process_image_name(pid):
    if os.name != "nt" or not pid or int(pid) <= 0:
        return ""
    k32 = _get_kernel32()
    h = k32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not h:
        return ""
    try:
        from ctypes import wintypes
        buf = ctypes.create_unicode_buffer(1024)
        sz = wintypes.DWORD(1024)
        if k32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(sz)):
            return buf.value
        return ""
    finally:
        k32.CloseHandle(h)


def _is_python_listening_on_port(pid, port):
    """Verify via Win32 image name + netstat that pid is a Python process listening on port."""
    if os.name != "nt" or not pid or not port:
        return False
    img_name = os.path.basename(_get_process_image_name(pid)).lower()
    if img_name not in ("python.exe", "pythonw.exe"):
        return False
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        net = subprocess.run(
            ["netstat", "-aon"],
            stdin=subprocess.DEVNULL,
            capture_output=True,
            text=True,
            encoding="gbk",
            errors="replace",
            creationflags=flags,
            timeout=5,
        )
        port_token = f":{int(port)}"
        pid_str = str(int(pid))
        for line in (net.stdout or "").splitlines():
            parts = line.split()
            if len(parts) >= 5 and port_token in parts[1] and "LISTENING" in parts[3].upper() and parts[4] == pid_str:
                return True
        return False
    except Exception:
        return False


def _kill_pid(pid):
    if not pid or int(pid) <= 0 or int(pid) == os.getpid():
        return False
    if os.name == "nt":
        k32 = _get_kernel32()
        h = k32.OpenProcess(_PROCESS_TERMINATE, False, int(pid))
        if h:
            try:
                if k32.TerminateProcess(h, 1):
                    return True
            finally:
                k32.CloseHandle(h)
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            res = subprocess.run(
                ["taskkill", "/F", "/PID", str(int(pid))],
                stdin=subprocess.DEVNULL,
                capture_output=True,
                creationflags=flags,
                timeout=3,
            )
            return res.returncode == 0
        except Exception:
            return False
    try:
        os.kill(int(pid), 9)
        return True
    except Exception:
        return False


def stop_running_instance(base_port=None, runtime_file=DEFAULT_RUNTIME_FILE, port_span=10):
    """
    Gracefully stop the running dermatology-system instance via /api/system/exit,
    falling back to targeted PID termination if needed.
    Returns (stopped: bool, message: str).
    """
    if base_port is None:
        try:
            from app.config import load_config
            base_port = int(load_config().get("server_port", 8765) or 8765)
        except Exception:
            base_port = 8765

    rt = read_runtime_info(runtime_file=runtime_file)
    candidate_ports = []
    if rt and rt.get("port"):
        candidate_ports.append(int(rt["port"]))
    for i in range(port_span):
        p = int(base_port) + i
        if p not in candidate_ports:
            candidate_ports.append(p)

    target_port = None
    status_info = None
    for p in candidate_ports:
        info = probe_instance(p, timeout=0.8)
        if info:
            target_port = p
            status_info = info
            break

    if target_port is not None:
        target_pid = status_info.get("pid") or (rt.get("pid") if rt else None)
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        req = urllib.request.Request(
            f"http://127.0.0.1:{target_port}/api/system/exit",
            data=b"{}",
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with opener.open(req, timeout=2.0):
                pass
        except Exception:
            pass

        exited = False
        if target_pid:
            exited = _wait_for_pid_exit(target_pid, timeout=2.5)
        else:
            deadline = time.monotonic() + 2.5
            while time.monotonic() < deadline:
                if probe_instance(target_port, timeout=0.4) is None:
                    exited = True
                    break
                time.sleep(0.15)

        if not exited and target_pid:
            _kill_pid(target_pid)
            _wait_for_pid_exit(target_pid, timeout=1.5)

        remove_runtime_info(runtime_file=runtime_file)
        pid_desc = f", PID: {target_pid}" if target_pid else ""
        return True, f"已安全停止皮肤病理报告系统服务 (端口: {target_port}{pid_desc})"

    # Fallback: runtime file points to a hung Python process still listening on that port
    if rt and rt.get("pid") and rt.get("port"):
        rt_pid = int(rt["pid"])
        rt_port = int(rt["port"])
        if _is_python_listening_on_port(rt_pid, rt_port):
            _kill_pid(rt_pid)
            _wait_for_pid_exit(rt_pid, timeout=1.5)
            remove_runtime_info(runtime_file=runtime_file)
            return True, f"已强制结束无响应的系统服务进程 (端口: {rt_port}, PID: {rt_pid})"
        remove_runtime_info(runtime_file=runtime_file)
        return False, "未发现运行中的皮肤病理报告系统服务（已清理残留状态文件）。"

    return False, "未发现运行中的皮肤病理报告系统服务。"


if __name__ == "__main__":
    if sys.stdout is not None and hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(errors="replace")
        except Exception:
            pass
    if "--stop" in sys.argv:
        stopped, msg = stop_running_instance()
        print(msg)
        sys.exit(0)
