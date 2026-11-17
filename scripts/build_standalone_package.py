# -*- coding: utf-8 -*-
"""
Dermatology Pathology Reporting System - Standalone Portable Package Builder
Builds a zero-install, zero-dependency portable distribution under:
dist/皮肤病理报告系统-绿色免安装版/
"""
import os
import sys
import shutil

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIST_DIR = os.path.join(BASE_DIR, "dist")
PACKAGE_DIR = os.path.join(DIST_DIR, "皮肤病理报告系统-绿色免安装版")
PYTHON_DIR = os.path.dirname(sys.executable)

def build_package():
    print(f"[Build] 开始构建独立便携软件包...")
    print(f"[Build] 源目录: {BASE_DIR}")
    print(f"[Build] 目标目录: {PACKAGE_DIR}")
    print(f"[Build] 检测到宿主 Python 目录: {PYTHON_DIR}")

    if os.path.exists(PACKAGE_DIR):
        print("[Build] 清理旧的构建产物...")
        shutil.rmtree(PACKAGE_DIR, ignore_errors=True)

    os.makedirs(PACKAGE_DIR, exist_ok=True)
    runtime_dir = os.path.join(PACKAGE_DIR, "runtime")
    os.makedirs(runtime_dir, exist_ok=True)

    # 1. Copy Core Python Binaries & DLLs
    binaries = [
        "python.exe", "pythonw.exe", "python38.dll", "python3.dll", "vcruntime140.dll"
    ]
    for b in binaries:
        src = os.path.join(PYTHON_DIR, b)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(runtime_dir, b))
            print(f"  -> 复制核心文件: {b}")

    # Copy DLLs
    dlls_src = os.path.join(PYTHON_DIR, "DLLs")
    if os.path.exists(dlls_src):
        dlls_dst = os.path.join(runtime_dir, "DLLs")
        os.makedirs(dlls_dst, exist_ok=True)
        for item in os.listdir(dlls_src):
            if item.endswith((".pyd", ".dll")):
                shutil.copy2(os.path.join(dlls_src, item), os.path.join(dlls_dst, item))
        print("  -> 复制 DLLs 依赖组件")

    # 2. Copy Lib (Standard Library & site-packages for python-docx)
    lib_src = os.path.join(PYTHON_DIR, "Lib")
    lib_dst = os.path.join(runtime_dir, "Lib")
    os.makedirs(lib_dst, exist_ok=True)

    # Required stdlib packages/modules
    required_stdlib = [
        "encodings", "http", "urllib", "json", "sqlite3", "xml", "ctypes", "logging", "asyncio",
        "collections", "concurrent", "distutils", "email", "html", "importlib", "wsgiref"
    ]
    for name in os.listdir(lib_src):
        p_src = os.path.join(lib_src, name)
        if os.path.isfile(p_src) and name.endswith(".py"):
            shutil.copy2(p_src, os.path.join(lib_dst, name))
        elif os.path.isdir(p_src) and name in required_stdlib:
            shutil.copytree(p_src, os.path.join(lib_dst, name), dirs_exist_ok=True)

    # Copy site-packages for docx
    sp_src = os.path.join(lib_src, "site-packages")
    sp_dst = os.path.join(lib_dst, "site-packages")
    os.makedirs(sp_dst, exist_ok=True)

    needed_packages = [
        "docx", "python_docx-1.1.2.dist-info", "lxml", "typing_extensions.py",
        "openpyxl", "openpyxl-3.1.5.dist-info", "et_xmlfile", "et_xmlfile-2.0.0.dist-info"
    ]
    for pkg in needed_packages:
        src_pkg = os.path.join(sp_src, pkg)
        if os.path.exists(src_pkg):
            if os.path.isdir(src_pkg):
                shutil.copytree(src_pkg, os.path.join(sp_dst, pkg), dirs_exist_ok=True)
            else:
                shutil.copy2(src_pkg, os.path.join(sp_dst, pkg))
            print(f"  -> 复制第三方依赖库: {pkg}")

    # 3. Copy Application Code & Assets (Simplified, Clean Portable Package)
    print("[Build] 复制系统源码及界面静态资源...")
    shutil.copytree(os.path.join(BASE_DIR, "app"), os.path.join(PACKAGE_DIR, "app"), dirs_exist_ok=True)
    shutil.copytree(os.path.join(BASE_DIR, "web"), os.path.join(PACKAGE_DIR, "web"), dirs_exist_ok=True)
    
    # 仅复制有效 docx 模板与图片占位符，不包含旧版二进制 .doc 文件
    os.makedirs(os.path.join(PACKAGE_DIR, "模板"), exist_ok=True)
    tmpl_src = os.path.join(BASE_DIR, "模板")
    if os.path.exists(tmpl_src):
        for item in os.listdir(tmpl_src):
            if item.endswith((".docx", ".jpeg", ".jpg", ".png")):
                shutil.copy2(os.path.join(tmpl_src, item), os.path.join(PACKAGE_DIR, "模板", item))
                print(f"  -> 复制标准排版模板: {item}")

    if os.path.exists(os.path.join(BASE_DIR, "文字模板")):
        shutil.copytree(os.path.join(BASE_DIR, "文字模板"), os.path.join(PACKAGE_DIR, "文字模板"), dirs_exist_ok=True)
        print("  -> 复制文字模板词库")

    # 规范工作目录结构（纯净初始工作环境，不含示例文档与冗余数据）
    shutil.copytree(os.path.join(BASE_DIR, "工作目录"), os.path.join(PACKAGE_DIR, "工作目录"), dirs_exist_ok=True)
    if os.path.exists(os.path.join(BASE_DIR, "文档")):
        shutil.copytree(os.path.join(BASE_DIR, "文档"), os.path.join(PACKAGE_DIR, "文档"), dirs_exist_ok=True)
        print("  -> 复制文档目录至安装包根目录")

    # 清理工作目录下可能存在的残留文档及已归档报告示例文档
    pkg_work_doc = os.path.join(PACKAGE_DIR, "工作目录", "文档")
    if os.path.exists(pkg_work_doc):
        shutil.rmtree(pkg_work_doc, ignore_errors=True)
    pkg_archived = os.path.join(PACKAGE_DIR, "工作目录", "已归档报告")
    if os.path.exists(pkg_archived):
        for f in os.listdir(pkg_archived):
            f_path = os.path.join(pkg_archived, f)
            if os.path.isfile(f_path):
                try:
                    os.remove(f_path)
                except Exception:
                    pass

    shutil.copy2(os.path.join(BASE_DIR, "main.py"), os.path.join(PACKAGE_DIR, "main.py"))
    if os.path.exists(os.path.join(BASE_DIR, "README.md")):
        shutil.copy2(os.path.join(BASE_DIR, "README.md"), os.path.join(PACKAGE_DIR, "README.md"))

    # 4. Generate Launcher Scripts in Package Root
    # Batch Launcher
    bat_content = """@echo off
chcp 65001 >nul
title 皮肤病理报告系统
cd /d "%~dp0"
echo ======================================================
echo           皮肤病理报告系统 (便携免安装版)
echo ======================================================
echo 正在启动服务并在默认浏览器打开系统页面...
echo.
"%~dp0runtime\\python.exe" main.py
if %errorlevel% neq 0 (
    echo.
    echo 系统运行异常退出，代码: %errorlevel%
    pause
)
"""
    with open(os.path.join(PACKAGE_DIR, "启动系统.bat"), "w", encoding="utf-8") as f:
        f.write(bat_content)

    # Stop System Batch
    stop_bat_content = """@echo off
chcp 65001 >nul
title 停止皮肤病理报告系统
echo ======================================================
echo           正在安全退出皮肤病理报告系统...
echo ======================================================
echo.
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8765"') do (
    if not "%%a"=="0" (
        taskkill /f /pid %%a >nul 2>nul
        echo 已成功终止占用 8765 端口的服务进程 (PID: %%a)
    )
)
echo.
echo 系统服务已安全停止！
timeout /t 2 >nul
"""
    with open(os.path.join(PACKAGE_DIR, "停止系统.bat"), "w", encoding="utf-8") as f:
        f.write(stop_bat_content)

    # VBS Silent Launcher (No black CMD box)
    vbs_content = """Set ws = CreateObject("Wscript.Shell")
ws.currentdirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
ws.run "runtime\\pythonw.exe main.py", 0
"""
    with open(os.path.join(PACKAGE_DIR, "启动系统(无黑框后台版).vbs"), "w", encoding="utf-8") as f:
        f.write(vbs_content)

    # Readme
    readme_content = """============================================================
              皮肤病理报告系统 (绿色免安装便携版)
============================================================

【系统简介】
本软件包为绿色便携版，已内置轻量独立运行环境：
- 无需安装 Python
- 无需配置环境变量
- 无需部署依赖库
- 无需联网即可完整使用

【快速启动】
1. 双击运行【启动系统.bat】（推荐，带运行状态控制台）；
   或双击【启动系统(无黑框后台版).vbs】（静默在后台启动）；
2. 系统会自动启动本地微服务，并自动调用系统默认浏览器打开界面：
   http://127.0.0.1:8765

【数据与模板】
- 模板文件存放在：.\\模板 目录（支持1图、2图、3图及以上模板）；
- 文字模板存放在：.\\文字模板 目录（内置249篇皮肤病理专业诊断词库）；
- 文档手册存放在：.\\文档 目录（系统全景架构手册与操作指南）；
- 归档报告存放于：.\\工作目录\\已归档报告\\ 目录（原报告平铺直存，独立免疫组化报告自动追加 _免疫组化报告 后缀并存）；
- 未归档工作区：.\\工作目录\\未归档报告\\ 目录（包含录入草稿、病理照片、输出初稿、临床照片）；
- 疑难/质控与组化目录：.\\工作目录\\特殊患者归档\\ 目录（待审核 / 重切 / 免疫组化 / 完成审核，仅复制病理照片）；
- 数据库位于：.\\工作目录\\数据库\\pathology.db（支持高并发WAL模式与安全回收站软删除）。

【关闭系统】
在命令行窗口中按 Ctrl+C，或直接双击运行【停止系统.bat】即可安全退出。
"""
    with open(os.path.join(PACKAGE_DIR, "使用说明.txt"), "w", encoding="utf-8") as f:
        f.write(readme_content)

    print("\n[Build] Standalone 便携软件包打包完成！")
    print(f"[Build] 软件包文件夹: {PACKAGE_DIR}")

    # Generate ZIP archive
    zip_target = os.path.join(DIST_DIR, "皮肤病理报告系统-绿色免安装版")
    print(f"[Build] 正在打包压缩为 ZIP 文件: {zip_target}.zip ...")
    shutil.make_archive(zip_target, "zip", root_dir=PACKAGE_DIR)
    print(f"[Build] ZIP 压缩包生成完成: {zip_target}.zip")

if __name__ == "__main__":
    build_package()
