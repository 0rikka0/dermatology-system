@echo off
chcp 65001 >nul
title 皮肤病理报告系统

echo ======================================================
echo           皮肤病理报告系统 正在启动...
echo ======================================================
echo.

cd /d "%~dp0"

:: 检查本地是否存在内嵌便携版 Python (dist/python/python.exe)
if exist "%~dp0dist\python\python.exe" (
    echo [启动模式] 使用内置便携 Python 环境...
    "%~dp0dist\python\python.exe" main.py
    goto end
)

:: 检查系统环境变量中的 Python
where python >nul 2>nul
if %errorlevel% equ 0 (
    echo [启动模式] 使用系统 Python 环境...
    python main.py
    goto end
)

:: 未检测到 Python 环境时的友好提示
echo [错误] 未在系统中检测到 Python 运行环境。
echo 请确保本机已安装 Python 3.8+，或将便携版 Python 放置于 dist\python 目录。
echo.
pause

:end
