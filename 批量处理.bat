@echo off
chcp 65001 >nul 2>&1
title WOW English 批量维权处理

echo.
echo ========================================
echo   WOW English 知识产权投诉 - 批量处理
echo ========================================
echo.

cd /d "%~dp0"

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 检查依赖
python -c "import playwright, openpyxl, easyocr" >nul 2>&1
if errorlevel 1 (
    echo [提示] 正在安装依赖库...
    pip install playwright openpyxl easyocr
    echo [提示] 安装浏览器驱动...
    playwright install chromium
)

:: 检查 Chrome 调试模式
netstat -an | findstr ":9222" >nul 2>&1
if errorlevel 1 (
    echo [提示] 未检测到 Chrome 调试模式
    echo   建议：先关闭 Chrome，双击「启动已登录Chrome.bat」再运行本脚本
    echo   或直接在浏览器登录电商平台，脚本会自动处理
    echo.
)

echo [开始] 启动批量处理...
echo.
python batch_ipp.py

echo.
echo ========================================
echo   处理完成！结果在 output/ 目录下
echo ========================================
pause
