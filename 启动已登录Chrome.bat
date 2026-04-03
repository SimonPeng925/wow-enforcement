@echo off
chcp 65001 >nul
echo ============================================
echo   WOW English 维权工具 - 启动调试Chrome
echo ============================================
echo.
echo 此脚本会用调试模式打开 Chrome 浏览器
echo 打开后请登录京东，然后运行提取脚本
echo.

REM 关闭已有的 Chrome 进程（可选，如果不想关闭请注释掉下面这行）
REM taskkill /F /IM chrome.exe >nul 2>&1

REM 启动 Chrome（带远程调试端口）
REM 优先使用常见安装路径
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\chrome-jd-debug"
    echo [OK] Chrome 已启动（Program Files）
    goto done
)
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\chrome-jd-debug"
    echo [OK] Chrome 已启动（Program Files x86）
    goto done
)
if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" (
    start "" "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\chrome-jd-debug"
    echo [OK] Chrome 已启动（LocalAppData）
    goto done
)

echo [ERROR] 未找到 Chrome，请手动打开 Chrome 并加上参数:
echo   --remote-debugging-port=9222
goto end

:done
echo.
echo [提示] Chrome 已打开，请：
echo   1. 在浏览器中登录京东
echo   2. 打开你要提取的侵权商品页
echo   3. 回到命令行运行：
echo      python jd_extract.py "商品链接"
echo.
echo ============================================

:end
pause
