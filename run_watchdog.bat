@echo off
chcp 65001 >nul
title 中控检测数据监控 - 正在运行 V3.2

echo ================================================
echo   中控检测数据监控报警器 V3.2
echo   按 Ctrl+C 停止监控
echo ================================================
echo.

:loop
python watchdog.py --once
echo.
echo ⏳ 等待 %WATCH_INTERVAL% 秒后重新扫描...
echo   按 Ctrl+C 退出

rem 从 config.json 读取间隔（默认60秒）
setlocal enabledelayedexpansion
for /f "tokens=2 delims=:," %%a in ('findstr "interval_seconds" config.json') do set "interval=%%a"
set "interval=%interval: =%"
if "%interval%"=="" set "interval=60"
endlocal & set "WATCH_INTERVAL=%interval%"

timeout /t %WATCH_INTERVAL% /nobreak >nul
goto loop
