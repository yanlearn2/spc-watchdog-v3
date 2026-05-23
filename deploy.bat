@echo off
chcp 65001 >nul
title 中控检测数据监控报警器 V3.2

echo ================================================
echo   中控检测数据监控报警器 V3.2
echo   部署脚本
echo ================================================
echo.

rem 检查 Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 未找到 Python，请先安装 Python 3.8+
    echo   下载: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✅ Python: 
python --version

rem 安装依赖
echo.
echo 📦 安装依赖...
python -m pip install -r requirements.txt -q
if %errorlevel% neq 0 (
    echo ⚠ pip 安装失败，尝试直接安装...
    python -m pip install openpyxl requests msoffcrypto-tool -q
)

echo ✅ 依赖就绪

rem 初始化数据库
echo.
echo 🗄️  初始化数据库...
python watchdog.py --init-db
if %errorlevel% neq 0 (
    echo ❌ 数据库初始化失败
    pause
    exit /b 1
)

rem 测试扫描
echo.
echo 🔍 测试扫描...
python watchdog.py --once
if %errorlevel% neq 0 (
    echo ⚠ 测试扫描出现错误，请检查 config.json 配置
    pause
    exit /b 1
)

echo.
echo ✅ 部署验证完成！
echo.
echo 📋 下一步：
echo   1. 编辑 config.json 配置 Excel 文件路径和白名单
echo   2. 运行 run_watchdog.bat 启动监控（前台循环）
echo   3. 或运行 watchdog.py --daemon 后台运行
echo   4. 或复制 watchdog.service 到 /etc/systemd/system/（Linux）
echo.
echo 🔧 V3.2 新特性:
echo   · 功能开关（CSV/并行/健康检查/锁文件均可开关）
echo   · 扫描范围控制（按日期/最大行数过滤）
echo   · 报警冷却 + DB持久化去重（重启后不丢失）
echo   · 启动配置校验（自动检查文件/白名单/Webhook）
echo   · 锁文件机制（防止多实例冲突）
echo   · 报警去重（同一异常不再重复通知）
echo   · 自动重连（数据库异常自动恢复）
echo   · 多文件并行扫描
echo.
pause
