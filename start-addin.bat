@echo off
rem 启动 Slilot的本地服务（保持此窗口开着，或最小化）
cd /d "%~dp0"
echo [Slilot] 正在启动本地服务 http://localhost:3010 ...
node server.js
pause
