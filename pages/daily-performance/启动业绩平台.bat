@echo off
chcp 65001 >nul
title 业绩日报表平台
cd /d "%~dp0"
echo ============================================
echo   各部门业绩完成情况跟踪日报表
echo   正在启动本地服务 (http://localhost:8800)...
echo   启动后请勿关闭本窗口，关闭即停止服务
echo ============================================
echo.
start http://localhost:8800/
"C:\Users\admin\.workbuddy\binaries\node\versions\22.22.2\node.exe" server.js
pause
