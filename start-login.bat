@echo off
setlocal enabledelayedexpansion
chcp 936 >nul 2>&1
title 秦巴药都数据平台 - 本地服务
cd /d "%~dp0"

echo.
echo   ============================================================
echo      秦巴药都数据平台 · 本地服务启动器
echo   ============================================================
echo.

if not exist "%~dp0server.ps1" (
  echo   [错误] 未找到 server.ps1，请确认它与本文件在同一目录。
  echo.
  pause
  exit /b 1
)
if not exist "%~dp0data\homeData.xlsx" (
  echo   [提示] 未找到 data\homeData.xlsx，页面可能无法显示数据。
  echo.
)

echo   正在检测可用端口...
set "PORT="
for /f "usebackq delims=" %%p in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "$ok=0; foreach($pt in 8000,8010,8020,8080,8090,8100,8110,8120,8130,8140,8150){ try { $l=[System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback,$pt); $l.Start(); $l.Stop(); $ok=$pt; break } catch {} }; Write-Output $ok"`) do set "PORT=%%p"

if "%PORT%"=="" (
  echo.
  echo   [错误] 8000-8150 之间的常用端口都被占用，无法启动本地服务。
  echo          请关闭占用端口的程序后重试。
  echo.
  pause
  exit /b 1
)
if "%PORT%"=="0" (
  echo.
  echo   [错误] 未找到空闲端口，无法启动本地服务。
  echo.
  pause
  exit /b 1
)

echo   使用端口: %PORT%
echo   登录地址: http://127.0.0.1:%PORT%/html/login.html
echo.
echo   浏览器将在 3 秒后自动打开；按 Ctrl+C 或直接关闭本窗口即可停止服务。
echo.

start "" /min cmd /c "timeout /t 3 /nobreak >nul & start "" http://127.0.0.1:%PORT%/html/login.html"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1" -Port %PORT%

echo.
echo   服务已停止。
echo.
pause
