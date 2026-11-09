@echo off
chcp 65001 >nul
rem 本地预览入口（已扩展）：统一由上一级 start-login.bat 启动服务（服务根=output 目录）
rem 并自动打开 http://localhost:8000/html/login.html ，可正常读取 data\usersData.xlsx

if exist "%~dp0..\start-login.bat" (
    call "%~dp0..\start-login.bat"
    exit /b 0
)

cd /d "%~dp0.."

echo ============================================
echo  秦巴药都 数据平台本地预览（登录页）
echo ============================================
echo.
echo  服务根目录：%~dp0..
echo  登录地址  ：http://localhost:8000/html/login.html
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 未检测到 python 命令，请先安装 Python。
    pause
    exit /b 1
)

start "" http://localhost:8000/html/login.html
python -m http.server 8000

pause
