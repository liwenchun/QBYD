@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo ==========================================
echo    syCode 图片清单生成工具  manifest.json
echo ==========================================
echo.

if exist "manifest.json" (
    del /f /q "manifest.json"
    echo [1/2] 已删除旧的 manifest.json
) else (
    echo [1/2] 未发现旧的 manifest.json，跳过删除
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$q=[char]34; $files=Get-ChildItem -File | Where-Object { $_.Name -ne 'manifest.json' -and $_.Extension.ToLower() -match '^\.(png|jpg|jpeg|bmp|gif|webp|svg)$' } | Sort-Object Name; $items=$files | ForEach-Object { $n=$_.Name -replace '\\','\\\\'; $n=$n -replace $q,('\'+$q); $q + $n + $q }; $json='[' + ($items -join ',') + ']'; [IO.File]::WriteAllText((Join-Path (Get-Location) 'manifest.json'),$json,(New-Object System.Text.UTF8Encoding($false))); Write-Host ('[2/2] manifest.json 已写入，图片数量：' + $files.Count)"

if exist "manifest.json" (
    echo.
    echo 生成成功：%CD%\manifest.json
    echo 内容如下：
    type "manifest.json"
    echo.
) else (
    echo.
    echo 生成失败：未生成 manifest.json，请确认 PowerShell 可用。
)

exit /b 0
