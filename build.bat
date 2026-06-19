@echo off
chcp 65001 >nul 2>&1
title InvoiceRenamer 打包脚本

echo ============================================
echo   InvoiceRenamer 打包脚本
echo ============================================
echo.

REM ── 定位 venv Python ──────────────────────────
set "VENV_PY=C:\Users\29292\.workbuddy\binaries\python\envs\invoice-renamer\Scripts\python.exe"

if not exist "%VENV_PY%" (
    echo [错误] 找不到 venv Python: %VENV_PY%
    echo 请确认虚拟环境已创建。
    pause
    exit /b 1
)

echo [1/3] 清理旧的打包产物...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo     已清理。
echo.

echo [2/3] 开始打包（无 UPX 压缩）...
"%VENV_PY%" -m PyInstaller InvoiceRenamer.spec --noconfirm
if errorlevel 1 (
    echo.
    echo [错误] 打包失败！
    pause
    exit /b 1
)
echo.

echo [3/3] 打包结果:
set "FOUND_EXE="
for %%f in ("dist\InvoiceRenamer_*.exe") do set "FOUND_EXE=%%f"
if defined FOUND_EXE (
    for %%A in ("%FOUND_EXE%") do (
        echo     文件: %FOUND_EXE%
        echo     大小: %%~zA 字节
    )
    REM 换算 MB
    for /f "delims=" %%S in ('powershell -nop -c "[math]::Round((Get-Item '%FOUND_EXE%').Length / 1MB, 1)"') do set SIZE_MB=%%S
    echo       = %SIZE_MB% MB
) else (
    echo [错误] 未找到输出文件 dist\InvoiceRenamer_*.exe
)
echo.
echo ============================================
echo   打包完成！
echo ============================================
pause
