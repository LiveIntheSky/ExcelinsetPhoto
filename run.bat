@echo off
chcp 65001 >nul
title SiC Excel 照片批次自動填入工具
cd /d "%~dp0"

echo ========================================================
echo   SiC Excel 照片批次自動填入工具 (純記憶體微伺服器)
echo   正在啟動本機伺服器並自動開啟瀏覽器介面...
echo ========================================================
echo.

if exist "%~dp0PhotoToExcelTool.exe" (
    "%~dp0PhotoToExcelTool.exe" %*
) else (
    echo [✕ 錯誤] 找不到 PhotoToExcelTool.exe！
    pause
)
