@echo off
REM Windows 快速啟動批次檔
REM 自動檢查 Python、安裝依賴、啟動 Streamlit

chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================================
echo.  立法委員發言自動統計系統 - Windows 啟動程序
echo ============================================================
echo.

REM 檢查 Python 是否安裝
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 找不到 Python！
    echo.
    echo 請先安裝 Python 3.8 以上版本：
    echo https://www.python.org/downloads/
    echo.
    echo 安裝時務必勾選「Add Python to PATH」
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version') do set PYTHON_VERSION=%%i
echo ✅ 偵測到 Python：%PYTHON_VERSION%
echo.

REM 安裝依賴
echo ⚙️  正在安裝所需套件...
python -m pip install -r requirement.txt
if errorlevel 1 (
    echo ❌ 套件安裝失敗
    pause
    exit /b 1
)
echo ✅ 套件安裝完成
echo.

REM 啟動 Streamlit
echo 🚀 啟動 Streamlit 應用...
python -m streamlit run app.py

pause
