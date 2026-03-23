#!/bin/bash
# macOS / Linux 快速啟動腳本
# 自動檢查 Python、安裝依賴、啟動 Streamlit

echo "============================================================"
echo "🏛️  立法委員發言自動統計系統 - 啟動程序"
echo "============================================================"
echo ""

# 檢查 Python 是否安裝
if ! command -v python3 &> /dev/null; then
    echo "❌ 找不到 Python 3！"
    echo ""
    echo "請先安裝 Python 3.8 以上版本："
    echo ""
    echo "macOS 使用者："
    echo "  brew install python"
    echo ""
    echo "Ubuntu/Debian 使用者："
    echo "  sudo apt-get install python3 python3-pip"
    echo ""
    exit 1
fi

PYTHON_VERSION=$(python3 --version 2>&1)
echo "✅ 偵測到 Python：$PYTHON_VERSION"
echo ""

# 安裝依賴
echo "⚙️  正在安裝所需套件..."
python3 -m pip install -r requirement.txt
if [ $? -ne 0 ]; then
    echo "❌ 套件安裝失敗"
    exit 1
fi
echo "✅ 套件安裝完成"
echo ""

# 啟動 Streamlit
echo "🚀 啟動 Streamlit 應用..."
python3 -m streamlit run app.py
