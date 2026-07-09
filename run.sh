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

# --- [修正開始] 自動使用虛擬環境安裝依賴 ---
VENV_DIR="./venv"
if [ ! -d "$VENV_DIR" ]; then
    echo "⚙️  正在創建虛擬環境..."
    python3 -m venv "$VENV_DIR"
fi

# 啟用虛擬環境並安裝依賴
echo "⚙️  正在使用虛擬環境安裝所需套件..."
source "$VENV_DIR/bin/activate"
if [ $? -ne 0 ]; then
    echo "❌ 套件安裝失敗 (Virtual Environment setup failed)"
    deactivate
    exit 1
fi

# 使用虛擬環境安裝依賴
pip install -r requirement.txt
if [ $? -ne 0 ]; then
    echo "❌ 套件安裝失敗"
    deactivate
    exit 1
fi
echo "✅ 套件安裝完成"
deactivate
# --- [修正結束] ---

# 啟動 Streamlit 應用 (明確使用虛擬環境內的 python)
echo "🚀 啟動 Streamlit 應用..."
"$VENV_DIR/bin/python" -m streamlit run app.py