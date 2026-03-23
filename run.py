#!/usr/bin/env python3
"""
自動啟動腳本：檢查環境 → 安裝依賴 → 啟動應用
支援三層降級：Streamlit → Flask → 命令行工具
適用於 Windows / macOS / Linux
"""
import sys
import subprocess
import os
from pathlib import Path


def check_python_version():
    """檢查 Python 版本是否 >= 3.8"""
    if sys.version_info < (3, 8):
        print(f"❌ Python 版本過低：{sys.version}")
        print("   請確保安裝 Python 3.8 以上")
        sys.exit(1)
    print(f"✅ Python 版本檢查通過：{sys.version.split()[0]}")


def install_requirements():
    """自動安裝 requirement.txt 中的所有套件"""
    requirements_file = Path(__file__).parent / "requirement.txt"
    
    if not requirements_file.exists():
        print(f"❌ 找不到 requirement.txt：{requirements_file}")
        sys.exit(1)
    
    print("\n📦 正在安裝所需套件...")
    
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "-r", str(requirements_file)],
            capture_output=False,
            text=True
        )
        if result.returncode != 0:
            print("❌ 套件安裝失敗")
            sys.exit(1)
        print("✅ 套件安裝完成")
    except Exception as e:
        print(f"❌ 套件安裝出錯：{e}")
        sys.exit(1)


def launch_streamlit():
    """嘗試啟動 Streamlit 應用"""
    app_file = Path(__file__).parent / "app.py"
    
    if not app_file.exists():
        return False
    
    print("\n🚀 嘗試啟動 Streamlit 應用...")
    print(f"   應用檔案：{app_file}")
    
    try:
        result = subprocess.run(
            [sys.executable, "-m", "streamlit", "run", str(app_file)],
            cwd=str(app_file.parent),
            capture_output=True,
            timeout=10
        )
        if result.returncode == 0:
            return True
    except Exception as e:
        print(f"   ⚠️  Streamlit 啟動失敗：{e}")
    
    return False


def launch_flask():
    """降級方案 1：啟動 Flask 簡化版應用"""
    app_file = Path(__file__).parent / "app_flask.py"
    
    if not app_file.exists():
        return False
    
    print("\n🔄 Streamlit 不可用，改用 Flask 簡化版...")
    print("   應用檔案：app_flask.py")
    
    try:
        subprocess.run(
            [sys.executable, str(app_file)],
            cwd=str(app_file.parent)
        )
        return True
    except Exception as e:
        print(f"   ⚠️  Flask 啟動失敗：{e}")
    
    return False


def suggest_offline_tool():
    """降級方案 2：建議使用命令行工具"""
    offline_file = Path(__file__).parent / "analyze_offline.py"
    
    if not offline_file.exists():
        return False
    
    print("\n" + "=" * 60)
    print("⚠️  網頁版本無法啟動")
    print("=" * 60)
    print("\n改用命令行離線分析工具：")
    print(f"\n  python {offline_file.name} --label <名單.xlsx> --transcript <逐字稿.txt> --output <輸出.xlsx>")
    print("\n或簡化版：")
    print(f"  python {offline_file.name}")
    print("\n詳見：README.md")
    print("=" * 60)
    
    return True


if __name__ == "__main__":
    print("=" * 60)
    print("🏛️  立法委員發言自動統計系統 - 啟動程序")
    print("=" * 60)
    
    # 步驟 1：檢查 Python 版本
    print("\n⚙️  檢查環境...")
    check_python_version()
    
    # 步驟 2：安裝依賴
    install_requirements()
    
    # 步驟 3：嘗試啟動應用（三層降級）
    print("\n" + "=" * 60)
    print("🎯 啟動應用")
    print("=" * 60)
    
    # 第一層：Streamlit（最佳體驗）
    if launch_streamlit():
        sys.exit(0)
    
    # 第二層：Flask（輕量級降級）
    if launch_flask():
        sys.exit(0)
    
    # 第三層：建議使用命令行工具
    suggest_offline_tool()
    sys.exit(1)
