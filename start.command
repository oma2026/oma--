#!/bin/bash
set -e
cd "$(dirname "$0")"

clear
echo "========================================"
echo " 進口報價系統二手車 v1.10 - 一鍵啟動"
echo " 多人登入版：完稅價格修正版"
echo "========================================"

if ! command -v python3 >/dev/null 2>&1; then
  echo "找不到 python3。請先安裝 Python 3.10 以上版本。"
  echo "可到 https://www.python.org/downloads/ 安裝。"
  read -n 1 -s -r -p "按任意鍵結束..."
  exit 1
fi

if [ ! -d ".venv" ]; then
  echo "第一次啟動：建立虛擬環境 .venv ..."
  python3 -m venv .venv
fi

source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo ""
echo "啟動系統中，瀏覽器將自動開啟。"
echo "本機網址：http://localhost:8501"
echo "多人使用：請讓其他業務連到這台電腦的區網 IP，例如 http://主機IP:8501"
echo ""
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501
