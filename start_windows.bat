@echo off
cd /d "%~dp0"
echo ========================================
echo  二手車進口報價系統 v1.10 - Windows 啟動
echo  多人登入版：完稅價格修正版
echo ========================================

where python >nul 2>nul
if errorlevel 1 (
  echo 找不到 Python。請先安裝 Python 3.10 以上版本。
  pause
  exit /b 1
)

if not exist .venv (
  echo 第一次啟動：建立虛擬環境 .venv ...
  python -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo.
echo 啟動系統中，瀏覽器將自動開啟。
echo 本機網址：http://localhost:8501
echo 多人使用：請讓其他業務連到這台電腦的區網 IP，例如 http://主機IP:8501
echo 預設老闆帳號 darren / 密碼 oma1688；業務密碼預設 1234。
echo.
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501
pause
