# 二手車進口報價系統 v1.11 一鍵式多人版

## 版本重點

本版依照最新確認修正：完稅價格公式改為「(國外報價金額 + 1,500) × 匯率 + 關稅」，且 2000cc 以下關稅基數預設改為 0.545。

v1.11 修正：

1. 車輛基本資料增加「CC 數級距」：2000cc 以下／2000cc 以上。
2. 關稅改由系統自動計算，不再由業務手動輸入。
3. 關稅公式：

   ```text
   關稅 =（國外報價金額 + 1,500）× 關稅基數 × 匯率
   ```

4. 關稅基數預設：
   - 2000cc 以下：0.545
   - 2000cc 以上：0.61

5. 新增奢侈稅：

   ```text
   完稅價格 =（國外報價金額 + 1,500）× 匯率 + 關稅
   若完稅價格 > 3,000,000：奢侈稅 = 完稅價格 × 10%
   ```

6. 新增服務費：
   - TWD：直接輸入台幣金額。
   - USD：輸入美金金額與匯率，系統換算台幣。

7. 其他既有功能保留：多人登入、帳號管理、特殊車測、不限項數加裝配備、報價紀錄、Excel/CSV 匯出。

## 啟動方式

### Mac

1. 解壓縮 ZIP。
2. 進入資料夾 `used_car_import_quote_app_v1_10`。
3. 雙擊 `start.command`。
4. 若 Mac 安全性阻擋，請用右鍵 / Control + 點擊 `start.command`，選「打開」。

### Windows

1. 解壓縮 ZIP。
2. 進入資料夾 `used_car_import_quote_app_v1_10`。
3. 雙擊 `start_windows.bat`。

## 登入帳號

老闆：

```text
darren / oma1688
```

業務：

```text
peter / 1234
cbc / 1234
lai / 1234
gary / 1234
```

老闆可在「帳號管理」新增、停用、刪除帳號與修改密碼。

## 使用網址

本機使用：

```text
http://localhost:8501
```

同一個公司內網／Wi-Fi 的其他業務使用：

```text
http://主機IP:8501
```

不同網路或外出使用，需要雲端主機、VPN、Tailscale 或 Cloudflare Tunnel；不能直接用 `localhost`。

## 資料位置

```text
報價資料庫：data/quotes.db
後台設定：config/settings.json
```

## 關舊版方式

如果開 `http://localhost:8501` 仍看到舊版，先在 Mac 終端機執行：

```bash
lsof -ti:8501 | xargs kill -9
pkill -f streamlit
```

再重新從 v1.11 資料夾啟動 `start.command`。


## v1.11 更新
- 業務模式不再顯示關稅公式與完稅價格／奢侈稅公式的輔助說明。
- 老闆模式仍可看到上述公式，方便核對計算邏輯。
