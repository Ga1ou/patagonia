# Taiwan Financial Collector (Desktop GUI)

這是一個 Python 桌面應用程式，用來蒐集台灣公司季度財報與管理 EPS。

## V1 功能

- 固定追蹤公司: `8299`、`2330`、`2454`
- 蒐集季度: 支援自訂，預設包含 `2025Q3`、`2025Q4`、`2026Q1`
- 可選擇「加上目前最新季度」
- 本地 SQLite 資料庫儲存 (upsert 更新)
- EPS 管理:
  - 表格點擊輸入
  - 圖表拖拉紅色點位調整
  - 一鍵自動預估 (0.5 上季 + 0.3 去年同季 + 0.2 近四季均值)
- 匯出 CSV

## 專案結構

```text
hedge_fund/
├─ main.py
├─ requirements.txt
├─ README.md
├─ data/
├─ exports/
└─ app/
   ├─ constants.py
   ├─ quarters.py
   ├─ database.py
   ├─ estimators.py
   ├─ services.py
   ├─ ui.py
   └─ providers/
      ├─ base.py
      └─ yahoo_provider.py
```

## 安裝與執行

```powershell
cd "C:\Users\shaun.sun\Desktop\hedge_fund"
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## 使用流程

1. 左側勾選公司
2. 在「財報蒐集」輸入目標季度後按「蒐集財報資料」
3. 切到「EPS 管理」可:
   - 在表格直接輸入 EPS 預估值
   - 拖拉圖中紅色點位調整 EPS
   - 點「自動預估 EPS」補上缺值
4. 回到蒐集頁按「匯出 CSV」

## 注意事項

- V1 使用 Yahoo Finance 作為資料來源，部分欄位可能會有缺值。
- 若季度資料尚未公開，系統會建立 `manual_required` 佔位資料，供你手動補值。
- 資料庫路徑: `data/financial_records.db`

