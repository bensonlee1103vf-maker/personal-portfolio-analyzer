# Personal Portfolio Analyzer

Personal Portfolio Analyzer 是一個使用 Streamlit 製作的個人投資組合分析工具。使用者可以上傳持倉 Excel（.xlsx/.xls），系統會自動讀取表格、抓取最新股價、將美股換算成台幣，並產生摘要卡片、持倉報表、配置圖表與 Excel 報表。

## 專案截圖

![Personal Portfolio Analyzer Dashboard](portfolio_dashboard.png)

## 主要功能

- 上傳個人持倉 Excel (.xlsx/.xls)
- 使用 yfinance 抓取最新股價
- 美股依 USD/TWD 匯率換算成台幣
- 顯示總市值、總成本、總損益、總報酬率等摘要卡片
- 顯示高集中持股提醒
- 顯示持倉明細報表
- 顯示市場配置與 category 類別配置
- 產生投資組合 Dashboard PNG
- 下載多工作表 Excel 報表
- 下載範例 Excel
- 若部分股票抓不到價格，會在報表中標示並暫用成本估算

## Excel 欄位格式（Sheet 第一列為欄位名稱）

Excel（Sheet）需要包含以下欄位：

| 欄位 | 說明 | 範例 |
| --- | --- | --- |
| symbol | 股票代號 | NVDA, TSLA, 0050 |
| market | 市場，只能填 US 或 TW | US |
| category | 投資分類 | AI基礎建設 |
| shares | 股數 | 10 |
| cost | 平均成本 | 150 |

範例（請將以下欄位放在 Excel 的第一列作為欄位名稱）：

```
symbol,market,category,shares,cost
NVDA,US,AI基礎建設,10,150
AAPL,US,大型科技,5,180
QQQ,US,美股ETF,3,500
0050,TW,台股ETF,1000,180
006208,TW,台股ETF,500,110
```

## 如何在本機執行

1. 安裝 Python 3.10 或以上版本。
2. 安裝套件：

```bash
pip install -r requirements.txt
```

**注意**：`requirements.txt` 包含以下與 Excel 相關的套件，用於讀取與寫入 Excel 檔案：
- `openpyxl`：讀取、寫入 .xlsx 檔案（必須）
- `xlsxwriter`：產生 .xlsx 檔案且提供進階格式化功能（建議）

若安裝過程出現問題，可嘗試分別安裝：

```bash
pip install openpyxl xlsxwriter
```

3. 啟動 Streamlit：

```bash
streamlit run app.py
```

4. 在瀏覽器打開 Streamlit 顯示的網址，通常是 `http://localhost:8501`。

## 如何部署到 Streamlit Community Cloud

1. 將專案推到 GitHub repository。
2. 到 Streamlit Community Cloud 建立新 app。
3. 選擇你的 GitHub repository。
4. Main file path 設定為 `app.py`。
5. 確認 repository 中包含 `requirements.txt`。
6. 部署完成後，打開 Streamlit 提供的公開網址。

## 使用技術

- Python
- Streamlit
- pandas
- yfinance
- matplotlib
- openpyxl
- xlsxwriter

## 隱私提醒

不要把真實持股資料放在 public GitHub repo。若要展示作品，建議使用 `sample_portfolio.xlsx` 或自行製作的假資料。

## 常見問題與排除

### Q: 上傳 Excel 時出現「ModuleNotFoundError: No module named 'xlsxwriter'」

**A**: 請安裝 Excel 相關套件：

```bash
pip install openpyxl xlsxwriter
```

若仍無法產生範例 Excel，app 會自動改以 CSV 格式提供下載（具有相同資料內容），使用者仍可正常上傳檔案。

### Q: 為什麼側邊欄下載的範例檔是 CSV 而不是 Excel？

**A**: 這表示執行環境沒有安裝 `openpyxl` 或 `xlsxwriter`。請執行：

```bash
pip install -r requirements.txt
```

之後重新啟動 app。

### Q: 上傳 Excel 時出現「缺少必要欄位」錯誤，但我的 Excel 確實有這些欄位

**A**: 請確認：
1. 欄位名稱在第一列
2. 欄位名稱沒有前後空白（app 會自動 strip，但欄位名稱本身要正確）
3. 必要欄位為：`symbol`、`market`、`shares`、`cost`
4. `category` 是選填，若缺少會自動補「未分類」

### Q: 上傳 Excel 時出現「market 只能填 US 或 TW」

**A**: 請檢查 `market` 欄位的值是否只包含 `US` 或 `TW`（大小寫會自動轉成大寫）。例如 `uk`、`hk` 或其他市場代碼不被支援。

### Q: shares 或 cost 欄位有小數點或千分號時無法讀取

**A**: App 應該可以自動處理以下格式：
- 數字：`1000`、`150.5`
- 千分號：`1,000`、`1,234.56`
- 貨幣符號：`NT$1,000`、`$150.5`

若仍無法識別，可能是欄位內容包含其他特殊字元。請下載範例 Excel 確認格式。

### Q: 如何上傳包含多個工作表的 Excel？

**A**: App 會偵測 Excel 中的所有工作表，上傳後會出現「選擇要分析的工作表」下拉選單。選定工作表後，app 會讀取該工作表內的資料進行分析。
