# Personal Portfolio Analyzer

Personal Portfolio Analyzer 是一個使用 Streamlit 製作的個人投資組合分析工具。使用者可以上傳持倉 CSV，系統會自動抓取最新股價、將美股換算成台幣，並產生摘要卡片、持倉報表、配置圖表與 Excel 報表。

## 專案截圖

![Personal Portfolio Analyzer Dashboard](portfolio_dashboard.png)

## 主要功能

- 上傳個人持倉 CSV
- 使用 yfinance 抓取最新股價
- 美股依 USD/TWD 匯率換算成台幣
- 顯示總市值、總成本、總損益、總報酬率等摘要卡片
- 顯示高集中持股提醒
- 顯示持倉明細報表
- 顯示市場配置與 category 類別配置
- 產生投資組合 Dashboard PNG
- 下載多工作表 Excel 報表
- 下載範例 CSV
- 若部分股票抓不到價格，會在報表中標示並暫用成本估算

## CSV 欄位格式

CSV 檔案需要包含以下欄位：

| 欄位 | 說明 | 範例 |
| --- | --- | --- |
| symbol | 股票代號 | NVDA, TSLA, 0050 |
| market | 市場，只能填 US 或 TW | US |
| category | 投資分類 | AI基礎建設 |
| shares | 股數 | 10 |
| cost | 平均成本 | 150 |

範例：

```csv
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

## 隱私提醒

不要把真實持股資料放在 public GitHub repo。若要展示作品，建議使用 `sample_portfolio.csv` 或自行製作的假資料。
