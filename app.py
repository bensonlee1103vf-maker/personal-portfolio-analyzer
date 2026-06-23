from io import BytesIO

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
import yfinance as yf


BACKGROUND_COLOR = "#14232b"
REQUIRED_COLUMNS = ["symbol", "market", "category", "shares", "cost"]


def load_portfolio(uploaded_file):
    """讀取使用者上傳的 CSV，並檢查必要欄位是否存在。"""
    df = pd.read_csv(uploaded_file, dtype={"symbol": str, "market": str, "category": str})

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        missing_text = ", ".join(missing_columns)
        raise ValueError(f"CSV 缺少必要欄位：{missing_text}")

    df = df[REQUIRED_COLUMNS].copy()

    # 把 market 統一成大寫，避免使用者輸入 us / tw 時判斷失敗。
    df["market"] = df["market"].str.upper().str.strip()
    df["symbol"] = df["symbol"].str.strip()
    df["category"] = df["category"].fillna("").str.strip()
    df.loc[df["category"] == "", "category"] = "未分類"

    # shares 和 cost 必須是數字，無法轉換的資料會變成 NaN，後面再檢查。
    df["shares"] = pd.to_numeric(df["shares"], errors="coerce")
    df["cost"] = pd.to_numeric(df["cost"], errors="coerce")

    if df[["shares", "cost"]].isna().any().any():
        raise ValueError("shares 和 cost 欄位必須是數字。")

    invalid_market = sorted(set(df["market"]) - {"US", "TW"})
    if invalid_market:
        invalid_text = ", ".join(invalid_market)
        raise ValueError(f"market 只支援 US 或 TW，目前有：{invalid_text}")

    return df


def to_yahoo_symbol(row):
    """依 market 轉成 Yahoo Finance 使用的代號。"""
    if row["market"] == "TW":
        return f"{row['symbol'].zfill(4)}.TW"
    return row["symbol"]


@st.cache_data(ttl=900, show_spinner=False)
def get_latest_price(yahoo_symbol):
    """使用 yfinance 抓最新收盤價。"""
    stock = yf.Ticker(yahoo_symbol)
    data = stock.history(period="5d")
    close_prices = data["Close"].dropna() if not data.empty else pd.Series(dtype=float)

    if close_prices.empty:
        return None

    return float(close_prices.iloc[-1])


@st.cache_data(ttl=900, show_spinner=False)
def get_usd_twd_rate():
    """抓 USD/TWD 匯率，作為美股換算台幣的依據。"""
    fx_data = yf.Ticker("USDTWD=X").history(period="5d")
    close_prices = fx_data["Close"].dropna() if not fx_data.empty else pd.Series(dtype=float)

    if close_prices.empty:
        raise ValueError("無法取得 USD/TWD 匯率。")

    return float(close_prices.iloc[-1])


def calculate_portfolio(df, usd_twd):
    """計算價格、台幣市值、損益、報酬率與權重。"""
    result = df.copy()

    result["currency"] = result["market"].map({"US": "USD", "TW": "TWD"})
    result["yahoo_symbol"] = result.apply(to_yahoo_symbol, axis=1)
    result["price"] = result["yahoo_symbol"].apply(get_latest_price)

    missing_prices = result[result["price"].isna()]["yahoo_symbol"].tolist()
    if missing_prices:
        missing_text = ", ".join(missing_prices)
        raise ValueError(f"以下代號無法取得價格：{missing_text}")

    # 美股價格與成本用 USD/TWD 換成台幣，台股本來就是台幣。
    is_us = result["market"] == "US"
    result["price_twd"] = result["price"]
    result.loc[is_us, "price_twd"] = result.loc[is_us, "price"] * usd_twd

    result["cost_twd"] = result["cost"]
    result.loc[is_us, "cost_twd"] = result.loc[is_us, "cost"] * usd_twd

    result["market_value_twd"] = result["shares"] * result["price_twd"]
    result["cost_value_twd"] = result["shares"] * result["cost_twd"]
    result["profit_twd"] = result["market_value_twd"] - result["cost_value_twd"]
    result["return_rate"] = result["profit_twd"] / result["cost_value_twd"]
    result["weight"] = result["market_value_twd"] / result["market_value_twd"].sum()

    return result.sort_values("market_value_twd", ascending=False).reset_index(drop=True)


def make_top_n_summary(data, name_col, value_col, top_n):
    """取前 top_n 大項目，其餘合併成「其他」。"""
    if data.empty:
        return pd.DataFrame(columns=[name_col, value_col, "weight"])

    sorted_data = data.sort_values(value_col, ascending=False).copy()
    top_items = sorted_data.head(top_n).copy()
    other_items = sorted_data.iloc[top_n:].copy()

    summary = top_items[[name_col, value_col]].copy()

    if not other_items.empty:
        other_row = pd.DataFrame(
            [{name_col: "其他", value_col: other_items[value_col].sum()}]
        )
        summary = pd.concat([summary, other_row], ignore_index=True)

    summary["weight"] = summary[value_col] / summary[value_col].sum()
    return summary


def plot_donut_on_ax(
    ax,
    summary,
    name_col,
    value_col,
    title,
    center_text,
    label_size=10,
    pct_size=9,
    center_size=16,
    title_size=14,
    title_y=1.04,
):
    """在指定的 matplotlib ax 上畫一張 donut chart。"""
    ax.set_facecolor(BACKGROUND_COLOR)

    if summary.empty or summary[value_col].sum() == 0:
        ax.text(0, 0, "無資料", ha="center", va="center", color="white", fontsize=13)
        ax.set_title(title, fontsize=title_size, color="white", fontweight="bold", y=title_y)
        ax.axis("off")
        return

    wedges, texts, autotexts = ax.pie(
        summary[value_col],
        labels=summary[name_col],
        autopct="%1.1f%%",
        startangle=90,
        pctdistance=0.80,
        labeldistance=1.07,
        radius=0.88,
        wedgeprops={
            "width": 0.38,
            "edgecolor": BACKGROUND_COLOR,
            "linewidth": 1,
        },
        textprops={
            "color": "white",
            "fontsize": label_size,
        },
    )

    for autotext in autotexts:
        autotext.set_color("white")
        autotext.set_fontsize(pct_size)

    ax.text(
        0,
        0,
        center_text,
        ha="center",
        va="center",
        fontsize=center_size,
        color="white",
        fontweight="bold",
    )

    ax.set_title(
        title,
        fontsize=title_size,
        color="white",
        fontweight="bold",
        y=title_y,
    )
    ax.axis("equal")


def plot_dashboard(df, top_n):
    """建立左一右三的投資組合 dashboard。"""
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    overview_summary = make_top_n_summary(
        data=df,
        name_col="symbol",
        value_col="market_value_twd",
        top_n=top_n,
    )

    market_summary = (
        df.groupby("market", as_index=False)["market_value_twd"]
        .sum()
        .sort_values("market_value_twd", ascending=False)
    )
    market_summary["market_name"] = market_summary["market"].map(
        {"US": "美股", "TW": "台股"}
    )

    tw_summary = make_top_n_summary(
        data=df[df["market"] == "TW"],
        name_col="symbol",
        value_col="market_value_twd",
        top_n=top_n,
    )

    us_summary = make_top_n_summary(
        data=df[df["market"] == "US"],
        name_col="symbol",
        value_col="market_value_twd",
        top_n=top_n,
    )

    fig = plt.figure(figsize=(15, 12))
    fig.patch.set_facecolor(BACKGROUND_COLOR)

    # 左邊放整體總覽，右邊由上到下放三張小圖。
    gs = fig.add_gridspec(
        3,
        2,
        width_ratios=[1.15, 1],
        height_ratios=[1, 1, 1],
        left=0.06,
        right=0.96,
        top=0.88,
        bottom=0.08,
        wspace=0.18,
        hspace=0.45,
    )

    ax_overview = fig.add_subplot(gs[:, 0])
    ax_market = fig.add_subplot(gs[0, 1])
    ax_tw = fig.add_subplot(gs[1, 1])
    ax_us = fig.add_subplot(gs[2, 1])

    plot_donut_on_ax(
        ax=ax_overview,
        summary=overview_summary,
        name_col="symbol",
        value_col="market_value_twd",
        title=f"投資組合前 {top_n} 大 + 其他",
        center_text="總覽",
        label_size=11,
        pct_size=10,
        center_size=18,
        title_size=15,
        title_y=0.92,
    )

    plot_donut_on_ax(
        ax=ax_market,
        summary=market_summary,
        name_col="market_name",
        value_col="market_value_twd",
        title="台股 vs 美股",
        center_text="市場",
        label_size=10,
        pct_size=9,
        center_size=14,
        title_size=13,
    )

    plot_donut_on_ax(
        ax=ax_tw,
        summary=tw_summary,
        name_col="symbol",
        value_col="market_value_twd",
        title=f"台股前 {top_n} 大 + 其他",
        center_text="台股",
        label_size=10,
        pct_size=9,
        center_size=14,
        title_size=13,
    )

    plot_donut_on_ax(
        ax=ax_us,
        summary=us_summary,
        name_col="symbol",
        value_col="market_value_twd",
        title=f"美股前 {top_n} 大 + 其他",
        center_text="美股",
        label_size=10,
        pct_size=9,
        center_size=14,
        title_size=13,
    )

    fig.suptitle(
        "Personal Portfolio Analyzer（TWD）",
        fontsize=24,
        color="white",
        fontweight="bold",
        y=0.95,
    )

    return fig


def create_excel_download(df):
    """把報表寫成 Excel bytes，給 Streamlit 下載按鈕使用。"""
    output = BytesIO()

    try:
        writer = pd.ExcelWriter(output, engine="xlsxwriter")
    except ImportError:
        writer = pd.ExcelWriter(output)

    with writer:
        df.to_excel(writer, index=False, sheet_name="Portfolio")

        workbook = writer.book
        worksheet = writer.sheets["Portfolio"]

        # xlsxwriter 支援欄位格式；如果環境使用 openpyxl，仍可正常輸出 Excel。
        supports_xlsxwriter_format = hasattr(workbook, "add_format") and hasattr(
            worksheet, "set_column"
        )
        money_format = None
        percent_format = None

        if supports_xlsxwriter_format:
            money_format = workbook.add_format({"num_format": '#,##0.00'})
            percent_format = workbook.add_format({"num_format": "0.00%"})

        for idx, col_name in enumerate(df.columns):
            if supports_xlsxwriter_format:
                worksheet.set_column(idx, idx, max(14, len(col_name) + 2))

            if supports_xlsxwriter_format and col_name in {
                "price",
                "price_twd",
                "cost",
                "cost_twd",
                "market_value_twd",
                "cost_value_twd",
                "profit_twd",
            }:
                worksheet.set_column(idx, idx, 16, money_format)

            if supports_xlsxwriter_format and col_name in {"return_rate", "weight"}:
                worksheet.set_column(idx, idx, 14, percent_format)

    return output.getvalue()


def create_png_download(fig):
    """把 matplotlib dashboard 存成 PNG bytes。"""
    output = BytesIO()
    fig.savefig(output, format="png", dpi=300, bbox_inches="tight", facecolor=BACKGROUND_COLOR)
    return output.getvalue()


def format_report_for_display(df, sort_option):
    """建立畫面上比較好閱讀的報表格式，不影響下載的原始數字。"""
    # 網頁表格只保留使用者容易閱讀的欄位；換算台幣欄位仍保留在背後供圖表與下載使用。
    display_columns = [
        "symbol",
        "market",
        "category",
        "shares",
        "cost",
        "price",
        "return_rate",
        "weight",
    ]
    display_df = df[display_columns].copy()

    # 只改變畫面上的報表排序，不影響下載檔案或圖表計算。
    if sort_option == "照字母開頭":
        display_df = display_df.sort_values("symbol", ascending=True)
    else:
        display_df = display_df.sort_values("weight", ascending=False)

    display_df.index = range(1, len(display_df) + 1)

    display_df["price"] = display_df["price"].map(lambda value: f"{value:,.2f}")
    display_df["cost"] = display_df["cost"].map(lambda value: f"{value:,.2f}")
    display_df["return_rate"] = display_df["return_rate"].map(lambda value: f"{value:.2%}")
    display_df["weight"] = display_df["weight"].map(lambda value: f"{value:.2%}")

    return display_df


def main():
    st.set_page_config(
        page_title="Personal Portfolio Analyzer",
        page_icon="📊",
        layout="wide",
    )

    st.markdown(
        f"""
        <style>
        .stApp {{
            background-color: {BACKGROUND_COLOR};
            color: white;
        }}
        [data-testid="stSidebar"] {{
            background-color: #0f1b22;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("Personal Portfolio Analyzer")

    top_n = st.sidebar.slider("圖表顯示前幾大", min_value=1, max_value=15, value=7)
    sort_option = st.sidebar.radio(
        "持倉報表排序",
        ["持倉比重", "照字母開頭"],
    )
    uploaded_file = st.file_uploader("上傳 portfolio.csv", type=["csv"])

    if uploaded_file is None:
        st.info("請上傳包含 symbol, market, category, shares, cost 欄位的 portfolio.csv。")
        return

    try:
        portfolio_df = load_portfolio(uploaded_file)

        with st.spinner("正在抓最新股價與 USD/TWD 匯率..."):
            usd_twd = get_usd_twd_rate()
            report = calculate_portfolio(portfolio_df, usd_twd)

        total_value = report["market_value_twd"].sum()
        total_profit = report["profit_twd"].sum()
        total_return = total_profit / report["cost_value_twd"].sum()

        metric_col_1, metric_col_2, metric_col_3 = st.columns(3)
        metric_col_1.metric("總市值（TWD）", f"NT${total_value:,.0f}")
        metric_col_2.metric("總損益（TWD）", f"NT${total_profit:,.0f}")
        metric_col_3.metric("總報酬率", f"{total_return:.2%}")

        st.caption(f"USD/TWD 匯率：{usd_twd:.4f}")

        st.subheader("持倉報表")
        st.dataframe(
            format_report_for_display(report, sort_option=sort_option),
            use_container_width=True,
        )

        st.subheader("Dashboard")
        fig = plot_dashboard(report, top_n)
        st.pyplot(fig, use_container_width=True)

        excel_bytes = create_excel_download(report)
        png_bytes = create_png_download(fig)

        download_col_1, download_col_2 = st.columns(2)
        download_col_1.download_button(
            label="下載 Excel 報表",
            data=excel_bytes,
            file_name="portfolio_report_twd.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        download_col_2.download_button(
            label="下載 Dashboard PNG",
            data=png_bytes,
            file_name="portfolio_dashboard.png",
            mime="image/png",
        )

        plt.close(fig)

    except Exception as exc:
        st.error(str(exc))


if __name__ == "__main__":
    main()
