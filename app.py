from io import BytesIO

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
import yfinance as yf


BACKGROUND_COLOR = "#14232b"
REQUIRED_COLUMNS = ["symbol", "market", "shares", "cost"]
PORTFOLIO_COLUMNS = ["symbol", "market", "category", "shares", "cost"]
CHART_FONT_SCALE = 1.5
CATEGORY_OPTIONS = [
    "未分類",
    "台股ETF",
    "美股ETF",
    "大型科技",
    "AI基礎建設",
    "AI軟體/資料分析",
    "半導體",
    "半導體記憶體",
    "加密貨幣相關",
    "金融科技",
    "醫療健康",
    "電商/消費平台",
    "電動車/未來交通",
    "太空科技",
    "核能/能源科技",
    "無人機/通訊科技",
    "資安",
    "電子零組件",
    "傳產/民生消費",
    "金融股",
    "其他",
]
SAMPLE_CSV = """symbol,market,category,shares,cost
NVDA,US,AI基礎建設,10,150
AAPL,US,大型科技,5,180
QQQ,US,美股ETF,3,500
0050,TW,台股ETF,1000,180
006208,TW,台股ETF,500,110
"""
MARKET_NAME_MAP = {"US": "美股", "TW": "台股"}
MARKET_CURRENCY_MAP = {"US": "USD", "TW": "TWD"}
DISPLAY_COLUMNS = [
    "symbol",
    "market",
    "category",
    "shares",
    "currency",
    "cost",
    "price",
    "profit",
    "profit_twd",
    "return_rate",
    "weight",
    "price_status",
]
DISPLAY_COLUMN_NAMES = {
    "symbol": "代號",
    "market": "市場",
    "category": "類別",
    "shares": "股數",
    "currency": "原始幣別",
    "cost": "原始均價",
    "price": "原始現價",
    "profit": "原始損益",
    "profit_twd": "台幣損益",
    "return_rate": "報酬率",
    "weight": "持倉比例",
    "price_status": "價格狀態",
}
NUMERIC_DISPLAY_COLUMNS = [
    "股數",
    "原始均價",
    "原始現價",
    "原始損益",
    "台幣損益",
    "報酬率",
    "持倉比例",
]
CATEGORY_NUMERIC_COLUMNS = ["台幣市值", "持倉比例"]
EXCEL_SHEET_NAMES = {
    "summary": "Summary",
    "holdings": "Holdings",
    "market": "Market Allocation",
    "category": "Category Allocation",
    "risk": "Risk Flags",
}


def load_portfolio(uploaded_file):
    """讀取使用者上傳的 CSV，並檢查必要欄位是否存在。"""
    df = pd.read_csv(uploaded_file, dtype={"symbol": str, "market": str, "category": str})

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        missing_text = ", ".join(missing_columns)
        raise ValueError(f"CSV 缺少必要欄位：{missing_text}")

    df = ensure_category_column(df)[PORTFOLIO_COLUMNS].copy()

    # 把 market 統一成大寫，避免使用者輸入 us / tw 時判斷失敗。
    df["market"] = df["market"].str.upper().str.strip()
    df["symbol"] = df["symbol"].str.strip()
    df["category"] = clean_category_series(df["category"])

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


def ensure_category_column(df):
    """確保資料一定有 category 欄位，空白分類會補成「未分類」。"""
    result = df.copy()

    # 如果使用者的 CSV 沒有 category 欄位，就直接新增一欄。
    if "category" not in result.columns:
        result["category"] = "未分類"
        return result

    result["category"] = clean_category_series(result["category"])

    return result


def clean_category_series(series):
    """把 category 欄位中的空值、None、空字串統一整理成「未分類」。"""
    cleaned = series.fillna("").astype(str).str.strip()
    return cleaned.mask(cleaned == "", "未分類")


def get_category_options(df):
    """建立下拉選單選項，保留 CSV 原本自訂的分類名稱。"""
    options = CATEGORY_OPTIONS.copy()

    for category in df["category"].dropna().astype(str).str.strip().unique():
        if category and category not in options:
            options.append(category)

    return options


def update_category_options(df, custom_category):
    """合併固定分類、CSV 自帶分類與使用者新增的自訂分類。"""
    options = get_category_options(df)
    custom_category = custom_category.strip() if custom_category else ""

    # 使用者新增的分類只加入選單，不會覆蓋原本資料。
    if custom_category and custom_category not in options:
        options.append(custom_category)

    return options


def to_yahoo_symbol(row):
    """依 market 轉成 Yahoo Finance 使用的代號。"""
    if row["market"] == "TW":
        return f"{row['symbol'].zfill(4)}.TW"
    return row["symbol"]


@st.cache_data(ttl=900, show_spinner=False)
def get_latest_price(yahoo_symbol):
    """使用 yfinance 抓最新收盤價。"""
    try:
        stock = yf.Ticker(yahoo_symbol)
        data = stock.history(period="5d")
        close_prices = data["Close"].dropna() if not data.empty else pd.Series(dtype=float)
    except Exception:
        return None

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

    result["currency"] = result["market"].map(MARKET_CURRENCY_MAP)
    result["yahoo_symbol"] = result.apply(to_yahoo_symbol, axis=1)
    result["price"] = result["yahoo_symbol"].apply(get_latest_price)
    result["price_status"] = result["price"].apply(
        lambda price: "價格正常" if pd.notna(price) else "價格抓取失敗，暫用成本估算"
    )
    result["price_for_calc"] = result["price"].fillna(result["cost"])

    # 美股價格與成本用 USD/TWD 換成台幣，台股本來就是台幣。
    is_us = result["market"] == "US"
    result["price_twd"] = result["price_for_calc"]
    result.loc[is_us, "price_twd"] = result.loc[is_us, "price_for_calc"] * usd_twd

    result["cost_twd"] = result["cost"]
    result.loc[is_us, "cost_twd"] = result.loc[is_us, "cost"] * usd_twd

    # 原始損益保留各市場原本的幣別：美股是美元，台股是台幣。
    result["profit"] = result["shares"] * (result["price_for_calc"] - result["cost"])
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


def create_symbol_color_map(df):
    """建立股票代號與顏色的對照表，讓不同圖中的同一支股票顏色一致。"""
    sorted_symbols = (
        df.sort_values("market_value_twd", ascending=False)["symbol"]
        .drop_duplicates()
        .tolist()
    )

    # tab20 顏色清楚、數量也夠多；超過 20 支時會循環使用。
    palette = list(plt.cm.tab20.colors)
    color_map = {
        symbol: palette[index % len(palette)]
        for index, symbol in enumerate(sorted_symbols)
    }
    color_map["其他"] = "#8a8a8a"

    return color_map


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
    radius=0.88,
    color_map=None,
):
    """在指定的 matplotlib ax 上畫一張 donut chart。"""
    ax.set_facecolor(BACKGROUND_COLOR)

    label_size *= CHART_FONT_SCALE
    pct_size *= 1.1
    center_size *= CHART_FONT_SCALE
    title_size *= CHART_FONT_SCALE

    if summary.empty or summary[value_col].sum() == 0:
        ax.text(
            0,
            0,
            "無資料",
            ha="center",
            va="center",
            color="white",
            fontsize=13 * CHART_FONT_SCALE,
        )
        ax.set_title(title, fontsize=title_size, color="white", fontweight="bold", y=title_y)
        ax.axis("off")
        return

    # 使用同一份 color_map，確保同一支股票在不同 donut chart 中顏色一致。
    colors = None
    if color_map is not None:
        colors = [color_map.get(name, "#9aa4ad") for name in summary[name_col]]

    wedges, texts, autotexts = ax.pie(
        summary[value_col],
        labels=summary[name_col],
        colors=colors,
        autopct="%1.1f%%",
        startangle=90,
        pctdistance=0.80,
        labeldistance=1.22,
        radius=radius,
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


def plot_market_bar_on_ax(ax, market_summary):
    """在 dashboard 底部畫出台股與美股的市場比例長條圖。"""
    ax.set_facecolor(BACKGROUND_COLOR)

    if market_summary.empty or market_summary["market_value_twd"].sum() == 0:
        ax.text(
            0.5,
            0.5,
            "無市場資料",
            ha="center",
            va="center",
            color="white",
            fontsize=13 * CHART_FONT_SCALE,
            transform=ax.transAxes,
        )
        ax.axis("off")
        return

    summary = market_summary.copy()
    summary["weight"] = summary["market_value_twd"] / summary["market_value_twd"].sum()

    # 固定市場顯示順序，讓長條圖每次都容易比較。
    market_order = ["TW", "US"]
    summary["market"] = pd.Categorical(summary["market"], categories=market_order, ordered=True)
    summary = summary.sort_values("market")

    colors = {"TW": "#4cc9f0", "US": "#ff595e"}
    left = 0

    for _, row in summary.iterrows():
        width = row["weight"]
        market = row["market"]
        market_name = row["market_name"]

        ax.barh(
            y=0,
            width=width,
            left=left,
            height=0.42,
            color=colors.get(market, "#9aa4ad"),
            edgecolor=BACKGROUND_COLOR,
            linewidth=2,
        )

        if width >= 0.06:
            ax.text(
                left + width / 2,
                0,
                f"{market_name} {format_percent(width)}",
                ha="center",
                va="center",
                color="white",
                fontsize=11 * CHART_FONT_SCALE,
                fontweight="bold",
            )

        left += width

    ax.set_xlim(0, 1)
    ax.set_ylim(-0.5, 0.5)
    ax.set_title(
        "台股 vs 美股比例",
        fontsize=13 * CHART_FONT_SCALE,
        color="white",
        fontweight="bold",
        pad=24,
    )
    ax.set_xticks([])
    ax.set_yticks([])

    for spine in ax.spines.values():
        spine.set_visible(False)


def plot_dashboard(df, top_n):
    """建立左邊總覽、右邊台股/美股、底部市場比例的 dashboard。"""
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False
    symbol_color_map = create_symbol_color_map(df)

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
    market_summary["market_name"] = market_summary["market"].map(MARKET_NAME_MAP)

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

    fig = plt.figure(figsize=(15, 14))
    fig.patch.set_facecolor(BACKGROUND_COLOR)

    # 左邊放整體總覽，右邊台股與美股平分空間，底部放市場比例長條圖。
    gs = fig.add_gridspec(
        3,
        2,
        width_ratios=[1.15, 1],
        height_ratios=[1, 1, 0.46],
        left=0.06,
        right=0.96,
        top=0.82,
        bottom=0.10,
        wspace=0.22,
        hspace=0.72,
    )

    ax_overview = fig.add_subplot(gs[0:2, 0])
    ax_tw = fig.add_subplot(gs[0, 1])
    ax_us = fig.add_subplot(gs[1, 1])
    ax_market = fig.add_subplot(gs[2, :])

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
        title_y=1.03,
        color_map=symbol_color_map,
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
        title_y=1.18,
        radius=0.78,
        color_map=symbol_color_map,
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
        title_y=1.18,
        radius=0.78,
        color_map=symbol_color_map,
    )

    plot_market_bar_on_ax(ax_market, market_summary)

    fig.suptitle(
        "投資組合總覽",
        fontsize=24 * CHART_FONT_SCALE,
        color="white",
        fontweight="bold",
        y=0.94,
    )

    return fig


def create_summary_sheet(summary):
    """把摘要卡片資料整理成 Excel 的 Summary sheet。"""
    return pd.DataFrame(
        [
            {"metric": "總市值", "value": summary["total_value"]},
            {"metric": "總成本", "value": summary["total_cost"]},
            {"metric": "總損益", "value": summary["total_profit"]},
            {"metric": "總報酬率", "value": summary["total_return"]},
            {"metric": "最大持股", "value": summary["largest_holding"]},
            {"metric": "美股佔比", "value": summary["us_weight"]},
            {"metric": "台股佔比", "value": summary["tw_weight"]},
        ]
    )


def create_multi_sheet_excel(
    holdings_df,
    summary,
    market_allocation,
    category_allocation,
    risk_flags,
):
    """建立包含多個工作表的 Excel 報表。"""
    output = BytesIO()
    export_holdings = holdings_df.drop(columns=["price_for_calc"], errors="ignore")
    sheets = {
        EXCEL_SHEET_NAMES["summary"]: create_summary_sheet(summary),
        EXCEL_SHEET_NAMES["holdings"]: export_holdings,
        EXCEL_SHEET_NAMES["market"]: market_allocation,
        EXCEL_SHEET_NAMES["category"]: category_allocation,
        EXCEL_SHEET_NAMES["risk"]: risk_flags,
    }

    try:
        writer = pd.ExcelWriter(output, engine="xlsxwriter")
    except ImportError:
        writer = pd.ExcelWriter(output)

    with writer:
        write_excel_sheets(writer, sheets)
        apply_excel_formatting(writer, sheets)

    return output.getvalue()


def write_excel_sheets(writer, sheets):
    """把每個 DataFrame 寫入指定的 Excel sheet。"""
    for sheet_name, sheet_df in sheets.items():
        sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)


def apply_excel_formatting(writer, sheets):
    """若使用 xlsxwriter，替 Excel 欄位加上簡單欄寬與百分比格式。"""
    workbook = writer.book

    # openpyxl 沒有 add_format；沒有格式能力時直接略過，Excel 仍可正常下載。
    if not hasattr(workbook, "add_format"):
        return

    money_format = workbook.add_format({"num_format": '#,##0.00'})
    percent_format = workbook.add_format({"num_format": "0.00%"})
    formatted_sheets = {
        EXCEL_SHEET_NAMES["holdings"],
        EXCEL_SHEET_NAMES["market"],
        EXCEL_SHEET_NAMES["category"],
    }

    for sheet_name, sheet_df in sheets.items():
        worksheet = writer.sheets[sheet_name]
        worksheet.set_column(0, 20, 18)

        if sheet_name in formatted_sheets:
            worksheet.set_column(0, 20, 18, money_format)

        for idx, col_name in enumerate(sheet_df.columns):
            if col_name in {"return_rate", "weight"}:
                worksheet.set_column(idx, idx, 14, percent_format)


def create_png_download(fig):
    """把 matplotlib dashboard 存成 PNG bytes。"""
    output = BytesIO()
    fig.savefig(output, format="png", dpi=300, bbox_inches="tight", facecolor=BACKGROUND_COLOR)
    return output.getvalue()


def render_download_button(label, data, file_name, mime, container=st):
    """統一建立下載按鈕，避免同樣的參數寫法重複出現。"""
    container.download_button(
        label=label,
        data=data,
        file_name=file_name,
        mime=mime,
    )


def format_twd(value):
    """把數字格式化成台幣金額，例如 NT$1,234,567。"""
    if pd.isna(value):
        return "N/A"
    return f"NT${value:,.0f}"


def format_decimal(value, digits=2):
    """把一般數字格式化成固定小數位，例如 1,234.56。"""
    if pd.isna(value):
        return "N/A"
    return f"{value:,.{digits}f}"


def format_percent(value):
    """把小數格式化成百分比，例如 0.1234 會變成 12.34%。"""
    if pd.isna(value):
        return "N/A"
    return f"{value:.2%}"


def create_market_allocation(df):
    """依市場彙總台幣市值，建立台股 / 美股配置表。"""
    allocation = (
        df.groupby("market", as_index=False)["market_value_twd"]
        .sum()
        .sort_values("market_value_twd", ascending=False)
    )
    total_value = allocation["market_value_twd"].sum()
    allocation["weight"] = allocation["market_value_twd"] / total_value if total_value else 0
    allocation["market_name"] = allocation["market"].map(MARKET_NAME_MAP)
    return allocation


def create_category_allocation(df):
    """依 category 彙總台幣市值，建立類別配置分析資料。"""
    allocation = (
        df.groupby("category", as_index=False)["market_value_twd"]
        .sum()
        .sort_values("market_value_twd", ascending=False)
    )
    total_value = allocation["market_value_twd"].sum()
    allocation["weight"] = allocation["market_value_twd"] / total_value if total_value else 0
    return allocation


def create_category_display_table(category_allocation):
    """把類別配置表轉成適合畫面閱讀的格式。"""
    display_df = category_allocation.copy()
    display_df.index = range(1, len(display_df) + 1)
    display_df["market_value_twd"] = display_df["market_value_twd"].map(format_twd)
    display_df["weight"] = display_df["weight"].map(format_percent)
    return display_df.rename(
        columns={
            "category": "類別",
            "market_value_twd": "台幣市值",
            "weight": "持倉比例",
        }
    )


def plot_category_allocation(category_allocation, category_top_n):
    """畫出 category 類別配置 donut chart。"""
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    fig, ax = plt.subplots(figsize=(7, 5.6))
    fig.patch.set_facecolor(BACKGROUND_COLOR)
    ax.set_facecolor(BACKGROUND_COLOR)

    if category_allocation.empty:
        ax.text(0.5, 0.5, "無類別資料", ha="center", va="center", color="white")
        ax.axis("off")
        return fig

    chart_df = make_top_n_summary(
        data=category_allocation,
        name_col="category",
        value_col="market_value_twd",
        top_n=category_top_n,
    )

    plot_donut_on_ax(
        ax=ax,
        summary=chart_df,
        name_col="category",
        value_col="market_value_twd",
        title=f"類別配置前 {category_top_n} 大 + 其他",
        center_text="類別",
        label_size=10,
        pct_size=9,
        center_size=15,
        title_size=14,
        title_y=1.08,
        radius=0.82,
    )

    fig.tight_layout()
    return fig


def create_risk_flags(df):
    """建立風險提醒資料，例如高集中持股或股價抓取失敗。"""
    flags = []

    for _, row in df[df["weight"] > 0.3].iterrows():
        flags.append(
            {
                "type": "高集中持股",
                "symbol": row["symbol"],
                "detail": f"單一持股比例為 {format_percent(row['weight'])}，超過 30%。",
            }
        )

    for _, row in df[df["price_status"] != "價格正常"].iterrows():
        flags.append(
            {
                "type": "價格抓取失敗",
                "symbol": row["symbol"],
                "detail": "此股票暫用成本估算市值與損益。",
            }
        )

    if not flags:
        flags.append({"type": "無", "symbol": "", "detail": "目前沒有需要提醒的風險旗標。"})

    return pd.DataFrame(flags)


def create_summary_metrics(df):
    """計算摘要卡片需要用到的投資組合指標。"""
    total_value = df["market_value_twd"].sum()
    total_cost = df["cost_value_twd"].sum()
    total_profit = df["profit_twd"].sum()
    total_return = total_profit / total_cost if total_cost else 0

    # 最大持股用持倉比例最高的那一筆資料。
    largest_position = df.sort_values("weight", ascending=False).iloc[0]
    largest_holding = (
        f"{largest_position['symbol']} ({format_percent(largest_position['weight'])})"
    )

    if total_value:
        market_weights = df.groupby("market")["market_value_twd"].sum() / total_value
        us_weight = market_weights.get("US", 0)
        tw_weight = market_weights.get("TW", 0)
    else:
        us_weight = 0
        tw_weight = 0

    return {
        "total_value": total_value,
        "total_cost": total_cost,
        "total_profit": total_profit,
        "total_return": total_return,
        "largest_holding": largest_holding,
        "us_weight": us_weight,
        "tw_weight": tw_weight,
    }


def create_display_table(df, sort_option):
    """建立畫面上使用的中文報表，不影響下載的原始數字資料。"""
    # 先用完整資料排序，排序完成後才挑出要顯示的欄位。
    # 這樣即使 market_value_twd 不顯示在表格中，也能用它做預設排序。
    sorted_df = sort_holdings_for_display(df, sort_option)

    display_df = sorted_df[DISPLAY_COLUMNS].copy()

    display_df.index = range(1, len(display_df) + 1)

    # 市場代碼轉成中文，讓第一次使用的人更容易讀懂。
    display_df["market"] = display_df["market"].map(MARKET_NAME_MAP)

    display_df["price"] = display_df["price"].map(
        lambda value: "抓取失敗" if pd.isna(value) else format_decimal(value)
    )
    display_df["shares"] = display_df["shares"].map(format_decimal)
    display_df["cost"] = display_df["cost"].map(format_decimal)
    display_df["profit"] = display_df["profit"].map(format_decimal)
    display_df["profit_twd"] = display_df["profit_twd"].map(format_twd)
    display_df["return_rate"] = display_df["return_rate"].map(format_percent)
    display_df["weight"] = display_df["weight"].map(format_percent)

    return display_df.rename(columns=DISPLAY_COLUMN_NAMES)


def sort_holdings_for_display(df, sort_option):
    """依使用者選擇的排序方式排列持倉報表。"""
    sort_rules = {
        "照字母開頭": ("symbol", True),
        "報酬率": ("return_rate", False),
        "台幣損益": ("profit_twd", False),
        "持倉比重": ("market_value_twd", False),
    }
    sort_column, ascending = sort_rules.get(sort_option, ("market_value_twd", False))
    return df.sort_values(sort_column, ascending=ascending)


def apply_category_edits(report, edited_display_table, sort_option):
    """把持倉報表中編輯後的「類別」欄位寫回原始 report。"""
    updated_report = report.copy()

    # data_editor 回傳的資料列順序，會跟畫面上排序後的報表一致。
    sorted_index = sort_holdings_for_display(updated_report, sort_option).index.tolist()

    for row_position, original_index in enumerate(sorted_index):
        selected_category = edited_display_table.iloc[row_position]["類別"]
        updated_report.at[original_index, "category"] = selected_category

    return updated_report


def style_holdings_table(display_table):
    """替持倉報表加上損益與報酬率的條件色彩。"""

    def parse_display_number(value):
        """把 NT$1,234 或 12.34% 這類顯示文字轉回數字，方便判斷正負。"""
        if pd.isna(value):
            return 0

        text = str(value)
        text = text.replace("NT$", "").replace(",", "").replace("%", "").strip()

        if text in {"", "N/A", "抓取失敗"}:
            return 0

        try:
            return float(text)
        except ValueError:
            return 0

    def profit_color(value):
        number = parse_display_number(value)

        if number > 0:
            return "color: #4ade80; font-weight: 700;"
        if number < 0:
            return "color: #f87171; font-weight: 700;"
        return ""

    numeric_columns = [col for col in NUMERIC_DISPLAY_COLUMNS if col in display_table.columns]

    return (
        display_table.style
        .map(profit_color, subset=["台幣損益", "報酬率"])
        .set_properties(subset=numeric_columns, **{"text-align": "right"})
        .set_table_styles(
            [
                {"selector": "th", "props": [("text-align", "left")]},
            ]
        )
    )


def style_category_table(display_table):
    """讓類別配置表中的數字欄位靠右對齊。"""
    numeric_columns = [col for col in CATEGORY_NUMERIC_COLUMNS if col in display_table.columns]
    return (
        display_table.style
        .set_properties(subset=numeric_columns, **{"text-align": "right"})
        .set_table_styles(
            [
                {"selector": "th", "props": [("text-align", "left")]},
            ]
        )
    )


def render_styled_html_table(styler, width_percent=100):
    """用 HTML 顯示 styled table，確保數字靠右與條件色彩都會生效。"""
    st.markdown(
        f"""
        <div class="styled-table-scroll" style="width: {width_percent}%;">
            {styler.to_html()}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_summary_metrics(summary):
    """用較不擁擠的方式顯示摘要卡片，手機版也比較容易閱讀。"""
    metric_col_1, metric_col_2 = st.columns(2)
    metric_col_1.metric("總市值", format_twd(summary["total_value"]))
    metric_col_2.metric("總成本", format_twd(summary["total_cost"]))

    metric_col_3, metric_col_4 = st.columns(2)
    metric_col_3.metric("總損益", format_twd(summary["total_profit"]))
    metric_col_4.metric("總報酬率", format_percent(summary["total_return"]))

    metric_col_5, metric_col_6, metric_col_7 = st.columns(3)
    metric_col_5.metric("最大持股", summary["largest_holding"])
    metric_col_6.metric("美股佔比", format_percent(summary["us_weight"]))
    metric_col_7.metric("台股佔比", format_percent(summary["tw_weight"]))


def render_risk_messages(report):
    """顯示股價抓取失敗與高集中持股提醒。"""
    failed_price_symbols = report[report["price_status"] != "價格正常"]["symbol"].tolist()
    if failed_price_symbols:
        st.warning(
            "以下股票無法取得最新股價，系統已暫用成本估算市值與損益："
            + "、".join(failed_price_symbols)
        )

    concentrated_positions = report[report["weight"] > 0.3]["symbol"].tolist()
    if concentrated_positions:
        symbols_text = "、".join(concentrated_positions)
        st.warning(f"注意：以下持股超過 30%，投資組合可能過度集中：{symbols_text}")
    else:
        st.success("目前沒有單一持股超過 30%。")


def render_overview_tab(report, summary, usd_twd, top_n):
    """總覽 tab：摘要卡片、提醒與 Dashboard 圖。"""
    st.subheader("投資組合摘要")
    render_summary_metrics(summary)
    render_risk_messages(report)
    st.caption(f"USD/TWD 匯率：{usd_twd:.4f}")

    st.subheader("投資組合總覽")
    fig = plot_dashboard(report, top_n)
    st.pyplot(fig, use_container_width=True)
    render_download_button(
        label="下載 Dashboard PNG",
        data=create_png_download(fig),
        file_name="portfolio_dashboard.png",
        mime="image/png",
    )

    return fig


def render_holdings_tab(report, sort_option):
    """持倉 tab：在報表中調整類別，並顯示有條件色彩的完整報表。"""
    st.subheader("持倉報表")
    st.write("可以先新增自訂類別，再直接在下方完整持倉報表的「類別」欄位調整分類。")

    custom_category = st.text_input(
        "新增自訂 category",
        placeholder="例如：高股息ETF、債券、現金部位",
    )
    editable_category_options = update_category_options(report, custom_category)

    display_table = create_display_table(report, sort_option=sort_option)
    locked_columns = [col for col in display_table.columns if col != "類別"]

    st.markdown("#### 完整持倉報表")
    edited_display_table = st.data_editor(
        style_holdings_table(display_table),
        use_container_width=True,
        disabled=locked_columns,
        column_config={
            "類別": st.column_config.SelectboxColumn(
                "類別",
                options=editable_category_options,
                required=True,
            )
        },
        hide_index=False,
    )
    updated_report = apply_category_edits(report, edited_display_table, sort_option)

    if (updated_report["category"] == "未分類").any():
        st.warning("有持股尚未分類，建議補上 category，類別配置圖會更準確。")
    else:
        st.success("所有持股都已完成分類。")

    return updated_report


def render_category_tab(category_allocation, category_top_n):
    """類別 tab：類別配置表、donut chart 與簡短說明。"""
    st.subheader("類別配置")
    st.write("類別配置會依台幣市值由大到小排序，圖表會顯示前幾大，其餘合併成「其他」。")
    render_styled_html_table(
        style_category_table(create_category_display_table(category_allocation)),
        width_percent=75,
    )

    category_fig = plot_category_allocation(category_allocation, category_top_n)
    st.pyplot(category_fig, use_container_width=True)

    return category_fig


def render_excel_report_download(report, summary, market_allocation, category_allocation, risk_flags):
    """在持倉報表下方提供 Excel 多 sheet 報表下載。"""
    excel_bytes = create_multi_sheet_excel(
        report,
        summary,
        market_allocation,
        category_allocation,
        risk_flags,
    )
    render_download_button(
        label="下載 Excel 報表",
        data=excel_bytes,
        file_name="portfolio_report_twd.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_page_intro():
    """顯示頁面標題、說明與使用步驟。"""
    st.title("個人投資組合分析器")
    st.write(
        "這個工具可以上傳持倉 CSV，自動抓取股價，統一換算成台幣，"
        "並產生持倉報表與投資組合圖表。"
    )

    st.subheader("使用步驟")
    st.markdown(
        """
        1. Step 1：準備 CSV 檔
        2. Step 2：上傳 CSV
        3. Step 3：系統自動抓股價並換算台幣
        4. Step 4：查看報表與圖表
        5. Step 5：下載 Excel 報表或 PNG 圖片
        """
    )


def render_sidebar_controls():
    """顯示側邊欄設定，並回傳使用者選擇的控制值。"""
    top_n = st.sidebar.slider("圖表顯示前幾大", min_value=1, max_value=15, value=7)
    category_top_n = st.sidebar.slider(
        "類別配置顯示前幾大",
        min_value=1,
        max_value=15,
        value=7,
    )
    sort_option = st.sidebar.radio(
        "持倉報表排序",
        ["持倉比重", "報酬率", "台幣損益", "照字母開頭"],
    )

    st.sidebar.subheader("CSV 格式說明")
    st.sidebar.markdown(
        """
        必要欄位：

        - symbol：股票代號，例如 NVDA、TSLA、0050
        - market：市場，只能填 US 或 TW
        - shares：股數
        - cost：平均成本

        選填欄位：

        - category：分類，例如 AI基礎建設、台股ETF、金融科技。若留空或沒有此欄，會自動填入未分類。
        """
    )
    render_download_button(
        label="下載範例持倉 CSV",
        data=SAMPLE_CSV.encode("utf-8-sig"),
        file_name="sample_portfolio.csv",
        mime="text/csv",
        container=st.sidebar,
    )

    return top_n, category_top_n, sort_option


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
        div[data-testid="stMetric"] {{
            background-color: rgba(255, 255, 255, 0.03);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 8px;
            padding: 12px;
        }}
        @media (max-width: 768px) {{
            .block-container {{
                padding-left: 1rem;
                padding-right: 1rem;
            }}
            div[data-testid="stMetric"] {{
                padding: 10px;
            }}
            .stDownloadButton button {{
                width: 100%;
            }}
        }}
        .styled-table-scroll {{
            overflow-x: auto;
            max-height: 680px;
            border: 1px solid rgba(255, 255, 255, 0.16);
            border-radius: 8px;
        }}
        .styled-table-scroll table {{
            border-collapse: collapse;
            min-width: 100%;
            color: white;
            background-color: #0e1117;
        }}
        .styled-table-scroll th,
        .styled-table-scroll td {{
            padding: 10px 12px;
            border: 1px solid rgba(255, 255, 255, 0.10);
            white-space: nowrap;
        }}
        .styled-table-scroll th {{
            color: #aeb6c2;
            background-color: #171a22;
            font-weight: 700;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    render_page_intro()
    top_n, category_top_n, sort_option = render_sidebar_controls()

    uploaded_file = st.file_uploader("上傳 portfolio.csv", type=["csv"])

    if uploaded_file is None:
        st.info("請先上傳 portfolio.csv，或下載範例 CSV 試用。")
        return

    try:
        portfolio_df = load_portfolio(uploaded_file)

        with st.spinner("正在抓最新股價與 USD/TWD 匯率..."):
            usd_twd = get_usd_twd_rate()
            report = calculate_portfolio(portfolio_df, usd_twd)

        summary = create_summary_metrics(report)
        overview_tab, holdings_tab, category_tab = st.tabs(["總覽", "持倉", "類別"])

        with overview_tab:
            fig = render_overview_tab(report, summary, usd_twd, top_n)

        with holdings_tab:
            report = render_holdings_tab(report, sort_option)
            market_allocation = create_market_allocation(report)
            category_allocation = create_category_allocation(report)
            risk_flags = create_risk_flags(report)
            render_excel_report_download(
                report,
                summary,
                market_allocation,
                category_allocation,
                risk_flags,
            )

        with category_tab:
            category_fig = render_category_tab(category_allocation, category_top_n)

        plt.close(fig)
        plt.close(category_fig)

    except Exception as exc:
        st.error(str(exc))


if __name__ == "__main__":
    main()
