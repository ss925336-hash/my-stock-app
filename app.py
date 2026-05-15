"""
台股自動化篩選 Streamlit Web App
對應原始 Excel「總表」與「總表 2」的邏輯

Excel 總表欄位完整對應：
  股價、配息、配股、除息日、發息日
  殖利率(%)                  → 近一年現金股利 / 股價
  配息率(%)                  → 現金股利 / 去年EPS
  去年EPS                    → 前4季EPS加總
  3年/6年/10年平均股利
  10年股利次數
  本益比                     → 股價 / 近4季EPS
  股價淨值比
  (月)累積營收年增率(%)
  用預估EPS推隔年殖利率(平均加總) → (3年均EPS × 配息率) / 股價
  用預估EPS推隔年殖利率(前季推估) → (近4季累積EPS × 配息率) / 股價
  推隔年本益比(平均加總)     → 股價 / (3年均EPS × 配息率 / 殖利率門檻)
  昂貴(4%) / 合理(5.5%) / 便宜(7%) → 3年平均股利 / 殖利率門檻
"""

import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import requests
from datetime import datetime, timedelta
import time

# ─────────────────────────────────────────────────────
# Token
# ─────────────────────────────────────────────────────
try:
    _SECRET_TOKEN = st.secrets.get("FINMIND_TOKEN", "")
except Exception:
    _SECRET_TOKEN = ""

# ─────────────────────────────────────────────────────
# 頁面設定
# ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="台股分析儀",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.card {
    background: #1e2130;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 12px;
    border-left: 4px solid #4a9eff;
}
.card-red   { border-left-color: #ff4b4b; }
.card-green { border-left-color: #21c55d; }
.card-gold  { border-left-color: #f5a623; }
.card-blue  { border-left-color: #4a9eff; }
.card-purple{ border-left-color: #a855f7; }

.card-title { font-size:0.72rem; color:#8b9ab0; text-transform:uppercase; letter-spacing:0.08em; margin-bottom:4px; }
.card-value { font-size:1.7rem; font-weight:700; color:#ffffff; line-height:1.1; }
.card-value-red    { color:#ff4b4b; }
.card-value-green  { color:#21c55d; }
.card-value-gold   { color:#f5a623; }
.card-value-purple { color:#a855f7; }
.card-sub   { font-size:0.76rem; color:#6b7a8d; margin-top:4px; }

.section-title {
    font-size:1rem; font-weight:600; color:#c9d1d9;
    padding:10px 0 4px 0;
    border-bottom:1px solid #30363d;
    margin-bottom:12px; margin-top:4px;
}
</style>
""", unsafe_allow_html=True)

FINMIND_BASE = "https://api.finmindtrade.com/api/v4/data"


# ─────────────────────────────────────────────────────
# 工具函式
# ─────────────────────────────────────────────────────

def finmind_get(dataset, stock_id, start_date, token=""):
    params = {"dataset": dataset, "data_id": stock_id, "start_date": start_date}
    if token:
        params["token"] = token
    try:
        r = requests.get(FINMIND_BASE, params=params, timeout=20)
        raw = r.json()
        if raw.get("status") == 200 and raw.get("data"):
            return pd.DataFrame(raw["data"]), raw
        return pd.DataFrame(), raw
    except Exception as e:
        return pd.DataFrame(), {"error": str(e)}


def get_price(stock_id):
    """yfinance 取台股現價"""
    try:
        t = yf.Ticker(f"{stock_id}.TW")
        p = t.fast_info.last_price
        if p and p > 0:
            return float(p)
        hist = t.history(period="5d")
        if not hist.empty:
            return float(hist["Close"].iloc[-1])
    except Exception:
        pass
    return np.nan


def extract_cash_dividend(df: pd.DataFrame) -> pd.Series:
    """
    自動偵測 FinMind TaiwanStockDividend 的現金股利欄位。
    支援兩種資料結構：
    (A) 寬表格：CashEarningsDistribution + CashStatutorySurplus（加總）
    (B) 長表格：stock_or_cash / dividend_type 等分類欄 + cash_dividend 金額欄

    只要 FinMind 有回傳資料，不論欄位大小寫，都能正確加總。
    回傳 Series（每列對應現金股利金額，非現金列為 0）。
    """
    if df.empty:
        return pd.Series(dtype=float)

    # ── 結構 A：寬表格（有 CashEarningsDistribution）────
    cash_cols = [c for c in df.columns
                 if any(k in c for k in ["CashEarnings", "CashStatutory",
                                          "cashearnings", "cashstatutory"])]
    if cash_cols:
        result = pd.Series(0.0, index=df.index)
        for col in cash_cols:
            result += pd.to_numeric(df[col], errors="coerce").fillna(0)
        return result

    # ── 結構 B：長表格（有 stock_or_cash 分類欄）─────────
    type_col = next(
        (c for c in df.columns
         if c.lower() in ["stock_or_cash", "stockorcash", "dividend_type",
                          "dividendtype", "type"]),
        None
    )
    amt_col = next(
        (c for c in df.columns
         if c.lower() in ["cash_dividend", "cashdividend",
                          "dividend", "amount", "value"]),
        None
    )
    if type_col and amt_col:
        is_cash = df[type_col].astype(str).str.lower().str.contains("cash")
        amounts = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
        return amounts.where(is_cash, 0.0)

    # ── Fallback：嘗試所有數字欄中最像股利的那欄 ─────────
    num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    likely = [c for c in num_cols
              if any(k in c.lower() for k in ["div", "cash", "yield", "amount"])]
    if likely:
        return pd.to_numeric(df[likely[0]], errors="coerce").fillna(0)

    return pd.Series(0.0, index=df.index)


def fmt(val, dec=2, suffix="", prefix="", na="N/A"):
    try:
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return na
        return f"{prefix}{val:.{dec}f}{suffix}"
    except Exception:
        return na


def card(title, value, sub="", color="blue"):
    val_cls = {"red": "card-value-red", "green": "card-value-green",
               "gold": "card-value-gold", "purple": "card-value-purple"}.get(color, "")
    return f"""
<div class="card card-{color}">
  <div class="card-title">{title}</div>
  <div class="card-value {val_cls}">{value}</div>
  {"<div class='card-sub'>" + sub + "</div>" if sub else ""}
</div>"""


def section(title):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────────────
# 個股完整資料抓取（對應 Excel 總表）
# ─────────────────────────────────────────────────────

@st.cache_data(ttl=1800, show_spinner=False)
def get_stock_data(stock_id: str, token: str) -> dict:
    """
    抓取並計算個股所有指標，完整對應 Excel 總表欄位。
    """
    d = {}

    # ── 股價 ──────────────────────────────────────────
    price = get_price(stock_id)
    d["price"] = price

    # ── 股價歷史（一年，用於走勢圖）──────────────────
    try:
        hist = yf.Ticker(f"{stock_id}.TW").history(period="1y")
        d["price_hist"] = hist[["Close"]].rename(columns={"Close": "收盤價"}) if not hist.empty else pd.DataFrame()
    except Exception:
        d["price_hist"] = pd.DataFrame()

    # ── yfinance 基本資訊 ─────────────────────────────
    try:
        info = yf.Ticker(f"{stock_id}.TW").info
        d["company"]  = info.get("longName", stock_id)
        d["industry"] = info.get("industry", "—")
        d["pbr"]      = info.get("priceToBook", np.nan)
    except Exception:
        d["company"]  = stock_id
        d["industry"] = "—"
        d["pbr"]      = np.nan

    # ── 股利資料（近10年）─────────────────────────────
    start_10y = (datetime.now() - timedelta(days=3650)).strftime("%Y-%m-%d")
    df_div, _ = finmind_get("TaiwanStockDividend", stock_id, start_10y, token)

    if not df_div.empty:
        df_div["_date"]     = pd.to_datetime(df_div["date"])
        df_div["_cash"]     = extract_cash_dividend(df_div)
        df_div["_stock"]    = pd.to_numeric(
            df_div.get("StockEarningsDistribution",
            df_div.get("stock_dividend", 0)), errors="coerce").fillna(0)

        # 近一年現金股利加總（用於殖利率計算）
        cutoff_1y = pd.Timestamp(datetime.now() - timedelta(days=500))
        div_1y    = df_div[df_div["_date"] >= cutoff_1y]
        cash_sum  = div_1y["_cash"].sum()
        d["cash_div"]  = cash_sum
        d["stock_div"] = div_1y["_stock"].sum()

        # 除息日、發息日
        has_cash = df_div[df_div["_cash"] > 0].sort_values("_date", ascending=False)
        if not has_cash.empty:
            latest = has_cash.iloc[0]
            d["ex_div_date"] = str(latest["_date"])[:10]
            pay_col = next((c for c in ["CashDividendPaymentDate", "pay_date"] if c in df_div.columns), None)
            d["pay_date"] = str(latest[pay_col])[:10] if pay_col else "—"
        else:
            d["ex_div_date"] = "—"
            d["pay_date"]    = "—"

        # 10年股利次數
        d["div_count_10y"] = int((df_div["_cash"] > 0).sum())

        # 3/6/10年平均現金股利（按年分組）
        df_yr = df_div[df_div["_cash"] > 0].copy()
        df_yr["_yr"] = df_yr["_date"].dt.year
        yr_sum = df_yr.groupby("_yr")["_cash"].sum().sort_index()
        d["avg_div_3y"]  = yr_sum.tail(3).mean()  if len(yr_sum) >= 1 else np.nan
        d["avg_div_6y"]  = yr_sum.tail(6).mean()  if len(yr_sum) >= 1 else np.nan
        d["avg_div_10y"] = yr_sum.tail(10).mean() if len(yr_sum) >= 1 else np.nan
    else:
        d.update({
            "cash_div": np.nan, "stock_div": np.nan,
            "ex_div_date": "—", "pay_date": "—", "div_count_10y": 0,
            "avg_div_3y": np.nan, "avg_div_6y": np.nan, "avg_div_10y": np.nan,
        })

    # ── EPS（近8季）──────────────────────────────────
    start_eps = (datetime.now() - timedelta(days=1100)).strftime("%Y-%m-%d")
    df_eps, _ = finmind_get("TaiwanStockFinancialStatements", stock_id, start_eps, token)

    if not df_eps.empty and "type" in df_eps.columns:
        e = df_eps[df_eps["type"] == "EPS"][["date", "value"]].copy()
        e["value"] = pd.to_numeric(e["value"], errors="coerce")
        e["date"]  = pd.to_datetime(e["date"])
        e = e.dropna().sort_values("date").drop_duplicates("date")

        # 各季 EPS（用月份判斷）
        def last_q_eps(months):
            sub = e[e["date"].dt.month.isin(months)]
            return float(sub.tail(1)["value"].iloc[0]) if not sub.empty else np.nan

        d["eps_q1"] = last_q_eps([3, 4, 5])
        d["eps_q2"] = last_q_eps([6, 7, 8])
        d["eps_q3"] = last_q_eps([9, 10, 11])
        d["eps_q4"] = last_q_eps([12, 1, 2])

        # 今年累積EPS（近4季）= 對應 Excel「今年累積EPS」
        d["eps_ytd"]  = float(e.tail(4)["value"].sum()) if len(e) >= 1 else np.nan
        # 去年EPS（前4季）= 對應 Excel「去年EPS」
        d["eps_prev"] = float(e.iloc[-8:-4]["value"].sum()) if len(e) >= 8 else np.nan

        # EPS 年度比較
        if not np.isnan(d.get("eps_ytd", np.nan)) and not np.isnan(d.get("eps_prev", np.nan)) \
                and d["eps_prev"] != 0:
            d["eps_yoy"] = d["eps_ytd"] / d["eps_prev"]
        else:
            d["eps_yoy"] = np.nan

        # 本益比 = 股價 / 近4季EPS
        if not np.isnan(price) and not np.isnan(d.get("eps_ytd", np.nan)) and d["eps_ytd"] > 0:
            d["pe_ratio"] = price / d["eps_ytd"]
        else:
            d["pe_ratio"] = np.nan

        # 近3年年度EPS加總 → 平均（用於推估）
        e_yr = e.groupby(e["date"].dt.year)["value"].sum()
        d["avg_eps_3y"] = e_yr.tail(3).mean() if not e_yr.empty else np.nan
    else:
        d.update({
            "eps_q1": np.nan, "eps_q2": np.nan, "eps_q3": np.nan, "eps_q4": np.nan,
            "eps_ytd": np.nan, "eps_prev": np.nan, "eps_yoy": np.nan,
            "pe_ratio": np.nan, "avg_eps_3y": np.nan,
        })

    # ── 配息率 = 近一年現金股利 / 去年EPS ────────────
    # 對應 Excel 總表「配息率(%)」欄
    cash_div = d.get("cash_div", np.nan)
    eps_prev = d.get("eps_prev", np.nan)
    avg_div_3y = d.get("avg_div_3y", np.nan)
    avg_eps_3y = d.get("avg_eps_3y", np.nan)

    if not np.isnan(cash_div) and not np.isnan(eps_prev) and eps_prev > 0:
        d["payout_ratio"] = cash_div / eps_prev
    else:
        d["payout_ratio"] = np.nan

    # 近3年平均配息率（用於推估殖利率）
    if not np.isnan(avg_div_3y) and not np.isnan(avg_eps_3y) and avg_eps_3y > 0:
        payout_3y = min(avg_div_3y / avg_eps_3y, 1.5)
    else:
        payout_3y = d.get("payout_ratio", 0.6)
        if np.isnan(payout_3y):
            payout_3y = 0.6
    d["payout_3y"] = payout_3y

    # ── 殖利率 = 近一年現金股利 / 股價 ───────────────
    if not np.isnan(cash_div) and not np.isnan(price) and price > 0:
        d["div_yield"] = cash_div / price
    else:
        d["div_yield"] = np.nan

    # ── 推估EPS推隔年殖利率（平均加總）──────────────
    # 對應 Excel 總表紅字「用預估EPS推隔年殖利率(平均加總)」
    # 公式：(近3年平均EPS × 近3年平均配息率) / 股價
    if not np.isnan(avg_eps_3y) and avg_eps_3y > 0 and not np.isnan(price) and price > 0:
        est_div_avg      = avg_eps_3y * payout_3y
        d["est_yield_avg"] = est_div_avg / price
    else:
        d["est_yield_avg"] = np.nan

    # ── 推估EPS推隔年殖利率（前季推估）──────────────
    # 對應 Excel 總表紅字「用預估EPS推隔年殖利率(前季推估)」
    # 公式：近4季累積EPS × 近3年平均配息率 / 股價
    eps_ytd = d.get("eps_ytd", np.nan)
    if not np.isnan(eps_ytd) and eps_ytd > 0 and not np.isnan(price) and price > 0:
        d["est_yield_latest"] = (eps_ytd * payout_3y) / price
    else:
        d["est_yield_latest"] = np.nan

    # ── 推隔年本益比（平均加總）──────────────────────
    # 對應 Excel 截圖紅字「推隔年本益比(平均加總)」
    # 公式：股價 / (3年均EPS)  → 用未來預估EPS算本益比
    if not np.isnan(avg_eps_3y) and avg_eps_3y > 0 and not np.isnan(price):
        d["est_pe_avg"] = price / avg_eps_3y
    else:
        d["est_pe_avg"] = np.nan

    # ── 推隔年本益比（前季推估）──────────────────────
    if not np.isnan(eps_ytd) and eps_ytd > 0 and not np.isnan(price):
        d["est_pe_latest"] = price / eps_ytd
    else:
        d["est_pe_latest"] = np.nan

    # ── 價格推估區間（對應 Excel 昂貴/合理/便宜）────
    # 計算基礎：3年平均現金股利（與 Excel 總表一致）
    # 昂貴(4%) = 3年均股利 / 4%
    # 合理(5.5%) = 3年均股利 / 5.5%
    # 便宜(7%) = 3年均股利 / 7%
    if not np.isnan(avg_div_3y) and avg_div_3y > 0:
        d["price_expensive"] = avg_div_3y / 0.04
        d["price_fair"]      = avg_div_3y / 0.055
        d["price_cheap"]     = avg_div_3y / 0.07
    else:
        d["price_expensive"] = np.nan
        d["price_fair"]      = np.nan
        d["price_cheap"]     = np.nan

    # ── 月營收 ────────────────────────────────────────
    start_rev = (datetime.now() - timedelta(days=760)).strftime("%Y-%m-%d")
    df_rev, _ = finmind_get("TaiwanStockMonthRevenue", stock_id, start_rev, token)

    if not df_rev.empty:
        df_rev["date"]    = pd.to_datetime(df_rev["date"])
        df_rev["revenue"] = pd.to_numeric(df_rev["revenue"], errors="coerce")
        df_rev = df_rev.dropna(subset=["revenue"]).sort_values("date")

        cur_year  = datetime.now().year
        cur_df    = df_rev[df_rev["date"].dt.year == cur_year]
        prev_df   = df_rev[df_rev["date"].dt.year == cur_year - 1]

        d["rev_ytd_yoy"] = np.nan
        if not cur_df.empty and not prev_df.empty:
            months    = set(cur_df["date"].dt.month)
            prev_same = prev_df[prev_df["date"].dt.month.isin(months)]
            if not prev_same.empty and prev_same["revenue"].sum() != 0:
                d["rev_ytd_yoy"] = (cur_df["revenue"].sum() / prev_same["revenue"].sum() - 1) * 100

        d["rev_mom"] = np.nan
        if len(df_rev) >= 2:
            c, p = df_rev.iloc[-1]["revenue"], df_rev.iloc[-2]["revenue"]
            if p != 0:
                d["rev_mom"] = (c - p) / p * 100

        d["df_rev"] = df_rev   # 保留供明細顯示
    else:
        d["rev_ytd_yoy"] = np.nan
        d["rev_mom"]     = np.nan
        d["df_rev"]      = pd.DataFrame()

    d["df_div_raw"] = df_div if not df_div.empty else pd.DataFrame()
    return d


# ─────────────────────────────────────────────────────
# 批次篩選（輕量版）
# ─────────────────────────────────────────────────────

def calc_metrics_batch(stock_id, token):
    price = get_price(stock_id)

    # 股利
    start_10y = (datetime.now() - timedelta(days=3650)).strftime("%Y-%m-%d")
    df_div, _ = finmind_get("TaiwanStockDividend", stock_id, start_10y, token)
    div_yield, div_count, avg_eps_3y_b = np.nan, 0, np.nan

    if not df_div.empty:
        df_div["_date"] = pd.to_datetime(df_div["date"])
        df_div["_cash"] = extract_cash_dividend(df_div)
        div_1y   = df_div[df_div["_date"] >= pd.Timestamp(datetime.now() - timedelta(days=500))]
        cash_sum = div_1y["_cash"].sum()
        if not np.isnan(price) and price > 0 and cash_sum > 0:
            div_yield = cash_sum / price
        div_count = int((df_div["_cash"] > 0).sum())
        # 3年均股利
        df_yr = df_div[df_div["_cash"] > 0].copy()
        df_yr["_yr"] = df_yr["_date"].dt.year
        yr_sum = df_yr.groupby("_yr")["_cash"].sum()
        avg_div_3y_b = yr_sum.tail(3).mean() if not yr_sum.empty else np.nan
    else:
        avg_div_3y_b = np.nan

    # 月營收
    start_rev = (datetime.now() - timedelta(days=760)).strftime("%Y-%m-%d")
    df_rev, _ = finmind_get("TaiwanStockMonthRevenue", stock_id, start_rev, token)
    ytd, mom  = np.nan, np.nan
    if not df_rev.empty:
        df_rev["date"]    = pd.to_datetime(df_rev["date"])
        df_rev["revenue"] = pd.to_numeric(df_rev["revenue"], errors="coerce")
        df_rev = df_rev.dropna(subset=["revenue"]).sort_values("date")
        cur_year = datetime.now().year
        cur_df   = df_rev[df_rev["date"].dt.year == cur_year]
        prev_df  = df_rev[df_rev["date"].dt.year == cur_year - 1]
        if not cur_df.empty and not prev_df.empty:
            months = set(cur_df["date"].dt.month)
            ps = prev_df[prev_df["date"].dt.month.isin(months)]
            if not ps.empty and ps["revenue"].sum() != 0:
                ytd = (cur_df["revenue"].sum() / ps["revenue"].sum() - 1) * 100
        if len(df_rev) >= 2:
            c, p = df_rev.iloc[-1]["revenue"], df_rev.iloc[-2]["revenue"]
            if p != 0:
                mom = (c - p) / p * 100

    # EPS
    start_eps = (datetime.now() - timedelta(days=1100)).strftime("%Y-%m-%d")
    df_eps, _ = finmind_get("TaiwanStockFinancialStatements", stock_id, start_eps, token)
    eps_yoy, est_yield = np.nan, np.nan
    if not df_eps.empty and "type" in df_eps.columns:
        e = df_eps[df_eps["type"] == "EPS"][["date", "value"]].copy()
        e["value"] = pd.to_numeric(e["value"], errors="coerce")
        e["date"]  = pd.to_datetime(e["date"])
        e = e.dropna().sort_values("date").drop_duplicates("date")
        if len(e) >= 8:
            eps_yoy = e.tail(4)["value"].sum() / e.iloc[-8:-4]["value"].sum() \
                      if e.iloc[-8:-4]["value"].sum() != 0 else np.nan
        e_yr = e.groupby(e["date"].dt.year)["value"].sum()
        avg_eps_3y_b = e_yr.tail(3).mean() if not e_yr.empty else np.nan
        if not np.isnan(avg_eps_3y_b) and avg_eps_3y_b > 0 and not np.isnan(price) and price > 0:
            payout = min(avg_div_3y_b / avg_eps_3y_b, 1.5) \
                     if not np.isnan(avg_div_3y_b) and avg_eps_3y_b > 0 else 0.6
            est_yield = (avg_eps_3y_b * payout) / price

    return div_yield, div_count, ytd, mom, eps_yoy, est_yield


def get_tw_stock_list(token):
    """
    回傳所有台股股票的 (stock_id, stock_name) 清單。
    供批次篩選用（只取代號）與個股搜尋用（代號+名稱）。
    """
    try:
        params = {"dataset": "TaiwanStockInfo"}
        if token:
            params["token"] = token
        r = requests.get(FINMIND_BASE, params=params, timeout=20)
        data = r.json()
        if data.get("status") == 200 and data.get("data"):
            df = pd.DataFrame(data["data"])
            df["stock_id"] = df["stock_id"].astype(str)
            # 篩選4~5碼（上市上櫃一般股 + 部分ETF如00878）
            df = df[df["stock_id"].str.match(r"^\d{4,5}$")].copy()
            name_col = next((c for c in ["stock_name", "name", "Name"] if c in df.columns), None)
            if name_col:
                df["_label"] = df["stock_id"] + "  " + df[name_col].fillna("")
            else:
                df["_label"] = df["stock_id"]
            pairs = list(zip(df["stock_id"].tolist(), df["_label"].tolist()))
            return sorted(pairs, key=lambda x: x[0])
    except Exception:
        pass
    return []


@st.cache_data(ttl=3600, show_spinner=False)
def load_stock_options(token: str):
    """
    啟動時快取完整股票清單，回傳：
      options_labels : ["2330  台積電", "2882  國泰金", ...]  ← selectbox 用
      label_to_id   : {"2330  台積電": "2330", ...}         ← 反查代號用
      id_list       : ["2330", "2882", ...]                 ← 批次掃描用
    """
    pairs = get_tw_stock_list(token)
    if not pairs:
        return [], {}, []
    labels    = [p[1] for p in pairs]
    label2id  = {p[1]: p[0] for p in pairs}
    id_list   = [p[0] for p in pairs]
    return labels, label2id, id_list


@st.cache_data(ttl=3600, show_spinner=False)
def run_screening(min_yield, min_div_count, min_rev_ytd_yoy, min_rev_mom,
                  min_eps_yoy, min_eps_est_yield, token, stock_ids):
    results = []
    progress = st.progress(0, text="開始掃描...")
    total = len(stock_ids)
    for i, sid in enumerate(stock_ids):
        progress.progress((i + 1) / total, text=f"分析中：{sid}（{i+1}/{total}）")
        try:
            dy, dc, ytd, mom, ey, eey = calc_metrics_batch(sid, token)
            ok = True
            if min_yield > -1         and (np.isnan(dy)  or dy * 100 < min_yield):             ok = False
            if min_div_count > -1     and dc < min_div_count:                                   ok = False
            if min_rev_ytd_yoy > -100 and (np.isnan(ytd) or ytd < min_rev_ytd_yoy):            ok = False
            if min_rev_mom > -100     and (np.isnan(mom) or mom < min_rev_mom):                 ok = False
            if min_eps_yoy > -1       and (np.isnan(ey)  or ey  < min_eps_yoy):                ok = False
            if min_eps_est_yield > -1 and (np.isnan(eey) or eey * 100 < min_eps_est_yield):     ok = False
            if ok:
                results.append({
                    "代號":              sid,
                    "殖利率(%)":         round(dy * 100, 2)  if not np.isnan(dy)  else None,
                    "10年股利次數":       dc,
                    "累計營收年增率(%)":  round(ytd, 2)       if not np.isnan(ytd) else None,
                    "營收與前月比(%)":    round(mom, 2)       if not np.isnan(mom) else None,
                    "EPS年度比較":        round(ey, 4)        if not np.isnan(ey)  else None,
                    "EPS推隔年殖利率(%)": round(eey * 100, 2) if not np.isnan(eey) else None,
                })
        except Exception:
            pass
        time.sleep(0.25)
    progress.empty()
    cols = ["代號","殖利率(%)","10年股利次數","累計營收年增率(%)","營收與前月比(%)","EPS年度比較","EPS推隔年殖利率(%)"]
    return pd.DataFrame(results) if results else pd.DataFrame(columns=cols)


# ─────────────────────────────────────────────────────
# 側邊欄
# ─────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 篩選條件")
    st.caption("對應 Excel「總表 2」第一列")
    st.divider()
    min_yield         = st.number_input("殖利率 ≥ (%)",         min_value=-1.0,   max_value=100.0, value=5.0,    step=0.5,  help="-1=不篩選")
    min_div_count     = st.number_input("10年股利次數 ≥",        min_value=-1,     max_value=40,    value=9,      step=1,    help="-1=不篩選")
    min_rev_ytd_yoy   = st.number_input("累計營收年增率 ≥ (%)",  min_value=-100.0, max_value=200.0, value=-100.0, step=5.0,  help="-100=不篩選")
    min_rev_mom       = st.number_input("營收與前月比 ≥ (%)",    min_value=-100.0, max_value=200.0, value=-100.0, step=5.0,  help="-100=不篩選")
    min_eps_yoy       = st.number_input("EPS年度比較 ≥",         min_value=-1.0,   max_value=5.0,   value=0.7,    step=0.05, help="-1=不篩選")
    min_eps_est_yield = st.number_input("EPS推隔年殖利率 ≥ (%)", min_value=-1.0,   max_value=30.0,  value=6.0,    step=0.5,  help="-1=不篩選")
    st.divider()
    if _SECRET_TOKEN:
        finmind_token = _SECRET_TOKEN
        st.success("✅ Token 已從 Secrets 載入", icon="🔑")
    else:
        finmind_token = st.text_input("FinMind Token（選填）", type="password")
    max_stocks = st.slider("掃描上限（支）", 20, 500, 100, 20)
    st.divider()
    run_btn = st.button("🚀 RUN 篩選", use_container_width=True, type="primary")


# ─────────────────────────────────────────────────────
# 主頁面
# ─────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📊 批次篩選（RUN）", "🔍 個股儀表板", "🛠 API 診斷"])


# ══════════════════════════════════════════════════════
# Tab1：批次篩選
# ══════════════════════════════════════════════════════
with tab1:
    st.title("📊 批次篩選")
    st.markdown("""
| Excel 欄位 | 篩選邏輯 | 預設門檻 |
|---|---|---|
| 殖利率 | ≥ 門檻 | 5% |
| 10年股利次數 | ≥ 門檻 | 9次 |
| 累計營收年增率 | ≥ 門檻 | -100（不篩選）|
| 營收增率與前月比 | ≥ 門檻 | -100（不篩選）|
| EPS年度比較 | ≥ 門檻 | 0.7 |
| EPS推隔年殖利率 | ≥ 門檻 | 6% |
""")
    if run_btn:
        with st.spinner("取得股票清單..."):
            _, __, all_ids = load_stock_options(finmind_token)
            if not all_ids:
                all_ids = [str(i) for i in range(1101, 3700)] + [str(i) for i in range(4100, 6900)]
            stock_sample = all_ids[:max_stocks]
        st.info(f"掃描 {len(stock_sample)} 支，約需 {len(stock_sample)//4} 秒...")
        result_df = run_screening(
            min_yield, min_div_count, min_rev_ytd_yoy, min_rev_mom,
            min_eps_yoy, min_eps_est_yield, finmind_token, tuple(stock_sample)
        )
        if result_df.empty:
            st.warning("無符合條件的股票，請嘗試放寬門檻。")
        else:
            st.success(f"✅ 共 **{len(result_df)}** 支通過篩選")
            st.dataframe(result_df, use_container_width=True, hide_index=True)
            st.download_button("⬇️ 下載 CSV",
                result_df.to_csv(index=False, encoding="utf-8-sig"), "result.csv", "text/csv")


# ══════════════════════════════════════════════════════
# Tab2：個股儀表板
# ══════════════════════════════════════════════════════
with tab2:
    st.title("🔍 個股儀表板")

    # ── 啟動時載入完整股票清單（快取，不重複請求）────────────
    with st.spinner("載入股票清單..."):
        _labels, _label2id, _id_list = load_stock_options(finmind_token)

    if _labels:
        # st.selectbox 支援鍵盤輸入即時過濾，在手機/桌面都能搜尋
        # 使用者輸入「國泰」→ 自動列出含「國泰」的所有選項
        selected_label = st.selectbox(
            "搜尋股票（輸入代號或名稱關鍵字）",
            options=[""] + _labels,          # 第一項空白 = 未選取
            index=0,
            placeholder="輸入代號或關鍵字，例如：2882 或 國泰",
            key="stock_selectbox",
        )
        sid = _label2id.get(selected_label, "").strip() if selected_label else ""
    else:
        # FinMind 無法取得清單時，退回手動輸入
        st.warning("⚠️ 無法載入股票清單（FinMind API 限流或無 Token），請手動輸入代號")
        manual = st.text_input("手動輸入股票代號", placeholder="例如：2882", key="stock_manual")
        sid = manual.strip()

    # 有選到股票就立即查詢（不需按鈕，選即查）
    if sid:
        with st.spinner(f"正在抓取 {sid} 數據..."):
            d = get_stock_data(sid, finmind_token)

        price   = d.get("price", np.nan)
        company = d.get("company", sid)

        # ── 標題 ──────────────────────────────────────
        st.markdown(f"## {sid}・{company}")
        st.caption(d.get("industry", ""))

        # ══ 股價走勢圖（一年，手機可視） ══════════════
        section("📉 近一年股價走勢")
        price_hist = d.get("price_hist", pd.DataFrame())
        if not price_hist.empty:
            # 標記最高/最低點
            max_p = price_hist["收盤價"].max()
            min_p = price_hist["收盤價"].min()
            now_p = price_hist["收盤價"].iloc[-1] if not price_hist.empty else np.nan
            pct_from_low  = (now_p - min_p) / min_p * 100 if min_p > 0 else np.nan
            pct_from_high = (now_p - max_p) / max_p * 100 if max_p > 0 else np.nan

            # 三欄顯示高低點資訊
            sc1, sc2, sc3 = st.columns(3)
            sc1.metric("52週最高", f"{max_p:.1f}", delta=None)
            sc2.metric("52週最低", f"{min_p:.1f}", delta=None)
            sc3.metric("距最低",
                       f"+{pct_from_low:.1f}%" if not np.isnan(pct_from_low) else "N/A",
                       delta=f"{pct_from_high:.1f}% (距最高)" if not np.isnan(pct_from_high) else None)

            st.line_chart(price_hist, height=220, use_container_width=True)

            # 便宜/合理/昂貴參考線說明
            pe_chp  = d.get("price_cheap", np.nan)
            pe_fair = d.get("price_fair", np.nan)
            pe_exp  = d.get("price_expensive", np.nan)
            if not np.isnan(pe_chp):
                st.caption(
                    f"💚 便宜價 {pe_chp:.1f}　　🟡 合理價 {pe_fair:.1f}　　🔴 昂貴價 {pe_exp:.1f}　　"
                    f"（基準：3年均股利 {fmt(d.get('avg_div_3y', np.nan))} 元）"
                )
        else:
            st.info("無法取得歷史股價，請確認股票代號格式（如：2330、2882）")

        # ══ 即時行情 ══════════════════════════════════
        section("📡 即時行情")

        # 判斷價格區間
        pe_chp  = d.get("price_cheap", np.nan)
        pe_fair = d.get("price_fair", np.nan)
        pe_exp  = d.get("price_expensive", np.nan)
        if not np.isnan(price) and not np.isnan(pe_chp):
            if price <= pe_chp:
                zone, zone_color = "✅ 便宜區", "green"
            elif price <= pe_fair:
                zone, zone_color = "🟡 合理偏低區", "green"
            elif price <= pe_exp:
                zone, zone_color = "🟠 合理偏高區", "gold"
            else:
                zone, zone_color = "🔴 昂貴區", "red"
        else:
            zone, zone_color = "—", "blue"

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(card("股價（元）", fmt(price, 1), sub=zone, color=zone_color), unsafe_allow_html=True)
        with c2:
            pe = d.get("pe_ratio", np.nan)
            st.markdown(card("本益比", fmt(pe, 2), sub="股價/近4季EPS", color="blue"), unsafe_allow_html=True)
        with c3:
            pbr = d.get("pbr", np.nan)
            st.markdown(card("股價淨值比", fmt(pbr, 2), color="blue"), unsafe_allow_html=True)

        c4, c5 = st.columns(2)
        with c4:
            ytd_val = d.get("rev_ytd_yoy", np.nan)
            col = "green" if (not np.isnan(ytd_val) and ytd_val >= 0) else "red"
            st.markdown(card("累積營收年增率", fmt(ytd_val, 2, "%"), sub="今年累積 vs 去年同期", color=col), unsafe_allow_html=True)
        with c5:
            mom_val = d.get("rev_mom", np.nan)
            col = "green" if (not np.isnan(mom_val) and mom_val >= 0) else "red"
            st.markdown(card("營收與前月比", fmt(mom_val, 2, "%"), sub="最新月 vs 上月", color=col), unsafe_allow_html=True)

        # ══ 股利政策 ══════════════════════════════════
        section("💰 股利政策（當年度）")

        c6, c7, c8 = st.columns(3)
        with c6:
            dy = d.get("div_yield", np.nan)
            dy_str = f"{dy*100:.2f}%" if not np.isnan(dy) else "N/A"
            st.markdown(card("殖利率", dy_str, sub="近一年現金股利/股價", color="gold"), unsafe_allow_html=True)
        with c7:
            cd = d.get("cash_div", np.nan)
            st.markdown(card("配息（元/股）", fmt(cd), sub="近一年現金股利", color="gold"), unsafe_allow_html=True)
        with c8:
            pr = d.get("payout_ratio", np.nan)
            pr_str = f"{pr*100:.1f}%" if not np.isnan(pr) else "N/A"
            st.markdown(card("配息率", pr_str, sub="現金股利/去年EPS", color="gold"), unsafe_allow_html=True)

        c9, c10, c11 = st.columns(3)
        with c9:
            st.markdown(card("3年平均股利", fmt(d.get("avg_div_3y")), sub="價格推估基準", color="gold"), unsafe_allow_html=True)
        with c10:
            st.markdown(card("6年平均股利", fmt(d.get("avg_div_6y")), color="gold"), unsafe_allow_html=True)
        with c11:
            st.markdown(card("10年平均股利", fmt(d.get("avg_div_10y")), color="gold"), unsafe_allow_html=True)

        c12, c13, c14 = st.columns(3)
        with c12:
            st.markdown(card("除息日", d.get("ex_div_date", "—"), color="blue"), unsafe_allow_html=True)
        with c13:
            st.markdown(card("發息日", d.get("pay_date", "—"), color="blue"), unsafe_allow_html=True)
        with c14:
            st.markdown(card("10年股利次數", str(d.get("div_count_10y", 0)), color="blue"), unsafe_allow_html=True)

        # ══ EPS 表現 ══════════════════════════════════
        section("📊 EPS 表現")

        c15, c16, c17, c18 = st.columns(4)
        with c15:
            st.markdown(card("1Q EPS", fmt(d.get("eps_q1")), color="blue"), unsafe_allow_html=True)
        with c16:
            st.markdown(card("2Q EPS", fmt(d.get("eps_q2")), color="blue"), unsafe_allow_html=True)
        with c17:
            st.markdown(card("3Q EPS", fmt(d.get("eps_q3")), color="blue"), unsafe_allow_html=True)
        with c18:
            st.markdown(card("去年 EPS", fmt(d.get("eps_prev")), sub="前4季加總", color="blue"), unsafe_allow_html=True)

        c19, c20 = st.columns(2)
        with c19:
            st.markdown(card("今年累積 EPS", fmt(d.get("eps_ytd")), sub="近4季加總", color="blue"), unsafe_allow_html=True)
        with c20:
            ey_val = d.get("eps_yoy", np.nan)
            ey_col = "green" if (not np.isnan(ey_val) and ey_val >= 1) else "red"
            st.markdown(card("EPS年度比較", fmt(ey_val, 4), sub="近4季/前4季（≥1=成長）", color=ey_col), unsafe_allow_html=True)

        # ══ 價格推估 ══════════════════════════════════
        section("🏷️ 價格推估（以3年平均股利 / 殖利率反推）")

        pc1, pc2, pc3 = st.columns(3)
        with pc1:
            st.markdown(card("昂貴（4%）", fmt(d.get("price_expensive"), 1),
                             sub=f"3年均股利{fmt(d.get('avg_div_3y'))}÷4%", color="red"), unsafe_allow_html=True)
        with pc2:
            st.markdown(card("合理（5.5%）", fmt(d.get("price_fair"), 1),
                             sub=f"3年均股利{fmt(d.get('avg_div_3y'))}÷5.5%", color="gold"), unsafe_allow_html=True)
        with pc3:
            st.markdown(card("便宜（7%）", fmt(d.get("price_cheap"), 1),
                             sub=f"3年均股利{fmt(d.get('avg_div_3y'))}÷7%", color="green"), unsafe_allow_html=True)

        # ══ 未來展望推估（截圖紅字重點） ══════════════
        section("🔮 未來展望推估（對應 Excel 紅字欄位）")

        ea  = d.get("est_yield_avg", np.nan)
        el  = d.get("est_yield_latest", np.nan)
        epa = d.get("est_pe_avg", np.nan)
        epl = d.get("est_pe_latest", np.nan)
        ae  = d.get("avg_eps_3y", np.nan)
        p3y = d.get("payout_3y", np.nan)

        f1, f2 = st.columns(2)
        with f1:
            ea_str = f"{ea*100:.2f}%" if not np.isnan(ea) else "N/A"
            st.markdown(card(
                "推隔年殖利率（平均加總）",
                ea_str,
                sub=f"3年均EPS {fmt(ae)} × 配息率{fmt(p3y*100,1,'%')} ÷ 股價",
                color="red"
            ), unsafe_allow_html=True)
        with f2:
            el_str = f"{el*100:.2f}%" if not np.isnan(el) else "N/A"
            st.markdown(card(
                "推隔年殖利率（前季推估）",
                el_str,
                sub=f"累積EPS {fmt(d.get('eps_ytd'))} × 配息率{fmt(p3y*100,1,'%')} ÷ 股價",
                color="red"
            ), unsafe_allow_html=True)

        f3, f4 = st.columns(2)
        with f3:
            st.markdown(card(
                "推隔年本益比（平均加總）",
                fmt(epa, 2),
                sub=f"股價 ÷ 3年均EPS {fmt(ae)}",
                color="purple"
            ), unsafe_allow_html=True)
        with f4:
            st.markdown(card(
                "推隔年本益比（前季推估）",
                fmt(epl, 2),
                sub=f"股價 ÷ 累積EPS {fmt(d.get('eps_ytd'))}",
                color="purple"
            ), unsafe_allow_html=True)

        # 推估股價區間（以推估股利反推）
        st.markdown("**📌 以推估EPS計算的股價參考區間：**")
        if not np.isnan(ae) and not np.isnan(p3y):
            est_div = ae * p3y
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                st.markdown(card("推估昂貴（4%）", fmt(est_div / 0.04, 1),
                                 sub=f"推估股利{fmt(est_div)}÷4%", color="red"), unsafe_allow_html=True)
            with fc2:
                st.markdown(card("推估合理（5.5%）", fmt(est_div / 0.055, 1),
                                 sub=f"推估股利{fmt(est_div)}÷5.5%", color="gold"), unsafe_allow_html=True)
            with fc3:
                st.markdown(card("推估便宜（7%）", fmt(est_div / 0.07, 1),
                                 sub=f"推估股利{fmt(est_div)}÷7%", color="green"), unsafe_allow_html=True)

        # ── 摺疊明細 ──────────────────────────────────
        with st.expander("📋 股利明細（近10年）"):
            df_div_show = d.get("df_div_raw", pd.DataFrame())
            if not df_div_show.empty:
                df_div_show = df_div_show.copy()
                df_div_show["現金股利合計"] = extract_cash_dividend(df_div_show)
                base_cols = ["date", "year", "現金股利合計"]
                extra_cols = [c for c in ["CashEarningsDistribution", "CashStatutorySurplus",
                               "StockEarningsDistribution", "CashExDividendTradingDate",
                               "CashDividendPaymentDate"] if c in df_div_show.columns]
                show_cols = base_cols + extra_cols
                st.dataframe(df_div_show[show_cols].sort_values("date", ascending=False),
                             use_container_width=True, hide_index=True)
            else:
                st.info("無股利資料")

        with st.expander("📋 月營收明細（近2年）"):
            df_rev_show = d.get("df_rev", pd.DataFrame())
            if not df_rev_show.empty:
                st.dataframe(df_rev_show.sort_values("date", ascending=False).head(24),
                             use_container_width=True, hide_index=True)
            else:
                st.info("無月營收資料")

    elif not sid:
        st.info("👆 請在上方搜尋並選擇股票")


# ══════════════════════════════════════════════════════
# Tab3：API 診斷
# ══════════════════════════════════════════════════════
with tab3:
    st.title("🛠 API 診斷")
    st.markdown("**出現 N/A 請先在此診斷，確認 FinMind API 實際回傳的欄位名稱與資料內容。**")

    diag_id  = st.text_input("診斷股票代號", value="2330", key="diag")
    diag_btn = st.button("執行診斷", key="diag_btn")

    if diag_btn:
        start_diag = (datetime.now() - timedelta(days=400)).strftime("%Y-%m-%d")

        st.markdown("### 1️⃣ yfinance 股價與走勢")
        try:
            t = yf.Ticker(f"{diag_id}.TW")
            p = t.fast_info.last_price
            st.success(f"✅ 現價：{p}")
            hist_d = t.history(period="1y")
            st.write(f"歷史資料筆數：{len(hist_d)}，欄位：{list(hist_d.columns)}")
        except Exception as e:
            st.error(f"❌ {e}")

        st.markdown("### 2️⃣ TaiwanStockDividend（股利）— 欄位自動偵測")
        df2, raw2 = finmind_get("TaiwanStockDividend", diag_id, start_diag, finmind_token)
        st.write(f"status={raw2.get('status')} msg={raw2.get('msg','')}")
        if not df2.empty:
            st.success(f"✅ {len(df2)} 筆")
            st.write("**實際欄位：**", list(df2.columns))
            df2["_自動偵測現金股利"] = extract_cash_dividend(df2)
            st.write("**自動偵測結果（前5筆）：**")
            preview_cols = ["date"] + [c for c in ["year","_自動偵測現金股利",
                "CashEarningsDistribution","CashStatutorySurplus","cash_dividend",
                "stock_or_cash"] if c in df2.columns or c == "_自動偵測現金股利"]
            st.dataframe(df2[preview_cols].head(5), use_container_width=True)
        else:
            st.error("❌ 無資料")
            st.json(raw2)

        st.markdown("### 3️⃣ TaiwanStockMonthRevenue（月營收）")
        df3, raw3 = finmind_get("TaiwanStockMonthRevenue", diag_id, start_diag, finmind_token)
        st.write(f"status={raw3.get('status')} msg={raw3.get('msg','')}")
        if not df3.empty:
            st.success(f"✅ {len(df3)} 筆 ｜ 欄位：{list(df3.columns)}")
            st.dataframe(df3.head(5), use_container_width=True)
        else:
            st.error("❌ 無資料"); st.json(raw3)

        st.markdown("### 4️⃣ TaiwanStockFinancialStatements（EPS）")
        start_eps = (datetime.now() - timedelta(days=1100)).strftime("%Y-%m-%d")
        df4, raw4 = finmind_get("TaiwanStockFinancialStatements", diag_id, start_eps, finmind_token)
        st.write(f"status={raw4.get('status')} msg={raw4.get('msg','')}")
        if not df4.empty:
            st.success(f"✅ {len(df4)} 筆 ｜ 欄位：{list(df4.columns)}")
            if "type" in df4.columns:
                st.write("type 唯一值：", df4["type"].unique().tolist())
                st.dataframe(df4[df4["type"] == "EPS"].head(8), use_container_width=True)
        else:
            st.error("❌ 無資料"); st.json(raw4)

        st.markdown("### 5️⃣ 完整指標計算結果")
        with st.spinner("計算中..."):
            d_diag = get_stock_data(diag_id, finmind_token)

        rows = [
            ("股價",                  fmt(d_diag.get("price"))),
            ("殖利率",                f"{d_diag['div_yield']*100:.2f}%" if not np.isnan(d_diag.get("div_yield", np.nan)) else "❌"),
            ("配息（元）",             fmt(d_diag.get("cash_div"))),
            ("10年股利次數",           d_diag.get("div_count_10y", 0)),
            ("3年平均股利",            fmt(d_diag.get("avg_div_3y"))),
            ("去年EPS",               fmt(d_diag.get("eps_prev"))),
            ("今年累積EPS",            fmt(d_diag.get("eps_ytd"))),
            ("3年平均EPS",             fmt(d_diag.get("avg_eps_3y"))),
            ("配息率(3年均)",          fmt(d_diag.get("payout_3y"), 3)),
            ("累積營收年增率",         fmt(d_diag.get("rev_ytd_yoy"), 2, "%")),
            ("EPS年度比較",            fmt(d_diag.get("eps_yoy"), 4)),
            ("推隔年殖利率(均)",        f"{d_diag['est_yield_avg']*100:.2f}%" if not np.isnan(d_diag.get("est_yield_avg", np.nan)) else "❌"),
            ("推隔年殖利率(前季)",      f"{d_diag['est_yield_latest']*100:.2f}%" if not np.isnan(d_diag.get("est_yield_latest", np.nan)) else "❌"),
            ("推隔年本益比(均)",        fmt(d_diag.get("est_pe_avg"), 2)),
            ("推隔年本益比(前季)",      fmt(d_diag.get("est_pe_latest"), 2)),
            ("昂貴價(4%)",             fmt(d_diag.get("price_expensive"), 1)),
            ("合理價(5.5%)",           fmt(d_diag.get("price_fair"), 1)),
            ("便宜價(7%)",             fmt(d_diag.get("price_cheap"), 1)),
        ]
        st.dataframe(pd.DataFrame(rows, columns=["指標", "計算值"]),
                     use_container_width=True, hide_index=True)


st.divider()
st.caption("⚠️ 本工具數據僅供參考，不構成投資建議。數據來源：FinMind / yfinance。")
