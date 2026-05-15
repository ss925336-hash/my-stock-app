"""
台股自動化篩選 Streamlit Web App  ─ 最終完整版
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
  推隔年本益比(平均加總)     → 股價 / 3年均EPS
  昂貴(4%) / 合理(5.5%) / 便宜(7%) → 3年平均股利 / 殖利率門檻

架構說明：
  按鈕A「🔄 更新市場數據」→ 抓成交前500個股 + 預算所有指標（24h快取）
  按鈕B「🎯 執行條件篩選」→ 僅對快取數據本地過濾，不耗API
  Tab2 個股儀表板 → 三步驟搜尋：輸入關鍵字 → 選清單 → 確認執行
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
/* ── 卡片基礎樣式 ── */
.card {
    background: #1e2130;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 12px;
    border-left: 4px solid #4a9eff;
}
.card-red    { border-left-color: #ff4b4b; }
.card-green  { border-left-color: #21c55d; }
.card-gold   { border-left-color: #f5a623; }
.card-blue   { border-left-color: #4a9eff; }
.card-purple { border-left-color: #a855f7; }

.card-title { font-size:0.72rem; color:#8b9ab0; text-transform:uppercase;
              letter-spacing:0.08em; margin-bottom:4px; }
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

/* ── 手機優化 ── */
@media (max-width: 768px) {
    .card-value { font-size:1.3rem; }
    .card { padding: 12px 14px; }
}
</style>
""", unsafe_allow_html=True)

FINMIND_BASE = "https://api.finmindtrade.com/api/v4/data"


# ═══════════════════════════════════════════════════════
# 工具函式
# ═══════════════════════════════════════════════════════

def finmind_get(dataset, stock_id, start_date, token=""):
    """通用 FinMind API 請求"""
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


def finmind_get_no_id(dataset, start_date, end_date, token=""):
    """不帶 stock_id 的全市場請求（用於掃描池）"""
    params = {
        "dataset": dataset,
        "start_date": start_date,
        "end_date":   end_date,
    }
    if token:
        params["token"] = token
    try:
        r = requests.get(FINMIND_BASE, params=params, timeout=30)
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
    """
    if df.empty:
        return pd.Series(dtype=float)

    # ── 結構 A：寬表格 ────────────────────────────────
    cash_cols = [c for c in df.columns
                 if any(k in c for k in ["CashEarnings", "CashStatutory",
                                          "cashearnings", "cashstatutory"])]
    if cash_cols:
        result = pd.Series(0.0, index=df.index)
        for col in cash_cols:
            result += pd.to_numeric(df[col], errors="coerce").fillna(0)
        return result

    # ── 結構 B：長表格 ────────────────────────────────
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

    # ── Fallback ──────────────────────────────────────
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


# ═══════════════════════════════════════════════════════
# 個股完整資料抓取（對應 Excel 總表）
# ═══════════════════════════════════════════════════════

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
        d["price_hist"] = (hist[["Close"]].rename(columns={"Close": "收盤價"})
                           if not hist.empty else pd.DataFrame())
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
        df_div["_date"]  = pd.to_datetime(df_div["date"])
        df_div["_cash"]  = extract_cash_dividend(df_div)
        df_div["_stock"] = pd.to_numeric(
            df_div.get("StockEarningsDistribution",
            df_div.get("stock_dividend", 0)), errors="coerce").fillna(0)

        # 近一年現金股利加總（用於殖利率計算）
        cutoff_1y = pd.Timestamp(datetime.now() - timedelta(days=500))
        div_1y    = df_div[df_div["_date"] >= cutoff_1y]
        cash_sum  = div_1y["_cash"].sum()
        d["cash_div"]  = cash_sum if cash_sum > 0 else np.nan
        d["stock_div"] = div_1y["_stock"].sum()

        # 除息日、發息日
        has_cash = df_div[df_div["_cash"] > 0].sort_values("_date", ascending=False)
        if not has_cash.empty:
            latest = has_cash.iloc[0]
            d["ex_div_date"] = str(latest["_date"])[:10]
            pay_col = next((c for c in ["CashDividendPaymentDate", "pay_date"]
                            if c in df_div.columns), None)
            d["pay_date"] = str(latest[pay_col])[:10] if pay_col else "—"
        else:
            d["ex_div_date"] = "—"
            d["pay_date"]    = "—"

        # 10年股利次數
        d["div_count_10y"] = int((df_div["_cash"] > 0).sum())

        # 3/6/10年平均現金股利（按年分組）
        df_yr  = df_div[df_div["_cash"] > 0].copy()
        df_yr["_yr"] = df_yr["_date"].dt.year
        yr_sum = df_yr.groupby("_yr")["_cash"].sum().sort_index()
        d["avg_div_3y"]  = float(yr_sum.tail(3).mean())  if len(yr_sum) >= 1 else np.nan
        d["avg_div_6y"]  = float(yr_sum.tail(6).mean())  if len(yr_sum) >= 1 else np.nan
        d["avg_div_10y"] = float(yr_sum.tail(10).mean()) if len(yr_sum) >= 1 else np.nan
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
        if (not np.isnan(d.get("eps_ytd", np.nan))
                and not np.isnan(d.get("eps_prev", np.nan))
                and d["eps_prev"] != 0):
            d["eps_yoy"] = d["eps_ytd"] / d["eps_prev"]
        else:
            d["eps_yoy"] = np.nan

        # 本益比 = 股價 / 近4季EPS
        if not np.isnan(price) and not np.isnan(d.get("eps_ytd", np.nan)) and d["eps_ytd"] > 0:
            d["pe_ratio"] = price / d["eps_ytd"]
        else:
            d["pe_ratio"] = np.nan

        # 近3年年度EPS → 平均（用於推估）
        e_yr = e.groupby(e["date"].dt.year)["value"].sum()
        d["avg_eps_3y"] = float(e_yr.tail(3).mean()) if not e_yr.empty else np.nan
    else:
        d.update({
            "eps_q1": np.nan, "eps_q2": np.nan,
            "eps_q3": np.nan, "eps_q4": np.nan,
            "eps_ytd": np.nan, "eps_prev": np.nan,
            "eps_yoy": np.nan, "pe_ratio": np.nan, "avg_eps_3y": np.nan,
        })

    # ── 配息率 = 近一年現金股利 / 去年EPS ────────────
    cash_div   = d.get("cash_div",   np.nan)
    eps_prev   = d.get("eps_prev",   np.nan)
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
    # 對應 Excel 紅字：(近3年平均EPS × 近3年平均配息率) / 股價
    if not np.isnan(avg_eps_3y) and avg_eps_3y > 0 and not np.isnan(price) and price > 0:
        est_div_avg        = avg_eps_3y * payout_3y
        d["est_yield_avg"] = est_div_avg / price
    else:
        d["est_yield_avg"] = np.nan

    # ── 推估EPS推隔年殖利率（前季推估）──────────────
    # 對應 Excel 紅字：近4季累積EPS × 近3年平均配息率 / 股價
    eps_ytd = d.get("eps_ytd", np.nan)
    if not np.isnan(eps_ytd) and eps_ytd > 0 and not np.isnan(price) and price > 0:
        d["est_yield_latest"] = (eps_ytd * payout_3y) / price
    else:
        d["est_yield_latest"] = np.nan

    # ── 推隔年本益比（平均加總）──────────────────────
    # 對應 Excel 紅字：股價 / 3年均EPS
    if not np.isnan(avg_eps_3y) and avg_eps_3y > 0 and not np.isnan(price):
        d["est_pe_avg"] = price / avg_eps_3y
    else:
        d["est_pe_avg"] = np.nan

    # ── 推隔年本益比（前季推估）──────────────────────
    if not np.isnan(eps_ytd) and eps_ytd > 0 and not np.isnan(price):
        d["est_pe_latest"] = price / eps_ytd
    else:
        d["est_pe_latest"] = np.nan

    # ── 價格推估區間（昂貴4% / 合理5.5% / 便宜7%）──
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

        cur_year = datetime.now().year
        cur_df   = df_rev[df_rev["date"].dt.year == cur_year]
        prev_df  = df_rev[df_rev["date"].dt.year == cur_year - 1]

        d["rev_ytd_yoy"] = np.nan
        if not cur_df.empty and not prev_df.empty:
            months    = set(cur_df["date"].dt.month)
            prev_same = prev_df[prev_df["date"].dt.month.isin(months)]
            if not prev_same.empty and prev_same["revenue"].sum() != 0:
                d["rev_ytd_yoy"] = (cur_df["revenue"].sum() / prev_same["revenue"].sum() - 1) * 100

        d["rev_mom"] = np.nan
        if len(df_rev) >= 2:
            c_r = df_rev.iloc[-1]["revenue"]
            p_r = df_rev.iloc[-2]["revenue"]
            if p_r != 0:
                d["rev_mom"] = (c_r - p_r) / p_r * 100

        d["df_rev"] = df_rev
    else:
        d["rev_ytd_yoy"] = np.nan
        d["rev_mom"]     = np.nan
        d["df_rev"]      = pd.DataFrame()

    d["df_div_raw"] = df_div if not df_div.empty else pd.DataFrame()
    return d


# ═══════════════════════════════════════════════════════
# 批次掃描 ── 按鈕A 系列（24h 快取）
# ═══════════════════════════════════════════════════════

@st.cache_data(ttl=86400, show_spinner=False)
def get_top500_scan_pool(token: str) -> list:
    """
    抓取最近一個交易日全市場成交數據，取成交金額前500名個股。
    排除所有 ETF（00開頭）、B/N結尾、含字母的權證。
    24小時全域快取。
    """
    for days_back in range(1, 8):
        target_date = (datetime.now() - timedelta(days=days_back)).strftime("%Y-%m-%d")
        try:
            df, raw = finmind_get_no_id("TaiwanStockPrice", target_date, target_date, token)
            if df.empty:
                continue

            df["stock_id"] = df["stock_id"].astype(str)
            # 嚴格過濾：4碼純數字、首碼1~9（排除00開頭ETF）
            valid_mask = (
                df["stock_id"].str.match(r"^[1-9]\d{3}$") &
                ~df["stock_id"].str.endswith("B") &
                ~df["stock_id"].str.endswith("N")
            )
            df = df[valid_mask].copy()
            if df.empty:
                continue

            close_col  = next((c for c in ["close", "Close", "closing_price"]
                               if c in df.columns), None)
            volume_col = next((c for c in ["Trading_Volume", "volume", "Volume"]
                               if c in df.columns), None)

            if close_col and volume_col:
                df["_amount"] = (
                    pd.to_numeric(df[close_col],  errors="coerce").fillna(0) *
                    pd.to_numeric(df[volume_col], errors="coerce").fillna(0)
                )
            else:
                num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
                df["_amount"] = df[num_cols[0]] if num_cols else 0

            top500 = (
                df.sort_values("_amount", ascending=False)
                .drop_duplicates("stock_id")
                .head(500)["stock_id"]
                .tolist()
            )
            if top500:
                return top500
        except Exception:
            continue
    return []


def _calc_single_metrics(stock_id, token):
    """輕量批次計算，供 batch_compute_all_metrics 呼叫"""
    price = get_price(stock_id)

    # 股利
    start_10y = (datetime.now() - timedelta(days=3650)).strftime("%Y-%m-%d")
    df_div, _ = finmind_get("TaiwanStockDividend", stock_id, start_10y, token)
    div_yield = avg_div_3y = np.nan
    div_count = 0

    if not df_div.empty:
        df_div["_date"] = pd.to_datetime(df_div["date"])
        df_div["_cash"] = extract_cash_dividend(df_div)
        div_1y   = df_div[df_div["_date"] >= pd.Timestamp(datetime.now() - timedelta(days=500))]
        cash_sum = div_1y["_cash"].sum()
        if not np.isnan(price) and price > 0 and cash_sum > 0:
            div_yield = cash_sum / price
        div_count = int((df_div["_cash"] > 0).sum())

        df_yr  = df_div[df_div["_cash"] > 0].copy()
        df_yr["_yr"] = df_yr["_date"].dt.year
        yr_sum = df_yr.groupby("_yr")["_cash"].sum()
        avg_div_3y = float(yr_sum.tail(3).mean()) if not yr_sum.empty else np.nan

    # 月營收
    start_rev = (datetime.now() - timedelta(days=760)).strftime("%Y-%m-%d")
    df_rev, _ = finmind_get("TaiwanStockMonthRevenue", stock_id, start_rev, token)
    ytd = mom = np.nan
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
            c_r = df_rev.iloc[-1]["revenue"]
            p_r = df_rev.iloc[-2]["revenue"]
            if p_r != 0:
                mom = (c_r - p_r) / p_r * 100

    # EPS
    start_eps = (datetime.now() - timedelta(days=1100)).strftime("%Y-%m-%d")
    df_eps, _ = finmind_get("TaiwanStockFinancialStatements", stock_id, start_eps, token)
    eps_yoy = est_yield = avg_eps_3y = np.nan
    if not df_eps.empty and "type" in df_eps.columns:
        e = df_eps[df_eps["type"] == "EPS"][["date", "value"]].copy()
        e["value"] = pd.to_numeric(e["value"], errors="coerce")
        e["date"]  = pd.to_datetime(e["date"])
        e = e.dropna().sort_values("date").drop_duplicates("date")
        if len(e) >= 8:
            s1 = e.tail(4)["value"].sum()
            s2 = e.iloc[-8:-4]["value"].sum()
            eps_yoy = s1 / s2 if s2 != 0 else np.nan
        e_yr = e.groupby(e["date"].dt.year)["value"].sum()
        avg_eps_3y = float(e_yr.tail(3).mean()) if not e_yr.empty else np.nan
        if (not np.isnan(avg_eps_3y) and avg_eps_3y > 0
                and not np.isnan(price) and price > 0):
            payout = (min(avg_div_3y / avg_eps_3y, 1.5)
                      if not np.isnan(avg_div_3y) and avg_eps_3y > 0 else 0.6)
            est_yield = (avg_eps_3y * payout) / price

    return {
        "price":       price,
        "div_yield":   div_yield,
        "div_count":   div_count,
        "rev_ytd_yoy": ytd,
        "rev_mom":     mom,
        "eps_yoy":     eps_yoy,
        "est_yield":   est_yield,
        "avg_div_3y":  avg_div_3y,
        "avg_eps_3y":  avg_eps_3y,
    }


@st.cache_data(ttl=86400, show_spinner=False)
def batch_compute_all_metrics(stock_ids: tuple, token: str) -> pd.DataFrame:
    """
    按鈕A 的核心：對整個掃描池預算所有指標，結果快取24小時。
    回傳 DataFrame，按鈕B 直接對此 DataFrame 做本地過濾。
    """
    results = []
    progress = st.progress(0, text="⏳ 預算市場數據中...")
    total = len(stock_ids)

    for i, sid in enumerate(stock_ids):
        progress.progress((i + 1) / total,
                          text=f"預算指標：{sid}（{i+1}/{total}）")
        try:
            m = _calc_single_metrics(sid, token)
            results.append({
                "代號":              sid,
                "股價":              m["price"],
                "殖利率(%)":         round(m["div_yield"] * 100, 2) if not np.isnan(m["div_yield"]) else np.nan,
                "10年股利次數":       m["div_count"],
                "累計營收年增率(%)":  round(m["rev_ytd_yoy"], 2) if not np.isnan(m["rev_ytd_yoy"]) else np.nan,
                "營收與前月比(%)":    round(m["rev_mom"], 2)       if not np.isnan(m["rev_mom"]) else np.nan,
                "EPS年度比較":        round(m["eps_yoy"], 4)       if not np.isnan(m["eps_yoy"]) else np.nan,
                "EPS推隔年殖利率(%)": round(m["est_yield"] * 100, 2) if not np.isnan(m["est_yield"]) else np.nan,
                "3年平均股利":        round(m["avg_div_3y"], 2)    if not np.isnan(m["avg_div_3y"]) else np.nan,
                "3年平均EPS":         round(m["avg_eps_3y"], 2)    if not np.isnan(m["avg_eps_3y"]) else np.nan,
            })
        except Exception:
            pass
        time.sleep(0.2)

    progress.empty()
    cols = ["代號","股價","殖利率(%)","10年股利次數",
            "累計營收年增率(%)","營收與前月比(%)","EPS年度比較",
            "EPS推隔年殖利率(%)","3年平均股利","3年平均EPS"]
    return pd.DataFrame(results, columns=cols) if results else pd.DataFrame(columns=cols)


def local_filter(df: pd.DataFrame,
                 min_yield, min_div_count, min_rev_ytd_yoy,
                 min_rev_mom, min_eps_yoy, min_eps_est_yield) -> pd.DataFrame:
    """按鈕B：僅對快取 DataFrame 做本地過濾，0 API 額度。"""
    mask = pd.Series([True] * len(df), index=df.index)
    if min_yield > -1:
        mask &= df["殖利率(%)"].notna() & (df["殖利率(%)"] >= min_yield)
    if min_div_count > -1:
        mask &= df["10年股利次數"] >= min_div_count
    if min_rev_ytd_yoy > -100:
        mask &= df["累計營收年增率(%)"].notna() & (df["累計營收年增率(%)"] >= min_rev_ytd_yoy)
    if min_rev_mom > -100:
        mask &= df["營收與前月比(%)"].notna() & (df["營收與前月比(%)"] >= min_rev_mom)
    if min_eps_yoy > -1:
        mask &= df["EPS年度比較"].notna() & (df["EPS年度比較"] >= min_eps_yoy)
    if min_eps_est_yield > -1:
        mask &= df["EPS推隔年殖利率(%)"].notna() & (df["EPS推隔年殖利率(%)"] >= min_eps_est_yield)
    return df[mask].sort_values("殖利率(%)", ascending=False).reset_index(drop=True)


# ═══════════════════════════════════════════════════════
# 股票清單（全市場，供 Tab2 搜尋）
# ═══════════════════════════════════════════════════════

def get_tw_stock_list(token):
    try:
        params = {"dataset": "TaiwanStockInfo"}
        if token:
            params["token"] = token
        r = requests.get(FINMIND_BASE, params=params, timeout=20)
        data = r.json()
        if data.get("status") == 200 and data.get("data"):
            df = pd.DataFrame(data["data"])
            df["stock_id"] = df["stock_id"].astype(str)
            mask = (
                df["stock_id"].str.match(r"^[1-9]\d{3}$") &
                ~df["stock_id"].str.endswith("B") &
                ~df["stock_id"].str.endswith("N")
            )
            df = df[mask].copy()
            name_col = next((c for c in ["stock_name", "name", "Name"]
                             if c in df.columns), None)
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
    pairs = get_tw_stock_list(token)
    if not pairs:
        return [], {}, []
    labels   = [p[1] for p in pairs]
    label2id = {p[1]: p[0] for p in pairs}
    id_list  = [p[0] for p in pairs]
    return labels, label2id, id_list


# ═══════════════════════════════════════════════════════
# 側邊欄
# ═══════════════════════════════════════════════════════
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

    # ── 雙按鈕 ─────────────────────────────────────────
    st.markdown("**市場數據操作**")
    update_btn = st.button("🔄 更新市場數據", use_container_width=True, type="primary",
                           help="抓取成交前500個股並預算所有指標（24h快取）")
    filter_btn = st.button("🎯 執行條件篩選", use_container_width=True,
                           help="對快取數據本地過濾，不耗API額度")
    st.divider()
    st.caption("⚠️ 數據僅供參考，不構成投資建議")


# ─────────────────────────────────────────────────────
# 主頁面
# ─────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📊 批次篩選", "🔍 個股儀表板", "🛠 API 診斷"])


# ═══════════════════════════════════════════════════════
# Tab1：批次篩選
# ═══════════════════════════════════════════════════════
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

    # ── Session State 管理快取 DataFrame ──────────────
    if "market_df" not in st.session_state:
        st.session_state["market_df"]        = pd.DataFrame()
        st.session_state["market_pool_size"] = 0
        st.session_state["market_updated"]   = None

    # ── 按鈕A：更新市場數據 ───────────────────────────
    if update_btn:
        with st.spinner("🔍 抓取市場成交排行，建立掃描池..."):
            top500 = get_top500_scan_pool(finmind_token)

        if top500:
            stock_sample = tuple(top500[:max_stocks])
            st.success(
                f"📊 掃描池：成交金額前 {len(top500)} 名中，取前 **{len(stock_sample)}** 支"
                "（已排除 ETF / 權證）"
            )
        else:
            _, __, all_ids = load_stock_options(finmind_token)
            if not all_ids:
                all_ids = [str(i) for i in range(1101, 3700)] + [str(i) for i in range(4100, 6900)]
            stock_sample = tuple(all_ids[:max_stocks])
            st.warning(f"⚠️ 無法取得成交排行，改用全市場清單（前 {len(stock_sample)} 支）")

        st.info(f"🔄 開始預算 {len(stock_sample)} 支個股指標，約需 {len(stock_sample) // 5} 秒...")

        with st.spinner("預算所有指標中（24h快取）..."):
            mdf = batch_compute_all_metrics(stock_sample, finmind_token)

        st.session_state["market_df"]        = mdf
        st.session_state["market_pool_size"] = len(stock_sample)
        st.session_state["market_updated"]   = datetime.now().strftime("%Y-%m-%d %H:%M")

        st.success(f"✅ 完成！共預算 **{len(mdf)}** 支，已快取24小時。"
                   f"請點「🎯 執行條件篩選」查看結果。")

    # ── 按鈕B：執行條件篩選（純本地，0 API）─────────
    if filter_btn:
        mdf = st.session_state.get("market_df", pd.DataFrame())
        if mdf.empty:
            st.warning("⚠️ 尚無快取數據，請先點「🔄 更新市場數據」。")
        else:
            updated_at = st.session_state.get("market_updated", "—")
            pool_size  = st.session_state.get("market_pool_size", 0)
            st.caption(f"📦 快取池：{pool_size} 支 ｜ 更新時間：{updated_at}")

            result_df = local_filter(
                mdf, min_yield, min_div_count, min_rev_ytd_yoy,
                min_rev_mom, min_eps_yoy, min_eps_est_yield
            )

            if result_df.empty:
                st.warning("無符合條件的股票，請嘗試放寬門檻。")
            else:
                st.success(f"✅ 共 **{len(result_df)}** 支通過篩選（本地過濾，0 API 額度）")
                st.dataframe(result_df, use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ 下載 CSV",
                    result_df.to_csv(index=False, encoding="utf-8-sig"),
                    "result.csv", "text/csv"
                )

    # ── 顯示目前快取狀況 ──────────────────────────────
    if (not update_btn and not filter_btn
            and not st.session_state.get("market_df", pd.DataFrame()).empty):
        updated_at = st.session_state.get("market_updated", "—")
        pool_size  = st.session_state.get("market_pool_size", 0)
        st.info(
            f"💾 已有快取數據（{pool_size} 支，更新於 {updated_at}）。"
            "點「🎯 執行條件篩選」立即過濾，或「🔄 更新市場數據」重新抓取。"
        )


# ═══════════════════════════════════════════════════════
# Tab2：個股儀表板（三步驟搜尋）
# ═══════════════════════════════════════════════════════
with tab2:
    st.title("🔍 個股儀表板")

    # ── 啟動時載入完整股票清單 ────────────────────────
    with st.spinner("載入股票清單..."):
        _labels, _label2id, _id_list = load_stock_options(finmind_token)

    # ══ 三步驟搜尋 ════════════════════════════════════
    # Step 1：輸入關鍵字
    keyword = st.text_input(
        "① 輸入代號或名稱關鍵字",
        placeholder="例如：2882 或 國泰",
        key="stock_kw",
    )

    sid = ""  # 最終確認的股票代號

    if keyword.strip():
        kw = keyword.strip()

        # Step 2：列出包含關鍵字的選項
        if _labels:
            matched = [lb for lb in _labels if kw in lb]
        else:
            matched = []

        if not matched:
            # 若完全匹配代號（4碼數字），直接接受
            if kw.isdigit() and len(kw) == 4:
                matched = [kw]
            else:
                st.warning(f"找不到包含「{kw}」的股票，請確認代號或名稱。")

        if matched:
            selected_label = st.selectbox(
                f"② 搜尋結果（共 {len(matched)} 筆），請選取目標：",
                options=[""] + matched,
                index=0,
                key="stock_select",
            )

            # Step 3：確認執行
            confirm_btn = st.button("✅ 確認執行", type="primary", key="confirm_stock")

            if selected_label and confirm_btn:
                sid = _label2id.get(selected_label, selected_label.split()[0]).strip()
            elif confirm_btn and not selected_label:
                st.warning("請先選擇一支股票。")

    # ── 個股儀表板主體 ────────────────────────────────
    if sid:
        with st.spinner(f"正在抓取 {sid} 數據..."):
            d = get_stock_data(sid, finmind_token)

        price   = d.get("price", np.nan)
        company = d.get("company", sid)

        st.markdown(f"## {sid}・{company}")
        st.caption(d.get("industry", ""))

        # ══ 近一年股價走勢圖 ══════════════════════════
        section("📉 近一年股價走勢")
        price_hist = d.get("price_hist", pd.DataFrame())
        if not price_hist.empty:
            max_p = price_hist["收盤價"].max()
            min_p = price_hist["收盤價"].min()
            now_p = price_hist["收盤價"].iloc[-1] if not price_hist.empty else np.nan
            pct_from_low  = (now_p - min_p) / min_p * 100 if min_p > 0 else np.nan
            pct_from_high = (now_p - max_p) / max_p * 100 if max_p > 0 else np.nan

            sc1, sc2, sc3 = st.columns(3)
            sc1.metric("52週最高", f"{max_p:.1f}")
            sc2.metric("52週最低", f"{min_p:.1f}")
            sc3.metric(
                "距最低",
                f"+{pct_from_low:.1f}%" if not np.isnan(pct_from_low) else "N/A",
                delta=f"{pct_from_high:.1f}% (距最高)" if not np.isnan(pct_from_high) else None
            )

            st.line_chart(price_hist, height=220, use_container_width=True)

            pe_chp  = d.get("price_cheap", np.nan)
            pe_fair = d.get("price_fair",  np.nan)
            pe_exp  = d.get("price_expensive", np.nan)
            if not np.isnan(pe_chp):
                st.caption(
                    f"💚 便宜價 {pe_chp:.1f}　　🟡 合理價 {pe_fair:.1f}　　🔴 昂貴價 {pe_exp:.1f}"
                    f"　　（基準：3年均股利 {fmt(d.get('avg_div_3y', np.nan))} 元）"
                )
        else:
            st.info("無法取得歷史股價，請確認股票代號格式（如：2330、2882）")

        # ══ 即時行情 ══════════════════════════════════
        section("📡 即時行情")

        pe_chp  = d.get("price_cheap",     np.nan)
        pe_fair = d.get("price_fair",      np.nan)
        pe_exp  = d.get("price_expensive", np.nan)
        if not np.isnan(price) and not np.isnan(pe_chp):
            if price <= pe_chp:
                zone, zone_color = "✅ 便宜區",   "green"
            elif price <= pe_fair:
                zone, zone_color = "🟡 合理偏低區", "green"
            elif price <= pe_exp:
                zone, zone_color = "🟠 合理偏高區", "gold"
            else:
                zone, zone_color = "🔴 昂貴區",   "red"
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
            st.markdown(card("累積營收年增率", fmt(ytd_val, 2, "%"),
                             sub="今年累積 vs 去年同期", color=col), unsafe_allow_html=True)
        with c5:
            mom_val = d.get("rev_mom", np.nan)
            col = "green" if (not np.isnan(mom_val) and mom_val >= 0) else "red"
            st.markdown(card("營收與前月比", fmt(mom_val, 2, "%"),
                             sub="最新月 vs 上月", color=col), unsafe_allow_html=True)

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
            st.markdown(card("EPS年度比較", fmt(ey_val, 4),
                             sub="近4季/前4季（≥1=成長）", color=ey_col), unsafe_allow_html=True)

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

        # ══ 未來展望推估（對應 Excel 紅字欄位）══════
        section("🔮 未來展望推估（對應 Excel 紅字欄位）")

        ea  = d.get("est_yield_avg",    np.nan)
        el  = d.get("est_yield_latest", np.nan)
        epa = d.get("est_pe_avg",       np.nan)
        epl = d.get("est_pe_latest",    np.nan)
        ae  = d.get("avg_eps_3y",       np.nan)
        p3y = d.get("payout_3y",        np.nan)

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
        if not np.isnan(ae) and not np.isnan(p3y) and ae > 0:
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
        else:
            st.info("EPS或配息率資料不足，無法推估股價區間")

        # ── 摺疊明細 ──────────────────────────────────
        with st.expander("📋 股利明細（近10年）"):
            df_div_show = d.get("df_div_raw", pd.DataFrame())
            if not df_div_show.empty:
                df_div_show = df_div_show.copy()
                df_div_show["現金股利合計"] = extract_cash_dividend(df_div_show)
                base_cols  = [c for c in ["date", "year", "現金股利合計"] if c in df_div_show.columns or c == "現金股利合計"]
                extra_cols = [c for c in [
                    "CashEarningsDistribution", "CashStatutorySurplus",
                    "StockEarningsDistribution", "CashExDividendTradingDate",
                    "CashDividendPaymentDate"
                ] if c in df_div_show.columns]
                show_cols = [c for c in base_cols + extra_cols if c in df_div_show.columns or c == "現金股利合計"]
                st.dataframe(
                    df_div_show.assign(**{"現金股利合計": df_div_show["現金股利合計"]})[show_cols]
                    .sort_values("date", ascending=False),
                    use_container_width=True, hide_index=True
                )
            else:
                st.info("無股利資料")

        with st.expander("📋 月營收明細（近2年）"):
            df_rev_show = d.get("df_rev", pd.DataFrame())
            if not df_rev_show.empty:
                st.dataframe(
                    df_rev_show.sort_values("date", ascending=False).head(24),
                    use_container_width=True, hide_index=True
                )
            else:
                st.info("無月營收資料")

    elif not keyword.strip():
        st.info("👆 請在上方輸入股票代號或名稱關鍵字搜尋")


# ═══════════════════════════════════════════════════════
# Tab3：API 診斷
# ═══════════════════════════════════════════════════════
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
            preview_cols = ["date"] + [c for c in [
                "year", "_自動偵測現金股利",
                "CashEarningsDistribution", "CashStatutorySurplus",
                "cash_dividend", "stock_or_cash"
            ] if c in df2.columns or c == "_自動偵測現金股利"]
            st.write("**自動偵測結果（前5筆）：**")
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
            st.error("❌ 無資料")
            st.json(raw3)

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
            st.error("❌ 無資料")
            st.json(raw4)

        st.markdown("### 5️⃣ 完整指標計算結果")
        with st.spinner("計算中..."):
            d_diag = get_stock_data(diag_id, finmind_token)

        rows = [
            ("股價",                  fmt(d_diag.get("price"))),
            ("殖利率",                f"{d_diag['div_yield']*100:.2f}%"
                                      if not np.isnan(d_diag.get("div_yield", np.nan)) else "❌"),
            ("配息（元）",             fmt(d_diag.get("cash_div"))),
            ("10年股利次數",           d_diag.get("div_count_10y", 0)),
            ("3年平均股利",            fmt(d_diag.get("avg_div_3y"))),
            ("6年平均股利",            fmt(d_diag.get("avg_div_6y"))),
            ("10年平均股利",           fmt(d_diag.get("avg_div_10y"))),
            ("去年EPS",               fmt(d_diag.get("eps_prev"))),
            ("今年累積EPS",            fmt(d_diag.get("eps_ytd"))),
            ("3年平均EPS",             fmt(d_diag.get("avg_eps_3y"))),
            ("配息率(去年EPS基礎)",     fmt(d_diag.get("payout_ratio"), 3)),
            ("配息率(3年均基礎)",       fmt(d_diag.get("payout_3y"), 3)),
            ("累積營收年增率",          fmt(d_diag.get("rev_ytd_yoy"), 2, "%")),
            ("EPS年度比較",            fmt(d_diag.get("eps_yoy"), 4)),
            ("本益比",                 fmt(d_diag.get("pe_ratio"), 2)),
            ("股價淨值比",             fmt(d_diag.get("pbr"), 2)),
            ("推隔年殖利率(均)",        f"{d_diag['est_yield_avg']*100:.2f}%"
                                       if not np.isnan(d_diag.get("est_yield_avg", np.nan)) else "❌"),
            ("推隔年殖利率(前季)",      f"{d_diag['est_yield_latest']*100:.2f}%"
                                       if not np.isnan(d_diag.get("est_yield_latest", np.nan)) else "❌"),
            ("推隔年本益比(均)",        fmt(d_diag.get("est_pe_avg"), 2)),
            ("推隔年本益比(前季)",      fmt(d_diag.get("est_pe_latest"), 2)),
            ("昂貴價(4%)",             fmt(d_diag.get("price_expensive"), 1)),
            ("合理價(5.5%)",           fmt(d_diag.get("price_fair"), 1)),
            ("便宜價(7%)",             fmt(d_diag.get("price_cheap"), 1)),
        ]
        st.dataframe(
            pd.DataFrame(rows, columns=["指標", "計算值"]),
            use_container_width=True, hide_index=True
        )

        st.markdown("### 6️⃣ 掃描池成交排行（前10筆）")
        with st.spinner("抓取市場成交資料..."):
            pool = get_top500_scan_pool(finmind_token)
        if pool:
            st.success(f"✅ 掃描池共 {len(pool)} 支")
            st.write("前10支（依成交金額排序）：", pool[:10])
        else:
            st.error("❌ 無法取得掃描池，請確認 Token 或網路")


st.divider()
st.caption("⚠️ 本工具數據僅供參考，不構成投資建議。數據來源：FinMind / yfinance。")
