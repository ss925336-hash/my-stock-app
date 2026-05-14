"""
台股自動化篩選 Streamlit Web App
對應原始 Excel「總表 2」的 VBA RUN 按鈕邏輯

篩選欄位對應說明：
  殖利率             → yfinance: trailingAnnualDividendYield  (現金股利 / 股價)
  10年股利次數        → FinMind DividendResult: 近10年發放次數 >= 門檻
  累計營收年增率(%)   → FinMind TaiwanStockMonthRevenue: 累積YTD年增率 >= 門檻
  營收增率%(與前月比) → FinMind: 當月營收 vs 上月營收 >= 門檻
  EPS年度比較        → FinMind TaiwanStockFinancialStatements: 今年EPS / 去年EPS >= 門檻
  EPS推隔年殖利率    → 預估股利 / 股價 = 近3年平均EPS * 配息率 / 股價 >= 門檻
"""

import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import requests
from datetime import datetime, timedelta
import time

# ──────────────────────────────────────────────
# 頁面設定
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="台股自動化篩選",
    page_icon="📈",
    layout="wide",
)

st.title("📈 台股自動化篩選系統")
st.caption("對應 Excel「總表 2」的篩選邏輯，數據來源：FinMind / yfinance")

# ──────────────────────────────────────────────
# FinMind API 工具函式
# ──────────────────────────────────────────────
FINMIND_BASE = "https://api.finmindtrade.com/api/v4/data"

def finmind_get(dataset: str, stock_id: str, start_date: str, token: str = "") -> pd.DataFrame:
    """通用 FinMind 查詢，回傳 DataFrame；出錯回傳空 DataFrame。"""
    params = {
        "dataset": dataset,
        "data_id": stock_id,
        "start_date": start_date,
        "token": token,
    }
    try:
        r = requests.get(FINMIND_BASE, params=params, timeout=15)
        r.raise_for_status()
        data = r.json()
        if data.get("status") == 200 and data.get("data"):
            return pd.DataFrame(data["data"])
    except Exception:
        pass
    return pd.DataFrame()


def get_tw_stock_list(token: str = "") -> list[str]:
    """從 FinMind 取得上市上櫃股票代碼清單。"""
    try:
        r = requests.get(
            FINMIND_BASE,
            params={"dataset": "TaiwanStockInfo", "token": token},
            timeout=20,
        )
        data = r.json()
        if data.get("status") == 200:
            df = pd.DataFrame(data["data"])
            # 只保留4位數字代碼（上市/上櫃一般股）
            codes = df["stock_id"].dropna().astype(str)
            codes = codes[codes.str.match(r"^\d{4}$")].tolist()
            return sorted(set(codes))
    except Exception:
        pass
    return []


# ──────────────────────────────────────────────
# 單一股票指標計算
# ──────────────────────────────────────────────

def calc_dividend_yield(stock_id: str, token: str) -> float:
    """
    殖利率 = 現金股利 / 當前股價
    對應 Excel 欄位「殖利率」（總表2 H欄）
    """
    tw = f"{stock_id}.TW"
    try:
        ticker = yf.Ticker(tw)
        info = ticker.fast_info
        price = info.last_price
        # 嘗試從 FinMind 取得最新現金股利
        start = (datetime.now() - timedelta(days=400)).strftime("%Y-%m-%d")
        df = finmind_get("TaiwanStockDividend", stock_id, start, token)
        if not df.empty:
            cash = df[df["stock_or_cash"] == "Cash"]["cash_dividend"].astype(float)
            if not cash.empty and price and price > 0:
                return float(cash.iloc[-1]) / price
        # fallback: yfinance
        div_yield = getattr(ticker, "dividends", None)
        annual = ticker.info.get("trailingAnnualDividendYield", None)
        if annual:
            return float(annual)
    except Exception:
        pass
    return np.nan


def calc_dividend_count_10y(stock_id: str, token: str) -> int:
    """
    10年股利次數 = 近10年（含配股）發放次數
    對應 Excel 欄位「10年股利次數」（總表2 P欄）
    """
    start = (datetime.now() - timedelta(days=3660)).strftime("%Y-%m-%d")
    df = finmind_get("TaiwanStockDividend", stock_id, start, token)
    if df.empty:
        return 0
    # 只計算現金股利次數（與 Excel 邏輯相同）
    cash_div = df[df["stock_or_cash"] == "Cash"]
    return len(cash_div)


def calc_revenue_yoy(stock_id: str, token: str) -> tuple[float, float]:
    """
    累計營收年增率 & 當月營收與前月比
    對應 Excel 欄位「累計營收年增率(%)」「營收增率(%) (與前月比)」（總表2 Z、AA欄）

    回傳 (ytd_yoy_pct, mom_pct)
    """
    start = (datetime.now() - timedelta(days=400)).strftime("%Y-%m-%d")
    df = finmind_get("TaiwanStockMonthRevenue", stock_id, start, token)
    if df.empty:
        return np.nan, np.nan

    df = df.sort_values("date")
    df["revenue"] = df["revenue"].astype(float)

    # 累積YTD年增率：FinMind 欄位 revenue_year 與 revenue_year_growth_percent
    if "revenue_year_growth_percent" in df.columns:
        ytd_yoy = df.iloc[-1]["revenue_year_growth_percent"]
    else:
        ytd_yoy = np.nan

    # 與前月比 (MoM)
    if len(df) >= 2:
        curr = df.iloc[-1]["revenue"]
        prev = df.iloc[-2]["revenue"]
        mom = (curr - prev) / prev * 100 if prev != 0 else np.nan
    else:
        mom = np.nan

    return float(ytd_yoy) if ytd_yoy is not None else np.nan, mom


def calc_eps_yoy(stock_id: str, token: str) -> float:
    """
    EPS年度比較 = 今年累積EPS / 去年同期EPS
    對應 Excel 欄位「EPS年度比較」（總表2 AB欄）
    """
    start = (datetime.now() - timedelta(days=800)).strftime("%Y-%m-%d")
    df = finmind_get("TaiwanStockFinancialStatements", stock_id, start, token)
    if df.empty:
        return np.nan

    eps_df = df[df["type"] == "EPS"][["date", "value"]].copy()
    eps_df["value"] = pd.to_numeric(eps_df["value"], errors="coerce")
    eps_df = eps_df.dropna().sort_values("date")
    if len(eps_df) < 2:
        return np.nan

    latest = eps_df.iloc[-1]["value"]
    prev = eps_df.iloc[-2]["value"]
    if prev == 0:
        return np.nan
    return latest / prev


def calc_eps_est_yield(stock_id: str, token: str) -> float:
    """
    EPS推隔年殖利率 = (近3年平均EPS * 近3年平均配息率) / 當前股價
    對應 Excel「用預估EPS推隔年殖利率 (平均加總)」（總表2 AC欄）
    """
    start_div = (datetime.now() - timedelta(days=1200)).strftime("%Y-%m-%d")
    start_eps = (datetime.now() - timedelta(days=1200)).strftime("%Y-%m-%d")

    # 取股價
    tw = f"{stock_id}.TW"
    try:
        price = yf.Ticker(tw).fast_info.last_price
    except Exception:
        return np.nan
    if not price or price <= 0:
        return np.nan

    # 取近3年EPS
    eps_df = finmind_get("TaiwanStockFinancialStatements", stock_id, start_eps, "")
    if eps_df.empty:
        return np.nan
    eps_rows = eps_df[eps_df["type"] == "EPS"][["date", "value"]].copy()
    eps_rows["value"] = pd.to_numeric(eps_rows["value"], errors="coerce")
    eps_rows = eps_rows.dropna().sort_values("date").tail(3)
    if eps_rows.empty:
        return np.nan
    avg_eps = eps_rows["value"].mean()

    # 取近3年配息率
    div_df = finmind_get("TaiwanStockDividend", stock_id, start_div, "")
    if div_df.empty:
        payout = 0.6  # 預設60%
    else:
        cash = div_df[div_df["stock_or_cash"] == "Cash"].copy()
        cash["cash_dividend"] = pd.to_numeric(cash["cash_dividend"], errors="coerce")
        if "EPS" in cash.columns:
            cash["EPS"] = pd.to_numeric(cash["EPS"], errors="coerce")
            valid = cash[(cash["EPS"] > 0) & (cash["cash_dividend"] > 0)].tail(3)
            payout = (valid["cash_dividend"] / valid["EPS"]).mean() if not valid.empty else 0.6
        else:
            payout = 0.6

    est_div = avg_eps * payout
    return est_div / price


# ──────────────────────────────────────────────
# 批次篩選（RUN 按鈕邏輯）
# ──────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner=False)
def run_screening(
    min_yield: float,
    min_div_count: int,
    min_rev_ytd_yoy: float,
    min_rev_mom: float,
    min_eps_yoy: float,
    min_eps_est_yield: float,
    token: str,
    stock_ids: list,
) -> pd.DataFrame:
    """
    對每檔股票計算六項指標並套用門檻篩選。
    等同 Excel VBA RUN 按鈕對「總表2」第一列門檻值的 autofilter。
    -100 表示「不篩選此項」（與 Excel 說明一致）。
    """
    results = []
    progress = st.progress(0, text="開始掃描...")
    total = len(stock_ids)

    for i, sid in enumerate(stock_ids):
        progress.progress((i + 1) / total, text=f"分析中：{sid} ({i+1}/{total})")
        try:
            dy = calc_dividend_yield(sid, token)
            dc = calc_dividend_count_10y(sid, token)
            ytd, mom = calc_revenue_yoy(sid, token)
            eps_yoy = calc_eps_yoy(sid, token)
            eps_ey = calc_eps_est_yield(sid, token)

            # 套用門檻（-100 代表跳過，對應 Excel「針對不要的可以輸入-100」提示）
            ok = True
            if min_yield > -1 and (np.isnan(dy) or dy * 100 < min_yield):
                ok = False
            if min_div_count > -1 and dc < min_div_count:
                ok = False
            if min_rev_ytd_yoy > -100 and (np.isnan(ytd) or ytd < min_rev_ytd_yoy):
                ok = False
            if min_rev_mom > -100 and (np.isnan(mom) or mom < min_rev_mom):
                ok = False
            if min_eps_yoy > -1 and (np.isnan(eps_yoy) or eps_yoy < min_eps_yoy):
                ok = False
            if min_eps_est_yield > -1 and (np.isnan(eps_ey) or eps_ey * 100 < min_eps_est_yield):
                ok = False

            if ok:
                results.append({
                    "代號": sid,
                    "殖利率(%)": round(dy * 100, 2) if not np.isnan(dy) else None,
                    "10年股利次數": dc,
                    "累計營收年增率(%)": round(ytd, 2) if not np.isnan(ytd) else None,
                    "營收與前月比(%)": round(mom, 2) if not np.isnan(mom) else None,
                    "EPS年度比較": round(eps_yoy, 4) if not np.isnan(eps_yoy) else None,
                    "EPS推隔年殖利率(%)": round(eps_ey * 100, 2) if not np.isnan(eps_ey) else None,
                })
        except Exception:
            pass
        time.sleep(0.3)  # 避免 API 限流

    progress.empty()
    return pd.DataFrame(results) if results else pd.DataFrame(
        columns=["代號", "殖利率(%)", "10年股利次數", "累計營收年增率(%)",
                 "營收與前月比(%)", "EPS年度比較", "EPS推隔年殖利率(%)"]
    )


# ──────────────────────────────────────────────
# 個股查詢
# ──────────────────────────────────────────────

def show_stock_detail(stock_id: str, token: str):
    """功能二：輸入股票代號，顯示所有指標詳細數值。"""
    st.subheader(f"📋 個股詳細指標：{stock_id}")

    with st.spinner("抓取數據中..."):
        dy = calc_dividend_yield(stock_id, token)
        dc = calc_dividend_count_10y(stock_id, token)
        ytd, mom = calc_revenue_yoy(stock_id, token)
        eps_yoy = calc_eps_yoy(stock_id, token)
        eps_ey = calc_eps_est_yield(stock_id, token)

    # 六大指標卡片
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("殖利率 (%)",
                  f"{dy*100:.2f}%" if not np.isnan(dy) else "N/A",
                  help="現金股利 / 股價（對應 Excel 殖利率欄）")
        st.metric("10年股利次數",
                  dc,
                  help="近10年現金股利發放次數（對應 Excel 10年股利次數欄）")
    with col2:
        st.metric("累計營收年增率 (%)",
                  f"{ytd:.2f}%" if not np.isnan(ytd) else "N/A",
                  help="當年度累積營收 vs 去年同期（對應 Excel 累計營收年增率欄）")
        st.metric("營收與前月比 (%)",
                  f"{mom:.2f}%" if not np.isnan(mom) else "N/A",
                  help="最新月營收 vs 上月（對應 Excel 營收增率% 與前月比欄）")
    with col3:
        st.metric("EPS年度比較",
                  f"{eps_yoy:.4f}" if not np.isnan(eps_yoy) else "N/A",
                  help="最新季EPS / 去年同季EPS（對應 Excel EPS年度比較欄）")
        st.metric("EPS推隔年殖利率 (%)",
                  f"{eps_ey*100:.2f}%" if not np.isnan(eps_ey) else "N/A",
                  help="預估股利 / 當前股價（對應 Excel EPS推隔年殖利率欄）")

    # 歷史股利明細
    st.divider()
    st.markdown("#### 近期股利明細")
    start = (datetime.now() - timedelta(days=3660)).strftime("%Y-%m-%d")
    div_df = finmind_get("TaiwanStockDividend", stock_id, start, token)
    if not div_df.empty:
        show_cols = [c for c in ["date", "stock_or_cash", "cash_dividend", "stock_dividend"] if c in div_df.columns]
        st.dataframe(div_df[show_cols].sort_values("date", ascending=False).head(20), use_container_width=True)
    else:
        st.info("無股利資料（可能為新股或尚未公告）")

    # 近期月營收
    st.markdown("#### 近期月營收")
    start_rev = (datetime.now() - timedelta(days=400)).strftime("%Y-%m-%d")
    rev_df = finmind_get("TaiwanStockMonthRevenue", stock_id, start_rev, token)
    if not rev_df.empty:
        show_cols = [c for c in ["date", "revenue", "revenue_month_growth_percent",
                                  "revenue_year_growth_percent"] if c in rev_df.columns]
        st.dataframe(rev_df[show_cols].sort_values("date", ascending=False).head(24), use_container_width=True)
    else:
        st.info("無月營收資料")

    # yfinance 補充資訊
    st.markdown("#### 基本資訊（yfinance）")
    try:
        info = yf.Ticker(f"{stock_id}.TW").info
        basics = {k: info.get(k) for k in
                  ["longName", "sector", "industry", "marketCap", "trailingPE",
                   "priceToBook", "trailingEps", "currentPrice"]}
        st.json({k: v for k, v in basics.items() if v is not None})
    except Exception as e:
        st.warning(f"yfinance 資料不可用：{e}")


# ──────────────────────────────────────────────
# 側邊欄 UI（與截圖一致）
# ──────────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ 篩選條件設定")
    st.caption("對應 Excel「總表 2」第一列輸入區")

    st.divider()

    # ─ 殖利率 ─────────────────────────────────
    # Excel預設：0.05（5%）
    min_yield = st.number_input(
        "殖利率 ≥ (%)",
        min_value=-100.0, max_value=100.0, value=5.0, step=0.5,
        help="對應 Excel 殖利率欄門檻（預設5%）。輸入-100表示不篩選。",
    )

    # ─ 10年股利次數 ───────────────────────────
    # Excel預設：9
    min_div_count = st.number_input(
        "10年股利次數 ≥",
        min_value=-1, max_value=10, value=9, step=1,
        help="近10年現金股利發放次數門檻（預設9次）。輸入-1表示不篩選。",
    )

    # ─ 累計營收年增率 ─────────────────────────
    # Excel預設：-100（不篩選）
    min_rev_ytd_yoy = st.number_input(
        "累計營收年增率 ≥ (%)",
        min_value=-100.0, max_value=200.0, value=-100.0, step=5.0,
        help="當年累積營收 vs 去年同期。預設-100表示不篩選（與Excel一致）。",
    )

    # ─ 營收增率（與前月比）─────────────────────
    # Excel預設：-100（不篩選）
    min_rev_mom = st.number_input(
        "營收與前月比 ≥ (%)",
        min_value=-100.0, max_value=200.0, value=-100.0, step=5.0,
        help="最新月營收 vs 上月。預設-100表示不篩選（與Excel一致）。",
    )

    # ─ EPS年度比較 ────────────────────────────
    # Excel預設：0.7（70%，即今年/去年 >= 0.7）
    min_eps_yoy = st.number_input(
        "EPS年度比較 ≥",
        min_value=-1.0, max_value=5.0, value=0.7, step=0.05,
        help="今年EPS / 去年EPS >= 門檻（預設0.7，即不衰退超過30%）。輸入-1表示不篩選。",
    )

    # ─ EPS推隔年殖利率 ───────────────────────
    # Excel預設：0.06（6%）
    min_eps_est_yield = st.number_input(
        "EPS推隔年殖利率 ≥ (%)",
        min_value=-1.0, max_value=30.0, value=6.0, step=0.5,
        help="以平均EPS * 配息率 / 股價估算的殖利率（預設6%）。輸入-1表示不篩選。",
    )

    st.divider()
    st.markdown("**數據來源設定**")
    finmind_token = st.text_input(
        "FinMind API Token（選填）",
        type="password",
        help="免費帳號有限流限制。登入 finmindtrade.com 可取得 token。",
    )

    max_stocks = st.slider(
        "掃描股票數量上限",
        min_value=20, max_value=500, value=100, step=20,
        help="股票越多，等待時間越長。建議先從100開始測試。",
    )

    st.divider()
    run_btn = st.button("🚀 RUN 篩選", use_container_width=True, type="primary")

# ──────────────────────────────────────────────
# 主頁面 Tab
# ──────────────────────────────────────────────

tab1, tab2 = st.tabs(["📊 批次篩選（RUN）", "🔍 個股查詢"])

# ── Tab1：批次篩選 ─────────────────────────────
with tab1:
    st.markdown("""
    **使用說明：**
    1. 在左側設定篩選條件（預設值與 Excel 截圖一致）
    2. 點擊「RUN 篩選」開始掃描
    3. 輸入 `-100`（殖利率輸入 `-1`）代表跳過此項篩選（與 Excel 邏輯相同）
    """)

    st.markdown("| Excel 欄位 | 篩選邏輯 | 預設門檻 |")
    st.markdown("|---|---|---|")
    st.markdown("| 殖利率 | ≥ 門檻 | 5% |")
    st.markdown("| 10年股利次數 | ≥ 門檻 | 9次 |")
    st.markdown("| 累計營收年增率 | ≥ 門檻 | -100（不篩選）|")
    st.markdown("| 營收增率與前月比 | ≥ 門檻 | -100（不篩選）|")
    st.markdown("| EPS年度比較 | ≥ 門檻 | 0.7 |")
    st.markdown("| EPS推隔年殖利率 | ≥ 門檻 | 6% |")

    if run_btn:
        with st.spinner("取得股票清單..."):
            all_stocks = get_tw_stock_list(finmind_token)
            if not all_stocks:
                # fallback：使用常見台股代碼範圍
                all_stocks = [str(i) for i in range(1101, 4000)] + [str(i) for i in range(4100, 6800)]
            stock_sample = all_stocks[:max_stocks]

        st.info(f"共掃描 {len(stock_sample)} 支股票，請耐心等待...")

        result_df = run_screening(
            min_yield=min_yield,
            min_div_count=min_div_count,
            min_rev_ytd_yoy=min_rev_ytd_yoy,
            min_rev_mom=min_rev_mom,
            min_eps_yoy=min_eps_yoy,
            min_eps_est_yield=min_eps_est_yield,
            token=finmind_token,
            stock_ids=tuple(stock_sample),
        )

        if result_df.empty:
            st.warning("無股票符合目前條件，請嘗試放寬篩選門檻。")
        else:
            st.success(f"✅ 篩選完成！共 **{len(result_df)}** 支股票通過。")
            st.dataframe(result_df, use_container_width=True, hide_index=True)
            csv = result_df.to_csv(index=False, encoding="utf-8-sig")
            st.download_button("⬇️ 下載 CSV", csv, "screening_result.csv", "text/csv")

# ── Tab2：個股查詢 ─────────────────────────────
with tab2:
    st.markdown("輸入台股代號查詢該股所有篩選指標的詳細數值。")
    col_input, col_btn = st.columns([3, 1])
    with col_input:
        query_id = st.text_input("股票代號（例如 2330、0056）", placeholder="2330")
    with col_btn:
        query_btn = st.button("查詢", use_container_width=True)

    if query_btn and query_id.strip():
        show_stock_detail(query_id.strip(), finmind_token)
    elif query_btn:
        st.warning("請輸入股票代號。")

# ──────────────────────────────────────────────
# 頁尾說明
# ──────────────────────────────────────────────
st.divider()
st.caption(
    "⚠️ 本工具數據僅供參考，不構成投資建議。"
    "  數據來源：FinMind（月營收、股利）、yfinance（即時股價）。"
    "  FinMind 免費方案有 API 呼叫頻率限制，建議申請 token 以獲得更穩定的服務。"
)
