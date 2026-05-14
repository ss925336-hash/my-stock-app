"""
台股自動化篩選 Streamlit Web App
對應原始 Excel「總表 2」的 VBA RUN 按鈕邏輯

FinMind TaiwanStockDividend 實際欄位（寬表格，每列一次除息）：
  date                      = 除息日
  year                      = 年度（例如 113年第4季）
  CashEarningsDistribution  = 盈餘分配現金股利（元/股）← 主要現金股利
  CashStatutorySurplus      = 法定公積現金股利（元/股）
  現金股利合計 = CashEarningsDistribution + CashStatutorySurplus

Excel 欄位對應：
  殖利率             → H欄  = 現金股利合計 / 當前股價
  10年股利次數        → P欄  = 近10年有發現金股利的次數（列數）
  累計營收年增率(%)   → Z欄  = 今年累積營收 / 去年同期 - 1
  營收增率%(與前月比) → AA欄 = 當月營收 / 上月營收 - 1
  EPS年度比較        → AB欄 = 近4季EPS / 前4季EPS
  EPS推隔年殖利率    → AC欄 = (近3年平均EPS × 配息率) / 股價
"""

import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import requests
from datetime import datetime, timedelta
import time

# ─────────────────────────────────────────────────────
# Token：優先從 Streamlit Cloud Secrets 讀取
# ─────────────────────────────────────────────────────
try:
    _SECRET_TOKEN = st.secrets.get("FINMIND_TOKEN", "")
except Exception:
    _SECRET_TOKEN = ""

# ─────────────────────────────────────────────────────
# 頁面設定
# ─────────────────────────────────────────────────────
st.set_page_config(page_title="台股自動化篩選", page_icon="📈", layout="wide")
st.title("📈 台股自動化篩選系統")
st.caption("對應 Excel「總表 2」篩選邏輯 ｜ 數據來源：FinMind / yfinance")

FINMIND_BASE = "https://api.finmindtrade.com/api/v4/data"


# ─────────────────────────────────────────────────────
# 工具函式
# ─────────────────────────────────────────────────────

def finmind_get(dataset: str, stock_id: str, start_date: str, token: str = ""):
    """FinMind API 查詢，回傳 (DataFrame, raw_dict)。"""
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


def get_price(stock_id: str) -> float:
    """yfinance 取台股現價，格式 XXXX.TW"""
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


def get_cash_dividend(df: pd.DataFrame) -> pd.Series:
    """
    從 TaiwanStockDividend 寬表計算每列的現金股利合計。
    現金股利 = CashEarningsDistribution + CashStatutorySurplus
    """
    result = pd.Series(0.0, index=df.index)
    for col in ["CashEarningsDistribution", "CashStatutorySurplus"]:
        if col in df.columns:
            result += pd.to_numeric(df[col], errors="coerce").fillna(0)
    return result


# ─────────────────────────────────────────────────────
# 六大指標計算
# ─────────────────────────────────────────────────────

def calc_dividend_yield(stock_id: str, token: str) -> float:
    """
    殖利率 = 近一年現金股利合計 / 當前股價
    對應 Excel H欄「殖利率」
    現金股利 = CashEarningsDistribution + CashStatutorySurplus
    """
    price = get_price(stock_id)
    if np.isnan(price):
        return np.nan

    # 抓近500天（涵蓋近一年除息）
    start = (datetime.now() - timedelta(days=500)).strftime("%Y-%m-%d")
    df, _ = finmind_get("TaiwanStockDividend", stock_id, start, token)
    if df.empty:
        return np.nan

    df["cash_div"] = get_cash_dividend(df)
    # 只保留有發現金股利的列
    df = df[df["cash_div"] > 0].sort_values("date")
    if df.empty:
        return np.nan

    # 近一年現金股利加總（台積電等季配公司會有多筆）
    total_div = df["cash_div"].sum()
    return total_div / price


def calc_dividend_count_10y(stock_id: str, token: str) -> int:
    """
    10年股利次數 = 近10年有發現金股利的次數（列數）
    對應 Excel P欄「10年股利次數」
    """
    start = (datetime.now() - timedelta(days=3650)).strftime("%Y-%m-%d")
    df, _ = finmind_get("TaiwanStockDividend", stock_id, start, token)
    if df.empty:
        return 0
    df["cash_div"] = get_cash_dividend(df)
    return int((df["cash_div"] > 0).sum())


def calc_revenue_metrics(stock_id: str, token: str) -> tuple:
    """
    回傳 (累計營收年增率%, 當月vs前月%)
    對應 Excel Z欄「累計營收年增率」、AA欄「營收增率%與前月比」
    """
    start = (datetime.now() - timedelta(days=760)).strftime("%Y-%m-%d")
    df, _ = finmind_get("TaiwanStockMonthRevenue", stock_id, start, token)
    if df.empty:
        return np.nan, np.nan

    df["date"] = pd.to_datetime(df["date"])
    df["revenue"] = pd.to_numeric(df["revenue"], errors="coerce")
    df = df.dropna(subset=["revenue"]).sort_values("date")

    cur_year = datetime.now().year
    cur_df  = df[df["date"].dt.year == cur_year]
    prev_df = df[df["date"].dt.year == cur_year - 1]

    # 累積YTD年增率
    ytd = np.nan
    if not cur_df.empty and not prev_df.empty:
        months = set(cur_df["date"].dt.month)
        prev_same = prev_df[prev_df["date"].dt.month.isin(months)]
        if not prev_same.empty and prev_same["revenue"].sum() != 0:
            ytd = (cur_df["revenue"].sum() / prev_same["revenue"].sum() - 1) * 100

    # MoM 月增率
    mom = np.nan
    if len(df) >= 2:
        curr_rev = df.iloc[-1]["revenue"]
        prev_rev = df.iloc[-2]["revenue"]
        if prev_rev != 0:
            mom = (curr_rev - prev_rev) / prev_rev * 100

    return ytd, mom


def calc_eps_yoy(stock_id: str, token: str) -> float:
    """
    EPS年度比較 = 近4季EPS加總 / 前4季EPS加總
    對應 Excel AB欄「EPS年度比較」
    FinMind TaiwanStockFinancialStatements: type='EPS'
    """
    start = (datetime.now() - timedelta(days=900)).strftime("%Y-%m-%d")
    df, _ = finmind_get("TaiwanStockFinancialStatements", stock_id, start, token)
    if df.empty or "type" not in df.columns:
        return np.nan

    eps_df = df[df["type"] == "EPS"][["date", "value"]].copy()
    eps_df["value"] = pd.to_numeric(eps_df["value"], errors="coerce")
    eps_df = eps_df.dropna().sort_values("date").drop_duplicates("date")

    if len(eps_df) < 8:
        return np.nan

    recent4 = eps_df.iloc[-4:]["value"].sum()
    prev4   = eps_df.iloc[-8:-4]["value"].sum()
    return recent4 / prev4 if prev4 != 0 else np.nan


def calc_eps_est_yield(stock_id: str, token: str) -> float:
    """
    EPS推隔年殖利率 = (近3年平均年度EPS × 近3年平均配息率) / 股價
    對應 Excel AC欄「EPS推隔年殖利率」
    配息率 = 近3年平均現金股利 / 近3年平均EPS
    """
    price = get_price(stock_id)
    if np.isnan(price):
        return np.nan

    # 近3年年度EPS（各年4季加總）
    start_eps = (datetime.now() - timedelta(days=1200)).strftime("%Y-%m-%d")
    df_eps, _ = finmind_get("TaiwanStockFinancialStatements", stock_id, start_eps, token)
    if df_eps.empty or "type" not in df_eps.columns:
        return np.nan

    eps_df = df_eps[df_eps["type"] == "EPS"][["date", "value"]].copy()
    eps_df["value"] = pd.to_numeric(eps_df["value"], errors="coerce")
    eps_df["date"]  = pd.to_datetime(eps_df["date"])
    eps_df = eps_df.dropna().sort_values("date")
    eps_by_year = eps_df.groupby(eps_df["date"].dt.year)["value"].sum()
    if eps_by_year.empty:
        return np.nan
    avg_eps = eps_by_year.tail(3).mean()
    if np.isnan(avg_eps) or avg_eps <= 0:
        return np.nan

    # 近3年每年現金股利（CashEarningsDistribution + CashStatutorySurplus）
    start_div = (datetime.now() - timedelta(days=1200)).strftime("%Y-%m-%d")
    df_div, _ = finmind_get("TaiwanStockDividend", stock_id, start_div, token)
    if df_div.empty:
        payout = 0.6
    else:
        df_div["date"] = pd.to_datetime(df_div["date"])
        df_div["cash_div"] = get_cash_dividend(df_div)
        # 以除息年度分組，取近3年年度現金股利加總的平均
        div_by_year = df_div[df_div["cash_div"] > 0].groupby(
            df_div["date"].dt.year)["cash_div"].sum()
        if div_by_year.empty:
            payout = 0.6
        else:
            avg_div = div_by_year.tail(3).mean()
            payout  = min(avg_div / avg_eps, 1.5)  # 上限150%防異常值

    return (avg_eps * payout) / price


# ─────────────────────────────────────────────────────
# 批次篩選
# ─────────────────────────────────────────────────────

def get_tw_stock_list(token: str) -> list:
    """從 FinMind TaiwanStockInfo 取得 4 碼股票代碼清單。"""
    try:
        params = {"dataset": "TaiwanStockInfo"}
        if token:
            params["token"] = token
        r = requests.get(FINMIND_BASE, params=params, timeout=20)
        data = r.json()
        if data.get("status") == 200 and data.get("data"):
            df = pd.DataFrame(data["data"])
            codes = df["stock_id"].dropna().astype(str)
            return sorted(set(codes[codes.str.match(r"^\d{4}$")].tolist()))
    except Exception:
        pass
    return []


@st.cache_data(ttl=3600, show_spinner=False)
def run_screening(
    min_yield, min_div_count, min_rev_ytd_yoy, min_rev_mom,
    min_eps_yoy, min_eps_est_yield, token, stock_ids
) -> pd.DataFrame:
    """
    批次掃描並篩選。
    -100（次數欄用 -1）= 跳過此項，對應 Excel「輸入-100表示不篩選」。
    """
    results = []
    progress = st.progress(0, text="開始掃描...")
    total = len(stock_ids)

    for i, sid in enumerate(stock_ids):
        progress.progress((i + 1) / total, text=f"分析中：{sid}（{i+1}/{total}）")
        try:
            dy       = calc_dividend_yield(sid, token)
            dc       = calc_dividend_count_10y(sid, token)
            ytd, mom = calc_revenue_metrics(sid, token)
            ey       = calc_eps_yoy(sid, token)
            eey      = calc_eps_est_yield(sid, token)

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
    cols = ["代號", "殖利率(%)", "10年股利次數", "累計營收年增率(%)",
            "營收與前月比(%)", "EPS年度比較", "EPS推隔年殖利率(%)"]
    return pd.DataFrame(results) if results else pd.DataFrame(columns=cols)


# ─────────────────────────────────────────────────────
# 側邊欄
# ─────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 篩選條件")
    st.caption("對應 Excel「總表 2」第一列輸入區")
    st.divider()

    min_yield         = st.number_input("殖利率 ≥ (%)",         min_value=-1.0,   max_value=100.0, value=5.0,    step=0.5,  help="-1 = 不篩選")
    min_div_count     = st.number_input("10年股利次數 ≥",        min_value=-1,     max_value=40,    value=9,      step=1,    help="-1 = 不篩選（台積電等季配股每年4次，10年共40次）")
    min_rev_ytd_yoy   = st.number_input("累計營收年增率 ≥ (%)",  min_value=-100.0, max_value=200.0, value=-100.0, step=5.0,  help="-100 = 不篩選")
    min_rev_mom       = st.number_input("營收與前月比 ≥ (%)",    min_value=-100.0, max_value=200.0, value=-100.0, step=5.0,  help="-100 = 不篩選")
    min_eps_yoy       = st.number_input("EPS年度比較 ≥",         min_value=-1.0,   max_value=5.0,   value=0.7,    step=0.05, help="-1 = 不篩選")
    min_eps_est_yield = st.number_input("EPS推隔年殖利率 ≥ (%)", min_value=-1.0,   max_value=30.0,  value=6.0,    step=0.5,  help="-1 = 不篩選")

    st.divider()
    if _SECRET_TOKEN:
        finmind_token = _SECRET_TOKEN
        st.success("✅ Token 已從 Secrets 載入", icon="🔑")
    else:
        finmind_token = st.text_input(
            "FinMind Token（選填）", type="password",
            help="Streamlit Cloud → Settings → Secrets → FINMIND_TOKEN"
        )

    max_stocks = st.slider("掃描上限（支）", 20, 500, 100, 20)
    st.divider()
    run_btn = st.button("🚀 RUN 篩選", use_container_width=True, type="primary")


# ─────────────────────────────────────────────────────
# 主頁面
# ─────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📊 批次篩選（RUN）", "🔍 個股查詢", "🛠 API 診斷"])


# ── Tab1：批次篩選 ────────────────────────────────────
with tab1:
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
            all_stocks = get_tw_stock_list(finmind_token)
            if not all_stocks:
                all_stocks = [str(i) for i in range(1101, 3700)] + [str(i) for i in range(4100, 6900)]
            stock_sample = all_stocks[:max_stocks]

        st.info(f"掃描 {len(stock_sample)} 支，約需 {len(stock_sample) // 4} 秒，請稍候...")
        result_df = run_screening(
            min_yield, min_div_count, min_rev_ytd_yoy, min_rev_mom,
            min_eps_yoy, min_eps_est_yield, finmind_token, tuple(stock_sample)
        )

        if result_df.empty:
            st.warning("無符合條件的股票，請嘗試放寬門檻。")
        else:
            st.success(f"✅ 共 **{len(result_df)}** 支通過篩選")
            st.dataframe(result_df, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ 下載 CSV",
                result_df.to_csv(index=False, encoding="utf-8-sig"),
                "result.csv", "text/csv"
            )


# ── Tab2：個股查詢 ────────────────────────────────────
with tab2:
    c1, c2 = st.columns([3, 1])
    with c1:
        query_id = st.text_input("股票代號", placeholder="例如：2330、0056")
    with c2:
        query_btn = st.button("查詢", use_container_width=True)

    if query_btn and query_id.strip():
        sid = query_id.strip()
        st.subheader(f"📋 {sid} 各項指標")

        with st.spinner("抓取數據..."):
            dy       = calc_dividend_yield(sid, finmind_token)
            dc       = calc_dividend_count_10y(sid, finmind_token)
            ytd, mom = calc_revenue_metrics(sid, finmind_token)
            ey       = calc_eps_yoy(sid, finmind_token)
            eey      = calc_eps_est_yield(sid, finmind_token)
            price    = get_price(sid)

        m1, m2, m3 = st.columns(3)
        m1.metric("股價",            f"{price:.2f}"     if not np.isnan(price) else "N/A")
        m1.metric("殖利率",           f"{dy*100:.2f}%"   if not np.isnan(dy)    else "N/A")
        m2.metric("10年股利次數",     dc)
        m2.metric("累計營收年增率",    f"{ytd:.2f}%"      if not np.isnan(ytd)   else "N/A")
        m3.metric("營收與前月比",      f"{mom:.2f}%"      if not np.isnan(mom)   else "N/A")
        m3.metric("EPS年度比較",       f"{ey:.4f}"        if not np.isnan(ey)    else "N/A")
        st.metric("EPS推隔年殖利率",   f"{eey*100:.2f}%"  if not np.isnan(eey)   else "N/A")

        st.divider()
        st.markdown("#### 股利明細（近10年）")
        ddf, _ = finmind_get(
            "TaiwanStockDividend", sid,
            (datetime.now() - timedelta(days=3650)).strftime("%Y-%m-%d"),
            finmind_token
        )
        if not ddf.empty:
            ddf["現金股利合計"] = get_cash_dividend(ddf)
            show_cols = [c for c in ["date", "year", "現金股利合計",
                                     "CashEarningsDistribution", "CashStatutorySurplus",
                                     "StockEarningsDistribution", "CashExDividendTradingDate",
                                     "CashDividendPaymentDate"] if c in ddf.columns or c == "現金股利合計"]
            st.dataframe(ddf[show_cols].sort_values("date", ascending=False).head(20),
                         use_container_width=True)
        else:
            st.info("無股利資料")

        st.markdown("#### 月營收（近2年）")
        rdf, _ = finmind_get(
            "TaiwanStockMonthRevenue", sid,
            (datetime.now() - timedelta(days=760)).strftime("%Y-%m-%d"),
            finmind_token
        )
        if not rdf.empty:
            st.dataframe(rdf.sort_values("date", ascending=False).head(24), use_container_width=True)
        else:
            st.info("無月營收資料")

    elif query_btn:
        st.warning("請輸入股票代號")


# ── Tab3：API 診斷 ────────────────────────────────────
with tab3:
    st.markdown("**出現 N/A 請先在此診斷**，可看到 API 實際回傳的欄位與資料。")
    diag_id  = st.text_input("診斷股票代號", value="2330", key="diag")
    diag_btn = st.button("執行診斷", key="diag_btn")

    if diag_btn:
        start_diag = (datetime.now() - timedelta(days=400)).strftime("%Y-%m-%d")

        # 1. yfinance
        st.markdown("### 1️⃣ yfinance 股價")
        try:
            p = yf.Ticker(f"{diag_id}.TW").fast_info.last_price
            st.success(f"✅ 股價：{p}")
        except Exception as e:
            st.error(f"❌ {e}")

        # 2. 股利
        st.markdown("### 2️⃣ TaiwanStockDividend（股利）")
        df2, raw2 = finmind_get("TaiwanStockDividend", diag_id, start_diag, finmind_token)
        st.write(f"**status:** `{raw2.get('status')}` ｜ **msg:** `{raw2.get('msg', '')}`")
        if not df2.empty:
            st.success(f"✅ {len(df2)} 筆 ｜ 欄位：`{list(df2.columns)}`")
            df2["現金股利合計"] = get_cash_dividend(df2)
            st.write("**CashEarningsDistribution 範例值：**",
                     pd.to_numeric(df2.get("CashEarningsDistribution", pd.Series()), errors="coerce").tolist()[:5])
            st.dataframe(df2[["date", "year", "現金股利合計",
                               "CashEarningsDistribution", "CashStatutorySurplus"]
                              ].head(10) if all(c in df2.columns for c in
                              ["CashEarningsDistribution", "CashStatutorySurplus"])
                              else df2.head(10),
                         use_container_width=True)
        else:
            st.error("❌ 無資料")
            st.json(raw2)

        # 3. 月營收
        st.markdown("### 3️⃣ TaiwanStockMonthRevenue（月營收）")
        df3, raw3 = finmind_get("TaiwanStockMonthRevenue", diag_id, start_diag, finmind_token)
        st.write(f"**status:** `{raw3.get('status')}` ｜ **msg:** `{raw3.get('msg', '')}`")
        if not df3.empty:
            st.success(f"✅ {len(df3)} 筆 ｜ 欄位：`{list(df3.columns)}`")
            st.dataframe(df3.head(5), use_container_width=True)
        else:
            st.error("❌ 無資料")
            st.json(raw3)

        # 4. EPS
        st.markdown("### 4️⃣ TaiwanStockFinancialStatements（EPS）")
        start_eps = (datetime.now() - timedelta(days=900)).strftime("%Y-%m-%d")
        df4, raw4 = finmind_get("TaiwanStockFinancialStatements", diag_id, start_eps, finmind_token)
        st.write(f"**status:** `{raw4.get('status')}` ｜ **msg:** `{raw4.get('msg', '')}`")
        if not df4.empty:
            st.success(f"✅ {len(df4)} 筆 ｜ 欄位：`{list(df4.columns)}`")
            if "type" in df4.columns:
                st.write("**type 唯一值：**", df4["type"].unique().tolist())
                st.dataframe(df4[df4["type"] == "EPS"].head(8), use_container_width=True)
            else:
                st.dataframe(df4.head(8), use_container_width=True)
        else:
            st.error("❌ 無資料")
            st.json(raw4)

        # 5. 計算結果
        st.markdown("### 5️⃣ 六大指標計算結果")
        with st.spinner("計算中..."):
            r_dy         = calc_dividend_yield(diag_id, finmind_token)
            r_dc         = calc_dividend_count_10y(diag_id, finmind_token)
            r_ytd, r_mom = calc_revenue_metrics(diag_id, finmind_token)
            r_ey         = calc_eps_yoy(diag_id, finmind_token)
            r_eey        = calc_eps_est_yield(diag_id, finmind_token)

        st.dataframe(pd.DataFrame([
            {"指標": "殖利率(%)",           "計算值": f"{r_dy*100:.2f}%"   if not np.isnan(r_dy)  else "❌ N/A", "Excel欄": "H"},
            {"指標": "10年股利次數",        "計算值": r_dc,                                                       "Excel欄": "P"},
            {"指標": "累計營收年增率(%)",   "計算值": f"{r_ytd:.2f}%"     if not np.isnan(r_ytd) else "❌ N/A", "Excel欄": "Z"},
            {"指標": "營收與前月比(%)",     "計算值": f"{r_mom:.2f}%"     if not np.isnan(r_mom) else "❌ N/A", "Excel欄": "AA"},
            {"指標": "EPS年度比較",         "計算值": f"{r_ey:.4f}"       if not np.isnan(r_ey)  else "❌ N/A", "Excel欄": "AB"},
            {"指標": "EPS推隔年殖利率(%)",  "計算值": f"{r_eey*100:.2f}%" if not np.isnan(r_eey) else "❌ N/A", "Excel欄": "AC"},
        ]), use_container_width=True, hide_index=True)
        st.caption("❌ N/A = 抓不到資料，請對照上方各節原始回傳排查。")


# ─────────────────────────────────────────────────────
st.divider()
st.caption("⚠️ 本工具數據僅供參考，不構成投資建議。數據來源：FinMind / yfinance。")
