"""
Mutual Fund Ranking & Analysis Dashboard — Single-file version
Designed for Streamlit Cloud deployment.
Place all 11 .xlsx files in the SAME folder as this app.py in your GitHub repo.
"""
import os
import sys
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

# =============================================================================
# PAGE CONFIG
# =============================================================================
st.set_page_config(
    page_title="Mutual Fund Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# DATA DIRECTORY — BULLETPROOF DETECTION
# =============================================================================
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

FILES = {
    "nav_flexi_multi": "flexiandmutlicap.xlsx",
    "nav_lm_mid_1": "largeand_midcsap_and_midcap_fund.xlsx",
    "nav_lm_mid_2": "large_mid_and_moidcp_fund_2.xlsx",
    "pe_flexi_multi": "flexiand_multipe.xlsx",
    "pe_lm_mid": "large_midand_midcap_pe_ratio.xlsx",
    "sector_flexi_multi": "flexi_and_multisecotr.xlsx",
    "sector_lm_mid": "large_miodsector_allocation.xlsx",
    "stock_flexi_multi": "flexiand_multistock.xlsx",
    "stock_lm_mid": "lege_mid_holding_allocations.xlsx",
    "asset_flexi_multi": "flexiand_multiassettpe.xlsx",
    "asset_lm_mid": "assettypelarge_midand_mid.xlsx",
}


def _find_data_dir():
    """Search multiple locations for the xlsx files."""
    candidates = [
        os.environ.get("FUND_DATA_DIR", ""),
        _SCRIPT_DIR,
        os.path.join(_SCRIPT_DIR, "data"),
        os.path.dirname(_SCRIPT_DIR),
        os.getcwd(),
        "/mount/src/analysis",
    ]
    for d in candidates:
        if not d or not os.path.isdir(d):
            continue
        xlsx = [f for f in os.listdir(d) if f.endswith(".xlsx")]
        if len(xlsx) >= 5:
            return d
    return _SCRIPT_DIR


DATA_DIR = _find_data_dir()


def _path(key):
    p = os.path.join(DATA_DIR, FILES[key])
    if not os.path.isfile(p):
        existing = []
        if os.path.isdir(DATA_DIR):
            existing = [f for f in sorted(os.listdir(DATA_DIR)) if f.endswith(".xlsx")]
        raise FileNotFoundError(
            f"Cannot find '{FILES[key]}' in '{DATA_DIR}'.\n"
            f"xlsx files actually found there: {existing}\n"
            f"Script directory: {_SCRIPT_DIR}\n"
            f"Working directory: {os.getcwd()}\n\n"
            f"FIX: Put all 11 xlsx files in the SAME folder as app.py in your GitHub repo."
        )
    return p


# =============================================================================
# CATEGORY CLASSIFICATION
# =============================================================================
def classify_fund(fund_name):
    if not isinstance(fund_name, str):
        return "Unknown"
    n = fund_name.lower()
    if "large" in n and "mid" in n:
        return "Large & Mid Cap"
    if "flexi" in n:
        return "Flexi Cap"
    if "multi" in n and "cap" in n:
        return "Multi Cap"
    if "mid" in n and ("cap" in n or "midcap" in n):
        return "Mid Cap"
    return "Unknown"


def clean_fund_name(name):
    if not isinstance(name, str):
        return name
    name = name.strip()
    if name.startswith("Scheme Name:"):
        name = name[len("Scheme Name:"):].strip()
    return name


# =============================================================================
# DATA LOADERS
# =============================================================================
def load_nav_file(path):
    df = pd.read_excel(path, sheet_name="sheet1", header=None)
    fund_names = df.iloc[2, 1:].tolist()
    data = df.iloc[4:, :].copy()
    data.columns = ["date"] + fund_names
    data["date"] = pd.to_datetime(data["date"], errors="coerce")
    data = data.dropna(subset=["date"])
    long = data.melt(id_vars="date", var_name="fund", value_name="nav")
    long["nav"] = pd.to_numeric(long["nav"], errors="coerce")
    long = long.dropna(subset=["nav"])
    long["fund"] = long["fund"].apply(clean_fund_name)
    return long


def load_all_nav():
    parts = [
        load_nav_file(_path("nav_flexi_multi")),
        load_nav_file(_path("nav_lm_mid_1")),
        load_nav_file(_path("nav_lm_mid_2")),
    ]
    nav = pd.concat(parts, ignore_index=True)
    nav = nav.drop_duplicates(subset=["date", "fund"]).sort_values(["fund", "date"])
    nav["category"] = nav["fund"].apply(classify_fund)
    return nav.reset_index(drop=True)


def load_ratios_file(path):
    df = pd.read_excel(path, sheet_name="sheet1", header=None)
    rows = []
    current_scheme = None
    for _, r in df.iterrows():
        col0 = r.iloc[0]
        if isinstance(col0, str) and col0.strip().startswith("Scheme Name:"):
            current_scheme = clean_fund_name(col0)
            continue
        d = pd.to_datetime(col0, errors="coerce")
        if pd.notna(d) and current_scheme is not None:
            rows.append({
                "fund": current_scheme, "month": d,
                "pe": pd.to_numeric(r.iloc[1], errors="coerce"),
                "pbv": pd.to_numeric(r.iloc[3], errors="coerce"),
                "div_yield": pd.to_numeric(r.iloc[5], errors="coerce"),
                "mcap_cr": pd.to_numeric(r.iloc[6], errors="coerce"),
            })
    return pd.DataFrame(rows)


def load_all_ratios():
    parts = [load_ratios_file(_path("pe_flexi_multi")),
             load_ratios_file(_path("pe_lm_mid"))]
    ratios = pd.concat(parts, ignore_index=True)
    ratios = ratios.drop_duplicates(subset=["fund", "month"]).sort_values(["fund", "month"])
    ratios["category"] = ratios["fund"].apply(classify_fund)
    return ratios.reset_index(drop=True)


def load_sector_file(path):
    df = pd.read_excel(path, sheet_name="sheet1", header=None, skiprows=4)
    df.columns = ["fund", "sector", "no_of_cos", "allocation"]
    df["fund"] = df["fund"].apply(clean_fund_name)
    df["allocation"] = pd.to_numeric(df["allocation"], errors="coerce")
    df["no_of_cos"] = pd.to_numeric(df["no_of_cos"], errors="coerce")
    df = df.dropna(subset=["fund", "sector", "allocation"])
    df = df[~df["fund"].str.contains("Accord Fintech", na=False)]
    return df.reset_index(drop=True)


def load_all_sector():
    parts = [load_sector_file(_path("sector_flexi_multi")),
             load_sector_file(_path("sector_lm_mid"))]
    sector = pd.concat(parts, ignore_index=True)
    sector["category"] = sector["fund"].apply(classify_fund)
    return sector.reset_index(drop=True)


def load_stock_file(path):
    df = pd.read_excel(path, sheet_name="sheet1", header=None, skiprows=4)
    df.columns = ["fund", "company", "asset", "sector", "allocation"]
    df["fund"] = df["fund"].apply(clean_fund_name)
    df["allocation"] = pd.to_numeric(df["allocation"], errors="coerce")
    df = df.dropna(subset=["fund", "company", "allocation"])
    df = df[~df["fund"].str.contains("Accord Fintech", na=False)]
    return df.reset_index(drop=True)


def load_all_stocks():
    parts = [load_stock_file(_path("stock_flexi_multi")),
             load_stock_file(_path("stock_lm_mid"))]
    stocks = pd.concat(parts, ignore_index=True)
    stocks["category"] = stocks["fund"].apply(classify_fund)
    return stocks.reset_index(drop=True)


def load_asset_file(path):
    df = pd.read_excel(path, sheet_name="sheet1", header=None, skiprows=4)
    df.columns = ["fund", "instrument", "allocation"]
    df["fund"] = df["fund"].apply(clean_fund_name)
    df["allocation"] = pd.to_numeric(df["allocation"], errors="coerce")
    df = df.dropna(subset=["fund", "instrument", "allocation"])
    df = df[~df["fund"].str.contains("Accord Fintech", na=False)]
    return df.reset_index(drop=True)


def load_all_assets():
    parts = [load_asset_file(_path("asset_flexi_multi")),
             load_asset_file(_path("asset_lm_mid"))]
    assets = pd.concat(parts, ignore_index=True)
    assets["category"] = assets["fund"].apply(classify_fund)
    return assets.reset_index(drop=True)


def build_fund_master(nav, ratios, sector, stocks, assets):
    all_funds = set()
    for df in [nav, ratios, sector, stocks, assets]:
        all_funds |= set(df["fund"].dropna().unique())
    out = pd.DataFrame({"fund": sorted(all_funds)})
    out["category"] = out["fund"].apply(classify_fund)
    out["has_nav"] = out["fund"].isin(nav["fund"].unique())
    out["has_ratios"] = out["fund"].isin(ratios["fund"].unique())
    out["has_sector"] = out["fund"].isin(sector["fund"].unique())
    out["has_stocks"] = out["fund"].isin(stocks["fund"].unique())
    out["has_assets"] = out["fund"].isin(assets["fund"].unique())
    return out


def load_all():
    nav = load_all_nav()
    ratios = load_all_ratios()
    sector = load_all_sector()
    stocks = load_all_stocks()
    assets = load_all_assets()
    master = build_fund_master(nav, ratios, sector, stocks, assets)
    return {"nav": nav, "ratios": ratios, "sector": sector,
            "stocks": stocks, "assets": assets, "master": master}


# =============================================================================
# METRICS ENGINE
# =============================================================================
RISK_FREE_RATE = 0.065
TRADING_DAYS = 252

HIGHER_IS_BETTER = [
    "cagr_1y", "cagr_3y", "cagr_5y", "sharpe", "sortino",
    "return_3m", "return_6m", "return_12m", "rank_accel",
]
LOWER_IS_BETTER = ["volatility", "max_drawdown", "downside_dev"]


def _returns_from_nav(nav_series):
    return np.log(nav_series / nav_series.shift(1)).dropna()


def _annualized_return(nav_series):
    if len(nav_series) < 2:
        return np.nan
    first, last = nav_series.iloc[0], nav_series.iloc[-1]
    if first <= 0 or last <= 0:
        return np.nan
    days = (nav_series.index[-1] - nav_series.index[0]).days
    if days < 30:
        return np.nan
    return (last / first) ** (1 / (days / 365.25)) - 1


def _period_cagr(nav_df, years):
    if len(nav_df) < 2:
        return np.nan
    end_date = nav_df.index[-1]
    start_date = end_date - pd.Timedelta(days=int(years * 365.25))
    window = nav_df[nav_df.index >= start_date]
    if len(window) < 20:
        return np.nan
    return _annualized_return(window["nav"])


def _max_drawdown(nav_series):
    if len(nav_series) < 2:
        return np.nan
    return ((nav_series - nav_series.cummax()) / nav_series.cummax()).min()


def _volatility(returns):
    return returns.std() * np.sqrt(TRADING_DAYS) if len(returns) >= 20 else np.nan


def _downside_deviation(returns, target=0.0):
    below = returns[returns < target]
    return np.sqrt((below ** 2).mean()) * np.sqrt(TRADING_DAYS) if len(below) >= 20 else np.nan


def _sharpe(returns, annual_return):
    vol = _volatility(returns)
    if vol is None or np.isnan(vol) or vol == 0:
        return np.nan
    return (annual_return - RISK_FREE_RATE) / vol


def _sortino(returns, annual_return):
    dd = _downside_deviation(returns, target=RISK_FREE_RATE / TRADING_DAYS)
    if dd is None or np.isnan(dd) or dd == 0:
        return np.nan
    return (annual_return - RISK_FREE_RATE) / dd


def _period_return(nav_df, months):
    if len(nav_df) < 2:
        return np.nan
    end_date = nav_df.index[-1]
    start_date = end_date - pd.DateOffset(months=months)
    window = nav_df[nav_df.index >= start_date]
    if len(window) < 5:
        return np.nan
    return window["nav"].iloc[-1] / window["nav"].iloc[0] - 1


def _rank_acceleration(fund_name, monthly_ret):
    if fund_name not in monthly_ret.columns:
        return np.nan
    last_6 = monthly_ret.tail(6)
    if len(last_6) < 3:
        return np.nan
    ranks = last_6.rank(axis=1, ascending=False)
    if fund_name not in ranks.columns:
        return np.nan
    fund_ranks = ranks[fund_name].dropna()
    if len(fund_ranks) < 3:
        return np.nan
    x = np.arange(len(fund_ranks))
    slope, _ = np.polyfit(x, fund_ranks.values, 1)
    return -slope


def compute_metrics(nav_long, min_history_days=365):
    results = []
    funds = nav_long["fund"].unique()
    nav_wide = nav_long.pivot(index="date", columns="fund", values="nav").sort_index()
    monthly = nav_wide.resample("ME").last()
    monthly_ret = monthly.pct_change()

    for fund in funds:
        fund_nav = nav_long[nav_long["fund"] == fund].set_index("date")[["nav"]].sort_index()
        if len(fund_nav) < 20:
            continue
        span_days = (fund_nav.index[-1] - fund_nav.index[0]).days
        if span_days < min_history_days:
            continue
        returns = _returns_from_nav(fund_nav["nav"])
        cagr_1y = _period_cagr(fund_nav, 1.0)
        cagr_3y = _period_cagr(fund_nav, 3.0)
        cagr_5y = _period_cagr(fund_nav, 5.0)
        cagr_incep = _annualized_return(fund_nav["nav"])
        vol = _volatility(returns)
        mdd = _max_drawdown(fund_nav["nav"])
        dd = _downside_deviation(returns)
        ann_ret = cagr_3y if not np.isnan(cagr_3y) else (cagr_1y if not np.isnan(cagr_1y) else cagr_incep)
        sharpe = _sharpe(returns, ann_ret)
        sortino = _sortino(returns, ann_ret)
        ret_3m = _period_return(fund_nav, 3)
        ret_6m = _period_return(fund_nav, 6)
        ret_12m = _period_return(fund_nav, 12)
        rank_accel = _rank_acceleration(fund, monthly_ret)

        results.append({
            "fund": fund, "history_years": round(span_days / 365.25, 1),
            "cagr_1y": cagr_1y, "cagr_3y": cagr_3y, "cagr_5y": cagr_5y,
            "cagr_incep": cagr_incep, "volatility": vol, "max_drawdown": mdd,
            "downside_dev": dd, "sharpe": sharpe, "sortino": sortino,
            "return_3m": ret_3m, "return_6m": ret_6m, "return_12m": ret_12m,
            "rank_accel": rank_accel,
        })
    return pd.DataFrame(results)


def _zscore_col(series):
    std = series.std()
    if std == 0 or np.isnan(std):
        return pd.Series(0, index=series.index)
    return (series - series.mean()) / std


def compute_composite_score(metrics_df, category_col="category", weights=None):
    if weights is None:
        weights = {"returns": 0.6, "risk": 0.2, "momentum": 0.2}
    out = metrics_df.copy()
    for col in HIGHER_IS_BETTER + LOWER_IS_BETTER:
        if col not in out.columns:
            continue
        out[col] = out.groupby(category_col)[col].transform(lambda s: s.fillna(s.median()))
    for col in HIGHER_IS_BETTER + LOWER_IS_BETTER:
        if col not in out.columns:
            continue
        out[f"z_{col}"] = out.groupby(category_col)[col].transform(_zscore_col)
        if col in LOWER_IS_BETTER:
            out[f"z_{col}"] = -out[f"z_{col}"]

    return_cols = [c for c in ["z_cagr_1y", "z_cagr_3y", "z_cagr_5y", "z_sharpe", "z_sortino"] if c in out.columns]
    risk_cols = [c for c in ["z_volatility", "z_max_drawdown", "z_downside_dev"] if c in out.columns]
    mom_cols = [c for c in ["z_return_3m", "z_return_6m", "z_return_12m", "z_rank_accel"] if c in out.columns]

    out["score_returns"] = out[return_cols].mean(axis=1) if return_cols else 0
    out["score_risk"] = out[risk_cols].mean(axis=1) if risk_cols else 0
    out["score_momentum"] = out[mom_cols].mean(axis=1) if mom_cols else 0

    out["composite"] = (weights["returns"] * out["score_returns"]
                        + weights["risk"] * out["score_risk"]
                        + weights["momentum"] * out["score_momentum"])
    out["rank"] = out.groupby(category_col)["composite"].rank(ascending=False, method="min").astype(int)
    return out.sort_values([category_col, "rank"])


# =============================================================================
# CHART COMPONENTS
# =============================================================================
CATEGORY_COLORS = {
    "Flexi Cap": "#534AB7",
    "Multi Cap": "#1D9E75",
    "Large & Mid Cap": "#BA7517",
    "Mid Cap": "#D85A30",
}


def nav_chart(nav_df, funds, normalize=True, start_date=None):
    fig = go.Figure()
    for fund in funds:
        sub = nav_df[nav_df["fund"] == fund].sort_values("date")
        if start_date is not None:
            sub = sub[sub["date"] >= pd.to_datetime(start_date)]
        if len(sub) == 0:
            continue
        y = sub["nav"].values
        if normalize and len(y) > 0:
            y = y / y[0] * 100
        fig.add_trace(go.Scatter(x=sub["date"], y=y, name=fund, mode="lines", line=dict(width=1.8)))
    fig.update_layout(
        height=420, margin=dict(t=30, b=30, l=10, r=10),
        yaxis_title="NAV (rebased to 100)" if normalize else "NAV (Rs.)",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="left", x=0),
    )
    return fig


def drawdown_chart(nav_df, funds):
    fig = go.Figure()
    for fund in funds:
        sub = nav_df[nav_df["fund"] == fund].sort_values("date")
        if len(sub) == 0:
            continue
        navv = sub["nav"].values
        peak = np.maximum.accumulate(navv)
        dd = (navv - peak) / peak * 100
        fig.add_trace(go.Scatter(x=sub["date"], y=dd, name=fund, mode="lines",
                                 fill="tozeroy", opacity=0.7, line=dict(width=1)))
    fig.update_layout(
        height=360, margin=dict(t=30, b=30, l=10, r=10),
        yaxis_title="Drawdown (%)",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="left", x=0),
    )
    return fig


def rolling_returns_chart(nav_df, fund, window_years=3):
    sub = nav_df[nav_df["fund"] == fund].set_index("date").sort_index()
    if len(sub) < 252 * window_years + 10:
        return None
    window = 252 * window_years
    roll = (sub["nav"] / sub["nav"].shift(window)) ** (1 / window_years) - 1
    roll = roll.dropna() * 100
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=roll.index, y=roll.values, mode="lines",
                             name=f"{window_years}Y rolling CAGR", line=dict(color="#534AB7", width=1.8)))
    fig.add_hline(y=roll.mean(), line_dash="dash", line_color="#888780",
                  annotation_text=f"Mean: {roll.mean():.1f}%", annotation_position="top right")
    fig.update_layout(height=340, margin=dict(t=30, b=30, l=10, r=10),
                      yaxis_title=f"{window_years}Y rolling CAGR (%)",
                      title=f"Rolling {window_years}Y CAGR — {fund}")
    return fig


def sector_bar(sector_df, fund, top_n=15):
    sub = sector_df[sector_df["fund"] == fund].sort_values("allocation", ascending=True).tail(top_n)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=sub["allocation"], y=sub["sector"], orientation="h",
                         marker_color="#534AB7",
                         text=[f"{v:.1f}%" for v in sub["allocation"]], textposition="outside"))
    fig.update_layout(height=max(300, 30 * len(sub) + 50),
                      margin=dict(t=30, b=30, l=10, r=30), xaxis_title="Allocation (%)")
    return fig


def sector_heatmap(sector_df, funds, top_n=12):
    sub = sector_df[sector_df["fund"].isin(funds)]
    top_sectors = (sub.groupby("sector")["allocation"].mean()
                   .sort_values(ascending=False).head(top_n).index.tolist())
    sub = sub[sub["sector"].isin(top_sectors)]
    pivot = sub.pivot_table(index="sector", columns="fund", values="allocation", fill_value=0)
    pivot = pivot.reindex(top_sectors)
    fig = go.Figure(data=go.Heatmap(
        z=pivot.values, x=pivot.columns, y=pivot.index, colorscale="Purples",
        text=[[f"{v:.1f}" for v in row] for row in pivot.values],
        texttemplate="%{text}", hovertemplate="%{y}<br>%{x}<br>%{z:.2f}%<extra></extra>",
        colorbar=dict(title="Alloc %"),
    ))
    fig.update_layout(height=max(350, 30 * len(pivot) + 50),
                      margin=dict(t=30, b=120, l=10, r=10), xaxis=dict(tickangle=-45))
    return fig


def top_holdings_bar(stocks_df, fund, top_n=10):
    sub = stocks_df[stocks_df["fund"] == fund].sort_values("allocation", ascending=True).tail(top_n)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=sub["allocation"], y=sub["company"], orientation="h",
        marker_color="#1D9E75",
        text=[f"{v:.2f}%" for v in sub["allocation"]], textposition="outside",
        customdata=sub["sector"],
        hovertemplate="<b>%{y}</b><br>Sector: %{customdata}<br>%{x:.2f}%<extra></extra>",
    ))
    fig.update_layout(height=max(280, 30 * len(sub) + 50),
                      margin=dict(t=30, b=30, l=10, r=30), xaxis_title="Allocation (%)")
    return fig


def stock_overlap_matrix(stocks_df, funds, top_n_stocks=15):
    sub = stocks_df[stocks_df["fund"].isin(funds)]
    counts = sub.groupby("company")["fund"].nunique()
    avg_alloc = sub.groupby("company")["allocation"].mean()
    combined = pd.concat([counts, avg_alloc], axis=1)
    combined.columns = ["fund_count", "avg_alloc"]
    combined = combined.sort_values(["fund_count", "avg_alloc"], ascending=False).head(top_n_stocks)
    top_stocks = combined.index.tolist()
    pivot = sub[sub["company"].isin(top_stocks)].pivot_table(
        index="company", columns="fund", values="allocation", fill_value=0)
    pivot = pivot.reindex(top_stocks)
    return pivot, combined


def overlap_heatmap(pivot):
    fig = go.Figure(data=go.Heatmap(
        z=pivot.values, x=pivot.columns, y=pivot.index, colorscale="Teal",
        text=[[f"{v:.1f}" if v > 0 else "" for v in row] for row in pivot.values],
        texttemplate="%{text}", hovertemplate="%{y}<br>%{x}<br>%{z:.2f}%<extra></extra>",
        colorbar=dict(title="Alloc %"),
    ))
    fig.update_layout(height=max(350, 30 * len(pivot) + 80),
                      margin=dict(t=30, b=150, l=10, r=10), xaxis=dict(tickangle=-45))
    return fig


def valuation_trend_chart(ratios_df, fund, metric="pe"):
    sub = ratios_df[ratios_df["fund"] == fund].sort_values("month")
    if len(sub) == 0:
        return None
    cat = sub["category"].iloc[0]
    cat_series = (ratios_df[ratios_df["category"] == cat]
                  .groupby("month")[metric].median().reset_index().sort_values("month"))
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=sub["month"], y=sub[metric], mode="lines+markers",
                             name=fund, line=dict(color="#534AB7", width=2)))
    fig.add_trace(go.Scatter(x=cat_series["month"], y=cat_series[metric], mode="lines",
                             name=f"{cat} median", line=dict(color="#888780", width=1.5, dash="dash")))
    label_map = {"pe": "PE (x)", "pbv": "PBV (x)", "div_yield": "Div Yield (%)", "mcap_cr": "MCap (Cr)"}
    fig.update_layout(height=320, margin=dict(t=30, b=30, l=10, r=10),
                      yaxis_title=label_map.get(metric, metric),
                      legend=dict(orientation="h", y=-0.2))
    return fig


def compute_hhi(allocations):
    shares = allocations / allocations.sum() * 100
    return (shares ** 2).sum()


# =============================================================================
# CACHED DATA + METRICS
# =============================================================================
@st.cache_data(show_spinner="Loading fund data from Excel files...")
def get_data():
    return load_all()


@st.cache_data(show_spinner="Computing metrics for every fund...")
def get_metrics(_nav_df):
    return compute_metrics(_nav_df)


@st.cache_data
def get_scored(_metrics_df, _master_df, weights_tuple):
    w_ret, w_risk, w_mom = weights_tuple
    weights = {"returns": w_ret, "risk": w_risk, "momentum": w_mom}
    merged = _metrics_df.merge(_master_df[["fund", "category"]], on="fund", how="left")
    merged = merged[merged["category"] != "Unknown"].reset_index(drop=True)
    return compute_composite_score(merged, weights=weights)


# =============================================================================
# LOAD DATA
# =============================================================================
try:
    data = get_data()
except FileNotFoundError as e:
    st.error(f"📁 **Data files not found**\n\n```\n{e}\n```")
    st.markdown(
        "**How to fix this:**\n\n"
        "1. Put all 11 `.xlsx` data files in the **same folder** as `app.py` in your GitHub repo root.\n"
        "2. Do NOT put them inside a subfolder.\n"
        "3. Make sure the filenames match exactly (case-sensitive).\n\n"
        "**Expected files:** " + ", ".join(f"`{v}`" for v in FILES.values())
    )
    st.stop()

nav = data["nav"]
ratios = data["ratios"]
sector = data["sector"]
stocks = data["stocks"]
assets = data["assets"]
master = data["master"]
metrics_df = get_metrics(nav)


# =============================================================================
# SIDEBAR
# =============================================================================
st.sidebar.title("📊 Fund Analysis")
st.sidebar.caption(f"Data as of {nav['date'].max().strftime('%d %b %Y')}")

page = st.sidebar.radio(
    "Navigation",
    ["Overview", "Rankings", "Top 10 Deep Dive", "Fund Compare"],
    label_visibility="collapsed",
)

st.sidebar.markdown("---")
st.sidebar.subheader("⚖️ Composite weights")
st.sidebar.caption("Adjust to retune the ranking. Auto-normalised if they don't sum to 1.")

w_ret = st.sidebar.slider("Returns weight", 0.0, 1.0, 0.60, 0.05)
w_risk = st.sidebar.slider("Risk weight", 0.0, 1.0, 0.20, 0.05)
w_mom = st.sidebar.slider("Momentum weight", 0.0, 1.0, 0.20, 0.05)

total_w = w_ret + w_risk + w_mom
if abs(total_w - 1.0) > 0.001:
    st.sidebar.warning(f"Weights sum to {total_w:.2f} — normalising.")
    if total_w > 0:
        w_ret /= total_w
        w_risk /= total_w
        w_mom /= total_w

scored = get_scored(metrics_df, master, (w_ret, w_risk, w_mom))


# =============================================================================
# PAGE 1 — OVERVIEW
# =============================================================================
if page == "Overview":
    st.title("Mutual Fund Analysis Dashboard")
    st.markdown(
        "Ranks mutual funds across **Flexi Cap**, **Multi Cap**, "
        "**Large & Mid Cap**, and **Mid Cap** using a composite score of "
        "return, risk, and momentum metrics. Top 10 funds per category are "
        "deep-dived on sector allocation, stock holdings, valuation trends, and drawdowns."
    )

    cols = st.columns(4)
    cols[0].metric("Total funds", f"{len(master)}")
    cols[1].metric("With NAV history", f"{int(master['has_nav'].sum())}")
    cols[2].metric("Date range", f"{(nav['date'].max() - nav['date'].min()).days // 365} yrs")
    cols[3].metric("Categories", "4")

    st.markdown("### Funds per category")
    summary = master[master["category"] != "Unknown"].groupby("category").agg(
        total=("fund", "count"),
        with_nav=("has_nav", "sum"),
        with_ratios=("has_ratios", "sum"),
        with_sector=("has_sector", "sum"),
        with_stocks=("has_stocks", "sum"),
    ).reset_index()
    summary.columns = ["Category", "Total", "With NAV", "With ratios", "With sector", "With holdings"]
    st.dataframe(summary, use_container_width=True, hide_index=True)

    st.info(
        "💡 Only funds with NAV history can be ranked. "
        "All funds with holdings data are inspectable on the Deep Dive page."
    )

    st.markdown("### Methodology")
    st.markdown(
        "**Returns** — 1Y/3Y/5Y CAGR, Sharpe, Sortino  \n"
        "**Risk** — Volatility, max drawdown, downside deviation  \n"
        "**Momentum** — 3M/6M/12M returns + rank acceleration  \n\n"
        "Each metric is z-scored within category. Sub-scores are weighted (tunable via sidebar) "
        "and combined into a composite. Funds are ranked per category."
    )


# =============================================================================
# PAGE 2 — RANKINGS
# =============================================================================
elif page == "Rankings":
    st.title("Fund Rankings")

    cat_choice = st.selectbox("Category", ["All categories"] + sorted(scored["category"].unique().tolist()))
    view = scored.copy() if cat_choice == "All categories" else scored[scored["category"] == cat_choice].copy()
    view = view.sort_values(["category", "rank"])

    display = view[[
        "rank", "fund", "category", "history_years",
        "cagr_1y", "cagr_3y", "cagr_5y", "sharpe", "sortino",
        "volatility", "max_drawdown",
        "return_3m", "return_6m", "return_12m",
        "score_returns", "score_risk", "score_momentum", "composite",
    ]].copy()

    pct_cols = ["cagr_1y", "cagr_3y", "cagr_5y", "volatility", "max_drawdown",
                "return_3m", "return_6m", "return_12m"]
    for c in pct_cols:
        display[c] = (display[c] * 100).round(2)
    for c in ["sharpe", "sortino", "score_returns", "score_risk", "score_momentum", "composite"]:
        display[c] = display[c].round(3)

    display.columns = [
        "Rank", "Fund", "Category", "Years",
        "1Y %", "3Y CAGR %", "5Y CAGR %", "Sharpe", "Sortino",
        "Vol %", "Max DD %", "3M %", "6M %", "12M %",
        "Ret score", "Risk score", "Mom score", "Composite",
    ]

    st.markdown(f"**{len(display)} funds** ranked by composite score.")
    st.dataframe(display, use_container_width=True, hide_index=True, height=600)

    csv = display.to_csv(index=False)
    st.download_button("📥 Download CSV", csv,
                       file_name=f"rankings_{cat_choice.replace(' ', '_')}.csv", mime="text/csv")

    st.markdown("### Top 10 by composite score")
    for cat in sorted(view["category"].unique()):
        top10 = view[view["category"] == cat].head(10)
        if len(top10) == 0:
            continue
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=top10["composite"], y=top10["fund"], orientation="h",
            marker_color=CATEGORY_COLORS.get(cat, "#534AB7"),
            text=[f"{v:.2f}" for v in top10["composite"]], textposition="outside",
        ))
        fig.update_layout(title=f"{cat} ({len(top10)} funds)",
                          height=max(280, 30 * len(top10) + 80),
                          margin=dict(t=40, b=20, l=10, r=40),
                          xaxis_title="Composite score", yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig, use_container_width=True)


# =============================================================================
# PAGE 3 — TOP 10 DEEP DIVE
# =============================================================================
elif page == "Top 10 Deep Dive":
    st.title("Top 10 Deep Dive")

    cat = st.selectbox("Category", sorted(scored["category"].unique().tolist()))
    top10 = scored[scored["category"] == cat].head(10)
    top10_funds = top10["fund"].tolist()

    st.caption(f"Top {len(top10_funds)} funds in {cat} by composite score.")

    summary_view = top10[["rank", "fund", "cagr_3y", "sharpe", "max_drawdown", "return_12m", "composite"]].copy()
    summary_view["cagr_3y"] = (summary_view["cagr_3y"] * 100).round(2)
    summary_view["max_drawdown"] = (summary_view["max_drawdown"] * 100).round(2)
    summary_view["return_12m"] = (summary_view["return_12m"] * 100).round(2)
    summary_view["sharpe"] = summary_view["sharpe"].round(3)
    summary_view["composite"] = summary_view["composite"].round(3)
    summary_view.columns = ["Rank", "Fund", "3Y CAGR %", "Sharpe", "Max DD %", "12M %", "Composite"]
    st.dataframe(summary_view, use_container_width=True, hide_index=True)

    tabs = st.tabs(["📈 NAV & Drawdown", "🏭 Sector", "📦 Holdings", "🥧 Asset type", "📊 Valuation"])

    with tabs[0]:
        st.markdown("#### NAV growth (rebased to 100)")
        years_back = st.slider("Lookback (years)", 1, 10, 5, key="nav_years")
        start = nav["date"].max() - pd.Timedelta(days=int(years_back * 365.25))
        st.plotly_chart(nav_chart(nav, top10_funds, normalize=True, start_date=start), use_container_width=True)

        st.markdown("#### Drawdown profile")
        nav_window = nav[nav["date"] >= start]
        st.plotly_chart(drawdown_chart(nav_window, top10_funds), use_container_width=True)

    with tabs[1]:
        st.markdown("#### Sector allocation heatmap")
        available = [f for f in top10_funds if f in sector["fund"].unique()]
        if not available:
            st.warning("No sector data available for these funds.")
        else:
            st.plotly_chart(sector_heatmap(sector, available, top_n=15), use_container_width=True)

            st.markdown("#### Concentration (HHI)")
            hhi_rows = []
            for f in available:
                sub = sector[sector["fund"] == f]["allocation"]
                hhi_rows.append({"Fund": f, "Sectors": int(sub.count()),
                                 "HHI": round(compute_hhi(sub), 0),
                                 "Top sector %": round(sub.max(), 2)})
            st.dataframe(pd.DataFrame(hhi_rows).sort_values("HHI", ascending=False),
                         use_container_width=True, hide_index=True)

            pick = st.selectbox("Zoom into fund", available, key="sector_zoom")
            st.plotly_chart(sector_bar(sector, pick), use_container_width=True)

    with tabs[2]:
        st.markdown("#### Stock overlap across top funds")
        available = [f for f in top10_funds if f in stocks["fund"].unique()]
        if not available:
            st.warning("No holdings data available.")
        else:
            pivot, combined = stock_overlap_matrix(stocks, available, top_n_stocks=15)
            st.plotly_chart(overlap_heatmap(pivot), use_container_width=True)

            st.markdown("#### Most-held stocks")
            cd = combined.reset_index()
            cd.columns = ["Company", "# funds", "Avg alloc %"]
            cd["Avg alloc %"] = cd["Avg alloc %"].round(2)
            st.dataframe(cd, use_container_width=True, hide_index=True)

            pick = st.selectbox("Fund holdings", available, key="stock_zoom")
            st.plotly_chart(top_holdings_bar(stocks, pick), use_container_width=True)

    with tabs[3]:
        st.markdown("#### Equity vs cash split")
        available = [f for f in top10_funds if f in assets["fund"].unique()]
        if not available:
            st.warning("No asset-type data available.")
        else:
            asset_wide = assets[assets["fund"].isin(available)].pivot_table(
                index="fund", columns="instrument", values="allocation", fill_value=0).round(2)
            st.dataframe(asset_wide.reset_index(), use_container_width=True, hide_index=True)

            fig = go.Figure()
            for col in asset_wide.columns:
                fig.add_trace(go.Bar(name=col, x=asset_wide.index, y=asset_wide[col]))
            fig.update_layout(barmode="stack", height=420,
                              margin=dict(t=30, b=120, l=10, r=10),
                              xaxis=dict(tickangle=-45), yaxis_title="Allocation (%)",
                              legend=dict(orientation="h", y=-0.4))
            st.plotly_chart(fig, use_container_width=True)

    with tabs[4]:
        st.markdown("#### Valuation ratios over time")
        available = [f for f in top10_funds if f in ratios["fund"].unique()]
        if not available:
            st.warning("No valuation data available.")
        else:
            pick = st.selectbox("Fund", available, key="val_zoom")
            c1, c2 = st.columns(2)
            with c1:
                fig = valuation_trend_chart(ratios, pick, "pe")
                if fig:
                    fig.update_layout(title="PE ratio")
                    st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = valuation_trend_chart(ratios, pick, "pbv")
                if fig:
                    fig.update_layout(title="PBV ratio")
                    st.plotly_chart(fig, use_container_width=True)

            st.markdown("#### Latest snapshot — all top 10")
            latest = ratios[ratios["fund"].isin(available)].sort_values("month").groupby("fund").tail(1)
            snap = latest[["fund", "month", "pe", "pbv", "div_yield", "mcap_cr"]].copy()
            snap["month"] = snap["month"].dt.strftime("%b %Y")
            snap[["pe", "pbv", "div_yield"]] = snap[["pe", "pbv", "div_yield"]].round(2)
            snap["mcap_cr"] = snap["mcap_cr"].round(0).astype(int)
            snap.columns = ["Fund", "As of", "PE", "PBV", "Div Yield %", "MCap (Cr)"]
            st.dataframe(snap, use_container_width=True, hide_index=True)


# =============================================================================
# PAGE 4 — FUND COMPARE
# =============================================================================
elif page == "Fund Compare":
    st.title("Fund Compare")
    st.caption("Pick 2–4 funds to compare side by side.")

    ranked_funds = scored["fund"].tolist()
    picks = st.multiselect("Select funds", ranked_funds,
                           default=ranked_funds[:2] if len(ranked_funds) >= 2 else ranked_funds,
                           max_selections=4)

    if len(picks) < 2:
        st.info("Select at least 2 funds to compare.")
        st.stop()

    sub = scored[scored["fund"].isin(picks)].copy()

    st.markdown("### Metrics comparison")
    cmp_cols = ["fund", "category", "rank", "cagr_1y", "cagr_3y", "cagr_5y",
                "sharpe", "sortino", "volatility", "max_drawdown",
                "return_3m", "return_6m", "return_12m", "composite"]
    view = sub[cmp_cols].copy()
    for c in ["cagr_1y", "cagr_3y", "cagr_5y", "volatility", "max_drawdown",
              "return_3m", "return_6m", "return_12m"]:
        view[c] = (view[c] * 100).round(2)
    for c in ["sharpe", "sortino", "composite"]:
        view[c] = view[c].round(3)
    view.columns = ["Fund", "Category", "Rank", "1Y %", "3Y %", "5Y %",
                    "Sharpe", "Sortino", "Vol %", "Max DD %", "3M %", "6M %", "12M %", "Composite"]
    st.dataframe(view.T, use_container_width=True)

    st.markdown("### NAV growth (rebased to 100)")
    years_back = st.slider("Lookback (years)", 1, 15, 5, key="compare_years")
    start = nav["date"].max() - pd.Timedelta(days=int(years_back * 365.25))
    st.plotly_chart(nav_chart(nav, picks, normalize=True, start_date=start), use_container_width=True)

    st.markdown("### Drawdown comparison")
    nav_window = nav[nav["date"] >= start]
    st.plotly_chart(drawdown_chart(nav_window, picks), use_container_width=True)

    if all(f in sector["fund"].unique() for f in picks):
        st.markdown("### Sector allocation")
        st.plotly_chart(sector_heatmap(sector, picks, top_n=12), use_container_width=True)

    st.markdown("### Rolling 3-year CAGR")
    for f in picks:
        fig = rolling_returns_chart(nav, f, window_years=3)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.caption(f"_{f} — not enough history for 3Y rolling CAGR_")
