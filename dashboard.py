#!/usr/bin/env python3
"""
Rune  —  Local Streamlit App
================================================
Nigerian fixed income & rates intelligence.

Reads from the ngx.db SQLite file that ngx_refresh.py creates,
renders a live web dashboard at http://localhost:8501

Setup (one time):
    pip3 install streamlit

Run:
    cd ~/Desktop/NGX_Dashboard
    python3 -m streamlit run dashboard.py
"""

import sqlite3
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import streamlit as st

# ── CONFIG ────────────────────────────────────────────────────────────────────
DB_PATH = Path(__file__).resolve().parent / "ngx.db"

TENOR_YEARS = {
    "NTB_91D":  0.25, "NTB_182D": 0.50, "NTB_364D": 1.00,
    "FGN_2Y":   2,    "FGN_3Y":   3,    "FGN_5Y":   5,
    "FGN_7Y":   7,    "FGN_10Y":  10,   "FGN_15Y":  15,
    "FGN_20Y":  20,   "FGN_30Y":  30,
    "10Y":      10,   "USD10Y":   10,
}

TENOR_LABEL = {
    "NTB_91D":  "91D",  "NTB_182D": "182D", "NTB_364D": "364D",
    "FGN_2Y":   "2Y",   "FGN_3Y":   "3Y",   "FGN_5Y":   "5Y",
    "FGN_7Y":   "7Y",   "FGN_10Y":  "10Y",  "FGN_15Y":  "15Y",
    "FGN_20Y":  "20Y",  "FGN_30Y":  "30Y",
}

NTB_INSTRUMENTS = ["NTB_91D", "NTB_182D", "NTB_364D"]
FGN_INSTRUMENTS = ["FGN_2Y", "FGN_3Y", "FGN_5Y", "FGN_7Y",
                   "FGN_10Y", "FGN_15Y", "FGN_20Y", "FGN_30Y"]


# ── DATA HELPERS ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=60)
def query(sql: str, params: tuple = ()) -> pd.DataFrame:
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql_query(sql, conn, params=params)
    finally:
        conn.close()
    return df


def latest_snapshot_date():
    df = query("SELECT MAX(snapshot_date) AS d FROM rates")
    if df.empty or df.iloc[0]["d"] is None:
        return None
    return df.iloc[0]["d"]


def latest_rates() -> dict:
    d = latest_snapshot_date()
    if not d:
        return {}
    df = query("SELECT instrument, rate FROM rates WHERE snapshot_date = ?", (d,))
    return dict(zip(df.instrument, df.rate))


def prior_rates() -> dict:
    df = query(
        "SELECT DISTINCT snapshot_date FROM rates ORDER BY snapshot_date DESC LIMIT 2"
    )
    if len(df) < 2:
        return {}
    prior_date = df.iloc[1]["snapshot_date"]
    pdf = query("SELECT instrument, rate FROM rates WHERE snapshot_date = ?", (prior_date,))
    return dict(zip(pdf.instrument, pdf.rate))


def format_pct(x, decimals=2):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "—"
    return f"{x*100:.{decimals}f}%"


def render_centered_table(df: pd.DataFrame, height=None) -> None:
    """Render a DataFrame using Streamlit's native renderer with all cells centered.
    Uses pandas Styler — reliable across Streamlit versions, respects dark/light mode."""
    styled = (
        df.style
          .set_properties(**{"text-align": "center"})
          .set_table_styles([
              {"selector": "th", "props": [("text-align", "center"), ("font-weight", "600")]},
              {"selector": "td", "props": [("text-align", "center")]},
          ])
    )
    st.dataframe(styled, width="stretch", hide_index=True, height=height)


# ── PAGE SETUP ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Rune", page_icon="🇳🇬", layout="wide")

st.title("Rune")
st.markdown("_Nigerian fixed income & rates intelligence_")
snap = latest_snapshot_date()
if snap:
    st.caption(f"Latest snapshot: **{snap}** · sources: CBN · DMO · TradingEconomics")
else:
    st.error(
        f"No data found in `{DB_PATH}`. "
        "Run `python3 ngx_refresh.py` first to populate the database."
    )
    st.stop()

with st.sidebar:
    st.header("Database")
    st.write(f"**Path:** `{DB_PATH.name}`")
    counts = query("""
        SELECT 'rates' AS tbl, COUNT(*) AS n FROM rates
        UNION ALL SELECT 'auctions', COUNT(*) FROM auctions
        UNION ALL SELECT 'eurobonds', COUNT(*) FROM eurobonds
    """)
    for _, r in counts.iterrows():
        st.write(f"**{r['tbl']}:** {r['n']:,} rows")
    if st.button("🔄 Reload data", width="stretch"):
        st.cache_data.clear()
        st.rerun()
    st.divider()
    st.caption("Refresh the underlying data by running:")
    st.code("python3 ngx_refresh.py", language="bash")


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1  —  KPI ROW
# ══════════════════════════════════════════════════════════════════════════════
curr = latest_rates()
prev = prior_rates()


def delta_bps(curr_val, prev_val):
    if curr_val is None or prev_val is None:
        return None
    bps = (curr_val - prev_val) * 10000
    if abs(bps) < 0.5:
        return "no change"
    sign = "+" if bps > 0 else ""
    return f"{sign}{bps:.0f} bps"


real_91d = (curr.get("NTB_91D") - curr.get("CPI")
            if curr.get("NTB_91D") is not None and curr.get("CPI") is not None else None)

kpis = [
    ("CBN MPR",        curr.get("MPR"),     prev.get("MPR")),
    ("91-Day T-Bill",  curr.get("NTB_91D"), prev.get("NTB_91D")),
    ("10Y FGN",        curr.get("FGN_10Y") or curr.get("10Y"),
                       prev.get("FGN_10Y") or prev.get("10Y")),
    ("CPI Inflation",  curr.get("CPI"),     prev.get("CPI")),
    ("Real Yield 91D", real_91d,            None),
]

cols = st.columns(len(kpis))
for col, (label, val, prev_val) in zip(cols, kpis):
    delta = delta_bps(val, prev_val)
    col.metric(
        label=label,
        value=format_pct(val) if val is not None else "—",
        delta=delta,
        delta_color="inverse" if label in ("CPI Inflation",) else "normal",
    )

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2  —  COMBINED YIELD CURVE
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("📈 Sovereign Yield Curve — Overview")
st.caption("Today's stop rates plotted across the full curve, 91 days to 30 years")

curve_rows = []
for instr, rate in curr.items():
    if instr in TENOR_YEARS and instr != "USD10Y":
        if instr == "10Y" and "FGN_10Y" in curr:
            continue
        curve_rows.append({
            "Tenor (yrs)": TENOR_YEARS[instr],
            "Tenor": TENOR_LABEL.get(instr, instr),
            "Yield %": rate * 100,
        })

if curve_rows:
    curve_df = pd.DataFrame(curve_rows).sort_values("Tenor (yrs)").reset_index(drop=True)
    chart_df = curve_df.set_index("Tenor (yrs)")[["Yield %"]]
    st.line_chart(chart_df, height=320)
else:
    st.info("No tenor data available for the latest snapshot.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3  —  NTB YIELD CURVE
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("🟦 NTB Yield Curve — Treasury Bills")
st.caption("Stop rates on 91-day, 182-day, and 364-day Nigerian Treasury Bills")

ntb_rows = []
for instr in NTB_INSTRUMENTS:
    if instr in curr:
        ntb_rows.append({
            "Tenor": TENOR_LABEL[instr],
            "Tenor (yrs)": TENOR_YEARS[instr],
            "Yield %": curr[instr] * 100,
        })

if ntb_rows:
    ntb_df = pd.DataFrame(ntb_rows).sort_values("Tenor (yrs)").reset_index(drop=True)

    col_chart, col_table = st.columns([2, 1])
    with col_chart:
        st.bar_chart(ntb_df.set_index("Tenor")[["Yield %"]], height=280)
    with col_table:
        st.markdown("**Latest NTB stops**")
        display_ntb = ntb_df[["Tenor", "Yield %"]].copy()
        display_ntb["Yield %"] = display_ntb["Yield %"].map(lambda x: f"{x:.2f}%")
        render_centered_table(display_ntb)
else:
    st.info("No NTB data available for the latest snapshot.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4  —  FGN BOND YIELD CURVE
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("🟫 FGN Bond Yield Curve — 2-Year through 30-Year")
st.caption("Stop rates and traded yields across the FGN sovereign bond curve")

fgn_rows = []
for instr in FGN_INSTRUMENTS:
    if instr in curr:
        fgn_rows.append({
            "Tenor": TENOR_LABEL[instr],
            "Tenor (yrs)": TENOR_YEARS[instr],
            "Yield %": curr[instr] * 100,
        })

if fgn_rows:
    fgn_df = pd.DataFrame(fgn_rows).sort_values("Tenor (yrs)").reset_index(drop=True)

    col_chart, col_table = st.columns([2, 1])
    with col_chart:
        st.line_chart(fgn_df.set_index("Tenor")[["Yield %"]], height=280)
    with col_table:
        st.markdown("**Latest FGN bond yields**")
        display_fgn = fgn_df[["Tenor", "Yield %"]].copy()
        display_fgn["Yield %"] = display_fgn["Yield %"].map(lambda x: f"{x:.2f}%")
        render_centered_table(display_fgn)
else:
    st.info("No FGN bond data available for the latest snapshot.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5  —  REAL YIELD BY TENOR
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("💹 Real Yield by Tenor")
cpi = curr.get("CPI")
if cpi is not None:
    st.caption(
        f"Nominal yield minus CPI inflation ({cpi*100:.2f}%). "
        "Positive = beating inflation. Negative = losing purchasing power."
    )
    real_rows = []
    for instr in NTB_INSTRUMENTS + FGN_INSTRUMENTS:
        if instr in curr:
            real_rows.append({
                "Tenor": TENOR_LABEL[instr],
                "Tenor (yrs)": TENOR_YEARS[instr],
                "Real Yield %": (curr[instr] - cpi) * 100,
            })

    if real_rows:
        real_df = pd.DataFrame(real_rows).sort_values("Tenor (yrs)").reset_index(drop=True)
        st.bar_chart(real_df.set_index("Tenor")[["Real Yield %"]], height=280)
    else:
        st.info("No tenor data available to compute real yields.")
else:
    st.info("CPI inflation rate not available — cannot compute real yields.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6  —  RATE HISTORY
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("📊 Rate History")
st.caption("Pick an instrument to see how its rate has moved over time")

all_instruments = query(
    "SELECT DISTINCT instrument FROM rates ORDER BY instrument"
)["instrument"].tolist()

if not all_instruments:
    st.info("No instruments in database yet.")
else:
    default_idx = all_instruments.index("NTB_91D") if "NTB_91D" in all_instruments else 0
    pick = st.selectbox("Instrument", all_instruments, index=default_idx)

    hist = query(
        "SELECT snapshot_date, rate FROM rates WHERE instrument = ? ORDER BY snapshot_date",
        (pick,),
    )

    if len(hist) >= 2:
        hist["snapshot_date"] = pd.to_datetime(hist["snapshot_date"])
        hist["Rate %"] = hist["rate"] * 100
        st.line_chart(hist.set_index("snapshot_date")[["Rate %"]], height=300)
    elif len(hist) == 1:
        st.info(
            f"Only **one** data point so far for {pick}: "
            f"{hist.iloc[0]['rate']*100:.2f}% on {hist.iloc[0]['snapshot_date']}. "
            "Run ngx_refresh.py for a few more days to build a trend."
        )
    else:
        st.info(f"No history found for {pick}.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7  —  EUROBONDS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("🌍 FGN Eurobonds — Latest Snapshot")

eu_date_row = query("SELECT MAX(snapshot_date) AS d FROM eurobonds")
eu_date = eu_date_row.iloc[0]["d"] if not eu_date_row.empty else None

if eu_date:
    st.caption(f"Data as of: **{eu_date}** (DMO daily Excel)")
    eu = query(
        "SELECT bond_name, price, yield_pct FROM eurobonds WHERE snapshot_date = ? "
        "ORDER BY bond_name",
        (eu_date,),
    )
    ust10y = curr.get("USD10Y")
    if ust10y is not None:
        eu["Spread_bps"] = ((eu["yield_pct"] - ust10y) * 10000).round(0)

    if not eu.empty and ust10y is not None:
        st.markdown(f"**Eurobond yields vs UST 10Y baseline ({ust10y*100:.2f}%)**")
        chart_eu = eu[["bond_name", "yield_pct"]].copy()
        chart_eu["Yield %"] = chart_eu["yield_pct"] * 100
        chart_eu["Bond"] = chart_eu["bond_name"].str.replace(r" \(US\$\)", "", regex=True).str.slice(0, 22)
        st.bar_chart(chart_eu.set_index("Bond")[["Yield %"]], height=280)

    st.markdown("**Eurobond detail**")
    eu_display = pd.DataFrame({
        "Bond":          eu["bond_name"].str.replace(r" \(US\$\)", "", regex=True),
        "Price (US$)":   eu["price"].map(lambda x: f"{x:.3f}" if pd.notna(x) else "—"),
        "Yield":         eu["yield_pct"].map(format_pct),
        "vs UST 10Y":    eu["Spread_bps"].map(
                            lambda x: f"+{int(x)} bps" if pd.notna(x) else "—"
                         ) if ust10y is not None else "—",
    })
    render_centered_table(eu_display)
else:
    st.info("No Eurobond data in database yet.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8  —  RECENT AUCTIONS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("🏛 Recent Auctions — NTBs and FGN Bonds")

auct = query("""
    SELECT auction_date, security_type, tenor,
           offered, subscribed, allotted, stop_rate, maturity_date
    FROM auctions
    ORDER BY auction_date DESC, security_type, tenor
    LIMIT 30
""")

if not auct.empty:
    auct["Bid-Cover"] = (auct["subscribed"] / auct["offered"]).round(2)
    auct_display = pd.DataFrame({
        "Auction Date": auct["auction_date"],
        "Type":         auct["security_type"],
        "Tenor":        auct["tenor"],
        "Offered (₦B)":     auct["offered"].map(lambda x: f"{x/1000:,.1f}" if pd.notna(x) else "—"),
        "Subscribed (₦B)":  auct["subscribed"].map(lambda x: f"{x/1000:,.1f}" if pd.notna(x) else "—"),
        "Allotted (₦B)":    auct["allotted"].map(lambda x: f"{x/1000:,.1f}" if pd.notna(x) else "—"),
        "Bid-Cover":        auct["Bid-Cover"].map(lambda x: f"{x:.2f}x" if pd.notna(x) else "—"),
        "Stop Rate":        auct["stop_rate"].map(format_pct),
        "Maturity":         auct["maturity_date"],
    })
    render_centered_table(auct_display)
else:
    st.info("No auction records in database yet.")

st.divider()
st.caption(
    f"**Rune** · Decoding Nigerian markets · "
    f"Page generated {datetime.now().strftime('%d %b %Y %H:%M:%S')}"
)
