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

Stop:
    Press Ctrl+C in the Terminal window
"""

import sqlite3
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import streamlit as st

# ── CONFIG ────────────────────────────────────────────────────────────────────
# DB sits next to this script — works both locally on your Mac and on
# Streamlit Cloud (where the repo is mounted at a different path).
DB_PATH = Path(__file__).resolve().parent / "ngx.db"

# Map every instrument code in the `rates` table to a position on the curve (years).
TENOR_YEARS = {
    "NTB_91D":  0.25,
    "NTB_182D": 0.50,
    "NTB_364D": 1.00,
    "FGN_2Y":   2,
    "FGN_3Y":   3,
    "FGN_5Y":   5,
    "FGN_7Y":   7,
    "FGN_10Y":  10,
    "FGN_15Y":  15,
    "FGN_20Y":  20,
    "FGN_30Y":  30,
    "10Y":      10,
    "USD10Y":   10,
}

TENOR_LABEL = {
    "NTB_91D":  "91D",  "NTB_182D": "182D", "NTB_364D": "364D",
    "FGN_2Y":   "2Y",   "FGN_3Y":   "3Y",   "FGN_5Y":   "5Y",
    "FGN_7Y":   "7Y",   "FGN_10Y":  "10Y",  "FGN_15Y":  "15Y",
    "FGN_20Y":  "20Y",  "FGN_30Y":  "30Y",
}

# ── DATA HELPERS ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=60)
def query(sql: str, params: tuple = ()) -> pd.DataFrame:
    """Run a SQL query against ngx.db and return a DataFrame."""
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql_query(sql, conn, params=params)
    finally:
        conn.close()
    return df


def latest_snapshot_date() -> str | None:
    df = query("SELECT MAX(snapshot_date) AS d FROM rates")
    if df.empty or df.iloc[0]["d"] is None:
        return None
    return df.iloc[0]["d"]


def latest_rates() -> dict:
    """Return {instrument: rate} for the most recent snapshot."""
    d = latest_snapshot_date()
    if not d:
        return {}
    df = query("SELECT instrument, rate FROM rates WHERE snapshot_date = ?", (d,))
    return dict(zip(df.instrument, df.rate))


def prior_rates() -> dict:
    """Return {instrument: rate} for the second-most-recent snapshot."""
    df = query(
        "SELECT DISTINCT snapshot_date FROM rates ORDER BY snapshot_date DESC LIMIT 2"
    )
    if len(df) < 2:
        return {}
    prior_date = df.iloc[1]["snapshot_date"]
    pdf = query("SELECT instrument, rate FROM rates WHERE snapshot_date = ?", (prior_date,))
    return dict(zip(pdf.instrument, pdf.rate))


def format_pct(x, decimals=2):
    """Render 0.1595 → '15.95%'. Returns '—' for None/NaN."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "—"
    return f"{x*100:.{decimals}f}%"


# ── PAGE SETUP ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Rune",
    page_icon="🇳🇬",
    layout="wide",
)

# Header
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

# Sidebar — meta info + manual refresh
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
# SECTION 1  —  KPI ROW  (top of page)
# ══════════════════════════════════════════════════════════════════════════════
curr = latest_rates()
prev = prior_rates()

def delta_bps(curr_val, prev_val):
    """Return change in basis points as a string, or None if no prior."""
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
# SECTION 2  —  YIELD CURVE
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("📈 Sovereign Yield Curve")
st.caption("Today's stop rates plotted against tenor (years to maturity)")

curve_rows = []
for instr, rate in curr.items():
    if instr in TENOR_YEARS and instr != "USD10Y":
        if instr == "10Y" and "FGN_10Y" in curr:
            continue
        curve_rows.append({
            "Tenor (yrs)": TENOR_YEARS[instr],
            "Tenor": TENOR_LABEL.get(instr, instr),
            "Yield": rate,
            "Yield %": rate * 100,
        })

if curve_rows:
    curve_df = pd.DataFrame(curve_rows).sort_values("Tenor (yrs)").reset_index(drop=True)
    chart_df = curve_df.set_index("Tenor (yrs)")[["Yield %"]]
    st.line_chart(chart_df, height=320)

    display = curve_df[["Tenor", "Yield %"]].copy()
    display["Yield %"] = display["Yield %"].map(lambda x: f"{x:.2f}%")
    st.dataframe(display.set_index("Tenor").T, width="stretch")
else:
    st.info("No tenor data available for the latest snapshot.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3  —  RATE HISTORY
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
# SECTION 4  —  EUROBONDS
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
    if ust10y:
        eu["Spread vs UST 10Y (bps)"] = ((eu["yield_pct"] - ust10y) * 10000).round(0)
    eu_display = pd.DataFrame({
        "Bond":          eu["bond_name"],
        "Price (US$)":   eu["price"].map(lambda x: f"{x:.3f}" if pd.notna(x) else "—"),
        "Yield":         eu["yield_pct"].map(format_pct),
        "vs UST 10Y":    eu["Spread vs UST 10Y (bps)"].map(
                            lambda x: f"+{int(x)} bps" if pd.notna(x) else "—"
                         ) if ust10y else "—",
    })
    st.dataframe(eu_display, width="stretch", hide_index=True)
else:
    st.info("No Eurobond data in database yet.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5  —  RECENT NTB / FGN AUCTIONS
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
    st.dataframe(auct_display, width="stretch", hide_index=True)
else:
    st.info("No auction records in database yet.")

# Footer
st.divider()
st.caption(
    f"**Rune** · Decoding Nigerian markets · "
    f"Page generated {datetime.now().strftime('%d %b %Y %H:%M:%S')}"
)
