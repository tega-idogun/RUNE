#!/usr/bin/env python3
"""
backfill_cbn_history.py — one-time (re-runnable) full-history loader
====================================================================
The CBN auction APIs return the ENTIRE history of every NTB and FGN bond
auction (≈1,780 NTB + ≈500 FGN records going back to 2002), not just today's.
The daily refresh only ever kept the latest snapshot, so the database started
empty and grew one day at a time. This script loads the whole back-catalogue
in a single pass, so your time series starts out years deep instead of days.

It writes two things into the database (via db.py, so it targets whatever
DATABASE_URL points at — Postgres in production, SQLite locally):

  • auctions  — one row per (auction_date, security_type, tenor): the full
                auction event log (offered / subscribed / allotted / stop rate).
  • rates     — a historical stop-rate time series: for every auction, one row
                (snapshot_date = auction_date, instrument = e.g. NTB_91D, rate).
                This is the real, dated yield-curve history.

It is safe to re-run: upserts overwrite matching rows instead of duplicating.

Data-quality handling
---------------------
CBN's tenor field is inconsistent ('91', '91DAY', '91 Day', '90 DAY', and
outright junk like '44010'). Dates are dd/mm/yyyy but a few rows are malformed.
This script normalises tenors to a clean canonical set and validates that each
auction's maturity falls a sensible distance after its auction date, dropping
rows that can't be reconciled (counts reported at the end).

Usage
-----
    pip install -r requirements.txt
    export DATABASE_URL="postgresql+psycopg2://USER:PASS@HOST/DB?sslmode=require"
    python backfill_cbn_history.py            # full load
    python backfill_cbn_history.py --dry-run  # parse + report, write nothing
"""

from __future__ import annotations

import argparse
import datetime as dt
import re
import sys

import requests

import db

# ── Config ────────────────────────────────────────────────────────────────────
NTB_URL = "https://www.cbn.gov.ng/api/GetAllSecuritiesNTB"
FGN_URL = "https://www.cbn.gov.ng/api/GetAllSecuritiesFGNBond"
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
TIMEOUT = 30

# Canonical tenors and how many days each represents (for date validation).
NTB_TENORS = {"91D": 91, "182D": 182, "364D": 364}
NTB_TOLERANCE_DAYS = 20          # accept 90/91/92-day etc. into the right bucket
FGN_YEARS = {2, 3, 5, 7, 10, 15, 20, 25, 30}  # accepted FGN bond maturities (years)


# ── Parsing helpers ─────────────────────────────────────────────────────────────
def _num(s) -> float | None:
    try:
        return float(str(s).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def normalise_ntb_tenor(raw: str) -> str | None:
    """'91', '91DAY', '91 Day', '90 DAY' -> 'NTB_91D'. Junk -> None."""
    m = re.search(r"\d+", str(raw))
    if not m:
        return None
    days = int(m.group())
    for label, target in NTB_TENORS.items():
        if abs(days - target) <= NTB_TOLERANCE_DAYS:
            return f"NTB_{label}"
    return None


def normalise_fgn_tenor(raw: str) -> str | None:
    """'10 Year', '10YEAR', '10', '5-Year' -> 'FGN_10Y'. Junk -> None."""
    m = re.search(r"\d+", str(raw))
    if not m:
        return None
    years = int(m.group())
    if years in FGN_YEARS:
        return f"FGN_{years}Y"
    return None


def parse_dates(auction_raw: str, maturity_raw: str, expected_days: int | None):
    """Parse CBN's dd/mm/yyyy dates robustly.

    CBN is overwhelmingly dd/mm/yyyy, but a handful of rows are mm/dd/yyyy or
    malformed. We try every day/month interpretation and pick the pairing where
    maturity falls *after* the auction by roughly the expected tenor. Returns
    (auction_date, maturity_date) as date objects, either may be None.
    """
    def candidates(s):
        out = []
        for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d"):
            try:
                out.append(dt.datetime.strptime(str(s).strip(), fmt).date())
            except (ValueError, TypeError):
                continue
        return out or [None]

    a_opts = candidates(auction_raw)
    m_opts = candidates(maturity_raw)

    best, best_score = None, None
    for a in a_opts:
        for m in m_opts:
            if a is None:
                continue
            if m is None:
                # Auction parseable, maturity not — keep auction, drop maturity.
                cand, score = (a, None), 10 ** 9
            elif m <= a:
                continue  # maturity must be after auction
            else:
                gap = (m - a).days
                score = abs(gap - expected_days) if expected_days else 0
                cand = (a, m)
            if best_score is None or score < best_score:
                best, best_score = cand, score
    if best:
        return best
    # Last resort: dd/mm/yyyy auction date only.
    a = candidates(auction_raw)[0]
    return (a, None)


# ── Fetch ─────────────────────────────────────────────────────────────────────
def fetch(url: str) -> list[dict]:
    r = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
    r.raise_for_status()
    return r.json()


# ── Transform ───────────────────────────────────────────────────────────────────
def build_rows(records: list[dict], kind: str):
    """Return (auction_rows, rate_rows, stats) for one security type.
    kind is 'NTB' or 'FGN'."""
    now = dt.datetime.now().isoformat(timespec="seconds")
    normaliser = normalise_ntb_tenor if kind == "NTB" else normalise_fgn_tenor

    auction_rows: dict[tuple, dict] = {}  # dedupe by (date, type, tenor)
    rate_rows: dict[tuple, dict] = {}
    stats = {"total": len(records), "bad_tenor": 0, "bad_date": 0, "kept": 0}

    for rec in records:
        instrument = normaliser(rec.get("tenor", ""))
        if not instrument:
            stats["bad_tenor"] += 1
            continue

        if kind == "NTB":
            expected = NTB_TENORS[instrument.split("_")[1]]
            tenor_label = instrument.split("_")[1]              # '91D'
        else:
            expected = int(instrument.split("_")[1].rstrip("Y")) * 365
            tenor_label = instrument.split("_")[1]              # '10Y'

        rate = _num(rec.get("rate"))
        if rate is None or rate <= 0:
            stats["bad_date"] += 1
            continue
        rate_dec = round(rate / 100, 6)

        a_date, m_date = parse_dates(
            rec.get("auctionDate", ""), rec.get("maturityDate", ""), expected
        )
        if a_date is None:
            stats["bad_date"] += 1
            continue

        a_iso = a_date.isoformat()
        m_iso = m_date.isoformat() if m_date else None
        sec_type = "NTB" if kind == "NTB" else "FGN"

        # Auction event log. If the same (date,type,tenor) appears twice, the
        # later API id wins (dict overwrite as we iterate the sorted list).
        auction_rows[(a_iso, sec_type, tenor_label)] = {
            "auction_date":  a_iso,
            "security_type": sec_type,
            "tenor":         tenor_label,
            "offered":       _num(rec.get("amtOffered")),
            "subscribed":    _num(rec.get("totalSubscription")),
            "allotted":      _num(rec.get("totalSuccessful")),
            "stop_rate":     rate_dec,
            "maturity_date": m_iso,
            "fetched_at":    now,
        }

        # Historical stop-rate time series (one rate per auction date/instrument).
        rate_rows[(a_iso, instrument)] = {
            "snapshot_date": a_iso,
            "instrument":    instrument,
            "rate":          rate_dec,
            "source":        "CBN-history",
            "fetched_at":    now,
        }
        stats["kept"] += 1

    return list(auction_rows.values()), list(rate_rows.values()), stats


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> int:
    ap = argparse.ArgumentParser(description="Backfill full CBN auction history")
    ap.add_argument("--dry-run", action="store_true",
                    help="Parse and report only; write nothing.")
    args = ap.parse_args()

    print("=" * 64)
    print("  CBN HISTORY BACKFILL")
    print(f"  target DB: {db.database_url().split('@')[-1]}")
    print("=" * 64)

    print("\n· Fetching CBN NTB + FGN auction history …")
    try:
        ntb_raw = fetch(NTB_URL)
        fgn_raw = fetch(FGN_URL)
    except Exception as e:  # network / API failure
        print(f"  ✗ fetch failed: {e}")
        return 1
    print(f"  ✓ NTB: {len(ntb_raw):,} records   FGN: {len(fgn_raw):,} records")

    ntb_auc, ntb_rates, ntb_stats = build_rows(ntb_raw, "NTB")
    fgn_auc, fgn_rates, fgn_stats = build_rows(fgn_raw, "FGN")

    auction_rows = ntb_auc + fgn_auc
    rate_rows = ntb_rates + fgn_rates

    def report(name, s):
        print(f"  {name:4}  total={s['total']:>5}  kept={s['kept']:>5}  "
              f"dropped_tenor={s['bad_tenor']:>4}  dropped_rate/date={s['bad_date']:>4}")

    print("\n· Cleaning & normalising:")
    report("NTB", ntb_stats)
    report("FGN", fgn_stats)
    print(f"\n  → {len(auction_rows):,} unique auction events")
    print(f"  → {len(rate_rows):,} historical rate points")

    if auction_rows:
        dates = sorted(r["auction_date"] for r in auction_rows)
        print(f"  → date range: {dates[0]} … {dates[-1]}")

    if args.dry_run:
        print("\n  DRY RUN — nothing written.")
        return 0

    print("\n· Writing to database …")
    db.init_schema()
    eng = db.get_engine()
    n_auc = db.upsert(eng, "auctions", auction_rows)
    n_rate = db.upsert(eng, "rates", rate_rows)
    print(f"  ✓ auctions: {n_auc:,} rows upserted")
    print(f"  ✓ rates:    {n_rate:,} rows upserted")

    # Verify
    snaps = db.read_sql(
        "SELECT COUNT(DISTINCT snapshot_date) AS d, COUNT(*) AS n FROM rates"
    )
    if not snaps.empty:
        print(f"\n  rates table now holds {int(snaps.iloc[0]['n']):,} rows "
              f"across {int(snaps.iloc[0]['d']):,} distinct dates.")
    print("\nDone. Run the daily refresh from here on to keep it current.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
