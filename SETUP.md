# Rune — Postgres migration & CBN history backfill

This replaces the ephemeral local SQLite database with a hosted Postgres one,
and seeds it with the **full CBN auction history** (≈2,200 points back to 2002)
instead of starting empty.

## What changed

| File | Change |
|------|--------|
| `db.py` | **New.** Shared SQLAlchemy layer. Targets Postgres when `DATABASE_URL` is set, falls back to local SQLite otherwise. Dialect-aware upsert (`INSERT … ON CONFLICT DO UPDATE`). |
| `backfill_cbn_history.py` | **New.** One-time (re-runnable) loader that pulls the entire NTB+FGN auction history, normalises messy tenors, fixes date parsing, and writes both the auction log and a dated stop-rate time series. |
| `ngx_refresh.py` | Persistence now goes through `db.py`. NTB/FGN rates are stamped with their *auction* date so daily runs extend the backfilled series. Excel-workbook logic is unchanged. |
| `dashboard.py` | Reads via `db.py` (named parameters, works on both engines). New `latest_curve()` builds the current curve correctly across sources. Added FGN 25Y. |
| `requirements.txt` | Added SQLAlchemy, psycopg2-binary, and the ingestion libs. |

## Why this was the critical fix

SQLite on Streamlit Cloud lives on an ephemeral disk that is wiped on every
redeploy/restart — that's why only two days of history ever survived. A hosted
Postgres persists across deploys, and is shared by the writer (refresh/backfill)
and the reader (dashboard).

---

## 1. Create a free Postgres database

Either works; both have free tiers and give you a connection string.

- **Neon** — https://neon.tech → create project → copy the connection string.
- **Supabase** — https://supabase.com → Project → Settings → Database → Connection string (URI).

You'll get something like:

```
postgresql://USER:PASSWORD@HOST/DBNAME?sslmode=require
```

Add the SQLAlchemy driver prefix so it reads:

```
postgresql+psycopg2://USER:PASSWORD@HOST/DBNAME?sslmode=require
```

## 2. Install & set the connection string

```bash
pip install -r requirements.txt
export DATABASE_URL="postgresql+psycopg2://USER:PASSWORD@HOST/DBNAME?sslmode=require"
```

(Windows PowerShell: `setx DATABASE_URL "postgresql+psycopg2://..."`)

## 3. Seed the full history (run once)

```bash
python backfill_cbn_history.py --dry-run   # preview counts, writes nothing
python backfill_cbn_history.py             # writes ~2,200 rows back to 2002
```

Expected output ends with something like:
`rates table now holds 2,190 rows across 976 distinct dates.`

## 4. Keep it current (daily)

```bash
python ngx_refresh.py --report   # DB only, no Excel
python ngx_refresh.py            # DB + Excel workbook
```

## 5. Point the dashboard at the same DB

Locally it picks up `DATABASE_URL` automatically. On **Streamlit Cloud**, add a
secret (Settings → Secrets):

```toml
DATABASE_URL = "postgresql+psycopg2://USER:PASSWORD@HOST/DBNAME?sslmode=require"
```

Then `streamlit run dashboard.py`. The sidebar shows **Backend: Postgres** when
connected correctly.

---

## 6. Schedule the daily refresh in the cloud (recommended)

Stop running it from your laptop. A simple GitHub Actions cron keeps the DB
fresh even when your machine is off. Add `.github/workflows/refresh.yml`:

```yaml
name: Daily refresh
on:
  schedule:
    - cron: "30 17 * * 1-5"   # 17:30 UTC weekdays (~18:30 WAT, after auctions)
  workflow_dispatch: {}
jobs:
  refresh:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "3.11" }
      - run: pip install -r requirements.txt
      - run: python ngx_refresh.py --report
        env:
          DATABASE_URL: ${{ secrets.DATABASE_URL }}
```

Add `DATABASE_URL` under the repo's **Settings → Secrets and variables →
Actions**. (The `--report` flag skips the Excel write, which doesn't make sense
in CI; the database still updates.)

---

## Notes & caveats

- **Re-running is safe.** Every write is an upsert on the primary key, so the
  backfill and the daily refresh overwrite rather than duplicate.
- **Data hygiene built in.** The backfill drops junk tenors (e.g. `44010`) and
  validates that each auction's maturity falls a sensible distance after its
  auction date — which fixes the earlier maturity-before-auction rows.
- **`allotted` can be noisy.** A minority of CBN records report
  `totalSuccessful` in a different unit; bid-cover is computed from
  `subscribed / offered`, which is reliable. Treat the allotted column as
  indicative.
- **TradingEconomics** is scraped via regex and its ToS restricts scraping;
  fine for a personal build, but not a foundation for a commercial product —
  plan to replace it (FMDQ / official feeds) before charging for the data.
