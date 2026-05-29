#!/usr/bin/env python3
"""
db.py — Shared database layer for Rune (Nigerian fixed income intelligence)
============================================================================
One module used by BOTH the refresh/backfill writers and the Streamlit reader.

Why this exists
---------------
The old setup wrote to a local SQLite file (ngx.db). On Streamlit Cloud that
file lives on an ephemeral disk that is wiped on every redeploy/restart — which
is why only two days of history ever survived. This layer lets you point the
whole app at a hosted Postgres (Neon / Supabase / RDS) by setting one env var,
while still falling back to local SQLite for offline development.

Configuration
-------------
Set DATABASE_URL to your Postgres connection string, e.g.

    export DATABASE_URL="postgresql+psycopg2://USER:PASS@HOST:5432/DBNAME?sslmode=require"

On Streamlit Cloud, add it under Settings → Secrets as:

    DATABASE_URL = "postgresql+psycopg2://..."

If DATABASE_URL is unset, this falls back to a local SQLite file next to this
module (./ngx.db) so nothing breaks while developing offline.

The public surface is intentionally tiny:

    get_engine()                      -> SQLAlchemy Engine (cached)
    init_schema(engine)               -> create tables/indexes if missing
    upsert(engine, table_name, rows)  -> dialect-aware INSERT ... ON CONFLICT
    read_sql(sql, params, engine)     -> pandas DataFrame
"""

from __future__ import annotations

import os
from pathlib import Path
from functools import lru_cache

import pandas as pd
from sqlalchemy import (
    create_engine, MetaData, Table, Column, String, Float, Index, text,
)
from sqlalchemy.engine import Engine

# ── Schema ──────────────────────────────────────────────────────────────────
# Dates are stored as ISO TEXT ('YYYY-MM-DD') to keep SQLite and Postgres
# byte-for-byte identical and avoid timezone surprises. Column names match the
# original schema exactly so existing dashboard SQL keeps working unchanged.
metadata = MetaData()

rates = Table(
    "rates", metadata,
    Column("snapshot_date", String, primary_key=True),
    Column("instrument",    String, primary_key=True),
    Column("rate",          Float,  nullable=False),
    Column("source",        String),
    Column("fetched_at",    String, nullable=False),
    Index("idx_rates_instrument", "instrument"),
)

auctions = Table(
    "auctions", metadata,
    Column("auction_date",  String, primary_key=True),
    Column("security_type", String, primary_key=True),
    Column("tenor",         String, primary_key=True),
    Column("offered",       Float),
    Column("subscribed",    Float),
    Column("allotted",      Float),
    Column("stop_rate",     Float),
    Column("maturity_date", String),
    Column("fetched_at",    String, nullable=False),
    Index("idx_auctions_security", "security_type", "tenor"),
)

eurobonds = Table(
    "eurobonds", metadata,
    Column("snapshot_date", String, primary_key=True),
    Column("bond_name",     String, primary_key=True),
    Column("price",         Float),
    Column("yield_pct",     Float),
    Column("fetched_at",    String, nullable=False),
    Index("idx_eurobonds_bond", "bond_name"),
)

_TABLES = {"rates": rates, "auctions": auctions, "eurobonds": eurobonds}


# ── Engine ────────────────────────────────────────────────────────────────────
def _default_sqlite_url() -> str:
    return f"sqlite:///{Path(__file__).resolve().parent / 'ngx.db'}"


def database_url() -> str:
    """Resolve the connection string. Prefers env var, then Streamlit secrets,
    then a local SQLite file."""
    url = os.environ.get("DATABASE_URL")
    if url:
        return url
    # Streamlit secrets, if running inside Streamlit and the key is present.
    try:
        import streamlit as st  # noqa: WPS433 (optional dependency)
        if "DATABASE_URL" in st.secrets:
            return str(st.secrets["DATABASE_URL"])
    except Exception:
        pass
    return _default_sqlite_url()


@lru_cache(maxsize=1)
def get_engine() -> Engine:
    """Return a cached SQLAlchemy engine. pool_pre_ping avoids stale
    connections on serverless Postgres (Neon/Supabase idle-timeout)."""
    url = database_url()
    kwargs = {"pool_pre_ping": True, "future": True}
    if url.startswith("sqlite"):
        # Allow cross-thread use (Streamlit reruns on multiple threads).
        kwargs["connect_args"] = {"check_same_thread": False}
    return create_engine(url, **kwargs)


def is_postgres(engine: Engine | None = None) -> bool:
    return (engine or get_engine()).dialect.name == "postgresql"


# ── Schema management ──────────────────────────────────────────────────────────
def init_schema(engine: Engine | None = None) -> None:
    """Create all tables and indexes if they do not already exist.
    Safe to call on every run."""
    metadata.create_all(engine or get_engine())


# ── Upsert ────────────────────────────────────────────────────────────────────
def upsert(engine: Engine, table_name: str, rows: list[dict], chunk: int = 500) -> int:
    """Dialect-aware bulk upsert (INSERT ... ON CONFLICT DO UPDATE).

    `rows` is a list of dicts keyed by column name. Conflicts on the table's
    primary key update the non-key columns, so re-running the same day or
    re-backfilling overwrites cleanly instead of duplicating.
    """
    if not rows:
        return 0

    table = _TABLES[table_name]
    pk_cols = [c.name for c in table.primary_key.columns]

    if engine.dialect.name == "postgresql":
        from sqlalchemy.dialects.postgresql import insert as _insert
    else:
        from sqlalchemy.dialects.sqlite import insert as _insert

    written = 0
    with engine.begin() as conn:
        for i in range(0, len(rows), chunk):
            batch = rows[i:i + chunk]
            stmt = _insert(table).values(batch)
            update_cols = {
                c.name: stmt.excluded[c.name]
                for c in table.columns
                if c.name not in pk_cols
            }
            stmt = stmt.on_conflict_do_update(index_elements=pk_cols, set_=update_cols)
            conn.execute(stmt)
            written += len(batch)
    return written


# ── Read ──────────────────────────────────────────────────────────────────────
def read_sql(sql: str, params: dict | None = None, engine: Engine | None = None) -> pd.DataFrame:
    """Run a parameterised SELECT and return a DataFrame.

    Use named parameters (:name) so the same SQL works on SQLite and Postgres.
    Returns an empty DataFrame if the tables don't exist yet.
    """
    eng = engine or get_engine()
    try:
        with eng.connect() as conn:
            return pd.read_sql_query(text(sql), conn, params=params or {})
    except Exception:
        return pd.DataFrame()
