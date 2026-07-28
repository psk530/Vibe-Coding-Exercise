"""Load the raw Amazon TV-category export (Dashboard-Dataset(min).xlsx) into SQLite.

Aggregates row-level (SKU x week) data up to (week, brand) and writes a single
`weekly_sales` table. Everything the dashboard needs -- brand totals, conversion
rate, period filters -- is derived from this one table at query time, so there
is no duplicated/precomputed aggregate to fall out of sync.

Runs unattended once a day (see etl/run_daily_refresh.bat + the Windows Task
Scheduler entry it's registered under), so it loads into a fresh
`weekly_sales_new` table and only swaps it into place once fully written --
the API never sees a dropped/half-populated table if a request lands mid-run.

Usage:
    python etl/load_excel_to_db.py
"""
import datetime
import sqlite3
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
XLSX_PATH = ROOT / "Dashboard-Dataset(min).xlsx"
DB_PATH = ROOT / "db" / "dashboard.db"

BRANDS = ["Samsung", "LG", "TCL", "Insignia", "HiSense", "Amazon", "Sony", "Roku", "Vizio", "Toshiba"]

SCHEMA = """
    week_id INTEGER NOT NULL,
    brand   TEXT NOT NULL,
    rev     INTEGER NOT NULL,
    units   INTEGER NOT NULL,
    traffic INTEGER NOT NULL,
    PRIMARY KEY (week_id, brand)
"""


def load():
    print(f"[{datetime.datetime.now().isoformat(timespec='seconds')}] Reading {XLSX_PATH} ...")
    df = pd.read_excel(XLSX_PATH, sheet_name=0)

    lower_map = {b.lower(): b for b in BRANDS}
    df["brand_norm"] = df["Brand2"].str.lower().map(lower_map)
    sub = df[df["brand_norm"].notna()]

    agg = (
        sub.groupby(["Week ID", "brand_norm"])
        .agg(
            units=("Units Sold", "sum"),
            rev=("Retail Sales (USD)", "sum"),
            traffic=("Total Traffic", "sum"),
        )
        .reset_index()
    )
    agg["rev"] = agg["rev"].round().astype(int)
    agg["units"] = agg["units"].astype(int)
    agg["traffic"] = agg["traffic"].astype(int)

    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DROP TABLE IF EXISTS weekly_sales_new")
    conn.execute(f"CREATE TABLE weekly_sales_new ({SCHEMA})")
    conn.executemany(
        "INSERT INTO weekly_sales_new (week_id, brand, rev, units, traffic) VALUES (?,?,?,?,?)",
        agg[["Week ID", "brand_norm", "rev", "units", "traffic"]].itertuples(index=False, name=None),
    )
    conn.execute("CREATE INDEX idx_weekly_sales_new_week ON weekly_sales_new(week_id)")
    conn.execute("CREATE INDEX idx_weekly_sales_new_brand ON weekly_sales_new(brand)")

    has_old = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='weekly_sales'"
    ).fetchone()
    if has_old:
        conn.execute("DROP TABLE IF EXISTS weekly_sales_old")
        conn.execute("ALTER TABLE weekly_sales RENAME TO weekly_sales_old")
        conn.execute("DROP TABLE weekly_sales_old")  # also drops its indexes, freeing the canonical names
    conn.execute("ALTER TABLE weekly_sales_new RENAME TO weekly_sales")
    conn.execute("DROP INDEX idx_weekly_sales_new_week")
    conn.execute("DROP INDEX idx_weekly_sales_new_brand")
    conn.execute("CREATE INDEX idx_weekly_sales_week ON weekly_sales(week_id)")
    conn.execute("CREATE INDEX idx_weekly_sales_brand ON weekly_sales(brand)")
    conn.commit()

    n = conn.execute("SELECT COUNT(*) FROM weekly_sales").fetchone()[0]
    wk_min, wk_max = conn.execute("SELECT MIN(week_id), MAX(week_id) FROM weekly_sales").fetchone()
    conn.close()
    print(f"[{datetime.datetime.now().isoformat(timespec='seconds')}] Loaded {n} (week, brand) rows into {DB_PATH} [{wk_min}-{wk_max}]")


if __name__ == "__main__":
    load()
