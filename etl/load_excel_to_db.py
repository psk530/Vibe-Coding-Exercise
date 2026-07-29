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
    organic INTEGER NOT NULL,
    PRIMARY KEY (week_id, brand)
"""

# Row-level (SKU x week) detail for the 10 focus brands -- backs the O/S share
# pie chart and the sortable model table. Unlike weekly_sales this is NOT
# pre-aggregated: the API groups by SKU at query time so period filtering
# (and future column additions) don't require another ETL pass.
SKU_SCHEMA = """
    week_id         INTEGER NOT NULL,
    brand           TEXT NOT NULL,
    retailer_sku    TEXT,
    model_number    TEXT,
    units_sold      INTEGER NOT NULL,
    retail_sales    REAL NOT NULL,
    total_traffic   INTEGER NOT NULL,
    organic_traffic INTEGER NOT NULL,
    os              TEXT
"""


def swap_in(conn, new_name, canonical_name, indexes):
    """Atomically replace `canonical_name` with the fully-built `new_name` table.

    `indexes` is a list of (suffix, column) pairs already indexed on the new
    table as idx_{new_name}_{suffix}; they get renamed to idx_{canonical_name}_{suffix}.
    """
    has_old = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (canonical_name,)
    ).fetchone()
    if has_old:
        old_name = f"{canonical_name}_old"
        conn.execute(f"DROP TABLE IF EXISTS {old_name}")
        conn.execute(f"ALTER TABLE {canonical_name} RENAME TO {old_name}")
        conn.execute(f"DROP TABLE {old_name}")  # also drops its indexes, freeing the canonical names
    conn.execute(f"ALTER TABLE {new_name} RENAME TO {canonical_name}")
    for suffix, column in indexes:
        conn.execute(f"DROP INDEX idx_{new_name}_{suffix}")
        conn.execute(f"CREATE INDEX idx_{canonical_name}_{suffix} ON {canonical_name}({column})")


def load_weekly_sales(df: pd.DataFrame, conn: sqlite3.Connection):
    lower_map = {b.lower(): b for b in BRANDS}
    df = df.copy()
    df["brand_norm"] = df["Brand2"].str.lower().map(lower_map)
    sub = df[df["brand_norm"].notna()]

    agg = (
        sub.groupby(["Week ID", "brand_norm"])
        .agg(
            units=("Units Sold", "sum"),
            rev=("Retail Sales (USD)", "sum"),
            traffic=("Total Traffic", "sum"),
            organic=("Organic Traffic", "sum"),
        )
        .reset_index()
    )
    agg["rev"] = agg["rev"].round().astype(int)
    agg["units"] = agg["units"].astype(int)
    agg["traffic"] = agg["traffic"].astype(int)
    agg["organic"] = agg["organic"].astype(int)

    conn.execute("DROP TABLE IF EXISTS weekly_sales_new")
    conn.execute(f"CREATE TABLE weekly_sales_new ({SCHEMA})")
    conn.executemany(
        "INSERT INTO weekly_sales_new (week_id, brand, rev, units, traffic, organic) VALUES (?,?,?,?,?,?)",
        agg[["Week ID", "brand_norm", "rev", "units", "traffic", "organic"]].itertuples(index=False, name=None),
    )
    conn.execute("CREATE INDEX idx_weekly_sales_new_week ON weekly_sales_new(week_id)")
    conn.execute("CREATE INDEX idx_weekly_sales_new_brand ON weekly_sales_new(brand)")

    swap_in(conn, "weekly_sales_new", "weekly_sales", [("week", "week_id"), ("brand", "brand")])

    n = conn.execute("SELECT COUNT(*) FROM weekly_sales").fetchone()[0]
    wk_min, wk_max = conn.execute("SELECT MIN(week_id), MAX(week_id) FROM weekly_sales").fetchone()
    print(f"[{datetime.datetime.now().isoformat(timespec='seconds')}] Loaded {n} (week, brand) rows into weekly_sales [{wk_min}-{wk_max}]")


def load_sku_weekly(df: pd.DataFrame, conn: sqlite3.Connection):
    lower_map = {b.lower(): b for b in BRANDS}
    df = df.copy()
    df["brand_norm"] = df["Brand2"].str.lower().map(lower_map)
    sub = df[df["brand_norm"].notna()].copy()

    sub["units_sold"] = sub["Units Sold"].fillna(0).astype(int)
    sub["retail_sales"] = sub["Retail Sales (USD)"].fillna(0.0).astype(float)
    sub["total_traffic"] = sub["Total Traffic"].fillna(0).astype(int)
    sub["organic_traffic"] = sub["Organic Traffic"].fillna(0).astype(int)
    sub["os"] = sub["O/S"].fillna("os_-")
    sub["retailer_sku"] = sub["Retailer SKU"].astype(object).where(sub["Retailer SKU"].notna(), None)
    sub["model_number"] = sub["Model Number"].astype(object).where(sub["Model Number"].notna(), None)

    conn.execute("DROP TABLE IF EXISTS sku_weekly_new")
    conn.execute(f"CREATE TABLE sku_weekly_new ({SKU_SCHEMA})")
    conn.executemany(
        """INSERT INTO sku_weekly_new
           (week_id, brand, retailer_sku, model_number, units_sold, retail_sales, total_traffic, organic_traffic, os)
           VALUES (?,?,?,?,?,?,?,?,?)""",
        sub[[
            "Week ID", "brand_norm", "retailer_sku", "model_number",
            "units_sold", "retail_sales", "total_traffic", "organic_traffic", "os",
        ]].itertuples(index=False, name=None),
    )
    conn.execute("CREATE INDEX idx_sku_weekly_new_week ON sku_weekly_new(week_id)")

    swap_in(conn, "sku_weekly_new", "sku_weekly", [("week", "week_id")])

    n = conn.execute("SELECT COUNT(*) FROM sku_weekly").fetchone()[0]
    print(f"[{datetime.datetime.now().isoformat(timespec='seconds')}] Loaded {n} sku-week rows into sku_weekly")


def load():
    print(f"[{datetime.datetime.now().isoformat(timespec='seconds')}] Reading {XLSX_PATH} ...")
    df = pd.read_excel(XLSX_PATH, sheet_name=0)

    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    load_weekly_sales(df, conn)
    load_sku_weekly(df, conn)
    conn.commit()
    conn.close()


if __name__ == "__main__":
    load()
