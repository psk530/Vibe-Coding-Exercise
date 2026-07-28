"""API + static file server for the US TV Market brand dashboard.

Serves brand-dashboard.html and two read-only JSON endpoints backed by
db/dashboard.db (built by etl/load_excel_to_db.py):

  GET /api/weekly  -> every (week, brand) row: the raw fact table
  GET /api/totals  -> all-time per-brand rollup (rev/units/traffic/conv)

The frontend fetches these once on load and does its own period/brand
filtering client-side, exactly as it did with the old embedded JS arrays --
only the data source changed.

Run with:
    uvicorn server.main:app --reload --port 8000
then open http://localhost:8000/
"""
import sqlite3
from pathlib import Path

from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

ROOT = Path(__file__).resolve().parent.parent
DB_PATH = ROOT / "db" / "dashboard.db"

app = FastAPI(title="Brand Dashboard API")


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


@app.get("/api/weekly")
def weekly():
    conn = get_conn()
    rows = conn.execute(
        "SELECT week_id AS week, brand, rev, units, traffic FROM weekly_sales ORDER BY week_id, brand"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


@app.get("/api/totals")
def totals():
    conn = get_conn()
    rows = conn.execute(
        """
        SELECT brand,
               SUM(rev) AS rev,
               SUM(units) AS units,
               SUM(traffic) AS traffic
        FROM weekly_sales
        GROUP BY brand
        """
    ).fetchall()
    conn.close()
    result = []
    for r in rows:
        conv = (r["units"] / r["traffic"] * 100) if r["traffic"] else 0
        result.append({
            "brand": r["brand"],
            "rev": r["rev"],
            "units": r["units"],
            "traffic": r["traffic"],
            "conv": round(conv, 2),
        })
    return result


@app.get("/")
def index():
    return FileResponse(ROOT / "brand-dashboard.html")


app.mount("/static", StaticFiles(directory=str(ROOT)), name="static")
