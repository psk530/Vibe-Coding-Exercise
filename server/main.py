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
import os
import sqlite3
from pathlib import Path

import anthropic
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

ROOT = Path(__file__).resolve().parent.parent
DB_PATH = ROOT / "db" / "dashboard.db"

app = FastAPI(title="Brand Dashboard API")
anthropic_client = anthropic.Anthropic()

CHAT_MODEL = "claude-sonnet-5"

CHAT_SYSTEM_TEMPLATE = """당신은 'US TV Market — Brand Competition Dashboard'의 데이터 분석 어시스턴트입니다.
아래는 2024~2026년 Amazon.com TV 카테고리의 주차별 브랜드 판매 데이터(CSV)입니다.
week 컬럼은 YYYYWW 형식입니다 (예: 202520 = 2025년 20주차). conv(전환율, %)는 units/traffic*100 입니다.
브랜드 목록: Samsung, LG, TCL, Insignia, HiSense, Amazon, Sony, Roku, Vizio, Toshiba.

답변 규칙:
1. 매출/트래픽/판매량/전환율 등 데이터에서 직접 확인 가능한 수치는 CSV를 근거로 정확히 계산해서 답하세요.
2. "왜 ~했는지", "원인이 뭐야" 같은 질문에는 먼저 CSV 데이터 안에서 근거(전후 주차 대비 변화, 다른 브랜드와의 비교, 계절성 등)를 찾아 분석하고,
   필요하면 web_search 도구로 관련 뉴스/기사(신제품 출시, 프로모션, 이벤트, 시즌 세일 등)를 검색해 근거를 보강하세요.
3. 답변은 간결한 한국어로 작성하고, 확실한 사실과 추정을 구분해서 표현하세요 (예: "~로 추정됩니다").

데이터(CSV):
{data_csv}
"""


class ChatMessage(BaseModel):
    role: str
    content: str


class ChatRequest(BaseModel):
    message: str
    history: list[ChatMessage] = []


def build_data_csv() -> str:
    conn = get_conn()
    rows = conn.execute(
        "SELECT week_id AS week, brand, rev, units, traffic FROM weekly_sales ORDER BY week_id, brand"
    ).fetchall()
    conn.close()
    lines = ["week,brand,rev,units,traffic"]
    for r in rows:
        lines.append(f"{r['week']},{r['brand']},{r['rev']},{r['units']},{r['traffic']}")
    return "\n".join(lines)


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


@app.post("/api/chat")
def chat(req: ChatRequest):
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise HTTPException(500, "ANTHROPIC_API_KEY가 서버에 설정되어 있지 않습니다.")

    system = CHAT_SYSTEM_TEMPLATE.format(data_csv=build_data_csv())
    messages = [{"role": m.role, "content": m.content} for m in req.history]
    messages.append({"role": "user", "content": req.message})

    try:
        resp = anthropic_client.messages.create(
            model=CHAT_MODEL,
            max_tokens=1024,
            system=system,
            tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 5}],
            messages=messages,
        )
    except Exception:
        resp = anthropic_client.messages.create(
            model=CHAT_MODEL,
            max_tokens=1024,
            system=system,
            messages=messages,
        )

    text = "".join(block.text for block in resp.content if block.type == "text")
    return {"reply": text or "답변을 생성하지 못했습니다."}


@app.get("/")
def index():
    return FileResponse(ROOT / "brand-dashboard.html")


app.mount("/static", StaticFiles(directory=str(ROOT)), name="static")
