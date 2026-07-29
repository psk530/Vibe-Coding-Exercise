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
import json
import os
import re
import sqlite3
from pathlib import Path

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from openai import OpenAI
from pydantic import BaseModel

ROOT = Path(__file__).resolve().parent.parent
DB_PATH = ROOT / "db" / "dashboard.db"

app = FastAPI(title="Brand Dashboard API")

CHAT_MODEL = os.environ.get("OPENROUTER_MODEL", "openai/gpt-4o")


def get_openrouter_client() -> OpenAI:
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        raise HTTPException(500, "OPENROUTER_API_KEY가 서버에 설정되어 있지 않습니다.")
    return OpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)

CHAT_SYSTEM_TEMPLATE = """당신은 'US TV Market — Brand Competition Dashboard'의 데이터 분석 어시스턴트입니다.
아래는 2024~2026년 Amazon.com TV 카테고리의 주차별 브랜드 판매 데이터(CSV)입니다.
week 컬럼은 YYYYWW 형식입니다 (예: 202520 = 2025년 20주차). conv(전환율, %)는 units/traffic*100 입니다.
브랜드 목록: Samsung, LG, TCL, Insignia, HiSense, Amazon, Sony, Roku, Vizio, Toshiba.

답변은 반드시 아래 JSON 스키마를 따르는 하나의 JSON 객체로만 출력하세요. JSON 앞뒤에 다른 텍스트를 붙이지 마세요.

{{
  "blocks": [
    {{"type": "text", "content": "설명 문장"}},
    {{"type": "table", "title": "표 제목(선택, 없으면 빈 문자열)", "headers": ["열1", "열2"], "rows": [["값1", "값2"]]}},
    {{"type": "chart", "chartType": "bar 또는 line", "title": "차트 제목(선택)", "xKey": "label", "series": [{{"key": "value", "name": "계열 이름"}}], "data": [{{"label": "202519", "value": 12345}}]}}
  ]
}}

블록 작성 규칙:
1. 매출/트래픽/판매량/전환율 등 데이터에서 직접 확인 가능한 수치는 CSV를 근거로 정확히 계산해서 답하세요.
2. "왜 ~했는지", "원인이 뭐야" 같은 질문에는 먼저 CSV 데이터 안에서 근거(전후 주차 대비 변화, 다른 브랜드와의 비교, 계절성 등)를 찾아 분석하고,
   검색으로 얻은 최신 뉴스/기사(신제품 출시, 프로모션, 이벤트, 시즌 세일 등)가 있다면 근거로 함께 활용하세요.
   출처를 밝힐 때는 "[텍스트](URL)" 같은 마크다운 링크 문법을 쓰지 말고, "OO 매체에 따르면"처럼 자연스러운 문장으로 녹여서 표현하세요.
3. 여러 주차나 여러 브랜드에 걸친 수치 비교/추이는 text로 나열하지 말고 table 또는 chart 블록으로 표현해서 가독성을 높이세요.
   시간 흐름(주차별 추이)에는 chart(line)를, 브랜드 간 비교에는 chart(bar) 또는 table을 우선 사용하세요.
   설명, 원인 분석, 결론처럼 표/차트로 표현하기 어려운 내용은 text 블록으로 쓰세요.
   한 답변 안에 여러 블록을 섞어 써도 됩니다 (예: text로 요약 → table/chart로 세부 수치 → text로 결론).
4. chart 블록의 data 안 숫자 값(value 등 series의 key에 해당하는 값)은 반드시 순수 숫자로 넣으세요 (문자열 금지).
5. 답변은 간결한 한국어로 작성하고, 확실한 사실과 추정을 구분해서 표현하세요 (예: "~로 추정됩니다").
6. text/table의 문자열 안에는 #, *, **, -, |, `, > 같은 마크다운 기호나 이모지를 쓰지 마세요. 강조하고 싶어도 별표 없이 평문으로 쓰세요.

데이터(CSV):
{data_csv}
"""


class ChatMessage(BaseModel):
    role: str
    content: str


class ChatRequest(BaseModel):
    message: str
    history: list[ChatMessage] = []


def strip_markdown(text: str) -> str:
    text = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'\1 (\2)', text)
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'__(.+?)__', r'\1', text)
    text = re.sub(r'(?<!\*)\*([^\*\n]+?)\*(?!\*)', r'\1', text)
    text = re.sub(r'`([^`]+)`', r'\1', text)
    text = re.sub(r'^\s{0,3}#{1,6}\s*', '', text, flags=re.MULTILINE)
    text = re.sub(r'^\s{0,3}>\s?', '', text, flags=re.MULTILINE)
    text = re.sub(r'^\s{0,3}[-*+]\s+', '', text, flags=re.MULTILINE)
    return text


def parse_blocks(raw: str) -> list:
    try:
        obj = json.loads(raw)
        blocks = obj.get("blocks")
        if not isinstance(blocks, list) or not blocks:
            raise ValueError("no blocks in response")
    except Exception:
        return [{"type": "text", "content": strip_markdown(raw or "")}]

    cleaned = []
    for b in blocks:
        if not isinstance(b, dict):
            continue
        btype = b.get("type")
        if btype == "text":
            content = strip_markdown(str(b.get("content", "")))
            if content.strip():
                cleaned.append({"type": "text", "content": content})
        elif btype == "table":
            headers = [strip_markdown(str(h)) for h in (b.get("headers") or [])]
            rows = [[strip_markdown(str(c)) for c in row] for row in (b.get("rows") or [])]
            if headers and rows:
                cleaned.append({
                    "type": "table",
                    "title": strip_markdown(str(b.get("title") or "")),
                    "headers": headers,
                    "rows": rows,
                })
        elif btype == "chart":
            data = b.get("data")
            series = b.get("series") or [{"key": "value", "name": "값"}]
            if isinstance(data, list) and data:
                cleaned.append({
                    "type": "chart",
                    "chartType": b.get("chartType") if b.get("chartType") in ("bar", "line") else "bar",
                    "title": strip_markdown(str(b.get("title") or "")),
                    "xKey": b.get("xKey") or "label",
                    "series": series,
                    "data": data,
                })
    return cleaned or [{"type": "text", "content": "답변을 생성하지 못했습니다."}]


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


def call_chat_model(client: OpenAI, messages: list):
    attempts = [
        dict(model=f"{CHAT_MODEL}:online", messages=messages, response_format={"type": "json_object"}),
        dict(model=CHAT_MODEL, messages=messages, response_format={"type": "json_object"}),
        dict(model=CHAT_MODEL, messages=messages),
    ]
    last_err = None
    for kwargs in attempts:
        try:
            return client.chat.completions.create(**kwargs)
        except Exception as e:
            last_err = e
    raise last_err


@app.post("/api/chat")
def chat(req: ChatRequest):
    client = get_openrouter_client()

    system = CHAT_SYSTEM_TEMPLATE.format(data_csv=build_data_csv())
    messages = [{"role": "system", "content": system}]
    messages += [{"role": m.role, "content": m.content} for m in req.history]
    messages.append({"role": "user", "content": req.message})

    try:
        resp = call_chat_model(client, messages)
    except Exception as e:
        print(f"chat error: {e!r}")
        raise HTTPException(500, "AI 응답 생성 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.")

    raw = resp.choices[0].message.content
    return {"blocks": parse_blocks(raw)}


@app.get("/")
def index():
    return FileResponse(ROOT / "brand-dashboard.html")


app.mount("/static", StaticFiles(directory=str(ROOT)), name="static")
