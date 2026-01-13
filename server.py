from __future__ import annotations

import asyncio
import contextlib
import hashlib
import json
import os
import time
from typing import Any, Dict, Optional, Set

import asyncpg
import httpx
from fastapi import Body, FastAPI, HTTPException, Request
from fastapi.encoders import jsonable_encoder
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

app = FastAPI(title="Realtime Dashboard Gateway")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
async def startup() -> None:
    global db_poll_task
    if db_configured():
        await init_db_pool()
        db_poll_task = asyncio.create_task(poll_db_changes())


@app.on_event("shutdown")
async def shutdown() -> None:
    global db_poll_task, db_pool
    if db_poll_task is not None:
        db_poll_task.cancel()
        with contextlib.suppress(asyncio.CancelledError):
            await db_poll_task
        db_poll_task = None
    if db_pool is not None:
        await db_pool.close()
        db_pool = None

clients: Set[asyncio.Queue] = set()
NOTIFY_TOKEN = os.getenv("NOTIFY_TOKEN", "").strip()
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
DB_HOST = os.getenv("DB_HOST", "").strip()
DB_PORT = int(os.getenv("DB_PORT", "5432").strip() or 5432)
DB_NAME = os.getenv("DB_NAME", "").strip()
DB_USER = os.getenv("DB_USER", "").strip()
DB_PASSWORD = os.getenv("DB_PASSWORD", "").strip()
DB_TABLE = os.getenv("DB_TABLE", "validades").strip() or "validades"
DB_SSLMODE = os.getenv("DB_SSLMODE", "disable").strip().lower()
DB_POLL_INTERVAL = float(os.getenv("DB_POLL_INTERVAL", "8").strip() or 8)

db_pool: Optional[asyncpg.Pool] = None
db_poll_task: Optional[asyncio.Task] = None
N8N_BASE_URL = os.getenv("N8N_BASE_URL", "https://myn8n.seommerce.shop").strip().rstrip("/")
N8N_WEBHOOK_PLANILHA = os.getenv(
    "N8N_WEBHOOK_PLANILHA",
    f"{N8N_BASE_URL}/webhook/planilha-atualizada",
).strip()
N8N_WEBHOOK_VALIDADE = os.getenv(
    "N8N_WEBHOOK_VALIDADE",
    f"{N8N_BASE_URL}/webhook/validade",
).strip()
N8N_WEBHOOK_BARCODE = os.getenv(
    "N8N_WEBHOOK_BARCODE",
    f"{N8N_BASE_URL}/webhook/barcode",
).strip()
N8N_WEBHOOK_ACOES = os.getenv(
    "N8N_WEBHOOK_ACOES",
    f"{N8N_BASE_URL}/webhook/a%C3%A7%C3%B5es",
).strip()


def normalize_payload(payload: Any) -> Dict[str, Any]:
    if payload is None:
        return {}
    if isinstance(payload, dict):
        return payload
    return {"data": payload}


def broadcast(payload: Dict[str, Any]) -> int:
    message = {
        "type": "update",
        "ts": time.time(),
        "payload": jsonable_encoder(payload),
    }
    for queue in list(clients):
        queue.put_nowait(message)
    return len(clients)


def require_notify_token(request: Request) -> None:
    if not NOTIFY_TOKEN:
        return
    auth_header = request.headers.get("authorization", "")
    token = ""
    if auth_header.lower().startswith("bearer "):
        token = auth_header.split(" ", 1)[1].strip()
    if not token:
        token = request.headers.get("x-notify-token", "").strip()
    if token != NOTIFY_TOKEN:
        raise HTTPException(status_code=401, detail="Unauthorized")


def db_configured() -> bool:
    return bool(DATABASE_URL or DB_HOST)


def resolve_table_name(raw_name: str) -> str:
    parts = [part for part in raw_name.split(".") if part]
    if not parts:
        raise HTTPException(status_code=500, detail="DB_TABLE inválida")
    for part in parts:
        if not part.replace("_", "").isalnum():
            raise HTTPException(status_code=500, detail="DB_TABLE inválida")
    return ".".join(parts)


def ssl_enabled() -> bool:
    return DB_SSLMODE in {"require", "verify-ca", "verify-full", "true", "1"}


async def init_db_pool() -> None:
    global db_pool
    if db_pool is not None or not db_configured():
        return

    if DATABASE_URL:
        db_pool = await asyncpg.create_pool(dsn=DATABASE_URL, ssl=ssl_enabled())
        return

    if not (DB_HOST and DB_NAME and DB_USER):
        raise HTTPException(status_code=500, detail="Configuração de banco incompleta")

    db_pool = await asyncpg.create_pool(
        host=DB_HOST,
        port=DB_PORT,
        database=DB_NAME,
        user=DB_USER,
        password=DB_PASSWORD,
        ssl=ssl_enabled(),
    )


async def fetch_table_rows() -> list[dict]:
    if db_pool is None:
        raise HTTPException(status_code=503, detail="Banco de dados indisponível")
    table_name = resolve_table_name(DB_TABLE)
    query = f"SELECT * FROM {table_name} ORDER BY 1"
    async with db_pool.acquire() as conn:
        rows = await conn.fetch(query)
    return [dict(row) for row in rows]


def rows_signature(rows: list[dict]) -> str:
    payload = json.dumps(rows, sort_keys=True, default=str)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


async def enqueue_snapshot(queue: asyncio.Queue) -> None:
    if not db_configured():
        return
    await init_db_pool()
    if db_pool is None:
        return
    rows = await fetch_table_rows()
    message = {
        "type": "snapshot",
        "ts": time.time(),
        "payload": {"data": jsonable_encoder(rows)},
    }
    queue.put_nowait(message)


async def poll_db_changes() -> None:
    last_signature: Optional[str] = None
    while True:
        await asyncio.sleep(DB_POLL_INTERVAL)
        if not clients:
            continue
        try:
            await init_db_pool()
            if db_pool is None:
                continue
            rows = await fetch_table_rows()
            signature = rows_signature(rows)
            if signature == last_signature:
                continue
            last_signature = signature
            broadcast({"data": rows})
        except Exception:
            continue


async def proxy_webhook(request: Request, url: str) -> Response:
    body = await request.body()
    headers = {}
    content_type = request.headers.get("content-type")
    if content_type:
        headers["content-type"] = content_type
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.post(url, content=body, headers=headers)
    except httpx.RequestError as exc:
        raise HTTPException(status_code=502, detail="Upstream request failed") from exc

    response_headers = {}
    response_content_type = resp.headers.get("content-type")
    if response_content_type:
        response_headers["content-type"] = response_content_type
    content_disposition = resp.headers.get("content-disposition")
    if content_disposition:
        response_headers["content-disposition"] = content_disposition
    return Response(content=resp.content, status_code=resp.status_code, headers=response_headers)


async def event_stream(queue: asyncio.Queue, request: Request):
    while True:
        if await request.is_disconnected():
            break
        try:
            payload = await asyncio.wait_for(queue.get(), timeout=15)
        except asyncio.TimeoutError:
            yield "event: ping\ndata: {}\n\n"
            continue
        yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"


@app.get("/events")
async def events(request: Request):
    queue: asyncio.Queue = asyncio.Queue()
    clients.add(queue)
    try:
        await enqueue_snapshot(queue)
    except Exception:
        pass

    async def generator():
        try:
            async for chunk in event_stream(queue, request):
                yield chunk
        finally:
            clients.discard(queue)

    return StreamingResponse(generator(), media_type="text/event-stream")


@app.post("/webhook")
async def webhook(request: Request, payload: Any = Body(default_factory=dict)):
    require_notify_token(request)
    data = normalize_payload(payload)
    clients_count = broadcast(data)
    return JSONResponse({"ok": True, "clients": clients_count})


@app.post("/api/notify")
async def notify(
    request: Request,
    payload: Any = Body(default_factory=dict),
    action: Optional[str] = None,
):
    require_notify_token(request)
    data = normalize_payload(payload)
    clients_count = broadcast(data)
    return JSONResponse({"ok": True, "clients": clients_count})


@app.post("/api/notify/{action}")
async def notify_action(
    action: str,
    request: Request,
    payload: Any = Body(default_factory=dict),
):
    require_notify_token(request)
    data = normalize_payload(payload)
    data.setdefault("action", action)
    clients_count = broadcast(data)
    return JSONResponse({"ok": True, "clients": clients_count})


@app.post("/api/planilha-atualizada")
async def planilha_atualizada(request: Request):
    if db_configured():
        await init_db_pool()
        rows = await fetch_table_rows()
        return JSONResponse(jsonable_encoder(rows))
    return await proxy_webhook(request, N8N_WEBHOOK_PLANILHA)


@app.post("/api/validade")
async def validade(request: Request):
    return await proxy_webhook(request, N8N_WEBHOOK_VALIDADE)


@app.post("/api/barcode")
async def barcode(request: Request):
    return await proxy_webhook(request, N8N_WEBHOOK_BARCODE)


@app.post("/api/acoes")
async def acoes(request: Request):
    return await proxy_webhook(request, N8N_WEBHOOK_ACOES)


@app.get("/health")
async def health():
    return {"ok": True, "clients": len(clients)}


app.mount("/", StaticFiles(directory=".", html=True), name="static")
