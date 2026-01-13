from __future__ import annotations

import asyncio
import json
import os
import time
from typing import Any, Dict, Optional, Set

import httpx
from fastapi import Body, FastAPI, HTTPException, Request
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

clients: Set[asyncio.Queue] = set()
NOTIFY_TOKEN = os.getenv("NOTIFY_TOKEN", "").strip()
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
    message = {"type": "update", "ts": time.time(), "payload": payload}
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
