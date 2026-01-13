from __future__ import annotations

import asyncio
import json
import os
import time
from typing import Any, Dict, Optional, Set

from fastapi import Body, FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
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


@app.get("/health")
async def health():
    return {"ok": True, "clients": len(clients)}


app.mount("/", StaticFiles(directory=".", html=True), name="static")
