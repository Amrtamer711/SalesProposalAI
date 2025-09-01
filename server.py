import asyncio
from datetime import datetime

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse

import config
from llm import main_llm_loop

app = FastAPI(title="Proposal Bot API")


@app.post("/slack/events")
async def slack_events(request: Request):
    body = await request.body()
    timestamp = request.headers.get("X-Slack-Request-Timestamp")
    signature = request.headers.get("X-Slack-Signature")

    if not config.signature_verifier.is_valid(body.decode(), timestamp, signature):
        raise HTTPException(status_code=403, detail="Invalid Slack signature")

    data = await request.json()
    if data.get("type") == "url_verification":
        return JSONResponse({"challenge": data["challenge"]})

    event = data.get("event", {})
    if event.get("type") == "message" and not event.get("bot_id"):
        asyncio.create_task(main_llm_loop(event["channel"], event["user"], event.get("text", ""), event))

    return JSONResponse({"status": "ok"})


@app.get("/health")
async def health():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()} 