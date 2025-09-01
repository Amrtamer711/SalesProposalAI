import asyncio
from datetime import datetime
import subprocess
import shutil

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse

import config
from llm import main_llm_loop
from font_utils import install_custom_fonts

# Install custom fonts on startup
install_custom_fonts()

# Check LibreOffice installation
logger = config.logger
logger.info("[STARTUP] Checking LibreOffice installation...")
libreoffice_found = False
for cmd in ['libreoffice', 'soffice', '/usr/bin/libreoffice']:
    if shutil.which(cmd) or subprocess.run(['which', cmd], capture_output=True).returncode == 0:
        try:
            result = subprocess.run([cmd, '--version'], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                logger.info(f"[STARTUP] LibreOffice found at '{cmd}': {result.stdout.strip()}")
                libreoffice_found = True
                break
        except Exception as e:
            logger.debug(f"[STARTUP] Error checking {cmd}: {e}")

if not libreoffice_found:
    logger.warning("[STARTUP] LibreOffice not found! PDF conversion will use fallback method.")
else:
    logger.info("[STARTUP] LibreOffice is ready for PDF conversion.")

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