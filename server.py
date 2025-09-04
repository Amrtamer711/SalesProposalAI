import asyncio
from datetime import datetime
import subprocess
import shutil
from contextlib import asynccontextmanager

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


async def periodic_cleanup():
    """Background task to clean up old data periodically"""
    while True:
        await asyncio.sleep(300)  # Every 5 minutes
        try:
            # Clean up old user histories
            from llm import user_history, pending_location_additions
            from datetime import timedelta
            
            # Clean user histories older than 1 hour
            cutoff = datetime.now() - timedelta(hours=1)
            expired_users = []
            for uid, history in user_history.items():
                if history and hasattr(history[-1], 'get'):
                    timestamp_str = history[-1].get('timestamp')
                    if timestamp_str:
                        try:
                            last_time = datetime.fromisoformat(timestamp_str)
                            if last_time < cutoff:
                                expired_users.append(uid)
                        except:
                            pass
            
            for uid in expired_users:
                del user_history[uid]
            
            if expired_users:
                logger.info(f"[CLEANUP] Removed {len(expired_users)} old user histories")
                
            # Clean pending locations older than 10 minutes
            location_cutoff = datetime.now() - timedelta(minutes=10)
            expired_locations = [
                uid for uid, data in pending_location_additions.items()
                if data.get("timestamp", datetime.now()) < location_cutoff
            ]
            for uid in expired_locations:
                del pending_location_additions[uid]
                
            if expired_locations:
                logger.info(f"[CLEANUP] Removed {len(expired_locations)} pending locations")
            
            # Clean up old temporary files
            import tempfile
            import time
            import os
            temp_dir = tempfile.gettempdir()
            now = time.time()
            
            cleaned_files = 0
            for filename in os.listdir(temp_dir):
                filepath = os.path.join(temp_dir, filename)
                # Clean files older than 1 hour that match our patterns
                if (filename.endswith(('.pptx', '.pdf', '.bin')) and 
                    os.path.isfile(filepath) and 
                    os.stat(filepath).st_mtime < now - 3600):
                    try:
                        os.unlink(filepath)
                        cleaned_files += 1
                    except:
                        pass
                        
            if cleaned_files > 0:
                logger.info(f"[CLEANUP] Removed {cleaned_files} old temporary files")
                
        except Exception as e:
            logger.error(f"[CLEANUP] Error in periodic cleanup: {e}")


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Manage startup and shutdown events"""
    # Startup
    cleanup_task = asyncio.create_task(periodic_cleanup())
    logger.info("[STARTUP] Started background cleanup task")
    
    yield
    
    # Shutdown
    cleanup_task.cancel()
    try:
        await cleanup_task
    except asyncio.CancelledError:
        logger.info("[SHUTDOWN] Background cleanup task cancelled")


app = FastAPI(title="Proposal Bot API", lifespan=lifespan)


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
        # Some Slack events might not have a user field
        user = event.get("user")
        channel = event.get("channel")
        if user and channel:
            asyncio.create_task(main_llm_loop(channel, user, event.get("text", ""), event))
        else:
            logger.warning(f"[SLACK_EVENT] Skipping event without user or channel: {event}")

    return JSONResponse({"status": "ok"})


@app.get("/health")
async def health():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}


@app.get("/metrics")
async def metrics():
    """Performance metrics endpoint for monitoring"""
    import psutil
    import os
    
    process = psutil.Process(os.getpid())
    memory_info = process.memory_info()
    
    # Get CPU count
    cpu_count = psutil.cpu_count()
    
    # Get current PDF conversion semaphore status
    from pdf_utils import _CONVERT_SEMAPHORE
    pdf_conversions_active = _CONVERT_SEMAPHORE._initial_value - _CONVERT_SEMAPHORE._value
    
    # Get user history size
    from llm import user_history, pending_location_additions
    
    return {
        "memory": {
            "rss_mb": round(memory_info.rss / 1024 / 1024, 2),
            "vms_mb": round(memory_info.vms / 1024 / 1024, 2),
        },
        "cpu": {
            "percent": process.cpu_percent(interval=0.1),
            "count": cpu_count,
        },
        "pdf_conversions": {
            "active": pdf_conversions_active,
            "max_concurrent": _CONVERT_SEMAPHORE._initial_value,
        },
        "cache_sizes": {
            "user_histories": len(user_history),
            "pending_locations": len(pending_location_additions),
            "templates_cached": len(config.get_location_mapping()),
        },
        "timestamp": datetime.now().isoformat()
    } 