import os
import logging
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import json

from dotenv import load_dotenv
from slack_sdk.web.async_client import AsyncWebClient
from slack_sdk.signature import SignatureVerifier
from openai import AsyncOpenAI

# Load environment
load_dotenv()

# Base paths
BASE_DIR = Path(__file__).parent

# Logger (set up early so we can use it)
logger = logging.getLogger("proposal-bot")
logging.basicConfig(level=logging.INFO)

# Use /data/ in production, local paths in development
if os.path.exists("/data/"):
    # Production paths
    TEMPLATES_DIR = Path(os.getenv("TEMPLATES_DIR", "/data/templates"))
    HOS_CONFIG_FILE = Path("/data/hos_config.json")
    logger.info("[STARTUP] Running in PRODUCTION mode - using /data/ paths")
else:
    # Development paths
    TEMPLATES_DIR = Path(os.getenv("TEMPLATES_DIR", str(BASE_DIR / "data" / "templates")))
    HOS_CONFIG_FILE = BASE_DIR / "data" / "hos_config.json"
    logger.info("[STARTUP] Running in DEVELOPMENT mode - using local paths")

logger.info(f"[STARTUP] Templates directory: {TEMPLATES_DIR}")
logger.info(f"[STARTUP] HOS config file: {HOS_CONFIG_FILE}")

# Clients and config
SLACK_BOT_TOKEN = os.getenv("SLACK_BOT_TOKEN", "")
SLACK_SIGNING_SECRET = os.getenv("SLACK_SIGNING_SECRET", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1")

slack_client = AsyncWebClient(token=SLACK_BOT_TOKEN)
signature_verifier = SignatureVerifier(SLACK_SIGNING_SECRET)
openai_client = AsyncOpenAI(api_key=OPENAI_API_KEY)

# Dynamic data populated from templates directory
UPLOAD_FEES_MAPPING: Dict[str, int] = {}
LOCATION_DETAILS: Dict[str, str] = {}
LOCATION_METADATA: Dict[str, Dict[str, object]] = {}

# Cache for templates
_MAPPING_CACHE: Optional[Dict[str, str]] = None
_DISPLAY_CACHE: Optional[List[str]] = None

# HOS config
_HOS_CONFIG: Dict[str, Dict[str, Dict[str, object]]] = {}


def load_hos_config() -> None:
    global _HOS_CONFIG
    try:
        logger.info(f"[HOS_CONFIG] Looking for config at: {HOS_CONFIG_FILE}")
        if HOS_CONFIG_FILE.exists():
            logger.info(f"[HOS_CONFIG] Found config at: {HOS_CONFIG_FILE}")
            _HOS_CONFIG = json.loads(HOS_CONFIG_FILE.read_text(encoding="utf-8"))
            logger.info(f"[HOS_CONFIG] Loaded config: {list(_HOS_CONFIG.keys())}")
        else:
            logger.warning(f"[HOS_CONFIG] Config file not found at: {HOS_CONFIG_FILE}")
            _HOS_CONFIG = {}
    except Exception as e:
        logger.warning(f"Failed to load hos_config.json: {e}")
        _HOS_CONFIG = {}


def can_manage_locations(slack_user_id: str) -> bool:
    if not _HOS_CONFIG:
        load_hos_config()
    groups = _HOS_CONFIG.get("permissions", {}).get("manage_locations", [])
    allowed_ids = set()
    for group in groups:
        members = _HOS_CONFIG.get(group, {})
        for _, info in members.items():
            if info.get("active") and info.get("slack_user_id"):
                allowed_ids.add(info["slack_user_id"])
    return slack_user_id in allowed_ids


def is_admin(slack_user_id: str) -> bool:
    """Check if user has admin privileges (can add locations and export database)."""
    if not _HOS_CONFIG:
        logger.info(f"[ADMIN_CHECK] Loading HOS config")
        load_hos_config()
    
    # Admin users are those in the 'admin' group
    admin_members = _HOS_CONFIG.get("admin", {})
    logger.info(f"[ADMIN_CHECK] Checking if {slack_user_id} is admin")
    logger.info(f"[ADMIN_CHECK] Admin members: {list(admin_members.keys())}")
    
    for name, info in admin_members.items():
        user_id = info.get("slack_user_id")
        is_active = info.get("active")
        logger.info(f"[ADMIN_CHECK] Checking {name}: user_id={user_id}, active={is_active}")
        
        if is_active and user_id == slack_user_id:
            logger.info(f"[ADMIN_CHECK] User {slack_user_id} is admin!")
            return True
    
    logger.info(f"[ADMIN_CHECK] User {slack_user_id} is NOT admin")
    return False


def _normalize_key(name: str) -> str:
    return os.path.splitext(name)[0].strip().lower()


def _parse_metadata_file(folder: Path) -> Dict[str, object]:
    meta: Dict[str, object] = {}
    path = folder / "metadata.txt"
    if not path.exists():
        return meta

    try:
        for line in path.read_text(encoding="utf-8").splitlines():
            if ":" not in line:
                continue
            k, v = line.split(":", 1)
            key = k.strip().lower().replace(" ", "_")
            val = v.strip()
            meta[key] = val
    except Exception as e:
        logger.warning(f"Failed to parse metadata at {path}: {e}")
        return meta

    # Parse specific fields
    upload_fee: Optional[int] = None
    uf = str(meta.get("upload_fee", "")).replace(",", "").strip()
    if uf.isdigit():
        upload_fee = int(uf)

    sov_text = str(meta.get("sov", "16.6%"))
    try:
        base_sov = float(sov_text.replace("%", "").strip())
    except Exception:
        base_sov = 16.6

    display_name = str(meta.get("display_name", meta.get("location_name", ""))).strip()
    
    # Parse new fields
    series = str(meta.get("series", "")).strip()
    height = str(meta.get("height", "")).strip()
    width = str(meta.get("width", "")).strip()
    number_of_faces = 1
    if meta.get("number_of_faces"):
        try:
            number_of_faces = int(meta.get("number_of_faces"))
        except:
            number_of_faces = 1
    
    display_type = str(meta.get("display_type", "Digital")).strip()
    spot_duration = 16
    if meta.get("spot_duration"):
        try:
            spot_duration = int(meta.get("spot_duration"))
        except:
            spot_duration = 16
    
    loop_duration = 96
    if meta.get("loop_duration"):
        try:
            loop_duration = int(meta.get("loop_duration"))
        except:
            loop_duration = 96

    return {
        "display_name": display_name,
        "upload_fee": upload_fee,
        "sov": f"{base_sov}%",
        "series": series,
        "height": height,
        "width": width,
        "number_of_faces": number_of_faces,
        "display_type": display_type,
        "spot_duration": spot_duration,
        "loop_duration": loop_duration,
        "folder": str(folder.name),
    }


def _discover_templates() -> Tuple[Dict[str, str], List[str]]:
    logger.info(f"[DISCOVER] Starting template discovery in '{TEMPLATES_DIR}'")
    key_to_relpath: Dict[str, str] = {}
    display_names: List[str] = []

    UPLOAD_FEES_MAPPING.clear()
    LOCATION_DETAILS.clear()
    LOCATION_METADATA.clear()

    if not TEMPLATES_DIR.exists():
        logger.warning(f"[DISCOVER] Templates directory does not exist: '{TEMPLATES_DIR}'")
        return key_to_relpath, display_names

    logger.info(f"[DISCOVER] Scanning for PPTX files in '{TEMPLATES_DIR}'")
    for pptx_path in TEMPLATES_DIR.rglob("*.pptx"):
        try:
            rel_path = pptx_path.relative_to(TEMPLATES_DIR)
        except Exception:
            rel_path = pptx_path
        key = _normalize_key(pptx_path.stem)
        key_to_relpath[key] = str(rel_path)
        logger.info(f"[DISCOVER] Found template: '{pptx_path}' -> key: '{key}'")

        meta = _parse_metadata_file(pptx_path.parent)
        logger.info(f"[DISCOVER] Metadata for '{key}': {meta}")
        
        display_name = meta.get("display_name") or pptx_path.stem
        description = meta.get("description") or f"{pptx_path.stem} - Digital Display - 1 Spot - 16 Seconds - 16.6% SOV - Total Loop is 6 spots"
        upload_fee = meta.get("upload_fee") or 3000
        base_sov = meta.get("base_sov_percent") or 16.6

        display_names.append(str(display_name))
        LOCATION_DETAILS[key] = str(description)
        
        # Use upload fee from metadata or default
        if meta.get("upload_fee") is not None:
            UPLOAD_FEES_MAPPING[key] = int(meta.get("upload_fee"))
        else:
            UPLOAD_FEES_MAPPING[key] = 3000
        
        # Store all metadata fields
        LOCATION_METADATA[key] = meta
        LOCATION_METADATA[key]["pptx_rel_path"] = str(rel_path)

    logger.info(f"[DISCOVER] Discovery complete. Found {len(key_to_relpath)} templates")
    logger.info(f"[DISCOVER] Location keys: {list(key_to_relpath.keys())}")
    logger.info(f"[DISCOVER] Display names: {display_names}")
    return key_to_relpath, display_names


def refresh_templates() -> None:
    global _MAPPING_CACHE, _DISPLAY_CACHE
    logger.info("[REFRESH] Refreshing templates cache")
    mapping, names = _discover_templates()
    _MAPPING_CACHE = mapping
    _DISPLAY_CACHE = names
    logger.info(f"[REFRESH] Templates cache refreshed: {len(mapping)} templates")
    logger.info(f"[REFRESH] Cached mapping: {mapping}")
    logger.info(f"[REFRESH] Upload fees: {UPLOAD_FEES_MAPPING}")
    logger.info(f"[REFRESH] Location metadata: {LOCATION_METADATA}")


def get_location_mapping() -> Dict[str, str]:
    global _MAPPING_CACHE
    if _MAPPING_CACHE is None:
        logger.info("[GET_MAPPING] Cache is empty, refreshing templates")
        refresh_templates()
    else:
        logger.info(f"[GET_MAPPING] Using cached mapping with {len(_MAPPING_CACHE)} entries")
    return _MAPPING_CACHE or {}


def available_location_names() -> List[str]:
    global _DISPLAY_CACHE
    if _DISPLAY_CACHE is None:
        refresh_templates()
    return _DISPLAY_CACHE or []


def get_location_key_from_display_name(display_name: str) -> Optional[str]:
    """Convert a display name back to its location key."""
    # Ensure metadata is loaded
    if not LOCATION_METADATA:
        refresh_templates()
    
    # Normalize the input
    display_name_lower = display_name.lower().strip()
    
    # First try exact match
    for key, meta in LOCATION_METADATA.items():
        if meta.get('display_name', '').lower() == display_name_lower:
            return key
    
    # Then try partial matches
    for key, meta in LOCATION_METADATA.items():
        meta_display = meta.get('display_name', '').lower()
        if display_name_lower in meta_display or meta_display in display_name_lower:
            return key
    
    # Also check if the display name is actually a key
    for key in LOCATION_METADATA.keys():
        if key == display_name_lower:
            return key
    
    return None


def markdown_to_slack(text: str) -> str:
    """Convert markdown formatting to Slack's mrkdwn format.
    
    Handles common markdown patterns and converts them to Slack equivalents:
    - **bold** -> *bold*
    - *italic* -> _italic_
    - ***bold italic*** -> *_bold italic_*
    - `code` -> `code`
    - ```code block``` -> ```code block```
    - [link](url) -> <url|link>
    - # Header -> *Header*
    - ## Subheader -> *Subheader*
    - - bullet -> • bullet
    - 1. numbered -> 1. numbered
    """
    import re
    
    # Convert headers
    text = re.sub(r'^### (.+)$', r'*\1*', text, flags=re.MULTILINE)
    text = re.sub(r'^## (.+)$', r'*\1*', text, flags=re.MULTILINE)
    text = re.sub(r'^# (.+)$', r'*\1*', text, flags=re.MULTILINE)
    
    # Convert bold italic (must come before bold/italic)
    text = re.sub(r'\*\*\*(.+?)\*\*\*', r'*_\1_*', text)
    
    # Convert bold
    text = re.sub(r'\*\*(.+?)\*\*', r'*\1*', text)
    
    # Convert italic (but not already converted bold)
    text = re.sub(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)', r'_\1_', text)
    
    # Convert links
    text = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<\2|\1>', text)
    
    # Convert bullet points
    text = re.sub(r'^- ', '• ', text, flags=re.MULTILINE)
    text = re.sub(r'^\* ', '• ', text, flags=re.MULTILINE)
    
    # Ensure proper line breaks for lists
    text = re.sub(r'\n(?=\d+\.|•)', '\n', text)
    
    return text 