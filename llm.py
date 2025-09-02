import json
import asyncio
from typing import Dict, Any
import os
from pathlib import Path
import aiohttp
from datetime import datetime, timedelta

import config
import db
from proposals import process_proposals
from slack_formatting import SlackResponses

user_history: Dict[str, list] = {}

# Global for temporary location uploads
temp_location_uploads: Dict[str, Dict[str, Any]] = {}


async def handle_edit_task_flow(channel: str, user_id: str, user_input: str, task_number: int, task_data: Dict[str, Any]) -> str:
    import textwrap

    def _load_mapping_config():
        return {
            "sales_people": {"Nourhan": {}, "Jason": {}, "James": {}, "Amr": {}},
            "location_mappings": {name: {} for name in config.available_location_names()},
            "videographers": {"James Sevillano": {}, "Jason Pieterse": {}, "Cesar Sierra": {}, "Amr Tamer": {}},
        }

    def _format_sales_people_hint(cfg):
        return ", ".join(cfg["sales_people"].keys())

    def _format_locations_hint(cfg):
        return ", ".join(cfg["location_mappings"].keys())

    mapping_cfg = _load_mapping_config()

    system_prompt = f"""
You are helping edit Task #{task_number}. The user said: "{user_input}"

Determine their intent and parse any field updates:
- If they want to save/confirm/done: action = 'save'
- If they want to cancel/stop/exit: action = 'cancel'
- If they want to see current values: action = 'view'
- If they're making changes: action = 'edit' and parse the field updates

Current task data: {json.dumps(task_data, indent=2)}

CRITICAL VALIDATION RULES - YOU MUST ENFORCE:

1. Sales Person - ONLY accept these exact values: {list(mapping_cfg.get('sales_people', {}).keys())}
   Auto-map: {_format_sales_people_hint(mapping_cfg)}
   Common: "Nour"→"Nourhan"
   If invalid: keep current value, tell user valid options

2. Location - ONLY accept these exact values: {list(mapping_cfg.get('location_mappings', {}).keys())}
   Valid: {_format_locations_hint(mapping_cfg)}
   If invalid: keep current value, tell user valid options

3. Videographer - ONLY accept these exact values: {list(mapping_cfg.get('videographers', {}).keys())}
   If invalid: keep current value, tell user valid options

Return JSON with: action, fields (only changed fields with VALID values), message.
In your message, explain any fields that couldn't be updated due to invalid values.
IMPORTANT: Use natural language in messages - say 'Sales Person' not 'sales_person', 'Location' not 'location'.
"""

    res = await config.openai_client.responses.create(
        model=config.OPENAI_MODEL,
        input=[{"role": "system", "content": system_prompt}],
        text={
            'format': {
                'type': 'json_schema',
                'name': 'edit_response',
                'strict': False,
                'schema': {
                    'type': 'object',
                    'properties': {
                        'action': {'type': 'string', 'enum': ['save', 'cancel', 'edit', 'view']},
                        'fields': {
                            'type': 'object',
                            'properties': {
                                'Brand': {'type': 'string'},
                                'Campaign Start Date': {'type': 'string'},
                                'Campaign End Date': {'type': 'string'},
                                'Reference Number': {'type': 'string'},
                                'Location': {'type': 'string'},
                                'Sales Person': {'type': 'string'},
                                'Status': {'type': 'string'},
                                'Filming Date': {'type': 'string'},
                                'Videographer': {'type': 'string'}
                            },
                            'additionalProperties': False
                        },
                        'message': {'type': 'string'}
                    },
                    'required': ['action'],
                    'additionalProperties': False
                }
            }
        },
        store=False
    )

    payload = {}
    try:
        if res.output and len(res.output) > 0 and hasattr(res.output[0], 'content'):
            content = res.output[0].content
            if content and len(content) > 0 and hasattr(content[-1], 'text'):
                payload = json.loads(content[-1].text)
    except Exception:
        payload = {"action": "view", "fields": {}, "message": "I couldn't parse your request. Showing current values."}

    action = payload.get("action", "view")
    message = payload.get("message", "")
    fields = payload.get("fields", {})

    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack(message or f"Action: {action}"))
    return action


async def _download_slack_file(file_info: Dict[str, Any]) -> Path:
    url = file_info.get("url_private_download") or file_info.get("url_private")
    if not url:
        raise ValueError("Missing file download URL")
    headers = {"Authorization": f"Bearer {config.SLACK_BOT_TOKEN}"}
    suffix = Path(file_info.get("name", "upload.bin")).suffix or ".bin"
    import tempfile
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.close()
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            resp.raise_for_status()
            with open(tmp.name, "wb") as f:
                f.write(await resp.read())
    return Path(tmp.name)


async def _persist_location_upload(location_key: str, pptx_path: Path, metadata_text: str) -> None:
    location_dir = config.TEMPLATES_DIR / location_key
    location_dir.mkdir(parents=True, exist_ok=True)
    target_pptx = location_dir / f"{location_key}.pptx"
    target_meta = location_dir / "metadata.txt"
    # Move/copy files
    import shutil
    shutil.move(str(pptx_path), str(target_pptx))
    target_meta.write_text(metadata_text, encoding="utf-8")


async def main_llm_loop(channel: str, user_id: str, user_input: str, slack_event: Dict[str, Any] = None):
    logger = config.logger

    available_names = ", ".join(config.available_location_names())
    
    # Get static locations for the prompt
    static_locations = []
    for key, meta in config.LOCATION_METADATA.items():
        if meta.get('display_type', '').lower() == 'static':
            static_locations.append(f"{key} ({meta.get('display_name', key)})")
    
    static_list = ", ".join(static_locations) if static_locations else "None"

    prompt = (
        f"You are a sales proposal bot for BackLite Media. You help create financial proposals for digital advertising locations.\n"
        f"You can handle SINGLE or MULTIPLE location proposals in one request.\n\n"
        f"PACKAGE TYPES:\n"
        f"1. SEPARATE PACKAGE (default): Each location gets its own proposal slide, multiple durations/rates allowed per location\n"
        f"2. COMBINED PACKAGE: All locations in ONE proposal slide, single duration per location, one combined net rate\n\n"
        
        f"AVAILABLE LOCATIONS: {available_names}\n"
        f"STATIC LOCATIONS (require production fee instead of upload fee): {static_list}\n\n"
        
        f"REQUIRED INFORMATION:\n"
        f"For SEPARATE PACKAGE (each location):\n"
        f"1. Location (must be one of the available locations)\n"
        f"2. Start Date\n"
        f"3. Duration Options (multiple allowed)\n"
        f"4. Net Rates for EACH duration\n"
        f"5. Production Fee (required ONLY for static locations, e.g., 'AED 5,000')\n"
        f"6. Client Name (required)\n"
        f"7. Submitted By (optional - defaults to current user)\n\n"
        f"For COMBINED PACKAGE:\n"
        f"1. All Locations\n"
        f"2. Start Date for EACH location\n"
        f"3. ONE Duration per location\n"
        f"4. ONE Combined Net Rate for entire package\n"
        f"5. Production Fee for EACH static location (if any)\n"
        f"6. Client Name (required)\n"
        f"7. Submitted By (optional - defaults to current user)\n\n"
        
        f"MULTIPLE PROPOSALS RULES:\n"
        f"- User can request proposals for multiple locations at once\n"
        f"- EACH location must have its own complete set of information\n"
        f"- EACH location must have matching number of durations and net rates\n"
        f"- Different locations can have different durations/rates\n"
        f"- Multiple proposals will be combined into a single PDF document\n\n"
        
        f"VALIDATION RULES:\n"
        f"- For EACH location, durations count MUST equal net rates count\n"
        f"- If a location has 3 duration options, it MUST have exactly 3 net rates\n"
        f"- DO NOT proceed until ALL locations have complete information\n"
        f"- Ask follow-up questions for any missing information\n"
        f"- ALWAYS ask for client name if not provided\n\n"
        
        f"PARSING EXAMPLES:\n"
        f"User: 'jawhara, oryx and triple crown special combined deal 2 mil, 2, 4 and 6 weeks respectively, 1st jan 2026, 2nd jan 2026 and 3rd'\n"
        f"Parse as: Combined package with Jawhara (2 weeks, Jan 1), Oryx (4 weeks, Jan 2), Triple Crown (6 weeks, Jan 3), total 2 million AED\n\n"
        
        f"SINGLE LOCATION EXAMPLE:\n"
        f"User: 'Proposal for landmark, Jan 1st, 2 weeks at 1.5M'\n"
        f"Bot confirms and generates one proposal\n\n"
        
        f"MULTIPLE LOCATIONS EXAMPLE:\n"
        f"User: 'I need proposals for landmark and gateway'\n"
        f"Bot: 'I'll help you create proposals for The Landmark and The Gateway. Let me get the details for each:\n\n"
        f"For THE LANDMARK:\n"
        f"- What's the campaign start date?\n"
        f"- What duration options do you want?\n"
        f"- What are the net rates for each duration?\n\n"
        f"For THE GATEWAY:\n"
        f"- What's the campaign start date?\n"
        f"- What duration options do you want?\n"
        f"- What are the net rates for each duration?'\n\n"
        
        f"COMBINED PACKAGE EXAMPLE:\n"
        f"User: 'I need a combined package for landmark, gateway, and oryx at 5 million total'\n"
        f"Bot: 'I'll create a combined package proposal. Let me confirm the details:\n\n"
        f"COMBINED PACKAGE:\n"
        f"- Locations: The Landmark, The Gateway, The Oryx\n"
        f"- Package Net Rate: AED 5,000,000\n\n"
        f"For each location, I need:\n"
        f"- Start date\n"
        f"- Duration (one per location for combined packages)\n\n"
        f"Please provide these details.'\n\n"
        
        f"ADDITIONAL FEATURES:\n"
        f"- You can ADD new locations interactively:\n"
        f"  1. User says 'add location <key>' and uploads PPTX file\n"
        f"  2. Bot collects all required metadata fields\n"
        f"  3. User confirms and location is saved\n"
        f"  4. PPTX is stored temporarily so user doesn't need to re-upload\n"
        f"- You can REFRESH templates to reload available locations\n"
        f"- You can LIST available locations\n"
        f"- You can EXPORT the backend database to Excel when user asks for 'excel backend' or similar\n"
        f"- You can GET STATISTICS about proposals generated\n"
        f"- You can EDIT tasks (for task management workflows)\n\n"
        
        f"IMPORTANT:\n"
        f"- Use get_separate_proposals for individual location proposals with multiple duration/rate options\n"
        f"- Use get_combined_proposal for special package deals with one total price\n"
        f"- For SEPARATE packages: each location gets its own proposal slide\n"
        f"- For COMBINED packages: all locations in ONE proposal slide with ONE net rate\n"
        f"- Single location always uses get_separate_proposals\n"
        f"- When user mentions 'combined deal' or 'special package' with total price, use get_combined_proposal\n"
        f"- Format all rates as 'AED X,XXX,XXX'\n"
        f"- Parse 'mil' or 'million' as 000,000 (e.g., '2 mil' = 'AED 2,000,000')\n"
        f"- Number of spots defaults to 1 if not specified\n"
        f"- For STATIC locations: MUST collect production fee (replaces upload fee)\n"
        f"- For DIGITAL locations: Use the pre-configured upload fee\n"
        f"- In COMBINED packages with both static and digital: collect production fees for static only\n"
        f"- ALWAYS collect client name - it's required for tracking"
    )

    history = user_history.get(user_id, [])
    history.append({"role": "user", "content": user_input})
    history = history[-10:]
    messages = [{"role": "developer", "content": prompt}] + history

    tools = [
        {
            "type": "function", 
            "name": "get_separate_proposals",
            "description": "Generate SEPARATE proposals - each location gets its own proposal slide with multiple duration/rate options. Returns individual PPTs and combined PDF.",
            "parameters": {
                "type": "object",
                "properties": {
                    "proposals": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "location": {"type": "string", "description": "The location name (e.g., landmark, gateway, oryx)"},
                                "start_date": {"type": "string", "description": "Start date for the campaign (e.g., 1st December 2025)"},
                                "durations": {
                                    "type": "array",
                                    "items": {"type": "string"},
                                    "description": "List of duration options (e.g., ['2 Weeks', '4 Weeks', '6 Weeks'])"
                                },
                                "net_rates": {
                                    "type": "array", 
                                    "items": {"type": "string"},
                                    "description": "List of net rates corresponding to each duration (e.g., ['AED 1,250,000', 'AED 2,300,000', 'AED 3,300,000'])"
                                },
                                "spots": {"type": "integer", "description": "Number of spots (default: 1)", "default": 1},
                                "production_fee": {"type": "string", "description": "Production fee for static locations (e.g., 'AED 5,000'). Required for static locations."}
                            },
                            "required": ["location", "start_date", "durations", "net_rates"]
                        },
                        "description": "Array of proposal objects. Each location can have multiple duration/rate options."
                    },
                    "client_name": {
                        "type": "string",
                        "description": "Name of the client (required)"
                    }
                },
                "required": ["proposals", "client_name"]
            }
        },
        {
            "type": "function", 
            "name": "get_combined_proposal",
            "description": "Generate COMBINED package proposal - all locations in ONE slide with single net rate. Use for special package deals.",
            "parameters": {
                "type": "object",
                "properties": {
                    "proposals": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "location": {"type": "string", "description": "The location name (e.g., landmark, gateway, oryx)"},
                                "start_date": {"type": "string", "description": "Start date for this location (e.g., 1st January 2026)"},
                                "duration": {"type": "string", "description": "Duration for this location (e.g., '2 Weeks')"},
                                "spots": {"type": "integer", "description": "Number of spots (default: 1)", "default": 1},
                                "production_fee": {"type": "string", "description": "Production fee for static locations (e.g., 'AED 5,000'). Required for static locations."}
                            },
                            "required": ["location", "start_date", "duration"]
                        },
                        "description": "Array of locations with their individual durations and start dates"
                    },
                    "combined_net_rate": {
                        "type": "string",
                        "description": "The total net rate for the entire package (e.g., 'AED 2,000,000')"
                    },
                    "client_name": {
                        "type": "string",
                        "description": "Name of the client (required)"
                    }
                },
                "required": ["proposals", "combined_net_rate", "client_name"]
            }
        },
        {"type": "function", "name": "refresh_templates", "parameters": {"type": "object", "properties": {}}},
        {"type": "function", "name": "edit_task_flow", "parameters": {"type": "object", "properties": {"task_number": {"type": "integer"}, "task_data": {"type": "object"}}, "required": ["task_number", "task_data"]}},
        {
            "type": "function", 
            "name": "add_location", 
            "description": "Add new location with metadata. Collect all required fields interactively.", 
            "parameters": {
                "type": "object", 
                "properties": {
                    "location_key": {"type": "string", "description": "Folder/key name to use (lowercase, no spaces)"},
                    "metadata": {
                        "type": "object",
                        "description": "Location metadata fields",
                        "properties": {
                            "location_name": {"type": "string", "description": "Full location name (e.g., 'The Oryx')"},
                            "display_name": {"type": "string", "description": "Display name for proposals"},
                            "series": {"type": "string", "description": "Series name (e.g., 'The Landmark Series')"},
                            "display_type": {"type": "string", "enum": ["Digital", "Static"], "description": "Display type"},
                            "height": {"type": "string", "description": "Height with units (e.g., '6m')"},
                            "width": {"type": "string", "description": "Width with units (e.g., '12m')"},
                            "number_of_faces": {"type": "integer", "description": "Number of display faces", "default": 1},
                            "spot_duration": {"type": "integer", "description": "Spot duration in seconds (for digital)"},
                            "loop_duration": {"type": "integer", "description": "Loop duration in seconds (for digital)"},
                            "sov": {"type": "string", "description": "Share of Voice percentage (e.g., '16.6%')"},
                            "upload_fee": {"type": "integer", "description": "Upload fee in AED"}
                        },
                        "required": ["location_name", "display_name", "display_type", "height", "width"]
                    },
                    "stage": {"type": "string", "enum": ["init", "collecting", "confirm", "save"], "description": "Current stage of the process"},
                    "confirm": {"type": "boolean", "description": "True only when user explicitly confirms"}
                }, 
                "required": ["location_key"]
            }
        },
        {"type": "function", "name": "list_locations", "description": "List the currently available locations to the user", "parameters": {"type": "object", "properties": {}}},
        {"type": "function", "name": "export_proposals_to_excel", "description": "Export all proposals from the backend database to Excel and send to user", "parameters": {"type": "object", "properties": {}}},
        {"type": "function", "name": "get_proposals_stats", "description": "Get summary statistics of proposals from the database", "parameters": {"type": "object", "properties": {}}}
    ]

    try:
        res = await config.openai_client.responses.create(model=config.OPENAI_MODEL, input=messages, tools=tools, tool_choice="auto")

        if not res.output or len(res.output) == 0:
            await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("I can help with proposals or add locations. Say 'add location'."))
            return

        msg = res.output[0]
        if msg.type == "function_call":
            if msg.name == "get_separate_proposals":
                args = json.loads(msg.arguments)
                proposals_data = args.get("proposals", [])
                client_name = args.get("client_name") or "Unknown Client"
                
                logger.info(f"[SEPARATE] Raw args: {args}")
                logger.info(f"[SEPARATE] Proposals data: {proposals_data}")
                logger.info(f"[SEPARATE] Client: {client_name}, User: {user_id}")

                if not proposals_data:
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("❌ **Error:** No proposals data provided"))
                    return
                
                result = await process_proposals(proposals_data, "separate", None, user_id, client_name)
            elif msg.name == "get_combined_proposal":
                args = json.loads(msg.arguments)
                proposals_data = args.get("proposals", [])
                combined_net_rate = args.get("combined_net_rate", None)
                client_name = args.get("client_name") or "Unknown Client"
                
                logger.info(f"[COMBINED] Raw args: {args}")
                logger.info(f"[COMBINED] Proposals data: {proposals_data}")
                logger.info(f"[COMBINED] Combined rate: {combined_net_rate}")
                logger.info(f"[COMBINED] Client: {client_name}, User: {user_id}")

                if not proposals_data:
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("❌ **Error:** No proposals data provided"))
                    return
                elif not combined_net_rate:
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("❌ **Error:** Combined package requires a combined net rate"))
                    return
                elif len(proposals_data) < 2:
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("❌ **Error:** Combined package requires at least 2 locations"))
                    return
                
                # Transform proposals data for combined package (add durations as list with single item)
                for proposal in proposals_data:
                    if "duration" in proposal:
                        proposal["durations"] = [proposal.pop("duration")]
                        logger.info(f"[COMBINED] Transformed proposal: {proposal}")
                        
                result = await process_proposals(proposals_data, "combined", combined_net_rate, user_id, client_name)
            
            # Handle result for both get_separate_proposals and get_combined_proposal
            if msg.name in ["get_separate_proposals", "get_combined_proposal"] and 'result' in locals():
                logger.info(f"[RESULT] Processing result: {result}")
                if result["success"]:
                    if result.get("is_combined"):
                        logger.info(f"[RESULT] Combined package - PDF: {result.get('pdf_filename')}")
                        await config.slack_client.files_upload_v2(channel=channel, file=result["pdf_path"], filename=result["pdf_filename"], initial_comment=config.markdown_to_slack(f"📦 **Combined Package Proposal**\n📍 Locations: {result['locations']}"))
                        try: os.unlink(result["pdf_path"])  # type: ignore
                        except: pass
                    elif result.get("is_single"):
                        logger.info(f"[RESULT] Single proposal - Location: {result.get('location')}")
                        # Send the formatted PPTX (with slides removed)
                        await config.slack_client.files_upload_v2(channel=channel, file=result["formatted_pptx_path"], filename=result["formatted_pptx_filename"], initial_comment=config.markdown_to_slack(f"📊 **PowerPoint Proposal (Formatted)**\n📍 Location: {result['location']}"))
                        # Send the PDF
                        await config.slack_client.files_upload_v2(channel=channel, file=result["pdf_path"], filename=result["pdf_filename"], initial_comment=config.markdown_to_slack(f"📄 **PDF Proposal**\n📍 Location: {result['location']}"))
                        try:
                            os.unlink(result["pptx_path"])  # type: ignore
                            os.unlink(result["formatted_pptx_path"])  # type: ignore
                            os.unlink(result["pdf_path"])  # type: ignore
                        except: pass
                    else:
                        logger.info(f"[RESULT] Multiple separate proposals - Count: {len(result.get('individual_files', []))}")
                        for f in result["individual_files"]:
                            await config.slack_client.files_upload_v2(channel=channel, file=f["path"], filename=f["filename"], initial_comment=config.markdown_to_slack(f"📊 **PowerPoint Proposal**\n📍 Location: {f['location']}"))
                        await config.slack_client.files_upload_v2(channel=channel, file=result["merged_pdf_path"], filename=result["merged_pdf_filename"], initial_comment=config.markdown_to_slack(f"📄 **Combined PDF**\n📍 All Locations: {result['locations']}"))
                        try:
                            for f in result["individual_files"]: os.unlink(f["path"])  # type: ignore
                            os.unlink(result["merged_pdf_path"])  # type: ignore
                        except: pass
                else:
                    logger.error(f"[RESULT] Error: {result.get('error')}")
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack(f"❌ **Error:** {result['error']}"))

            elif msg.name == "refresh_templates":
                config.refresh_templates()
                await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("✅ Templates refreshed successfully."))

            elif msg.name == "edit_task_flow":
                args = json.loads(msg.arguments)
                task_number = int(args.get("task_number"))
                task_data = args.get("task_data", {})
                await handle_edit_task_flow(channel, user_id, user_input, task_number, task_data)

            elif msg.name == "add_location":
                # Admin permission gate
                if not config.is_admin(user_id):
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("❌ **Error:** You need admin privileges to add locations."))
                    return

                args = json.loads(msg.arguments)
                location_key = args.get("location_key", "").strip().lower().replace(" ", "_")
                metadata = args.get("metadata", {})
                stage = args.get("stage", "init")
                confirm = bool(args.get("confirm", False))
                
                # Create unique session key for this user/location
                session_key = f"{user_id}_{location_key}"

                if not location_key:
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("Please provide a short key for the location (e.g., `oryx`)."))
                    return

                # Check if location already exists
                mapping = config.get_location_mapping()
                if location_key in mapping and stage == "init":
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack(f"⚠️ Location `{location_key}` already exists. Please provide a different key or say 'overwrite {location_key}' to replace it."))
                    return

                # Handle file uploads
                if slack_event and "files" in slack_event and stage in ["init", "collecting"]:
                    # Initialize session data if needed
                    if session_key not in temp_location_uploads:
                        temp_location_uploads[session_key] = {
                            "location_key": location_key,
                            "pptx_path": None,
                            "metadata": {},
                            "timestamp": datetime.now()
                        }
                    
                    for f in slack_event["files"]:
                        if f.get("filetype") == "pptx" or f.get("mimetype", "").endswith("powerpoint"):
                            # Download and save PPTX temporarily
                            pptx_temp = await _download_slack_file(f)
                            temp_location_uploads[session_key]["pptx_path"] = pptx_temp
                            await config.slack_client.chat_postMessage(
                                channel=channel, 
                                text=config.markdown_to_slack("✅ PowerPoint file received and saved temporarily.")
                            )

                # Different stages of the add location flow
                if stage == "init":
                    # Check if we have a PPTX file
                    if session_key not in temp_location_uploads or not temp_location_uploads[session_key].get("pptx_path"):
                        await config.slack_client.chat_postMessage(
                            channel=channel,
                            text=config.markdown_to_slack(
                                f"📎 **Adding location: {location_key}**\n\n"
                                f"Please upload the PowerPoint template file for this location."
                            )
                        )
                        return
                    
                    # Start collecting metadata
                    await config.slack_client.chat_postMessage(
                        channel=channel,
                        text=config.markdown_to_slack(
                            f"📋 **Location Metadata for {location_key}**\n\n"
                            f"I need the following information:\n"
                            f"• **Location Name**: Full name (e.g., 'The Oryx')\n"
                            f"• **Display Name**: Name for proposals\n"
                            f"• **Series**: Series name (e.g., 'The Landmark Series')\n"
                            f"• **Display Type**: Digital or Static\n"
                            f"• **Size**: Height x Width (e.g., '6m x 12m')\n"
                            f"• **Number of Faces**: How many display faces (default: 1)\n"
                            f"• **Upload Fee**: Fee in AED\n\n"
                            f"For digital displays also provide:\n"
                            f"• **Spot Duration**: Duration in seconds\n"
                            f"• **Loop Duration**: Total loop in seconds\n"
                            f"• **SOV**: Share of Voice percentage\n\n"
                            f"Please provide this information in your next message."
                        )
                    )
                    
                elif stage == "collecting":
                    # Update metadata from current call
                    if session_key in temp_location_uploads:
                        temp_location_uploads[session_key]["metadata"].update(metadata)
                        current_meta = temp_location_uploads[session_key]["metadata"]
                        
                        # Check what's missing
                        required = ["location_name", "display_name", "display_type", "height", "width"]
                        missing = [f for f in required if f not in current_meta or not current_meta[f]]
                        
                        if current_meta.get("display_type", "").lower() == "digital":
                            digital_required = ["spot_duration", "loop_duration", "sov"]
                            missing.extend([f for f in digital_required if f not in current_meta or not current_meta[f]])
                        
                        if missing:
                            missing_formatted = [f.replace("_", " ").title() for f in missing]
                            await config.slack_client.chat_postMessage(
                                channel=channel,
                                text=config.markdown_to_slack(
                                    f"⚠️ **Missing information:**\n\n" +
                                    "\n".join([f"• {field}" for field in missing_formatted]) +
                                    "\n\nPlease provide the missing fields."
                                )
                            )
                        else:
                            # All required fields collected, show summary
                            await config.slack_client.chat_postMessage(
                                channel=channel,
                                text=config.markdown_to_slack(
                                    f"📋 **Location Summary for {location_key}**\n\n" +
                                    f"• **Location Name**: {current_meta['location_name']}\n" +
                                    f"• **Display Name**: {current_meta['display_name']}\n" +
                                    f"• **Series**: {current_meta.get('series', 'Not specified')}\n" +
                                    f"• **Display Type**: {current_meta['display_type']}\n" +
                                    f"• **Size**: {current_meta['height']} x {current_meta['width']}\n" +
                                    f"• **Number of Faces**: {current_meta.get('number_of_faces', 1)}\n" +
                                    f"• **Upload Fee**: AED {current_meta.get('upload_fee', 'Not specified')}\n" +
                                    (f"• **Spot Duration**: {current_meta.get('spot_duration', 'N/A')} seconds\n" if current_meta.get('display_type', '').lower() == 'digital' else "") +
                                    (f"• **Loop Duration**: {current_meta.get('loop_duration', 'N/A')} seconds\n" if current_meta.get('display_type', '').lower() == 'digital' else "") +
                                    (f"• **SOV**: {current_meta.get('sov', 'N/A')}\n" if current_meta.get('display_type', '').lower() == 'digital' else "") +
                                    f"\n**Reply 'confirm' to save this location or 'edit' to make changes.**"
                                )
                            )
                    else:
                        await config.slack_client.chat_postMessage(
                            channel=channel,
                            text=config.markdown_to_slack("❌ Session expired. Please start over with 'add location'.")
                        )
                        
                elif stage == "confirm" and confirm:
                    if session_key in temp_location_uploads:
                        session_data = temp_location_uploads[session_key]
                        pptx_path = session_data["pptx_path"]
                        metadata_dict = session_data["metadata"]
                        
                        # Build metadata.txt content
                        metadata_lines = []
                        metadata_lines.append(f"Location Name: {metadata_dict.get('location_name', '')}")
                        metadata_lines.append(f"Display Name: {metadata_dict.get('display_name', '')}")
                        metadata_lines.append(f"Display Type: {metadata_dict.get('display_type', 'Digital')}")
                        metadata_lines.append(f"Number of Faces: {metadata_dict.get('number_of_faces', 1)}")
                        
                        if metadata_dict.get('display_type', '').lower() == 'digital':
                            metadata_lines.append(f"Spot Duration: {metadata_dict.get('spot_duration', 16)}")
                            metadata_lines.append(f"Loop Duration: {metadata_dict.get('loop_duration', 96)}")
                            metadata_lines.append(f"SOV: {metadata_dict.get('sov', '16.6%')}")
                        
                        metadata_lines.append(f"Upload Fee: {metadata_dict.get('upload_fee', 3000)}")
                        if metadata_dict.get('series'):
                            metadata_lines.append(f"Series: {metadata_dict.get('series')}")
                        metadata_lines.append(f"Height: {metadata_dict.get('height', '')}")
                        metadata_lines.append(f"Width: {metadata_dict.get('width', '')}")
                        
                        metadata_text = "\n".join(metadata_lines)
                        
                        # Save the location
                        await _persist_location_upload(location_key, pptx_path, metadata_text)
                        
                        # Clean up temporary data
                        del temp_location_uploads[session_key]
                        
                        # Refresh templates
                        config.refresh_templates()
                        
                        await config.slack_client.chat_postMessage(
                            channel=channel,
                            text=config.markdown_to_slack(
                                f"✅ **Successfully added location `{location_key}`**\n\n"
                                f"The location is now available for use in proposals."
                            )
                        )
                    else:
                        await config.slack_client.chat_postMessage(
                            channel=channel,
                            text=config.markdown_to_slack("❌ Session expired. Please start over with 'add location'.")
                        )
                        
                # Clean up old sessions (older than 30 minutes)
                cutoff = datetime.now() - timedelta(minutes=30)
                expired_sessions = [
                    key for key, data in temp_location_uploads.items()
                    if data.get("timestamp", datetime.now()) < cutoff
                ]
                for key in expired_sessions:
                    if "pptx_path" in temp_location_uploads[key] and temp_location_uploads[key]["pptx_path"]:
                        try:
                            os.unlink(temp_location_uploads[key]["pptx_path"])
                        except:
                            pass
                    del temp_location_uploads[key]

            elif msg.name == "list_locations":
                names = config.available_location_names()
                if not names:
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("📍 No locations available. Use **'add location'** to add one."))
                else:
                    listing = "\n".join(f"• {n}" for n in names)
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack(f"📍 **Current locations:**\n{listing}"))
            
            elif msg.name == "export_proposals_to_excel":
                # Admin permission gate
                if not config.is_admin(user_id):
                    await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("❌ **Error:** You need admin privileges to export the database."))
                    return
                    
                logger.info("[EXCEL_EXPORT] User requested Excel export")
                try:
                    excel_path = db.export_to_excel()
                    logger.info(f"[EXCEL_EXPORT] Created Excel file at {excel_path}")
                    
                    # Get file size for display
                    file_size = os.path.getsize(excel_path)
                    size_mb = file_size / (1024 * 1024)
                    
                    await config.slack_client.files_upload_v2(
                        channel=channel,
                        file=excel_path,
                        filename=f"proposals_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        initial_comment=config.markdown_to_slack(
                            f"📊 **Proposals Database Export**\n"
                            f"📅 Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                            f"📁 Size: {size_mb:.2f} MB"
                        )
                    )
                    
                    # Clean up temp file
                    try:
                        os.unlink(excel_path)
                    except:
                        pass
                        
                except Exception as e:
                    logger.error(f"[EXCEL_EXPORT] Error: {e}", exc_info=True)
                    await config.slack_client.chat_postMessage(
                        channel=channel,
                        text=config.markdown_to_slack("❌ **Error:** Failed to export database to Excel. Please try again.")
                    )
            
            elif msg.name == "get_proposals_stats":
                logger.info("[STATS] User requested proposals statistics")
                try:
                    stats = db.get_proposals_summary()
                    
                    # Format the statistics message
                    message = "📊 **Proposals Database Summary**\n\n"
                    message += f"**Total Proposals:** {stats['total_proposals']}\n\n"
                    
                    if stats['by_package_type']:
                        message += "**By Package Type:**\n"
                        for pkg_type, count in stats['by_package_type'].items():
                            message += f"• {pkg_type.title()}: {count}\n"
                        message += "\n"
                    
                    if stats['recent_proposals']:
                        message += "**Recent Proposals:**\n"
                        for proposal in stats['recent_proposals']:
                            date_str = datetime.fromisoformat(proposal['date']).strftime('%Y-%m-%d %H:%M')
                            message += f"• {proposal['client']} - {proposal['locations']} ({date_str})\n"
                    else:
                        message += "_No proposals generated yet._"
                    
                    await config.slack_client.chat_postMessage(
                        channel=channel,
                        text=config.markdown_to_slack(message)
                    )
                    
                except Exception as e:
                    logger.error(f"[STATS] Error: {e}", exc_info=True)
                    await config.slack_client.chat_postMessage(
                        channel=channel,
                        text=config.markdown_to_slack("❌ **Error:** Failed to retrieve statistics. Please try again.")
                    )

        else:
            reply = msg.content[-1].text if hasattr(msg, 'content') and msg.content else "How can I help you today?"
            # Format any markdown-style text from the LLM
            formatted_reply = reply
            # Ensure bullet points are properly formatted
            formatted_reply = formatted_reply.replace('\n- ', '\n• ')
            formatted_reply = formatted_reply.replace('\n* ', '\n• ')
            # Ensure headers are bolded
            import re
            formatted_reply = re.sub(r'^(For .+:)$', r'**\1**', formatted_reply, flags=re.MULTILINE)
            formatted_reply = re.sub(r'^([A-Z][A-Z\s]+:)$', r'**\1**', formatted_reply, flags=re.MULTILINE)
            await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack(formatted_reply))

        user_history[user_id] = history[-10:]

    except Exception as e:
        config.logger.error(f"LLM loop error: {e}", exc_info=True)
        await config.slack_client.chat_postMessage(channel=channel, text=config.markdown_to_slack("❌ **Error:** Something went wrong. Please try again.")) 