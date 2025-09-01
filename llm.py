import json
import asyncio
from typing import Dict, Any
import os
from pathlib import Path
import aiohttp

import config
from proposals import process_proposals

user_history: Dict[str, list] = {}


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

    await config.slack_client.chat_postMessage(channel=channel, text=message or f"Action: {action}")
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

    prompt = (
        f"You are a sales proposal bot for BackLite Media. You help create financial proposals for digital advertising locations.\n"
        f"You can handle SINGLE or MULTIPLE location proposals in one request.\n\n"
        f"PACKAGE TYPES:\n"
        f"1. SEPARATE PACKAGE (default): Each location gets its own proposal slide, multiple durations/rates allowed per location\n"
        f"2. COMBINED PACKAGE: All locations in ONE proposal slide, single duration per location, one combined net rate\n\n"
        
        f"AVAILABLE LOCATIONS: {available_names}\n\n"
        
        f"REQUIRED INFORMATION:\n"
        f"For SEPARATE PACKAGE (each location):\n"
        f"1. Location (must be one of the available locations)\n"
        f"2. Start Date\n"
        f"3. Duration Options (multiple allowed)\n"
        f"4. Net Rates for EACH duration\n"
        f"5. Client Name (required)\n"
        f"6. Submitted By (optional - defaults to current user)\n\n"
        f"For COMBINED PACKAGE:\n"
        f"1. All Locations\n"
        f"2. Start Date for EACH location\n"
        f"3. ONE Duration per location\n"
        f"4. ONE Combined Net Rate for entire package\n"
        f"5. Client Name (required)\n"
        f"6. Submitted By (optional - defaults to current user)\n\n"
        
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
        f"- You can ADD new locations by collecting a PPTX and metadata.txt\n"
        f"- You can REFRESH templates to reload available locations\n"
        f"- You can LIST available locations\n"
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
                                "spots": {"type": "integer", "description": "Number of spots (default: 1)", "default": 1}
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
                                "spots": {"type": "integer", "description": "Number of spots (default: 1)", "default": 1}
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
        {"type": "function", "name": "add_location", "description": "Conversationally add new location: gather location_key, files (pptx, metadata), dedupe, confirm, then persist and refresh", "parameters": {"type": "object", "properties": {"location_key": {"type": "string", "description": "Folder/key name to use (lowercase, no spaces)"}, "confirm": {"type": "boolean", "description": "True only when user explicitly confirms"}}, "required": ["location_key"]}},
        {"type": "function", "name": "list_locations", "description": "List the currently available locations to the user", "parameters": {"type": "object", "properties": {}}}
    ]

    try:
        res = await config.openai_client.responses.create(model=config.OPENAI_MODEL, input=messages, tools=tools, tool_choice="auto")

        if not res.output or len(res.output) == 0:
            await config.slack_client.chat_postMessage(channel=channel, text="I can help with proposals or add locations. Say 'add location'.")
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
                    await config.slack_client.chat_postMessage(channel=channel, text="❌ No proposals data provided")
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
                    await config.slack_client.chat_postMessage(channel=channel, text="❌ No proposals data provided")
                    return
                elif not combined_net_rate:
                    await config.slack_client.chat_postMessage(channel=channel, text="❌ Combined package requires a combined net rate")
                    return
                elif len(proposals_data) < 2:
                    await config.slack_client.chat_postMessage(channel=channel, text="❌ Combined package requires at least 2 locations")
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
                        await config.slack_client.files_upload_v2(channel=channel, file=result["pdf_path"], filename=result["pdf_filename"], initial_comment=f"Combined package for {result['locations']}")
                        try: os.unlink(result["pdf_path"])  # type: ignore
                        except: pass
                    elif result.get("is_single"):
                        logger.info(f"[RESULT] Single proposal - Location: {result.get('location')}")
                        await config.slack_client.files_upload_v2(channel=channel, file=result["pptx_path"], filename=result["pptx_filename"], initial_comment=f"PPT for {result['location']}")
                        await config.slack_client.files_upload_v2(channel=channel, file=result["pdf_path"], filename=result["pdf_filename"], initial_comment=f"PDF for {result['location']}")
                        try:
                            os.unlink(result["pptx_path"])  # type: ignore
                            os.unlink(result["pdf_path"])  # type: ignore
                        except: pass
                    else:
                        logger.info(f"[RESULT] Multiple separate proposals - Count: {len(result.get('individual_files', []))}")
                        for f in result["individual_files"]:
                            await config.slack_client.files_upload_v2(channel=channel, file=f["path"], filename=f["filename"], initial_comment=f"PPT for {f['location']}")
                        await config.slack_client.files_upload_v2(channel=channel, file=result["merged_pdf_path"], filename=result["merged_pdf_filename"], initial_comment=f"Combined PDF for {result['locations']}")
                        try:
                            for f in result["individual_files"]: os.unlink(f["path"])  # type: ignore
                            os.unlink(result["merged_pdf_path"])  # type: ignore
                        except: pass
                else:
                    logger.error(f"[RESULT] Error: {result.get('error')}")
                    await config.slack_client.chat_postMessage(channel=channel, text=f"❌ {result['error']}")

            elif msg.name == "refresh_templates":
                config.refresh_templates()
                await config.slack_client.chat_postMessage(channel=channel, text="🔄 Templates refreshed.")

            elif msg.name == "edit_task_flow":
                args = json.loads(msg.arguments)
                task_number = int(args.get("task_number"))
                task_data = args.get("task_data", {})
                await handle_edit_task_flow(channel, user_id, user_input, task_number, task_data)

            elif msg.name == "add_location":
                # Permission gate
                if not config.can_manage_locations(user_id):
                    await config.slack_client.chat_postMessage(channel=channel, text="❌ You are not authorized to manage locations.")
                    return

                args = json.loads(msg.arguments)
                location_key = args.get("location_key", "").strip().lower().replace(" ", "-")
                confirm = bool(args.get("confirm", False))

                if not location_key:
                    await config.slack_client.chat_postMessage(channel=channel, text="Please provide a short key for the location (e.g., 'oryx').")
                    return

                mapping = config.get_location_mapping()
                if location_key in mapping:
                    await config.slack_client.chat_postMessage(channel=channel, text=f"I already have a location '{location_key}'. Reply 'confirm overwrite' to replace it or provide a different key.")
                    return

                pptx_temp = None
                metadata_text = None
                if slack_event and "files" in slack_event:
                    for f in slack_event["files"]:
                        if f.get("filetype") == "pptx" or f.get("mimetype", "").endswith("powerpoint"):
                            pptx_temp = await _download_slack_file(f)
                        elif f.get("filetype") in ("txt",) or f.get("mimetype", "").startswith("text/"):
                            txt_path = await _download_slack_file(f)
                            metadata_text = Path(txt_path).read_text(encoding="utf-8")
                            try: os.unlink(txt_path)
                            except: pass

                if not pptx_temp or not metadata_text:
                    await config.slack_client.chat_postMessage(channel=channel, text="Upload both the PPTX and a metadata.txt in the same message, then say 'add location <key>'.")
                    return

                if not confirm:
                    await config.slack_client.chat_postMessage(channel=channel, text=f"Ready to add '{location_key}'. Reply 'confirm' to proceed or 'cancel'.")
                    return

                await _persist_location_upload(location_key, pptx_temp, metadata_text)
                config.refresh_templates()
                await config.slack_client.chat_postMessage(channel=channel, text=f"✅ Added location '{location_key}'. You can use it in proposals now.")

            elif msg.name == "list_locations":
                names = config.available_location_names()
                if not names:
                    await config.slack_client.chat_postMessage(channel=channel, text="No locations available. Use 'add location' to add one.")
                else:
                    listing = "\n".join(f"• {n}" for n in names)
                    await config.slack_client.chat_postMessage(channel=channel, text=f"Current locations:\n{listing}")

        else:
            reply = msg.content[-1].text if hasattr(msg, 'content') and msg.content else "How can I help you today?"
            await config.slack_client.chat_postMessage(channel=channel, text=reply)

        user_history[user_id] = history[-10:]

    except Exception as e:
        config.logger.error(f"LLM loop error: {e}", exc_info=True)
        await config.slack_client.chat_postMessage(channel=channel, text="❌ Something went wrong. Please try again.") 