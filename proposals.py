import os
import asyncio
from pathlib import Path
from typing import Dict, Any, List, Tuple

from pptx import Presentation

import config
import db
from pptx_utils import create_financial_proposal_slide, create_combined_financial_proposal_slide
from pdf_utils import convert_pptx_to_pdf, merge_pdfs, remove_slides_and_convert_to_pdf


def _template_path_for_key(key: str) -> Path:
    mapping = config.get_location_mapping()
    filename = mapping.get(key)
    if not filename:
        raise FileNotFoundError(f"Unknown location '{key}'. Available: {', '.join(config.available_location_names())}")
    return config.TEMPLATES_DIR / filename


def create_proposal_with_template(source_path: str, financial_data: dict) -> Tuple[str, List[str], List[str]]:
    import tempfile

    pres = Presentation(source_path)
    insert_position = max(len(pres.slides) - 1, 0)
    slide_width = pres.slide_width
    slide_height = pres.slide_height

    blank_layout = pres.slide_layouts[6] if len(pres.slide_layouts) > 6 else pres.slide_layouts[0]
    financial_slide = pres.slides.add_slide(blank_layout)

    vat_amounts, total_amounts = create_financial_proposal_slide(financial_slide, financial_data, slide_width, slide_height)

    if len(pres.slides) > 1 and insert_position < len(pres.slides) - 1:
        xml_slides = pres.slides._sldIdLst
        new_slide_element = xml_slides[-1]
        xml_slides.remove(new_slide_element)
        xml_slides.insert(insert_position, new_slide_element)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    pres.save(tmp.name)
    return tmp.name, vat_amounts, total_amounts


def create_combined_proposal_with_template(source_path: str, proposals_data: list, combined_net_rate: str) -> Tuple[str, str]:
    import tempfile

    pres = Presentation(source_path)
    insert_position = max(len(pres.slides) - 1, 0)
    slide_width = pres.slide_width
    slide_height = pres.slide_height

    layout = pres.slide_layouts[0]
    financial_slide = pres.slides.add_slide(layout)

    for shape in list(financial_slide.shapes):
        if hasattr(shape, "text_frame"):
            shape.text_frame.clear()

    total_combined = create_combined_financial_proposal_slide(financial_slide, proposals_data, combined_net_rate, slide_width, slide_height)

    xml_slides = pres.slides._sldIdLst
    slides_list = list(xml_slides)
    new_slide_element = slides_list[-1]
    xml_slides.remove(new_slide_element)
    xml_slides.insert(insert_position, new_slide_element)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    pres.save(tmp.name)
    return tmp.name, total_combined


async def process_combined_package(proposals_data: list, combined_net_rate: str, submitted_by: str, client_name: str) -> Dict[str, Any]:
    validated_proposals = []
    for idx, proposal in enumerate(proposals_data):
        location = proposal.get("location", "").lower().strip()
        start_date = proposal.get("start_date", "1st December 2025")
        durations = proposal.get("durations", [])
        spots = int(proposal.get("spots", 1))

        mapping = config.get_location_mapping()
        matched_key = None
        for key in mapping.keys():
            if key in location or location in key:
                matched_key = key
                break
        if not matched_key:
            return {"success": False, "error": f"Unknown location '{location}' in proposal {idx + 1}"}
        if not durations:
            return {"success": False, "error": f"No duration specified for {matched_key}"}

        validated_proposals.append({
            "location": matched_key,
            "start_date": start_date,
            "durations": durations,
            "spots": spots,
            "filename": mapping[matched_key],
        })

    loop = asyncio.get_event_loop()
    pdf_files: List[str] = []

    for idx, proposal in enumerate(validated_proposals):
        src = config.TEMPLATES_DIR / proposal["filename"]
        if not src.exists():
            return {"success": False, "error": f"{proposal['filename']} not found"}

        if idx == len(validated_proposals) - 1:
            pptx_file, total_combined = await loop.run_in_executor(
                None, create_combined_proposal_with_template, str(src), validated_proposals, combined_net_rate
            )
        else:
            pptx_file = str(src)
            total_combined = None

        remove_first = False
        remove_last = False
        if idx == 0:
            remove_last = True
        elif idx < len(validated_proposals) - 1:
            remove_first = True
            remove_last = True
        else:
            remove_first = True

        pdf_file = await remove_slides_and_convert_to_pdf(pptx_file, remove_first, remove_last)
        pdf_files.append(pdf_file)

        if idx == len(validated_proposals) - 1:
            try:
                os.unlink(pptx_file)
            except:
                pass

    merged_pdf = await loop.run_in_executor(None, merge_pdfs, pdf_files)
    for pdf_file in pdf_files:
        try:
            os.unlink(pdf_file)
        except:
            pass

    locations_str = ", ".join([p["location"].title() for p in validated_proposals])

    if total_combined is None:
        municipality_fee = 520
        total_upload_fees = sum(config.UPLOAD_FEES_MAPPING.get(p["location"].lower(), 3000) for p in validated_proposals)
        net_rate_numeric = float(combined_net_rate.replace("AED", "").replace(",", "").strip())
        subtotal = net_rate_numeric + total_upload_fees + municipality_fee
        vat = subtotal * 0.05
        total_combined = f"AED {subtotal + vat:,.0f}"

    db.log_proposal(
        submitted_by=submitted_by,
        client_name=client_name,
        package_type="combined",
        locations=locations_str,
        total_amount=total_combined,
    )

    return {
        "success": True,
        "is_combined": True,
        "pptx_path": None,
        "pdf_path": merged_pdf,
        "locations": locations_str,
        "pdf_filename": f"Combined_Package_{len(validated_proposals)}_Locations.pdf",
    }


async def process_proposals(
    proposals_data: list,
    package_type: str = "separate",
    combined_net_rate: str = None,
    submitted_by: str = "",
    client_name: str = "",
) -> Dict[str, Any]:
    if not proposals_data:
        return {"success": False, "error": "No proposals provided"}

    is_single = len(proposals_data) == 1 and package_type != "combined"

    if package_type == "combined" and len(proposals_data) > 1:
        return await process_combined_package(proposals_data, combined_net_rate, submitted_by, client_name)

    individual_files = []
    pdf_files = []
    locations = []

    loop = asyncio.get_event_loop()

    for idx, proposal in enumerate(proposals_data):
        location = proposal.get("location", "").lower().strip()
        start_date = proposal.get("start_date", "1st December 2025")
        durations = proposal.get("durations", [])
        net_rates = proposal.get("net_rates", [])
        spots = int(proposal.get("spots", 1))

        mapping = config.get_location_mapping()
        matched_key = None
        for key in mapping.keys():
            if key in location or location in key:
                matched_key = key
                break
        if not matched_key:
            return {"success": False, "error": f"Unknown location '{location}' in proposal {idx + 1}"}

        if len(durations) != len(net_rates):
            return {"success": False, "error": f"Mismatched durations and rates for {matched_key} - {len(durations)} durations but {len(net_rates)} rates"}
        if not durations:
            return {"success": False, "error": f"No duration specified for {matched_key}"}

        src = config.TEMPLATES_DIR / mapping[matched_key]
        if not src.exists():
            return {"success": False, "error": f"{mapping[matched_key]} not found"}

        financial_data = {
            "location": matched_key,
            "start_date": start_date,
            "durations": durations,
            "net_rates": net_rates,
            "spots": spots,
        }

        pptx_file, vat_amounts, total_amounts = await loop.run_in_executor(None, create_proposal_with_template, str(src), financial_data)

        individual_files.append({
            "path": pptx_file,
            "location": matched_key.title(),
            "filename": f"{matched_key.title()}_Proposal.pptx",
            "totals": total_amounts,
        })

        locations.append(matched_key.title())

        if is_single:
            pdf_file = await loop.run_in_executor(None, convert_pptx_to_pdf, pptx_file)
            individual_files[0]["pdf_path"] = pdf_file
            individual_files[0]["pdf_filename"] = f"{matched_key.title()}_Proposal.pdf"
        else:
            remove_first = False
            remove_last = False
            if idx == 0:
                remove_last = True
            elif idx < len(proposals_data) - 1:
                remove_first = True
                remove_last = True
            else:
                remove_first = True
            pdf_file = await remove_slides_and_convert_to_pdf(pptx_file, remove_first, remove_last)
            pdf_files.append(pdf_file)

    if is_single:
        totals = individual_files[0].get("totals", [])
        total_str = totals[0] if totals else "AED 0"
        db.log_proposal(
            submitted_by=submitted_by,
            client_name=client_name,
            package_type="single",
            locations=individual_files[0]["location"],
            total_amount=total_str,
        )
        return {
            "success": True,
            "is_single": True,
            "pptx_path": individual_files[0]["path"],
            "pdf_path": individual_files[0]["pdf_path"],
            "location": individual_files[0]["location"],
            "pptx_filename": individual_files[0]["filename"],
            "pdf_filename": individual_files[0]["pdf_filename"],
        }

    merged_pdf = await loop.run_in_executor(None, merge_pdfs, pdf_files)
    for pdf_file in pdf_files:
        try:
            os.unlink(pdf_file)
        except:
            pass

    first_totals = [files.get("totals", ["AED 0"])[0] for files in individual_files]
    summary_total = ", ".join(first_totals)
    db.log_proposal(
        submitted_by=submitted_by,
        client_name=client_name,
        package_type="separate",
        locations=", ".join(locations),
        total_amount=summary_total,
    )

    return {
        "success": True,
        "is_single": False,
        "individual_files": individual_files,
        "merged_pdf_path": merged_pdf,
        "locations": ", ".join(locations),
        "merged_pdf_filename": f"Combined_Proposal_{len(locations)}_Locations.pdf",
    } 