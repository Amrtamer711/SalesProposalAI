"""Utilities for extracting specific slides from PowerPoint to PDF without quality loss"""

import os
import tempfile
import asyncio
from typing import Tuple

from pdf_utils import convert_pptx_to_pdf, _CONVERT_SEMAPHORE
from pypdf import PdfReader, PdfWriter
import config


async def extract_first_and_last_slide_as_pdfs(file_path: str) -> Tuple[str, str]:
    """
    Extract first and last slides as separate PDFs.
    Can handle both PowerPoint files (converts to PDF first) and existing PDF files.
    
    Args:
        file_path: Path to either a .pptx or .pdf file
        
    Returns: (intro_pdf_path, outro_pdf_path)
    """
    logger = config.logger
    logger.info(f"[EXTRACT_SLIDES] Extracting first and last slides from: {file_path}")
    
    async with _CONVERT_SEMAPHORE:
        # Check if it's already a PDF
        if file_path.lower().endswith('.pdf'):
            logger.info(f"[EXTRACT_SLIDES] ✅ Input is already a PDF, using directly (no conversion needed)")
            logger.info(f"[EXTRACT_SLIDES] PDF path: {file_path}")
            full_pdf = file_path
            should_delete_full_pdf = False
        else:
            # Convert PowerPoint to PDF with HIGH QUALITY
            logger.info(f"[EXTRACT_SLIDES] 🔄 Converting PowerPoint to PDF with HIGH QUALITY")
            logger.info(f"[EXTRACT_SLIDES] PowerPoint path: {file_path}")
            full_pdf = await asyncio.get_event_loop().run_in_executor(
                None, convert_pptx_to_pdf, file_path, True  # high_quality=True
            )
            logger.info(f"[EXTRACT_SLIDES] 📄 Conversion complete: {full_pdf}")
            should_delete_full_pdf = True
        
        try:
            # Read the PDF
            reader = PdfReader(full_pdf)
            num_pages = len(reader.pages)
            
            if num_pages == 0:
                raise ValueError("PDF has no pages")
            
            # Extract first page
            intro_writer = PdfWriter()
            intro_writer.add_page(reader.pages[0])
            
            intro_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            intro_file.close()
            
            with open(intro_file.name, 'wb') as f:
                intro_writer.write(f)
            
            # Extract last page
            outro_writer = PdfWriter()
            outro_writer.add_page(reader.pages[-1])
            
            outro_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            outro_file.close()
            
            with open(outro_file.name, 'wb') as f:
                outro_writer.write(f)
            
            logger.info(f"[EXTRACT_SLIDES] Successfully extracted intro: {intro_file.name}, outro: {outro_file.name}")
            
            return intro_file.name, outro_file.name
            
        finally:
            # Only clean up the full PDF if we created it (from PowerPoint conversion)
            if should_delete_full_pdf:
                try:
                    os.unlink(full_pdf)
                except:
                    pass