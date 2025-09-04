"""Utilities for extracting specific slides from PowerPoint to PDF without quality loss"""

import os
import tempfile
import asyncio
from typing import Tuple

from pdf_utils import convert_pptx_to_pdf, _CONVERT_SEMAPHORE
from pypdf import PdfReader, PdfWriter
import config


async def extract_first_and_last_slide_as_pdfs(pptx_path: str) -> Tuple[str, str]:
    """
    Extract first and last slides as separate PDFs without re-saving the PowerPoint.
    This avoids quality degradation from python-pptx re-saving.
    
    Returns: (intro_pdf_path, outro_pdf_path)
    """
    logger = config.logger
    logger.info(f"[EXTRACT_SLIDES] Extracting first and last slides from: {pptx_path}")
    
    async with _CONVERT_SEMAPHORE:
        # First, convert the entire PowerPoint to PDF with HIGH QUALITY
        full_pdf = await asyncio.get_event_loop().run_in_executor(
            None, convert_pptx_to_pdf, pptx_path, True  # high_quality=True
        )
        
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
            # Clean up the full PDF
            try:
                os.unlink(full_pdf)
            except:
                pass