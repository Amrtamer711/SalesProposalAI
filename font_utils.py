import os
import shutil
from pathlib import Path
import logging

logger = logging.getLogger("proposal-bot")

def install_custom_fonts():
    """Install fonts from /data/Sofia-Pro Font directory for python-pptx usage."""
    data_fonts_dir = Path("/data/Sofia-Pro Font")
    
    if not data_fonts_dir.exists():
        logger.info("No fonts directory found in /data/Sofia-Pro Font")
        return
    
    # Multiple possible font directories
    font_dirs = [
        Path.home() / ".fonts",
        Path.home() / ".local/share/fonts",
        Path("/usr/share/fonts/truetype/custom")
    ]
    
    # Find all font files
    font_files = list(data_fonts_dir.glob("*.ttf")) + list(data_fonts_dir.glob("*.otf"))
    
    if not font_files:
        logger.info("No font files found in /data/Sofia-Pro Font")
        return
    
    # Try to copy fonts to available directories
    for font_dir in font_dirs:
        try:
            font_dir.mkdir(parents=True, exist_ok=True)
            for font_file in font_files:
                dest = font_dir / font_file.name
                if not dest.exists():
                    shutil.copy2(font_file, dest)
                    logger.info(f"Copied {font_file.name} to {font_dir}")
            break  # If successful, don't try other directories
        except Exception as e:
            logger.debug(f"Could not copy to {font_dir}: {e}")
            continue
    
    # Set environment variable for fontconfig
    os.environ['FONTCONFIG_PATH'] = '/data/Sofia-Pro Font'
    
    logger.info("Custom fonts installation completed")