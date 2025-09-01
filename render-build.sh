#!/usr/bin/env bash
# Build script for Render

echo "Starting Render build process..."

# Install Python dependencies
pip install -r requirements.txt

# Install LibreOffice for PDF conversion
echo "Installing LibreOffice..."
apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    libreoffice-draw \
    fonts-liberation \
    fonts-dejavu \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install custom fonts
echo "Installing custom fonts..."
if [ -d "/data/Sofia-Pro Font" ]; then
    mkdir -p ~/.local/share/fonts
    cp "/data/Sofia-Pro Font"/*.ttf ~/.local/share/fonts/ 2>/dev/null || true
    cp "/data/Sofia-Pro Font"/*.otf ~/.local/share/fonts/ 2>/dev/null || true
    fc-cache -f -v || true
fi

echo "Build complete!"