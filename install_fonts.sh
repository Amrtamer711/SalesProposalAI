#!/bin/bash

# Install fonts from /data/Sofia-Pro Font if they exist
if [ -d "/data/Sofia-Pro Font" ]; then
    echo "Installing fonts from /data/Sofia-Pro Font..."
    
    # Create user fonts directory
    mkdir -p ~/.local/share/fonts
    
    # Copy all font files
    cp "/data/Sofia-Pro Font"/*.ttf ~/.local/share/fonts/ 2>/dev/null || true
    cp "/data/Sofia-Pro Font"/*.otf ~/.local/share/fonts/ 2>/dev/null || true
    
    # Try to update font cache if fc-cache exists
    if command -v fc-cache &> /dev/null; then
        fc-cache -f -v
    else
        echo "fc-cache not available, fonts copied to ~/.local/share/fonts"
    fi
    
    echo "Fonts installed successfully"
else
    echo "No fonts directory found in /data/Sofia-Pro Font"
fi