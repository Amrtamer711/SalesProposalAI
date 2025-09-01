#!/bin/bash

# Install fonts from /data/fonts if they exist
if [ -d "/data/fonts" ]; then
    echo "Installing fonts from /data/fonts..."
    
    # Create user fonts directory
    mkdir -p ~/.local/share/fonts
    
    # Copy all font files
    cp /data/fonts/*.ttf ~/.local/share/fonts/ 2>/dev/null || true
    cp /data/fonts/*.otf ~/.local/share/fonts/ 2>/dev/null || true
    
    # Update font cache
    fc-cache -f -v
    
    echo "Fonts installed successfully"
else
    echo "No fonts directory found in /data/fonts"
fi