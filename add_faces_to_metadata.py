#!/usr/bin/env python3
"""Add Number of Faces to existing metadata files from Excel."""

import pandas as pd
from pathlib import Path
import re

def clean_folder_name(name):
    """Convert location name to folder name (same logic as generator)."""
    # Remove 'The' prefix if present
    if name.lower().startswith('the '):
        name = name[4:]
    
    # Remove special characters and replace spaces with underscores
    name = re.sub(r'[^\w\s-]', '', name)
    name = re.sub(r'[-\s]+', '_', name)
    return name.lower().strip('_')

# Read Excel file
df = pd.read_excel('metadata_excel.xlsx')
print(f"✓ Loaded {len(df)} locations from Excel")

templates_dir = Path('data/templates')
updated = 0
skipped = 0

for idx, row in df.iterrows():
    location_name = row['Location Name']
    folder_name = clean_folder_name(location_name)
    metadata_path = templates_dir / folder_name / 'metadata.txt'
    
    if not metadata_path.exists():
        print(f"⚠️  Skipping {location_name} - metadata file not found")
        skipped += 1
        continue
    
    # Read existing metadata
    content = metadata_path.read_text()
    
    # Check if Number of Faces already exists
    if 'Number of Faces:' in content:
        print(f"✓ {location_name} already has Number of Faces")
        continue
    
    # Get number of faces from Excel
    faces = int(row['No. of Faces']) if pd.notna(row.get('No. of Faces')) else 1
    
    # Find where to insert (after Display Type)
    lines = content.split('\n')
    new_lines = []
    inserted = False
    
    for line in lines:
        new_lines.append(line)
        if line.startswith('Display Type:') and not inserted:
            new_lines.append(f'Number of Faces: {faces}')
            inserted = True
    
    # Write back
    new_content = '\n'.join(new_lines)
    metadata_path.write_text(new_content)
    
    print(f"✓ Updated {location_name} - added Number of Faces: {faces}")
    updated += 1

print(f"\n📊 Summary:")
print(f"  Updated: {updated} files")
print(f"  Skipped: {skipped} files")
print(f"  Already had faces: {len(df) - updated - skipped} files")