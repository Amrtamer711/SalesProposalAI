#!/usr/bin/env python3
"""
Script to generate metadata.txt files for each location from an Excel spreadsheet.
Creates folder structure: data/templates/{location_name}/metadata.txt
"""

import pandas as pd
import os
from pathlib import Path
import argparse
import re


def clean_folder_name(name):
    """Convert location name to a valid folder name."""
    # Remove 'The' prefix if present
    if name.lower().startswith('the '):
        name = name[4:]
    
    # Remove special characters and replace spaces with underscores
    name = re.sub(r'[^\w\s-]', '', name)
    name = re.sub(r'[-\s]+', '_', name)
    return name.lower().strip('_')


def format_sov(sov_value, spot_length, loop_length, display_type):
    """Format SOV percentage, preferring Excel value if available."""
    if display_type.lower() == 'static':
        return "100%"  # Static displays have 100% SOV
    
    # If SOV is provided in Excel, use it
    if pd.notna(sov_value):
        try:
            # Convert decimal to percentage
            sov_percent = float(sov_value) * 100
            return f"{sov_percent:.1f}%"
        except:
            pass
    
    # Otherwise calculate from spot/loop
    try:
        spot = float(spot_length)
        loop = float(loop_length)
        if loop > 0:
            sov = (spot / loop) * 100
            return f"{sov:.1f}%"
        else:
            return "16.6%"  # Default
    except (ValueError, TypeError):
        return "16.6%"  # Default if conversion fails


def generate_metadata(row):
    """Generate metadata content from a DataFrame row."""
    metadata_lines = []
    
    # Location Name (use as display name)
    if pd.notna(row.get('Location Name')):
        metadata_lines.append(f"Location Name: {row['Location Name']}")
        metadata_lines.append(f"Display Name: {row['Location Name']}")
    
    # Display Type (Digital or Static) - check with spaces in column name
    display_type = 'Digital'  # Default
    spot_length_col = 'Spot Length (in seconds) '
    loop_length_col = 'Loop Length (in seconds) '
    
    if pd.notna(row.get(spot_length_col)):
        if str(row[spot_length_col]).lower() == 'static':
            display_type = 'Static'
        else:
            display_type = 'Digital'
    metadata_lines.append(f"Display Type: {display_type}")
    
    # For digital displays, add duration info
    if display_type == 'Digital':
        # Spot Duration
        spot_length = 16  # Default
        if pd.notna(row.get(spot_length_col)):
            try:
                spot_length = int(float(row[spot_length_col]))
            except:
                spot_length = 16
        metadata_lines.append(f"Spot Duration: {spot_length}")
        
        # Loop Duration
        loop_length = 96  # Default
        if pd.notna(row.get(loop_length_col)):
            try:
                loop_length = int(float(row[loop_length_col]))
            except:
                loop_length = 96
        metadata_lines.append(f"Loop Duration: {loop_length}")
    
    # For static displays, add Number of Faces
    if display_type == 'Static':
        if pd.notna(row.get('No. of Faces')):
            faces = int(row['No. of Faces'])
            metadata_lines.append(f"Number of Faces: {faces}")
    
    # SOV percentage - only for digital displays
    if display_type == 'Digital':
        sov = format_sov(
            row.get('SOV'), 
            row.get(spot_length_col, 16), 
            row.get(loop_length_col, 96), 
            display_type
        )
        metadata_lines.append(f"SOV: {sov}")
    
    # Upload Fee - only for digital displays
    if display_type == 'Digital':
        upload_fee = 3000  # Default
        metadata_lines.append(f"Upload Fee: {upload_fee}")
    
    # Series
    if pd.notna(row.get('Series')):
        metadata_lines.append(f"Series: {row['Series']}")
    
    # Height and Width as separate fields with units
    if pd.notna(row.get('Height')):
        metadata_lines.append(f"Height: {row['Height']}m")
    
    if pd.notna(row.get('Width')):
        metadata_lines.append(f"Width: {row['Width']}m")
    
    return '\n'.join(metadata_lines)


def main():
    parser = argparse.ArgumentParser(description='Generate metadata files from Excel')
    parser.add_argument('excel_file', help='Path to the Excel file')
    parser.add_argument('--output-dir', default='data/templates', 
                      help='Output directory for templates (default: data/templates)')
    parser.add_argument('--sheet', default=0, 
                      help='Sheet name or index to read (default: 0)')
    parser.add_argument('--dry-run', action='store_true',
                      help='Print what would be created without creating files')
    
    args = parser.parse_args()
    
    # Read Excel file
    try:
        df = pd.read_excel(args.excel_file, sheet_name=args.sheet)
        print(f"✓ Loaded {len(df)} rows from Excel file")
        print(f"  Columns found: {', '.join(df.columns.tolist())}")
    except Exception as e:
        print(f"✗ Error reading Excel file: {e}")
        return 1
    
    # Check required columns
    required_columns = ['Location Name']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print(f"✗ Missing required columns: {', '.join(missing_columns)}")
        print(f"  Available columns: {', '.join(df.columns.tolist())}")
        return 1
    
    # Process each row
    output_base = Path(args.output_dir)
    created_folders = []
    
    for idx, row in df.iterrows():
        if pd.isna(row.get('Location Name')):
            print(f"  Skipping row {idx + 1}: No location name")
            continue
        
        location_name = row['Location Name']
        folder_name = clean_folder_name(location_name)
        
        if not folder_name:
            print(f"  Skipping row {idx + 1}: Invalid location name '{location_name}'")
            continue
        
        # Create folder path
        folder_path = output_base / folder_name
        metadata_path = folder_path / 'metadata.txt'
        
        # Generate metadata content
        metadata_content = generate_metadata(row)
        
        if args.dry_run:
            print(f"\n→ Would create: {folder_path}/")
            print(f"  Location: {location_name}")
            print(f"  Metadata content:")
            print("  " + "\n  ".join(metadata_content.split('\n')))
        else:
            # Create directory
            folder_path.mkdir(parents=True, exist_ok=True)
            
            # Write metadata file
            with open(metadata_path, 'w', encoding='utf-8') as f:
                f.write(metadata_content)
            
            created_folders.append(folder_name)
            print(f"✓ Created {folder_path}/ with metadata.txt")
    
    if not args.dry_run:
        print(f"\n✓ Created {len(created_folders)} location folders")
        print(f"  Output directory: {output_base.absolute()}")
        print(f"\n📋 Next steps:")
        print(f"  1. Copy the corresponding PowerPoint files to each folder")
        print(f"  2. Rename each PowerPoint file to match the folder name")
        print(f"  3. Example: data/templates/landmark/landmark.pptx")
    
    return 0


if __name__ == "__main__":
    exit(main())