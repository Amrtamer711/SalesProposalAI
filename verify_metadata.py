#!/usr/bin/env python3
"""Verify all metadata files have the correct format."""

from pathlib import Path

templates_dir = Path("data/templates")

# Check all metadata files
issues = []
for metadata_file in templates_dir.rglob("metadata.txt"):
    content = metadata_file.read_text()
    location_name = metadata_file.parent.name
    
    # Check for old format issues
    if "Total Loop Length:" in content:
        issues.append(f"{location_name}: Still has 'Total Loop Length'")
    
    if "Description:" in content:
        issues.append(f"{location_name}: Still has 'Description'")
    
    # Check digital vs static format
    if "Display Type: Digital" in content:
        if "Spot Duration:" not in content:
            issues.append(f"{location_name}: Digital but missing 'Spot Duration'")
        if "Loop Duration:" not in content:
            issues.append(f"{location_name}: Digital but missing 'Loop Duration'")
    
    if "Display Type: Static" in content:
        if "Number of Faces:" not in content:
            issues.append(f"{location_name}: Static but missing 'Number of Faces'")
    
    # Check dimensions format
    if "Dimensions:" in content:
        issues.append(f"{location_name}: Has old 'Dimensions' format instead of separate Height/Width")

if issues:
    print("❌ Found issues:")
    for issue in issues:
        print(f"  - {issue}")
else:
    print("✅ All metadata files are in the correct format!")

# Show sample of each type
print("\n📄 Sample Digital (dubai_gateway):")
print((templates_dir / "dubai_gateway" / "metadata.txt").read_text())

print("\n📄 Sample Static (t3):")
print((templates_dir / "t3" / "metadata.txt").read_text())