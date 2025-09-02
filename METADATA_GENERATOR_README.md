# Metadata Generator from Excel

This script generates metadata.txt files for each location from your Excel spreadsheet.

## How to Use

### 1. Basic Usage
```bash
python generate_metadata_from_excel.py your_locations.xlsx
```

### 2. Dry Run (Preview without creating files)
```bash
python generate_metadata_from_excel.py your_locations.xlsx --dry-run
```

### 3. Custom Output Directory
```bash
python generate_metadata_from_excel.py your_locations.xlsx --output-dir custom/path
```

### 4. Specific Sheet
```bash
python generate_metadata_from_excel.py your_locations.xlsx --sheet "Sheet2"
```

## Excel Format

Your Excel should have these columns:
- **Location Name** (required): The name of the location
- **Spot Length**: Duration in seconds (or "Static" for static displays)
- **Loop Length**: Total loop duration in seconds
- **Number of Spots**: How many spots available
- **Upload Fee**: Fee amount (e.g., 3000 or "AED 3,000")
- **City**: City name (optional)
- **Area**: Area/district name (optional)
- **GPS Coordinates**: Location coordinates (optional)

## What It Creates

For each location, it creates:
```
data/templates/
├── the_landmark/
│   └── metadata.txt
├── the_gateway/
│   └── metadata.txt
└── dubai_mall_static/
    └── metadata.txt
```

## Metadata Format

Each metadata.txt contains:
```
Location Name: The Landmark
Display Name: The Landmark
Display Type: Digital
Description: The Landmark - Digital Display - 1 Spot - 16 Seconds - 16.6% SOV - Total Loop is 6 spots
SOV: 16.6%
Upload Fee: 3000
City: Dubai
Area: Sheikh Zayed Road
GPS: 25.0657° N 55.1713° E
```

## Special Cases

### Static Displays
If Spot Length = "Static", it creates:
- Display Type: Static
- SOV: 100%
- Description without spot/duration info

### Digital Displays
- Calculates SOV from spot length / loop length
- Includes spot count and duration in description
- Defaults: 16 seconds spot, 96 seconds loop

## Next Steps

After running the script:
1. Copy your PowerPoint templates to each created folder
2. Rename each .pptx file to match the folder name
   - Example: `data/templates/the_landmark/the_landmark.pptx`

## Example with CSV

```bash
# Using the example CSV file
python generate_metadata_from_excel.py example_locations.csv --dry-run

# Create the actual files
python generate_metadata_from_excel.py example_locations.csv
```