#!/usr/bin/env python3
import pandas as pd

# Read the Excel file
df = pd.read_excel('metadata_excel.xlsx')

# Print column names
print("Columns in the Excel file:")
print(df.columns.tolist())
print("\n")

# Print first few rows to understand the data
print("First 5 rows of data:")
print(df.head())
print("\n")

# Check data types
print("Data types:")
print(df.dtypes)
print("\n")

# Check unique values in Spot Length and Loop Length
print("Unique values in 'Spot Length (in seconds) ':")
print(df['Spot Length (in seconds) '].unique())
print("\n")

print("Unique values in 'Loop Length (in seconds) ':")
print(df['Loop Length (in seconds) '].unique())
print("\n")

# Print sample row with all columns
print("Sample complete row:")
print(df.iloc[0].to_dict())