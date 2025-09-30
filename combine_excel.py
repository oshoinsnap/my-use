import pandas as pd
import sys
import os

if len(sys.argv) < 3:
    print("Usage: python combine_excel.py <output_file.xlsx> <input_file1.xlsx> [<input_file2.xlsx> ...]")
    sys.exit(1)

output_file = sys.argv[1]
input_files = sys.argv[2:]

all_dataframes = []

# Loop through all input files
for file in input_files:
    if not os.path.exists(file):
        print(f"Warning: File '{file}' not found, skipping...")
        continue
    
    # Read all sheets from each Excel file
    sheets = pd.read_excel(file, sheet_name=None)
    
    # Append all sheet DataFrames to a list
    for df in sheets.values():
        all_dataframes.append(df)

if not all_dataframes:
    print("Error: No data found in the given files.")
    sys.exit(1)

# Combine everything into one DataFrame
combined = pd.concat(all_dataframes, ignore_index=True)

# Possible email column names (case insensitive)
possible_email_cols = ['email', 'email address', 'e-mail']

# Find the email column (case-insensitive)
email_col = None
for col in combined.columns:
    if col.strip().lower() in [p.lower() for p in possible_email_cols]:
        email_col = col
        break

if email_col is None:
    print("Error: No email column found. Possible names: email, email address, e-mail")
    sys.exit(1)

# Drop duplicates based on the email column
unique = combined.drop_duplicates(subset=email_col, keep="first")

# Save to output Excel
unique.to_excel(output_file, index=False)

print(f"Unique rows saved to {output_file} based on '{email_col}' column")
