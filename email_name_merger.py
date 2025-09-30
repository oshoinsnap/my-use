#!/usr/bin/env python3
"""
Email Sheet Merger Script (Extended)

This script compares emails from a source sheet/file to a target sheet/file.
It merges rows based on matching emails and writes a new Excel file.

Usage:
    python email_name_merger.py
"""

import csv
import json
import pandas as pd
from typing import List, Dict, Tuple

def read_list_from_excel(filename: str, sheet_name: str = 0) -> List[Dict[str, str]]:
    """Read a list of contacts from Excel file (.xlsx, .xls)."""
    try:
        df = pd.read_excel(filename, sheet_name=sheet_name)
        return df.to_dict('records')
    except FileNotFoundError:
        print(f"Error: File '{filename}' not found.")
        return []
    except Exception as e:
        print(f"Error reading Excel file '{filename}': {str(e)}")
        return []

def merge_by_email(
    source_list: List[Dict[str, str]],
    target_list: List[Dict[str, str]],
    email_key: str = 'email'
) -> Tuple[List[Dict[str, str]], int]:
    """Compare source and target lists based on matching emails and merge rows."""
    updated_count = 0
    email_to_data = {}

    for contact in source_list:
        email = contact.get(email_key, '').lower().strip()
        if email:
            email_to_data[email] = contact

    matched_list = []
    for contact in target_list:
        email = contact.get(email_key, '').lower().strip()
        if email in email_to_data:
            combined = {**email_to_data[email], **contact}
            matched_list.append(combined)
            updated_count += 1

    return matched_list, updated_count

def write_list_to_excel(contacts: List[Dict[str, str]], filename: str, sheet_name: str = 'Sheet1') -> bool:
    """Write list of contacts to Excel file."""
    if not contacts:
        print("Warning: No contacts to write.")
        return False
    try:
        df = pd.DataFrame(contacts)
        df.to_excel(filename, sheet_name=sheet_name, index=False)
        return True
    except Exception as e:
        print(f"Error writing to file '{filename}': {str(e)}")
        return False

def main():
    print("Email Sheet Merger Script (Supports multiple files or sheets)")
    print("=" * 50)

    # Ask user if they want to merge across files
    cross_file = input("Do you want to merge across two Excel files? (y/n): ").lower().strip()

    if cross_file == 'y':
        # File 1
        source_file = input("Enter source Excel file: ").strip()
        source_sheet = input("Enter source sheet name (leave blank for first): ").strip() or 0

        # File 2
        target_file = input("Enter target Excel file: ").strip()
        target_sheet = input("Enter target sheet name (leave blank for first): ").strip() or 0
    else:
        # Single file with two sheets
        excel_file = input("Enter Excel file name: ").strip() or 'newjonny.xlsx'
        try:
            xl = pd.ExcelFile(excel_file)
            sheets = xl.sheet_names
            print(f"Available sheets: {sheets}")
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return

        source_file = target_file = excel_file
        source_sheet = input("Enter source sheet name: ").strip() or sheets[0]
        target_sheet = input("Enter target sheet name: ").strip() or (sheets[1] if len(sheets) > 1 else 'Sheet2')

    output_file = input("Enter output Excel file name: ").strip() or 'merged.xlsx'
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'

    # Read data
    source_list = read_list_from_excel(source_file, source_sheet)
    target_list = read_list_from_excel(target_file, target_sheet)

    if not source_list or not target_list:
        print("Error: Unable to read source or target data.")
        return

    # Merge
    updated_list, matches = merge_by_email(source_list, target_list)

    print(f"\nProcessing complete:")
    print(f"- Source list: {len(source_list)} contacts")
    print(f"- Target list: {len(target_list)} contacts")
    print(f"- Matches found: {matches}")

    # Write output
    success = write_list_to_excel(updated_list, output_file)

    if success:
        print(f"- Merged data written to: {output_file}")
    else:
        print("Error: Failed to write output file.")

if __name__ == "__main__":
    main()
