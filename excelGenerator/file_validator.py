#!/usr/bin/env python3
"""
STARK096 — ALDEA: File Validator
Validate generated Excel files
"""

import os
import openpyxl


def validate_files():
    """Validate Excel files in output directory"""
    
    print("✅ Validating generated files...")
    print("")
    
    output_dir = "output"
    
    if not os.path.exists(output_dir):
        print("⚠️  Output directory not found. No files to validate.")
        return 0
    
    excel_files = [f for f in os.listdir(output_dir) if f.endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print("⚠️  No Excel files found in output directory.")
        return 0
    
    print(f"Found {len(excel_files)} Excel file(s) to validate:")
    print("")
    
    valid_count = 0
    error_count = 0
    
    for file in excel_files:
        file_path = os.path.join(output_dir, file)
        try:
            wb = openpyxl.load_workbook(file_path)
            sheets = wb.sheetnames
            print(f"  ✅ {file}")
            print(f"     Sheets: {', '.join(sheets)}")
            valid_count += 1
            wb.close()
        except Exception as e:
            print(f"  ❌ {file}")
            print(f"     Error: {e}")
            error_count += 1
    
    print("")
    print(f"Validation complete: {valid_count} valid, {error_count} errors")
    
    return 0 if error_count == 0 else 1


if __name__ == "__main__":
    try:
        exit(validate_files())
    except Exception as e:
        print(f"❌ Error: {e}")
        exit(1)
