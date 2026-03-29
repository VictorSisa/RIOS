#!/usr/bin/env python3
"""
Test script for Excel Extractor
Verifies category and vendor extraction
"""

from services.excel_extractor import extract_category_data, extract_vendor_data, extract_brand_data
from pathlib import Path
import sys

# Find Excel file
downloads = Path.home() / "Downloads"
excel_files = list(downloads.glob("*2026*Baby*.xlsm"))

if not excel_files:
    print("❌ Excel file not found")
    sys.exit(1)

file_path = str(excel_files[0])
print(f"\n✅ Using file: {Path(file_path).name}\n")

try:
    # Extract category data
    print("=" * 100)
    print("CATEGORY SUMMARY DATA")
    print("=" * 100)
    categories = extract_category_data(file_path)
    print(f"\nExtracted {len(categories)} categories")
    print(f"Columns: {list(categories.columns)}\n")
    print(categories.head(10).to_string())
    print("\n")
    
    # Extract vendor data
    print("=" * 100)
    print("VENDOR->BRAND SUMMARY DATA")
    print("=" * 100)
    vendors = extract_vendor_data(file_path)
    print(f"\nExtracted {len(vendors)} vendor-brand combinations")
    print(f"Columns: {list(vendors.columns)}\n")
    print(vendors.head(10).to_string())
    print("\n")
    
    # Extract brand data
    print("=" * 100)
    print("BRAND->VENDOR SUMMARY DATA")
    print("=" * 100)
    brands = extract_brand_data(file_path)
    print(f"\nExtracted {len(brands)} brand-vendor combinations")
    print(f"Columns: {list(brands.columns)}\n")
    print(brands.head(10).to_string())
    
    print("\n✅ Extraction complete!")
    
except Exception as e:
    print(f"❌ Error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

