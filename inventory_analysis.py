#!/usr/bin/env python3
"""
Inventory Turns Analysis & Explanation
Helps understand the data structure and calculation methods
"""

from services.excel_extractor import extract_category_data
from pathlib import Path
import pandas as pd
import glob

# Find Excel file
downloads = Path.home() / "Downloads"
excel_files = glob.glob(str(downloads / "*2026*Baby*.xlsm"))

if excel_files:
    file_path = excel_files[0]
    print("\n" + "=" * 120)
    print("UNDERSTANDING INVENTORY TURNS")
    print("=" * 120)
    print("")
    
    # Extract data
    categories = extract_category_data(file_path)
    cat = categories.copy()
    
    # Clean numeric columns
    cat['sales_ty'] = pd.to_numeric(cat['sales_ty'], errors='coerce')
    cat['units_ty'] = pd.to_numeric(cat['units_ty'], errors='coerce')
    cat['inventory'] = pd.to_numeric(cat['inventory'], errors='coerce')
    cat['aur'] = pd.to_numeric(cat['aur'], errors='coerce')
    
    print("1. WHAT DATA DO WE HAVE?")
    print("-" * 120)
    print("")
    print("Column: 'sales_ty' = This Year Sales $ (weekly)")
    print("Column: 'units_ty' = This Year Units (weekly)")
    print("Column: 'inventory' = EOP Inventory $ (End of Period)")
    print("Column: 'aur' = Average Unit Retail")
    print("")
    print("KEY QUESTION: Do we have BEGINNING inventory? Let me check sheet structure...")
    print("")
    
    # Check if there are more columns we haven't extracted
    print(f"Total columns in extracted data: {len(cat.columns)}")
    print(f"Columns: {list(cat.columns)}")
    print("")
    
    print("2. INVENTORY TURNS CALCULATION OPTIONS")
    print("-" * 120)
    print("")
    print("OPTION A: Weekly Turns (Current Calculation)")
    print("  Formula: Sales $ / Ending Inventory $")
    print("  Example:")
    cat_example = cat[cat['inventory'] > 0].head(1)
    if not cat_example.empty:
        row = cat_example.iloc[0]
        weekly_turns = row['sales_ty'] / row['inventory']
        print(f"    Category: {row['category']}")
        print(f"    Weekly Sales: ${row['sales_ty']:,.2f}")
        print(f"    EOP Inventory: ${row['inventory']:,.2f}")
        print(f"    Weekly Turns: {weekly_turns:.4f}")
        print(f"    Annualized Turns: {weekly_turns * 52:.2f}x")
    print("")
    
    print("  ✓ INTERPRETATION:")
    print("    - 0.08x weekly = ~4.2x annually (reasonable for baby products)")
    print("    - This means inventory turns over ~4 times per year")
    print("")
    
    print("OPTION B: If we had BEGINNING inventory (which we might not)")
    print("  Formula: Sales $ / AVERAGE Inventory")
    print("  Average = (Beginning Inventory + Ending Inventory) / 2")
    print("  This would give a more accurate picture mid-period")
    print("")
    
    print("3. CHECKING YOUR EXCEL STRUCTURE")
    print("-" * 120)
    print("")
    print("To understand what inventory data you have, let me show sample data:")
    print("")
    
    # Show category sample
    sample_cols = ['category', 'sales_ty', 'units_ty', 'aur', 'inventory', 'inventory_pct_change']
    sample_data = cat[sample_cols].head(10)
    print(sample_data.to_string())
    print("")
    
    print("4. QUESTIONS TO ASK YOURSELF")
    print("-" * 120)
    print("")
    print("□ Does 'inventory' column = Ending of Period (EOP) Inventory?")
    print("□ Is there a 'Beginning of Period' (BOP) inventory column we haven't extracted?")
    print("□ Is this weekly data or period-end snapshot?")
    print("")
    print("If YES to all three, we should calculate:")
    print("  Average Inventory = (BOP + EOP) / 2")
    print("  Then: Turns = Sales / Average Inventory")
    print("")
    
    print("5. RECOMMENDED APPROACH FOR REPORT")
    print("-" * 120)
    print("")
    print("SIMPLE (Current):")
    print("  'Inventory Turns: {weekly_turns:.2f}x weekly (~{annualized:.1f}x annually)'")
    print("")
    print("BETTER (If we had BOP):")
    print("  'Inventory Turns: {turns:.1f}x annually based on average inventory'")
    print("")
    print("BEST (With context):")
    print("  'Inventory Turns: {turns:.1f}x")
    print("   - Ending Inventory: ${eop:,.0f}")
    print("   - YoY Inventory Change: {inv_change:.1%}")
    print("   - Assessment: {'Efficient' if turns > 4 else 'Needs optimization'}'")
    print("")
    
    print("=" * 120)
    print("NEXT STEP: Can you check your Excel sheet and tell me:")
    print("  1. Does the data have a 'Beginning' inventory column?")
    print("  2. What's the exact column header for the inventory data?")
    print("  3. Is this weekly or period-based?")
    print("=" * 120)

