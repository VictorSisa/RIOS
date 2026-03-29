import pandas as pd
from pathlib import Path
import glob

downloads = Path.home() / "Downloads"
excel_files = glob.glob(str(downloads / "*2026*Baby*.xlsm"))
file_path = excel_files[0]

xls = pd.ExcelFile(file_path)
df = pd.read_excel(file_path, sheet_name='1010 Data')

print("\n" + "=" * 150)
print("1010 DATA ANALYSIS - WHAT CAN WE CALCULATE")
print("=" * 150)
print("")

print("SHEET INFO:")
print(f"  Rows: {len(df):,}")
print(f"  Columns: {len(df.columns)}")
print("")

print("=" * 150)
print("DATA AVAILABLE FOR CALCULATIONS")
print("=" * 150)
print("")

print("ENDING INVENTORY:")
print("  Column AP: Curr. Per. EoP Total Extended Retail $")
print("  Column AQ: Comp. Per. EoP Total Extended Retail $")
print("  Status: ✅ AVAILABLE")
print("")

print("BEGINNING INVENTORY:")
print("  Status: ❌ NOT DIRECTLY AVAILABLE")
print("  Note: Only have 1 week snapshot, need historical data")
print("")

print("SALES DATA:")
print("  Status: ⚠️ NEED TO SEARCH columns A-AK")
print("  Current location: Aggregated in Category-Summary sheet")
print("")

print("GROSS MARGIN:")
print("  Column AN/AO: E-commerce GM $ only")
print("  Status: ❌ INCOMPLETE (missing store + total GM)")
print("")

print("=" * 150)
print("RECOMMENDATION FOR REPORT")
print("=" * 150)
print("")

print("REMOVE from current report:")
print("  ❌ Inventory Turns (need BOP data)")
print("  ❌ GMROI (need complete GM data)")
print("")

print("KEEP in report:")
print("  ✅ EoP Inventory $ with YoY change")
print("  ✅ All other current metrics")
print("")

print("ADD to report (new data):")
print("  ✅ On Order inventory (Column BB/BD)")
print("  ✅ Receipts/Inbound data (Column AX/AY)")
print("  ✅ Store vs E-comm breakdown")
print("")

