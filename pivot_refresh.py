"""Optional local helper: open the generated workbook in Excel/LibreOffice and refresh PivotTables.
The Streamlit app already patches Pivot cache source ranges and sets refreshOnLoad.
"""
from pathlib import Path
import sys

if __name__ == "__main__":
    p = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("Auto_Pivot_Live.xlsx")
    print(f"Workbook ready: {p.resolve()}")
    print("Open in Excel and use Data -> Refresh All if a manual refresh is required.")
