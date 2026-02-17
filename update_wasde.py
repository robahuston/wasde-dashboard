#!/usr/bin/env python3
"""
WASDE Dashboard Auto-Updater
Downloads the latest WASDE Excel file from USDA and regenerates the dashboard data.
Run monthly after WASDE release (typically 10th-12th of each month).

Usage: python update_wasde.py [--year YYYY] [--month MM]
  Defaults to current year/month.
"""

import sys, os, json, re, argparse
from datetime import datetime
from urllib.request import urlretrieve
from pathlib import Path

try:
    import xlrd
except ImportError:
    print("Installing xlrd..."); os.system(f"{sys.executable} -m pip install xlrd --quiet"); import xlrd

SCRIPT_DIR = Path(__file__).parent
TEMPLATE_PATH = SCRIPT_DIR / "index.html"

def month_code(month):
    """WASDE file uses 2-digit month + 2-digit year: wasde0226.xls = Feb 2026"""
    return f"{month:02d}"

def download_wasde(year, month):
    yy = year % 100
    mm = month_code(month)
    filename = f"wasde{mm}{yy}.xls"
    url = f"https://www.usda.gov/oce/commodity/wasde/{filename}"
    local = SCRIPT_DIR / filename
    print(f"Downloading {url} ...")
    try:
        urlretrieve(url, local)
        print(f"  Saved {local} ({local.stat().st_size:,} bytes)")
        return local
    except Exception as e:
        print(f"  ERROR: {e}")
        return None

def safe(val):
    """Convert cell value to float, return 0 if empty"""
    try:
        return float(val) if val != '' else 0
    except (ValueError, TypeError):
        return 0

def extract_data(xls_path, year, month):
    wb = xlrd.open_workbook(str(xls_path))
    
    # Determine report number from first sheet
    txt_sheet = wb.sheet_by_name('WASDE Text')
    report_id = "WASDE"
    for r in range(min(10, txt_sheet.nrows)):
        for c in range(min(5, txt_sheet.ncols)):
            val = str(txt_sheet.cell_value(r, c))
            m = re.search(r'WASDE[- ]+(\d+)', val)
            if m:
                report_id = f"WASDE-{m.group(1)}"
                break

    month_names = ["", "January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"]
    report_date = f"{month_names[month]} {year}"

    # Helper to read a sheet into rows
    def read_sheet(name):
        s = wb.sheet_by_name(name)
        rows = []
        for r in range(s.nrows):
            row = [s.cell_value(r, c) for c in range(s.ncols)]
            rows.append(row)
        return rows

    # --- CORN (Page 12) ---
    corn_rows = read_sheet('Page 12')
    # Find the CORN section (starts after feed grains)
    corn_start = 0
    for i, row in enumerate(corn_rows):
        if str(row[0]).strip().startswith('CORN'):
            corn_start = i
            break
    
    def get_corn_row(label, start=corn_start):
        for i in range(start, len(corn_rows)):
            if str(corn_rows[i][0]).strip().startswith(label):
                return [safe(corn_rows[i][j]) for j in range(1, 5)]
        return [0, 0, 0, 0]

    # Columns: [0]=label, [1]=2023/24, [2]=2024/25 Est., [3]=prev month proj, [4]=cur month proj
    corn = {
        'price': get_corn_row('Avg. Farm Price'),
        'planted': get_corn_row('Area Planted'),
        'harvested': get_corn_row('Area Harvested'),
        'yield': get_corn_row('Yield per'),
        'begStocks': get_corn_row('Beginning Stocks'),
        'production': get_corn_row('Production'),
        'imports': get_corn_row('Imports'),
        'supplyTotal': get_corn_row('Supply, Total'),
        'feedResidual': get_corn_row('Feed and Residual'),
        'fsi': get_corn_row('Food, Seed & Industrial'),
        'ethanol': get_corn_row('Ethanol'),
        'domesticTotal': get_corn_row('Domestic, Total'),
        'exports': get_corn_row('Exports'),
        'useTotal': get_corn_row('Use, Total'),
        'endStocks': get_corn_row('Ending Stocks'),
    }

    # --- SOYBEANS (Page 15) ---
    soy_rows = read_sheet('Page 15')
    
    def find_section(rows, label):
        for i, row in enumerate(rows):
            if str(row[0]).strip().startswith(label):
                return i
        return 0
    
    soy_start = find_section(soy_rows, 'SOYBEANS')
    oil_start = find_section(soy_rows, 'SOYBEAN OIL')
    meal_start = find_section(soy_rows, 'SOYBEAN MEAL')
    
    def get_soy_row(label, start=soy_start, end=None):
        e = end or oil_start or len(soy_rows)
        for i in range(start, e):
            if str(soy_rows[i][0]).strip().startswith(label):
                return [safe(soy_rows[i][j]) for j in range(1, 5)]
        return [0, 0, 0, 0]

    soybeans = {
        'price': get_soy_row('Avg. Farm Price'),
        'planted': get_soy_row('Area Planted'),
        'harvested': get_soy_row('Area Harvested'),
        'yield': get_soy_row('Yield per'),
        'begStocks': get_soy_row('Beginning Stocks'),
        'production': get_soy_row('Production'),
        'imports': get_soy_row('Imports'),
        'supplyTotal': get_soy_row('Supply, Total'),
        'crushings': get_soy_row('Crushings'),
        'exports': get_soy_row('Exports'),
        'seed': get_soy_row('Seed'),
        'residual': get_soy_row('Residual'),
        'useTotal': get_soy_row('Use, Total'),
        'endStocks': get_soy_row('Ending Stocks'),
        'oil': {
            'price': get_soy_row('Avg. Price', oil_start, meal_start),
            'production': get_soy_row('Production', oil_start, meal_start),
            'domesticUse': get_soy_row('Domestic Disappearance', oil_start, meal_start),
            'biofuel': get_soy_row('Biofuel', oil_start, meal_start),
            'exports': get_soy_row('Exports', oil_start, meal_start),
            'endStocks': get_soy_row('Ending stocks', oil_start, meal_start),
        },
        'meal': {
            'price': get_soy_row('Avg. Price', meal_start),
            'production': get_soy_row('Production', meal_start),
            'domesticUse': get_soy_row('Domestic Disappearance', meal_start),
            'exports': get_soy_row('Exports', meal_start),
            'endStocks': get_soy_row('Ending Stocks', meal_start),
        }
    }

    # --- WHEAT (Page 11) ---
    wheat_rows = read_sheet('Page 11')
    
    # Top section uses different column layout (more cols with empties)
    # Find the column indices by looking at header row
    def get_wheat_row(label, start=0, end=30):
        for i in range(start, min(end, len(wheat_rows))):
            cell = str(wheat_rows[i][0]).strip()
            if cell.startswith(label):
                # Wheat top section has cols: 0=label, 4=23/24, 6=24/25, 9=prev, 11=cur
                vals = [safe(wheat_rows[i][4]), safe(wheat_rows[i][6]),
                        safe(wheat_rows[i][9]), safe(wheat_rows[i][11])]
                return vals
        return [0, 0, 0, 0]

    wheat = {
        'price': get_wheat_row('Avg. Farm Price'),
        'planted': get_wheat_row('Area Planted'),
        'harvested': get_wheat_row('Area Harvested'),
        'yield': get_wheat_row('Yield per'),
        'begStocks': get_wheat_row('Beginning Stocks'),
        'production': get_wheat_row('Production'),
        'imports': get_wheat_row('Imports'),
        'supplyTotal': get_wheat_row('Supply, Total'),
        'food': get_wheat_row('Food'),
        'seed': get_wheat_row('Seed'),
        'feedResidual': get_wheat_row('Feed and Residual'),
        'domesticTotal': get_wheat_row('Domestic, Total'),
        'exports': get_wheat_row('Exports'),
        'useTotal': get_wheat_row('Use, Total'),
        'endStocks': get_wheat_row('Ending Stocks'),
    }

    # Wheat by class (bottom of page 11)
    wbc_start = find_section(wheat_rows, 'U.S. Wheat by Class')
    # Find the 2025/26 proj row
    proj_row_start = 0
    for i in range(wbc_start, len(wheat_rows)):
        if '2025/26' in str(wheat_rows[i][0]) or '2025/26' in str(wheat_rows[i][0]):
            proj_row_start = i
            break

    def get_wbc_row(label, start=proj_row_start):
        for i in range(start, min(start+15, len(wheat_rows))):
            if str(wheat_rows[i][1]).strip().startswith(label):
                # Cols: 3=HRW, 5=HRS, 7=SRW, 8=White, 10=Durum
                return [safe(wheat_rows[i][3]), safe(wheat_rows[i][5]),
                        safe(wheat_rows[i][7]), safe(wheat_rows[i][8]),
                        safe(wheat_rows[i][10])]
        return [0, 0, 0, 0, 0]

    wheat['byClass'] = {
        'labels': ["HRW", "HRS", "SRW", "White", "Durum"],
        'production': get_wbc_row('Production'),
        'exports': get_wbc_row('Exports'),
        'endStocks': get_wbc_row('Ending Stocks, Total'),
    }

    # Build the data object — use 3 values: [23/24, 24/25, 25/26 cur month]
    def to3(arr):
        """From [23/24, 24/25, prev_month, cur_month] -> [23/24, 24/25, cur_month]"""
        if len(arr) == 4:
            return [arr[0], arr[1], arr[3]]
        return arr[:3]

    def to3_prev(arr):
        """Get the prev month value for change tracking"""
        return arr[2] if len(arr) == 4 else 0

    def process_dict(d):
        result = {}
        for k, v in d.items():
            if isinstance(v, dict):
                result[k] = process_dict(v)
            elif isinstance(v, list) and len(v) >= 4 and all(isinstance(x, (int, float)) for x in v):
                result[k] = to3(v)
            else:
                result[k] = v
        return result

    # Get prev month values for ending stocks
    corn_es_prev = corn['endStocks'][2] if len(corn['endStocks']) >= 3 else 0
    soy_es_prev = soybeans['endStocks'][2] if len(soybeans['endStocks']) >= 3 else 0
    wheat_es_prev = wheat['endStocks'][2] if len(wheat['endStocks']) >= 3 else 0

    # Determine prev month name
    prev_months = ["", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov"]
    cur_months = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    result = {
        'reportId': report_id,
        'reportDate': report_date,
        'years': ["2023/24", "2024/25 Est.", "2025/26 Proj."],
        'prevMonth': prev_months[month],
        'curMonth': cur_months[month],
        'corn': process_dict(corn),
        'soybeans': process_dict(soybeans),
        'wheat': process_dict(wheat),
    }
    result['corn']['endStocksPrev'] = corn_es_prev
    result['soybeans']['endStocksPrev'] = soy_es_prev
    result['wheat']['endStocksPrev'] = wheat_es_prev

    return result

def update_html(data):
    html = TEMPLATE_PATH.read_text(encoding='utf-8')
    
    # Replace the WASDE_DATA block
    data_json = json.dumps(data, indent=2)
    pattern = r'const WASDE_DATA = \{.*?\};\s*\n// ========== END DATA =========='
    replacement = f'const WASDE_DATA = {data_json};\n// ========== END DATA =========='
    
    new_html = re.sub(pattern, replacement, html, flags=re.DOTALL)
    
    if new_html == html:
        print("WARNING: Could not find WASDE_DATA block to replace!")
        return False
    
    TEMPLATE_PATH.write_text(new_html, encoding='utf-8')
    print(f"Updated {TEMPLATE_PATH}")
    return True

def main():
    parser = argparse.ArgumentParser(description='Update WASDE Dashboard')
    now = datetime.now()
    parser.add_argument('--year', type=int, default=now.year)
    parser.add_argument('--month', type=int, default=now.month)
    args = parser.parse_args()

    xls = download_wasde(args.year, args.month)
    if not xls:
        sys.exit(1)

    print("Extracting data...")
    data = extract_data(xls, args.year, args.month)
    
    print(f"\nReport: {data['reportId']} — {data['reportDate']}")
    print(f"Corn price: ${data['corn']['price'][2]}/bu  Ending stocks: {data['corn']['endStocks'][2]} mil bu")
    print(f"Soy price:  ${data['soybeans']['price'][2]}/bu  Ending stocks: {data['soybeans']['endStocks'][2]} mil bu")
    print(f"Wheat price: ${data['wheat']['price'][2]}/bu  Ending stocks: {data['wheat']['endStocks'][2]} mil bu")
    
    print("\nUpdating HTML...")
    if update_html(data):
        print("\n✅ Dashboard updated successfully!")
        # Clean up downloaded file
        xls.unlink()
        print(f"  Cleaned up {xls.name}")
    else:
        print("\n❌ Failed to update HTML")
        sys.exit(1)

if __name__ == '__main__':
    main()
