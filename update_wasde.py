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
    total_soy_rows = len(soy_rows)
    
    def find_section(rows, label):
        for i, row in enumerate(rows):
            if str(row[0]).strip().startswith(label):
                return i
        return 0
    
    soy_start = find_section(soy_rows, 'SOYBEANS')
    oil_start = find_section(soy_rows, 'SOYBEAN OIL')
    meal_start = find_section(soy_rows, 'SOYBEAN MEAL')
    
    print(f"  Soy sections — beans: row {soy_start}, oil: row {oil_start}, meal: row {meal_start}, total rows: {total_soy_rows}")

    # FIX: Use explicit end parameter; never default to a section that comes BEFORE start
    def get_soy_row(label, start, end):
        """Search for label between start and end rows (case-insensitive startswith)."""
        label_lower = label.lower()
        for i in range(start, end):
            cell = str(soy_rows[i][0]).strip()
            if cell.lower().startswith(label_lower):
                return [safe(soy_rows[i][j]) for j in range(1, 5)]
        return [0, 0, 0, 0]

    soybeans = {
        'price': get_soy_row('Avg. Farm Price', soy_start, oil_start),
        'planted': get_soy_row('Area Planted', soy_start, oil_start),
        'harvested': get_soy_row('Area Harvested', soy_start, oil_start),
        'yield': get_soy_row('Yield per', soy_start, oil_start),
        'begStocks': get_soy_row('Beginning Stocks', soy_start, oil_start),
        'production': get_soy_row('Production', soy_start, oil_start),
        'imports': get_soy_row('Imports', soy_start, oil_start),
        'supplyTotal': get_soy_row('Supply, Total', soy_start, oil_start),
        'crushings': get_soy_row('Crushings', soy_start, oil_start),
        'exports': get_soy_row('Exports', soy_start, oil_start),
        'seed': get_soy_row('Seed', soy_start, oil_start),
        'residual': get_soy_row('Residual', soy_start, oil_start),
        'useTotal': get_soy_row('Use, Total', soy_start, oil_start),
        'endStocks': get_soy_row('Ending Stocks', soy_start, oil_start),
        'oil': {
            'price': get_soy_row('Avg. Price', oil_start, meal_start),
            'production': get_soy_row('Production', oil_start, meal_start),
            'domesticUse': get_soy_row('Domestic Disappearance', oil_start, meal_start),
            'biofuel': get_soy_row('Biofuel', oil_start, meal_start),
            'exports': get_soy_row('Exports', oil_start, meal_start),
            'endStocks': get_soy_row('Ending Stocks', oil_start, meal_start),
        },
        # FIX: meal section searches from meal_start to END OF SHEET
        'meal': {
            'price': get_soy_row('Avg. Price', meal_start, total_soy_rows),
            'production': get_soy_row('Production', meal_start, total_soy_rows),
            'domesticUse': get_soy_row('Domestic Disappearance', meal_start, total_soy_rows),
            'exports': get_soy_row('Exports', meal_start, total_soy_rows),
            'endStocks': get_soy_row('Ending Stocks', meal_start, total_soy_rows),
        }
    }

    # --- WHEAT (Page 11) ---
    wheat_rows = read_sheet('Page 11')
    
    def get_wheat_row(label, start=0, end=30):
        for i in range(start, min(end, len(wheat_rows))):
            cell = str(wheat_rows[i][0]).strip()
            if cell.startswith(label):
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
    wbc_start = 0
    for i, row in enumerate(wheat_rows):
        if 'Wheat by Class' in str(row[0]) or 'by Class' in str(row[0]):
            wbc_start = i
            break

    proj_row_start = 0
    for i in range(wbc_start, len(wheat_rows)):
        if '2025/26' in str(wheat_rows[i][0]):
            proj_row_start = i
            break

    def get_wbc_row(label, start=proj_row_start):
        for i in range(start, min(start+15, len(wheat_rows))):
            if str(wheat_rows[i][1]).strip().startswith(label):
                return [safe(wheat_rows[i][3]), safe(wheat_rows[i][5]),
                        safe(wheat_rows[i][7]), safe(wheat_rows[i][8]),
                        safe(wheat_rows[i][10])]
        return [0, 0, 0, 0, 0]

    # FIX: Store wheat by class BEFORE process_dict so 5-element arrays stay intact
    wheat_by_class = {
        'labels': ["HRW", "HRS", "SRW", "White", "Durum"],
        'production': get_wbc_row('Production'),
        'exports': get_wbc_row('Exports'),
        'endStocks': get_wbc_row('Ending Stocks, Total'),
    }

    def to3(arr):
        """From [23/24, 24/25, prev_month, cur_month] -> [23/24, 24/25, cur_month]"""
        if len(arr) == 4:
            return [arr[0], arr[1], arr[3]]
        return arr[:3]

    def process_dict(d):
        result = {}
        for k, v in d.items():
            if isinstance(v, dict):
                result[k] = process_dict(v)
            elif isinstance(v, list) and len(v) == 4 and all(isinstance(x, (int, float)) for x in v):
                result[k] = to3(v)
            else:
                result[k] = v
        return result

    # Get prev month values for ending stocks
    corn_es_prev = corn['endStocks'][2] if len(corn['endStocks']) >= 3 else 0
    soy_es_prev = soybeans['endStocks'][2] if len(soybeans['endStocks']) >= 3 else 0
    wheat_es_prev = wheat['endStocks'][2] if len(wheat['endStocks']) >= 3 else 0

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

    # Assign wheat by class AFTER process_dict
    result['wheat']['byClass'] = wheat_by_class

    result['corn']['endStocksPrev'] = corn_es_prev
    result['soybeans']['endStocksPrev'] = soy_es_prev
    result['wheat']['endStocksPrev'] = wheat_es_prev

    return result

def update_html(data):
    html = TEMPLATE_PATH.read_text(encoding='utf-8')
    
    data_json = json.dumps(data, indent=2)
    
    # FIX: More robust regex - flexible whitespace, case-insensitive END DATA marker
    pattern = r'const\s+WASDE_DATA\s*=\s*\{.*?\}\s*;\s*\n\s*//\s*=+\s*END\s+DATA\s*=+'
    
    match = re.search(pattern, html, flags=re.DOTALL)
    if not match:
        # Diagnostic: show what's around WASDE_DATA in the file
        idx = html.find('WASDE_DATA')
        if idx >= 0:
            print(f"  Found 'WASDE_DATA' at position {idx}")
            snippet = html[idx:idx+80].replace('\n', '\\n').replace('\r', '\\r')
            print(f"  Snippet: {snippet}")
            # Find END DATA
            eidx = html.find('END DATA')
            if eidx >= 0:
                before_end = html[max(0,eidx-40):eidx+30].replace('\n', '\\n').replace('\r', '\\r')
                print(f"  Found 'END DATA' at position {eidx}")
                print(f"  Context: {before_end}")
            else:
                print("  'END DATA' NOT FOUND in file!")
        else:
            print("  'WASDE_DATA' NOT FOUND in file!")
        print("WARNING: Could not find WASDE_DATA block to replace!")
        return False
    
    replacement = f'const WASDE_DATA = {data_json};\n// ========== END DATA =========='
    new_html = html[:match.start()] + replacement + html[match.end():]
    
    TEMPLATE_PATH.write_text(new_html, encoding='utf-8')
    print(f"  Updated {TEMPLATE_PATH}")
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
    
    # Verify soybean meal
    meal = data['soybeans']['meal']
    print(f"\nSoy meal — Production: {meal['production']}  Price: {meal['price']}  Exports: {meal['exports']}")
    if all(v == 0 for sublist in [meal['production'], meal['price'], meal['exports']] for v in sublist):
        print("  ⚠️  Soybean meal still showing zeros — check label matching against USDA spreadsheet")
    else:
        print("  ✅ Soybean meal data populated!")
    
    # Verify wheat by class
    wbc = data['wheat']['byClass']
    print(f"Wheat by class production: {wbc['production']} ({len(wbc['production'])} values)")
    
    print("\nUpdating HTML...")
    if update_html(data):
        print("\n✅ Dashboard updated successfully!")
        xls.unlink()
        print(f"  Cleaned up {xls.name}")
    else:
        print("\n❌ Failed to update HTML")
        sys.exit(1)

if __name__ == '__main__':
    main()
