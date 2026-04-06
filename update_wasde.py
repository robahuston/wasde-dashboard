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
    """Convert cell value to float, return 0 if empty."""
    try:
        return float(val) if val != '' else 0
    except (ValueError, TypeError):
        return 0

def try_read_sheet(wb, name):
    """Try to read a named sheet; return list-of-rows or None if sheet missing."""
    try:
        s = wb.sheet_by_name(name)
        return [[s.cell_value(r, c) for c in range(s.ncols)] for r in range(s.nrows)]
    except Exception:
        return None

def find_section(rows, label):
    """Return index of first row whose col-0 starts with label (case-insensitive)."""
    label_lower = label.lower()
    for i, row in enumerate(rows):
        if str(row[0]).strip().lower().startswith(label_lower):
            return i
    return 0

def get_col_row(rows, label, start=0, end=None):
    """Search col-0 for label between start/end; return [safe(cols 1-4)]."""
    if end is None:
        end = len(rows)
    for i in range(start, end):
        if str(rows[i][0]).strip().lower().startswith(label.lower()):
            return [safe(rows[i][j]) for j in range(1, 5)]
    return [0, 0, 0, 0]

def parse_simple_crop(rows, section_label):
    """Generic parser for crops with corn-style layout (label in col 0, data cols 1-4)."""
    if not rows:
        return {k: [0, 0, 0, 0] for k in ['price', 'production', 'begStocks', 'imports',
                                            'supplyTotal', 'feedResidual', 'exports', 'useTotal', 'endStocks']}
    start = find_section(rows, section_label)

    def get(label):
        return get_col_row(rows, label, start)

    return {
        'price':        get('Avg. Farm Price'),
        'production':   get('Production'),
        'begStocks':    get('Beginning Stocks'),
        'imports':      get('Imports'),
        'supplyTotal':  get('Supply, Total'),
        'feedResidual': get('Feed and Residual'),
        'exports':      get('Exports'),
        'useTotal':     get('Use, Total'),
        'endStocks':    get('Ending Stocks'),
    }

def extract_data(xls_path, year, month):
    wb = xlrd.open_workbook(str(xls_path))

    # Determine report number from first sheet
    txt_sheet = wb.sheet_by_name('WASDE Text')
    report_id = "WASDE"
    for r in range(min(10, txt_sheet.nrows)):
        for c_idx in range(min(5, txt_sheet.ncols)):
            val = str(txt_sheet.cell_value(r, c_idx))
            m = re.search(r'WASDE[- ]+(\d+)', val)
            if m:
                report_id = f"WASDE-{m.group(1)}"
                break

    month_names = ["", "January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"]
    report_date = f"{month_names[month]} {year}"

    def read_sheet(name):
        rows = try_read_sheet(wb, name)
        if rows is None:
            print(f"  WARNING: Sheet '{name}' not found — zeros will be used")
            return []
        return rows

    # -----------------------------------------------------------------------
    # CORN (Page 12)
    # -----------------------------------------------------------------------
    corn_rows = read_sheet('Page 12')
    corn_start = find_section(corn_rows, 'CORN')

    corn = {
        'price':         get_col_row(corn_rows, 'Avg. Farm Price',        corn_start),
        'planted':       get_col_row(corn_rows, 'Area Planted',            corn_start),
        'harvested':     get_col_row(corn_rows, 'Area Harvested',          corn_start),
        'yield':         get_col_row(corn_rows, 'Yield per',               corn_start),
        'begStocks':     get_col_row(corn_rows, 'Beginning Stocks',        corn_start),
        'production':    get_col_row(corn_rows, 'Production',              corn_start),
        'imports':       get_col_row(corn_rows, 'Imports',                 corn_start),
        'supplyTotal':   get_col_row(corn_rows, 'Supply, Total',           corn_start),
        'feedResidual':  get_col_row(corn_rows, 'Feed and Residual',       corn_start),
        'fsi':           get_col_row(corn_rows, 'Food, Seed & Industrial', corn_start),
        'ethanol':       get_col_row(corn_rows, 'Ethanol',                 corn_start),
        'domesticTotal': get_col_row(corn_rows, 'Domestic, Total',         corn_start),
        'exports':       get_col_row(corn_rows, 'Exports',                 corn_start),
        'useTotal':      get_col_row(corn_rows, 'Use, Total',              corn_start),
        'endStocks':     get_col_row(corn_rows, 'Ending Stocks',           corn_start),
    }

    # -----------------------------------------------------------------------
    # SOYBEANS (Page 15)
    # -----------------------------------------------------------------------
    soy_rows = read_sheet('Page 15')
    total_soy_rows = len(soy_rows)
    soy_start  = find_section(soy_rows, 'SOYBEANS')
    oil_start  = find_section(soy_rows, 'SOYBEAN OIL')
    meal_start = find_section(soy_rows, 'SOYBEAN MEAL')
    print(f"  Soy sections — beans: row {soy_start}, oil: row {oil_start}, meal: row {meal_start}, total: {total_soy_rows}")

    def get_soy(label, start, end):
        label_lower = label.lower()
        for i in range(start, end):
            if str(soy_rows[i][0]).strip().lower().startswith(label_lower):
                return [safe(soy_rows[i][j]) for j in range(1, 5)]
        return [0, 0, 0, 0]

    soybeans = {
        'price':         get_soy('Avg. Farm Price',        soy_start, oil_start),
        'planted':       get_soy('Area Planted',            soy_start, oil_start),
        'harvested':     get_soy('Area Harvested',          soy_start, oil_start),
        'yield':         get_soy('Yield per',               soy_start, oil_start),
        'begStocks':     get_soy('Beginning Stocks',        soy_start, oil_start),
        'production':    get_soy('Production',              soy_start, oil_start),
        'imports':       get_soy('Imports',                 soy_start, oil_start),
        'supplyTotal':   get_soy('Supply, Total',           soy_start, oil_start),
        'crushings':     get_soy('Crushings',               soy_start, oil_start),
        'exports':       get_soy('Exports',                 soy_start, oil_start),
        'seed':          get_soy('Seed',                    soy_start, oil_start),
        'residual':      get_soy('Residual',                soy_start, oil_start),
        'useTotal':      get_soy('Use, Total',              soy_start, oil_start),
        'endStocks':     get_soy('Ending Stocks',           soy_start, oil_start),
        'oil': {
            'price':       get_soy('Avg. Price',             oil_start, meal_start),
            'production':  get_soy('Production',             oil_start, meal_start),
            'domesticUse': get_soy('Domestic Disappearance', oil_start, meal_start),
            'biofuel':     get_soy('Biofuel',                oil_start, meal_start),
            'exports':     get_soy('Exports',                oil_start, meal_start),
            'endStocks':   get_soy('Ending Stocks',          oil_start, meal_start),
        },
        'meal': {
            'price':       get_soy('Avg. Price',             meal_start, total_soy_rows),
            'production':  get_soy('Production',             meal_start, total_soy_rows),
            'domesticUse': get_soy('Domestic Disappearance', meal_start, total_soy_rows),
            'exports':     get_soy('Exports',                meal_start, total_soy_rows),
            'endStocks':   get_soy('Ending Stocks',          meal_start, total_soy_rows),
        },
    }

    # -----------------------------------------------------------------------
    # WHEAT (Page 11)
    # -----------------------------------------------------------------------
    wheat_rows = read_sheet('Page 11')

    def get_wheat_row(label, start=0, end=30):
        for i in range(start, min(end, len(wheat_rows))):
            if str(wheat_rows[i][0]).strip().startswith(label):
                return [safe(wheat_rows[i][4]), safe(wheat_rows[i][6]),
                        safe(wheat_rows[i][9]), safe(wheat_rows[i][11])]
        return [0, 0, 0, 0]

    wheat = {
        'price':         get_wheat_row('Avg. Farm Price'),
        'planted':       get_wheat_row('Area Planted'),
        'harvested':     get_wheat_row('Area Harvested'),
        'yield':         get_wheat_row('Yield per'),
        'begStocks':     get_wheat_row('Beginning Stocks'),
        'production':    get_wheat_row('Production'),
        'imports':       get_wheat_row('Imports'),
        'supplyTotal':   get_wheat_row('Supply, Total'),
        'food':          get_wheat_row('Food'),
        'seed':          get_wheat_row('Seed'),
        'feedResidual':  get_wheat_row('Feed and Residual'),
        'domesticTotal': get_wheat_row('Domestic, Total'),
        'exports':       get_wheat_row('Exports'),
        'useTotal':      get_wheat_row('Use, Total'),
        'endStocks':     get_wheat_row('Ending Stocks'),
    }

    # Wheat by class (bottom of Page 11)
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
        for i in range(start, min(start + 15, len(wheat_rows))):
            if str(wheat_rows[i][1]).strip().startswith(label):
                return [safe(wheat_rows[i][3]), safe(wheat_rows[i][5]),
                        safe(wheat_rows[i][7]), safe(wheat_rows[i][8]),
                        safe(wheat_rows[i][10])]
        return [0, 0, 0, 0, 0]

    wheat_by_class = {
        'labels':     ["HRW", "HRS", "SRW", "White", "Durum"],
        'production': get_wbc_row('Production'),
        'exports':    get_wbc_row('Exports'),
        'endStocks':  get_wbc_row('Ending Stocks, Total'),
    }

    # -----------------------------------------------------------------------
    # SORGHUM (Page 13)
    # -----------------------------------------------------------------------
    sorghum = parse_simple_crop(read_sheet('Page 13'), 'SORGHUM')

    # -----------------------------------------------------------------------
    # OATS (Page 14 — may be combined Barley & Oats)
    # -----------------------------------------------------------------------
    oats_rows = read_sheet('Page 14') or read_sheet('Page 14a') or []
    oats_start = find_section(oats_rows, 'OATS') if oats_rows else 0
    oats_end   = len(oats_rows)
    for i in range(oats_start + 1, len(oats_rows)):
        cell = str(oats_rows[i][0]).strip().upper()
        if cell and cell != cell.lower() and len(cell) > 3 and cell[0].isalpha():
            oats_end = i
            break

    def get_oats(label):
        return get_col_row(oats_rows, label, oats_start, oats_end)

    oats = {
        'price':          get_oats('Avg. Farm Price'),
        'production':     get_oats('Production'),
        'begStocks':      get_oats('Beginning Stocks'),
        'imports':        get_oats('Imports'),
        'supplyTotal':    get_oats('Supply, Total'),
        'foodIndustrial': get_oats('Food'),
        'feedResidual':   get_oats('Feed and Residual'),
        'exports':        get_oats('Exports'),
        'useTotal':       get_oats('Use, Total'),
        'endStocks':      get_oats('Ending Stocks'),
    }

    # -----------------------------------------------------------------------
    # RICE — U.S. (Page 18)
    # -----------------------------------------------------------------------
    rice_rows = read_sheet('Page 18')
    rice_start = find_section(rice_rows, 'ALL RICE') if rice_rows else 0
    if rice_start == 0:
        rice_start = find_section(rice_rows, 'RICE') if rice_rows else 0

    def get_rice(label):
        return get_col_row(rice_rows, label, rice_start)

    rice = {
        'price':       get_rice('Avg. Farm Price'),
        'production':  get_rice('Production'),
        'begStocks':   get_rice('Beginning Stocks'),
        'imports':     get_rice('Imports'),
        'supplyTotal': get_rice('Supply, Total'),
        'domesticUse': get_rice('Domestic'),
        'exports':     get_rice('Exports'),
        'useTotal':    get_rice('Use, Total'),
        'endStocks':   get_rice('Ending Stocks'),
    }

    # World Rice (Page 19)
    world_rice_rows = read_sheet('Page 19') or []
    w_rice_start = find_section(world_rice_rows, 'WORLD') if world_rice_rows else 0

    world_rice = {
        'production':  get_col_row(world_rice_rows, 'Production',   w_rice_start),
        'consumption': get_col_row(world_rice_rows, 'Consumption',  w_rice_start),
        'trade':       get_col_row(world_rice_rows, 'Trade',         w_rice_start),
        'endStocks':   get_col_row(world_rice_rows, 'Ending Stocks', w_rice_start),
    }

    # -----------------------------------------------------------------------
    # COTTON — U.S. (Page 20)
    # -----------------------------------------------------------------------
    cotton_rows = read_sheet('Page 20')
    cotton_start = find_section(cotton_rows, 'UPLAND') if cotton_rows else 0
    if cotton_start == 0:
        cotton_start = find_section(cotton_rows, 'COTTON') if cotton_rows else 0

    def get_cotton(label):
        return get_col_row(cotton_rows, label, cotton_start)

    cotton = {
        'price':       get_cotton('Avg. Farm Price'),
        'production':  get_cotton('Production'),
        'begStocks':   get_cotton('Beginning Stocks'),
        'imports':     get_cotton('Imports'),
        'supplyTotal': get_cotton('Supply, Total'),
        'domesticUse': get_cotton('Mill Use'),
        'exports':     get_cotton('Exports'),
        'useTotal':    get_cotton('Use, Total'),
        'endStocks':   get_cotton('Ending Stocks'),
    }

    # World Cotton (Page 21)
    world_cotton_rows = read_sheet('Page 21') or []
    w_cotton_start = find_section(world_cotton_rows, 'WORLD') if world_cotton_rows else 0

    world_cotton = {
        'production':  get_col_row(world_cotton_rows, 'Production',   w_cotton_start),
        'consumption': get_col_row(world_cotton_rows, 'Consumption',  w_cotton_start),
        'trade':       get_col_row(world_cotton_rows, 'Trade',         w_cotton_start),
        'endStocks':   get_col_row(world_cotton_rows, 'Ending Stocks', w_cotton_start),
    }

    # -----------------------------------------------------------------------
    # CANOLA / RAPESEED (world oilseeds — try several page names)
    # -----------------------------------------------------------------------
    canola = None
    for page_name in ['Page 28', 'Page 27', 'Page 26', 'Page 29', 'Page 30']:
        canola_rows = try_read_sheet(wb, page_name)
        if not canola_rows:
            continue
        rap_start = find_section(canola_rows, 'RAPESEED')
        if rap_start == 0:
            rap_start = find_section(canola_rows, 'CANOLA')
        if rap_start > 0:
            canola = {
                'production': get_col_row(canola_rows, 'Production',     rap_start),
                'begStocks':  get_col_row(canola_rows, 'Beginning Stocks', rap_start),
                'crushings':  get_col_row(canola_rows, 'Crush',           rap_start),
                'exports':    get_col_row(canola_rows, 'Exports',         rap_start),
                'useTotal':   get_col_row(canola_rows, 'Use, Total',      rap_start),
                'endStocks':  get_col_row(canola_rows, 'Ending Stocks',   rap_start),
            }
            print(f"  Canola/Rapeseed found on {page_name}")
            break
    if canola is None:
        print("  WARNING: Rapeseed/Canola sheet not found — existing dashboard data preserved")

    # -----------------------------------------------------------------------
    # Convert 4-element [23/24, 24/25, prev_mo, cur_mo] -> 3-element [23/24, 24/25, cur_mo]
    # -----------------------------------------------------------------------
    def to3(arr):
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

    # Save prev-month ending stocks before to3 conversion
    corn_es_prev  = corn['endStocks'][2]     if len(corn['endStocks']) >= 3     else 0
    soy_es_prev   = soybeans['endStocks'][2] if len(soybeans['endStocks']) >= 3 else 0
    wheat_es_prev = wheat['endStocks'][2]    if len(wheat['endStocks']) >= 3    else 0

    prev_months = ["", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov"]
    cur_months  = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    g2_year_labels = ["2023/24", "2024/25 Est.", f"2025/26 {cur_months[month]}"]
    g2_note = f"{report_id} {report_date}"

    def g2_wrap(raw, years, note, unit, extra=None):
        processed = process_dict(raw)
        processed['years']      = years
        processed['reportNote'] = note
        processed['unit']       = unit
        if extra:
            processed.update(extra)
        return processed

    result = {
        'reportId':   report_id,
        'reportDate': report_date,
        'years':      ["2023/24", "2024/25 Est.", "2025/26 Proj."],
        'prevMonth':  prev_months[month],
        'curMonth':   cur_months[month],
        'corn':       process_dict(corn),
        'soybeans':   process_dict(soybeans),
        'wheat':      process_dict(wheat),
        'rice':    g2_wrap(rice,    g2_year_labels, g2_note, 'Mil Cwt',
                           extra={'world': process_dict(world_rice)}),
        'cotton':  g2_wrap(cotton,  g2_year_labels, g2_note, 'Mil 480-lb Bales',
                           extra={'world': process_dict(world_cotton)}),
        'sorghum': g2_wrap(sorghum, g2_year_labels, g2_note, 'Mil Bu'),
        'oats':    g2_wrap(oats,    g2_year_labels, g2_note, 'Mil Bu'),
    }

    # Wheat by class kept separate (5-element arrays must not go through process_dict)
    result['wheat']['byClass'] = wheat_by_class
    result['corn']['endStocksPrev']     = corn_es_prev
    result['soybeans']['endStocksPrev'] = soy_es_prev
    result['wheat']['endStocksPrev']    = wheat_es_prev

    if canola is not None:
        result['canola'] = g2_wrap(canola, g2_year_labels, g2_note, 'MMT')
    # If canola is None, update_html will preserve existing canola data from the file

    return result


def update_html(data):
    html = TEMPLATE_PATH.read_text(encoding='utf-8')

    # If canola wasn't found, preserve existing block from the current HTML
    if 'canola' not in data:
        canola_match = re.search(
            r'"canola"\s*:\s*(\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\})',
            html, flags=re.DOTALL
        )
        if canola_match:
            try:
                data['canola'] = json.loads(canola_match.group(1))
            except Exception:
                pass

    data_json = json.dumps(data, indent=2)

    pattern = r'const\s+WASDE_DATA\s*=\s*\{.*?\}\s*;\s*\n\s*//\s*=+\s*END\s+DATA\s*=+'
    match = re.search(pattern, html, flags=re.DOTALL)
    if not match:
        idx = html.find('WASDE_DATA')
        if idx >= 0:
            print(f"  Found 'WASDE_DATA' at pos {idx}")
            eidx = html.find('END DATA')
            print(f"  'END DATA' at pos {eidx}" if eidx >= 0 else "  'END DATA' NOT FOUND")
        else:
            print("  'WASDE_DATA' NOT FOUND in file!")
        print("WARNING: Could not find WASDE_DATA block to replace!")
        return False

    replacement = f'const WASDE_DATA = {data_json};\n// ========== END DATA =========='
    TEMPLATE_PATH.write_text(html[:match.start()] + replacement + html[match.end():], encoding='utf-8')
    print(f"  Updated {TEMPLATE_PATH}")
    return True


def main():
    parser = argparse.ArgumentParser(description='Update WASDE Dashboard')
    now = datetime.now()
    parser.add_argument('--year',  type=int, default=now.year)
    parser.add_argument('--month', type=int, default=now.month)
    args = parser.parse_args()

    xls = download_wasde(args.year, args.month)
    if not xls:
        sys.exit(1)

    print("Extracting data...")
    data = extract_data(xls, args.year, args.month)

    print(f"\nReport: {data['reportId']} — {data['reportDate']}")
    print(f"Corn  price: ${data['corn']['price'][2]}/bu  Ending stocks: {data['corn']['endStocks'][2]} mil bu")
    print(f"Soy   price: ${data['soybeans']['price'][2]}/bu  Ending stocks: {data['soybeans']['endStocks'][2]} mil bu")
    print(f"Wheat price: ${data['wheat']['price'][2]}/bu  Ending stocks: {data['wheat']['endStocks'][2]} mil bu")

    meal = data['soybeans']['meal']
    print(f"\nSoy meal — Production: {meal['production']}  Price: {meal['price']}  Exports: {meal['exports']}")
    if all(v == 0 for lst in [meal['production'], meal['price'], meal['exports']] for v in lst):
        print("  ⚠️  Soybean meal still zeros — check label matching against USDA spreadsheet")
    else:
        print("  ✅ Soybean meal populated!")

    print(f"Rice     production:  {data['rice']['production']}")
    print(f"Cotton   endStocks:   {data['cotton']['endStocks']}")
    print(f"Sorghum  exports:     {data['sorghum']['exports']}")
    print(f"Oats     production:  {data['oats']['production']}")
    if 'canola' in data:
        print(f"Canola   production:  {data['canola']['production']}")
    else:
        print("  ⚠️  Canola not found — existing data preserved")

    wbc = data['wheat']['byClass']
    print(f"Wheat by class production: {wbc['production']}")

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
