"""Regenerate the dashboard data files from the master Excel workbook.

Reads the "Vendor Bucketing" sheet and writes:
  Vendor bucketing/vendors.json     - fetched by the dashboard (cache-busted)
  Vendor bucketing/vendors-data.js  - offline fallback for file:// use

Stdlib only: an .xlsx is a zip of XML, so no third-party packages are needed.

Usage:  python tools/build_vendors.py
"""

import html
import json
import os
import re
import sys
import zipfile
from datetime import datetime, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, 'Vendor bucketing')
XLSX = os.path.join(DATA_DIR, 'Vendor Bucketing-Master list-Final.xlsx')
SHEET_NAME = 'Vendor Bucketing'

# Excel column -> JSON field. Header text is verified before extracting.
COLUMNS = {
    'A': ('name', 'Vendor Name'),
    'D': ('ids', 'vendor code'),
    'E': ('orderCount', 'Total Order Count'),
    'F': ('currency', 'Currency'),
    'G': ('totalPOValue', 'Total PO Value'),
    'H': ('engagementLevel', 'Engagement Level'),
    'I': ('tier', 'Vendor Usage Tier'),
    'J': ('serviceTypes', 'Service Details'),
    'K': ('procurementCategory', 'Procurement Category'),
    'L': ('procurementSubCategory', 'Sub Category'),
    'M': ('website', 'Vendor website'),
    'S': ('msme', 'MSME/NON MSME'),
    'T': ('department', 'Department wise'),
}
NUMERIC_INT = {'orderCount'}
NUMERIC_FLOAT = {'totalPOValue'}


def load_shared_strings(z):
    try:
        xml = z.read('xl/sharedStrings.xml').decode('utf-8', errors='ignore')
    except KeyError:
        return []
    return [html.unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S)))
            for si in re.findall(r'<si>(.*?)</si>', xml, re.S)]


def find_sheet_path(z, sheet_name):
    wb = z.read('xl/workbook.xml').decode('utf-8', errors='ignore')
    rid = None
    for name, r in re.findall(r'<sheet[^>]*?name="([^"]+)"[^>]*?r:id="(rId\d+)"', wb):
        if html.unescape(name).strip().lower() == sheet_name.lower():
            rid = r
            break
    if not rid:
        raise SystemExit('Sheet "%s" not found in %s' % (sheet_name, XLSX))
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', errors='ignore')
    target = dict(re.findall(r'Id="(rId\d+)"[^>]*?Target="([^"]+)"', rels))[rid]
    return 'xl/' + target.lstrip('/').replace('xl/', '', 1)


def parse_rows(z, path, shared):
    xml = z.read(path).decode('utf-8', errors='ignore')
    for row in re.findall(r'<row[^>]*>(.*?)</row>', xml, re.S):
        cells = {}
        for m in re.finditer(r'<c[^>]*?r="([A-Z]+)\d+"([^>]*)>(.*?)</c>', row, re.S):
            col, attrs, body = m.groups()
            if 't="inlineStr"' in attrs:
                val = ''.join(re.findall(r'<t[^>]*>(.*?)</t>', body, re.S))
            else:
                v = re.search(r'<v>(.*?)</v>', body, re.S)
                val = v.group(1) if v else ''
                if 't="s"' in attrs and val.isdigit():
                    val = shared[int(val)]
            cells[col] = html.unescape(val).strip()
        if cells:
            yield cells


def to_number(raw, as_int):
    if raw in ('', None):
        return 0 if as_int else 0.0
    try:
        n = float(raw)
    except ValueError:
        return 0 if as_int else 0.0
    return int(round(n)) if as_int else n


def main():
    if not os.path.exists(XLSX):
        raise SystemExit('Workbook not found: %s' % XLSX)

    z = zipfile.ZipFile(XLSX)
    shared = load_shared_strings(z)
    rows = parse_rows(z, find_sheet_path(z, SHEET_NAME), shared)

    header = next(rows)
    for col, (field, expected) in COLUMNS.items():
        actual = header.get(col, '')
        if actual.lower() != expected.lower():
            raise SystemExit(
                'Column %s is "%s" but "%s" was expected. The sheet layout changed \u2014 '
                'update COLUMNS in this script.' % (col, actual, expected))

    vendors = []
    for cells in rows:
        rec = {}
        for col, (field, _) in COLUMNS.items():
            raw = cells.get(col, '')
            if field in NUMERIC_INT:
                rec[field] = to_number(raw, True)
            elif field in NUMERIC_FLOAT:
                rec[field] = to_number(raw, False)
            else:
                rec[field] = raw
        if rec['name']:
            vendors.append(rec)

    if not vendors:
        raise SystemExit('No vendor rows found \u2014 aborting so the live data is not wiped.')

    stamp = datetime.now(timezone.utc).isoformat(timespec='seconds')
    json_path = os.path.join(DATA_DIR, 'vendors.json')
    js_path = os.path.join(DATA_DIR, 'vendors-data.js')

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(vendors, f, indent=2, ensure_ascii=False)

    with open(js_path, 'w', encoding='utf-8') as f:
        f.write('/* Auto-generated from %s on %s \u2014 do not edit by hand */\n'
                % (os.path.basename(XLSX), stamp))
        f.write('window.__VENDOR_DATA__ = ')
        json.dump(vendors, f, ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
        f.write('window.__VENDOR_DATA_BUILT__ = %s;\n' % json.dumps(stamp))

    print('Wrote %d vendors' % len(vendors))
    print('  ' + json_path)
    print('  ' + js_path)
    return 0


if __name__ == '__main__':
    sys.exit(main())
