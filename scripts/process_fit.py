#!/usr/bin/env python3
"""
process_fit.py — FiT (Feed-in Tariff) update processing for the RES Database.

Takes the previous cycle's Fit_Pre-loading.xlsx (for its REGION.PC lookup tab)
plus the 3 freshly downloaded FiT Installation Report parts, and produces:

  1. FiT_Installation_tab_full_replace.csv  -> full replace of Fit_Pre-loading.xlsx 'Installation' tab
                                               (from row 6; includes pre-computed columns A-H,
                                               so no formula-dragging needed)
  2. RES_DB_paste_INPUT-FiT_K8.csv           -> paste as VALUES into RES DB .xlsb 'INPUT-FiT'!K8

Unlike RO, the FiT Installation Report is a FULL cumulative snapshot each
quarter (not incremental) — per the workbook's own 'Intro' tab instructions,
you delete all existing data and paste the fresh full download each cycle.

Column formulas replicated in Python (see references/column_mappings.md):
  A Tariff Code   = part of TariffCode after the 2nd "/"
  B PV Letter     = 'H'/'M'/'L' depending on which appears in Tariff Code, else ''
  C Tech & Cap    = part of TariffCode before the 2nd "/", + PV Letter
  D Capacity (MW) = Installed capacity (kW) / 1000
  E Start Month   = first-of-month(Commissioning date, else Application date),
                    floored at 2010-04-01
  F Clean postcode= postcode with spaces removed
  G Zone          = lookup of clean postcode against REGION.PC (only for PV/Wind)
  H Region        = lookup of postcode AREA (or GOR+LocalAuthority fallback)
                    against REGION.PC (only for PV/Wind)

The final OUTPUT-equivalent is a groupby (Technology, Region, Zone,
Start Month) summing Capacity (MW) — matching the workbook's pivot table.
'Tech & Cap' in this OUTPUT is NOT one of the groupby keys: it is forced to
the constant 'Utility-scale' for every row, overwriting whatever the
per-installation C_TechCap value would otherwise contribute. This only
affects the pivoted OUTPUT/INPUT-FiT paste — the Installation tab's own
column C (C_TechCap) is left as genuinely computed from TariffCode.

Usage:
    python process_fit.py \
        --old-workbook Fit_Pre-loading.xlsx \
        --report-parts part1.xlsx part2.xlsx part3.xlsx \
        --out-dir ./output
"""
import argparse
import pickle
from pathlib import Path

import openpyxl
import pandas as pd

RAW_COLS = [
    'Extension (Y/N)', 'PostCode', 'Technology', 'Installed capacity', 'Declared net capacity',
    'Application date', 'Commissioning date', 'MCS issue date', 'Export status', 'TariffCode',
    'Tariff Description', 'Installation Type', 'Installation Country', 'Government Office Region',
    'Local Authority', 'Constituency', 'Accreditation Route', 'MPAN Prefix',
    'Community school category', 'LLSOA Code', 'MLSOA Code', 'FIT ID',
]


def load_region_lookup(old_workbook_path):
    """Build the three lookup tables from the REGION.PC tab:
      area_lookup:      postcode area (e.g. 'NP') -> {'Photovoltaic': region, 'Wind': region}
      combined_lookup:  "GOR-LocalAuthority" -> {'Photovoltaic': region, 'Wind': region}  (fallback)
      postcode_zone:    full clean postcode -> Zone number
    """
    wb = openpyxl.load_workbook(old_workbook_path, read_only=True, data_only=True)
    ws = wb['REGION.PC']
    rows = list(ws.iter_rows(min_row=5, values_only=True))
    wb.close()

    area_lookup = {}
    combined_lookup = {}
    postcode_zone = {}
    for r in rows:
        area = r[0]
        if area:
            area_lookup[area] = {'Photovoltaic': r[3], 'Wind': r[4]}
        comb = r[11]
        if comb:
            combined_lookup[comb] = {'Photovoltaic': r[12], 'Wind': r[13]}
        pc = r[15]
        if pc:
            postcode_zone[pc] = r[16]
    return area_lookup, combined_lookup, postcode_zone


def load_report_parts(paths):
    dfs = []
    for p in paths:
        print(f'  reading {p}')
        df = pd.read_excel(p, sheet_name=0, header=4, engine='openpyxl')
        df.columns = RAW_COLS
        dfs.append(df)
        print(f'    {len(df)} rows')
    return pd.concat(dfs, ignore_index=True)


def tariff_code(tc):
    if not isinstance(tc, str):
        return ''
    parts = tc.split('/')
    return '/'.join(parts[2:]) if len(parts) >= 3 else tc


def pv_letter(a):
    if 'H' in a:
        return 'H'
    if 'M' in a:
        return 'M'
    if 'L' in a:
        return 'L'
    return ''


def tech_cap_base(tc):
    if not isinstance(tc, str):
        return ''
    parts = tc.split('/')
    return '/'.join(parts[:2]) if len(parts) >= 3 else tc


def start_month(row):
    comm = row['Commissioning date']
    app = row['Application date']
    d = comm if pd.notna(comm) else app
    if pd.isna(d):
        return pd.NaT
    d = pd.Timestamp(year=d.year, month=d.month, day=1)
    floor = pd.Timestamp(2010, 4, 1)
    return max(d, floor)


def get_postcode_area(postcode):
    """Postcode area = leading letters (1 if 2nd char is a digit, else 2)."""
    if not isinstance(postcode, str) or len(postcode) < 2:
        return postcode
    n = 1 if postcode[1].isdigit() else 2
    return postcode[:n]


def compute_derived_columns(full, area_lookup, combined_lookup, postcode_zone):
    full['A_TariffCode'] = full['TariffCode'].map(tariff_code)
    full['B_PVLetter'] = full['A_TariffCode'].map(pv_letter)
    full['C_TechCap'] = full['TariffCode'].map(tech_cap_base) + full['B_PVLetter']
    full['D_CapacityMW'] = full['Installed capacity'] / 1000.0
    full['F_CleanPostcode'] = full['PostCode'].fillna('').astype(str).str.replace(' ', '', regex=False)

    # E_StartMonth: vectorized (avoid per-row apply over 800k+ rows)
    comm = pd.to_datetime(full['Commissioning date'], errors='coerce')
    app = pd.to_datetime(full['Application date'], errors='coerce')
    chosen = comm.where(comm.notna(), app)
    first_of_month = chosen.values.astype('datetime64[M]')
    floor = pd.Timestamp(2010, 4, 1).to_datetime64()
    full['E_StartMonth'] = pd.Series(first_of_month, index=full.index).clip(lower=floor)
    full.loc[chosen.isna(), 'E_StartMonth'] = pd.NaT

    is_pv_wind = full['Technology'].isin(['Photovoltaic', 'Wind'])

    # G_Zone: vectorized map against postcode_zone dict
    zone_mapped = full['F_CleanPostcode'].map(postcode_zone)
    zone = pd.Series('', index=full.index, dtype=object)
    valid_pc = is_pv_wind & full['F_CleanPostcode'].notna() & ~full['F_CleanPostcode'].isin(['', 'Unknown', 'nan'])
    zone.loc[valid_pc] = zone_mapped.loc[valid_pc].apply(lambda z: z if pd.notna(z) else 'Missing postcode')
    full['G_Zone'] = zone

    # H_Region: vectorized two-stage lookup (postcode area, then GOR+LocalAuthority fallback)
    area = full['PostCode'].map(get_postcode_area)
    comb_key = full['Government Office Region'].astype(str) + '-' + full['Local Authority'].astype(str)

    def region_for_tech(tech_name):
        area_map = {k: v.get(tech_name) for k, v in area_lookup.items()}
        comb_map = {k: v.get(tech_name) for k, v in combined_lookup.items()}
        primary = area.map(area_map)
        fallback = comb_key.map(comb_map)
        # Faithful to the original workbook's IFERROR(primary_lookup, fallback_lookup):
        # only fall back when the primary KEY is missing (i.e. the MATCH would error in
        # Excel) - a successful lookup that happens to return 0 is a legitimate answer,
        # not a signal to fall back further.
        return primary.where(primary.notna(), fallback)

    region_pv = region_for_tech('Photovoltaic')
    region_wind = region_for_tech('Wind')
    region = pd.Series('', index=full.index, dtype=object)
    region.loc[full['Technology'] == 'Photovoltaic'] = region_pv.loc[full['Technology'] == 'Photovoltaic']
    region.loc[full['Technology'] == 'Wind'] = region_wind.loc[full['Technology'] == 'Wind']
    region = region.fillna('')
    full['H_Region'] = region

    return full


def build_output_pivot(full):
    # NOTE: 'Tech & Cap' is intentionally NOT a groupby key here. The
    # Installation tab's own Tech & Cap (C_TechCap, computed from TariffCode)
    # is left untouched, but the pivoted OUTPUT that feeds INPUT-FiT!K8 always
    # reports 'Utility-scale' for this column regardless of what C_TechCap
    # says - this collapses what would otherwise be separate Tech & Cap groups
    # within the same Technology/Region/Zone/Start Month into one row.
    grouped = full.groupby(
        ['Technology', 'H_Region', 'G_Zone', 'E_StartMonth'], dropna=False
    )['D_CapacityMW'].sum().reset_index()
    grouped.columns = ['Technology', 'Region', 'Zone', 'Start Month', 'Sum of Capacity']
    grouped.insert(3, 'Tech & Cap', 'Utility-scale')
    # Zone comes through as float (NaN-compatible) during grouping; render whole-number zones
    # as plain ints for a clean match to the workbook's own formatting, leaving text
    # ('Missing postcode') and blanks untouched.
    def clean_zone(z):
        if isinstance(z, float) and z.is_integer():
            return int(z)
        return z
    grouped['Zone'] = grouped['Zone'].map(clean_zone)
    return grouped.sort_values(['Technology', 'Region', 'Zone', 'Tech & Cap', 'Start Month'])


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--old-workbook', required=True,
                     help='Path to previous cycle Fit_Pre-loading.xlsx (used only for its REGION.PC tab)')
    ap.add_argument('--report-parts', required=True, nargs='+',
                     help='Paths to the 3 (or however many) freshly downloaded FiT Installation Report parts')
    ap.add_argument('--out-dir', default='.', help='Directory to write output CSVs into')
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    print('Loading REGION.PC lookup tables...')
    area_lookup, combined_lookup, postcode_zone = load_region_lookup(args.old_workbook)
    print(f'  {len(area_lookup)} postcode areas, {len(combined_lookup)} combined fallback keys, '
          f'{len(postcode_zone)} postcode-zone entries')

    print('Loading report parts...')
    full = load_report_parts(args.report_parts)
    print(f'  total rows: {len(full)}')

    print('Computing derived columns (A-H)...')
    full = compute_derived_columns(full, area_lookup, combined_lookup, postcode_zone)

    installation_path = out_dir / 'FiT_Installation_tab_full_replace.csv'
    cols_out = ['A_TariffCode', 'B_PVLetter', 'C_TechCap', 'D_CapacityMW', 'E_StartMonth',
                'F_CleanPostcode', 'G_Zone', 'H_Region'] + RAW_COLS
    full[cols_out].to_csv(installation_path, index=False)
    print(f'  wrote {installation_path} ({len(full)} rows)')

    print('Building OUTPUT pivot (groupby Technology/Region/Zone/StartMonth; '
          "Tech & Cap forced to 'Utility-scale')...")
    grouped = build_output_pivot(full)
    output_path = out_dir / 'RES_DB_paste_INPUT-FiT_K8.csv'
    grouped.to_csv(output_path, index=False)
    print(f'  wrote {output_path} ({len(grouped)} rows, total {grouped["Sum of Capacity"].sum():.1f} MW)')

    print(f'\nDone. Files written to {out_dir}:')
    print('  FiT_Installation_tab_full_replace.csv -> paste into Fit_Pre-loading.xlsx  Installation tab, '
          'replacing from row 6')
    print('  RES_DB_paste_INPUT-FiT_K8.csv          -> paste as VALUES into RES DB .xlsb  \'INPUT-FiT\'!K8')
    print('\nOr hand both to local_excel_update.py to write them automatically.')


if __name__ == '__main__':
    main()
