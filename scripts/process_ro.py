#!/usr/bin/env python3
"""
process_ro.py — RO (Renewables Obligation) update processing for the RES Database.

Takes the previous cycle's RO_Pre-loading.xlsx plus two freshly downloaded RER
exports (Accredited Stations, Certificates) and produces paste-ready CSVs:

  1. RO_new_stations_to_append.csv        -> append into RO_Pre-loading.xlsx  'INPUT-RO-Stations'
  2. RO_new_certificates_to_append.csv     -> append into RO_Pre-loading.xlsx  'INPUT-RO-Certificates'
  3. RES_DB_paste_INPUT-RO_U8.csv          -> paste as VALUES into RES DB .xlsb  'INPUT-RO'!U8

Why CSVs and not a rewritten .xlsx directly: the pre-loading workbook is large
enough (400k+ existing certificate rows) that loading + re-saving it with
openpyxl reliably exhausts memory in constrained environments. See
local_excel_update.py for the script that performs the actual writes locally
using real Excel (no such memory ceiling there).

Usage:
    python process_ro.py \
        --old-workbook RO_Pre-loading.xlsx \
        --new-stations RO_Accredited_Stations_<date>.csv \
        --new-certificates RO_Certificates_<date>.csv \
        --out-dir ./output

    (--capacity-cutoff-date is deliberately omitted: the default reuses the
    workbook's frozen baseline, which is what a normal cycle wants.)

Notes on the logic (see references/column_mappings.md for full detail):
  - Stations: append-only, keyed by 'Accreditation reference'. Excludes
    Northern Ireland (by the new export's own Country field) — RER's public
    export now includes NI and micro-scale (<=50kW) generation your database
    has never tracked; excluding NI removes ~all of that scope creep.
  - Certificates: append-only, keyed by 'Start certificate number'. RER's
    Certificates export only goes back to ~2015, so this MUST be append-only
    — a full rebuild would destroy pre-2015 history that only exists in your
    existing workbook. Also excludes NI by the certificate row's own Country
    field (independent of the station list, since some certs reference
    stations no longer in the current Accredited Stations snapshot).
  - OUTPUT-RO computation (per station, over the FULL combined station list):
      Capacity (MW)        = max(station capacity, latest cert capacity as of
                              cutoff date) / 1000
      Date Last Certificate = max(Issue Date) across all certs for that
                              accreditation number
      Date First Certificate= min(Issue Date) across all certs
      Historical avg gen    = average of "Valid Certificates Banding (MWh)"
                              summed per year, over up to 5 years ending the
                              year before Date Last Certificate, excluding
                              zero years, replicating the workbook's original
                              LET/MAXIFS array formula.
"""
import argparse
import csv
import pickle
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

import openpyxl


# ---------- parsing helpers ----------

MONTH_MAP = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
             'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}

EXCEL_EPOCH = date(1899, 12, 30)


def parse_output_period(s):
    """'Sep 2022' -> date(2022, 9, 1)"""
    try:
        mon, yr = s.split()
        return date(int(yr), MONTH_MAP[mon], 1)
    except Exception:
        return None


def parse_uk_date(s):
    if not s or s in ('N/A', ''):
        return None
    try:
        return datetime.strptime(s.strip(), '%d/%m/%Y').date()
    except Exception:
        return None


def parse_num(s):
    if s is None:
        return None
    s = str(s).replace(',', '').strip()
    if s in ('', 'N/A'):
        return None
    try:
        return float(s)
    except Exception:
        return None


def to_date(v):
    """Normalize an openpyxl cell value (datetime, date, or Excel serial int) to a date."""
    if v is None:
        return None
    if hasattr(v, 'date'):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        try:
            from datetime import timedelta
            return EXCEL_EPOCH + timedelta(days=int(v))
        except Exception:
            return None
    return None


def fmt_date(d):
    d = to_date(d) if not isinstance(d, date) or hasattr(d, 'date') else d
    return d.strftime('%d/%m/%Y') if d else ''


# ---------- loading existing workbook data ----------

def load_old_stations(path):
    """Returns dict: accreditation_number -> tuple(status, gen_station, scheme, capacity,
    country, tech, contract, acc_date, comm_date, org)"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['INPUT-RO-Stations']
    old_stations = {}
    for row in ws.iter_rows(min_row=8, values_only=True):
        acc = row[2]
        if not acc:
            continue
        old_stations[acc] = row[2:16]
    wb.close()
    return old_stations


def load_old_certs(path):
    """Returns list of raw row tuples from INPUT-RO-Certificates (0-indexed from col A)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['INPUT-RO-Certificates']
    rows = []
    for row in ws.iter_rows(min_row=9, values_only=True):
        if row[13]:  # Start Certificate No.
            rows.append(row)
    wb.close()
    return rows


# ---------- new download processing ----------

def load_new_stations(path):
    with open(path, encoding='utf-8-sig') as f:
        return list(csv.DictReader(f))


def load_new_certs(path):
    with open(path, encoding='utf-8-sig') as f:
        return list(csv.DictReader(f))


def build_appended_stations(old_stations, new_stations):
    """Append-only, excluding stations already present and excluding Northern Ireland."""
    appended = []
    for r in new_stations:
        acc = r['Accreditation reference']
        if acc in old_stations:
            continue
        if r['Country'] == 'Northern Ireland':
            continue
        appended.append({
            'Accreditation Number': acc,
            'Status': '',  # not present in the current RER export
            'Generating Station': r['Station name'],
            'Scheme': 'RO',
            'Capacity': parse_num(r['Station DNC in kW']),
            'Country': r['Country'],
            'Technology': r['Technology'],
            'Contract Type': r['Contract type'],
            'Accreditation Date': parse_uk_date(r['Effective date']),
            'Commission Date': parse_uk_date(r['Commissioning date']),
            'Organisation': r['Organisation name'],
        })
    return appended


def build_appended_certs(old_cert_keys, new_certs):
    """Append-only, excluding certs already present (by Start certificate number) and
    excluding Northern Ireland (by the cert row's own Country field)."""
    out = []
    for r in new_certs:
        key = r['Start certificate number']
        if key in old_cert_keys:
            continue
        if r['Country'] == 'Northern Ireland':
            continue
        out.append(r)
    return out


# ---------- OUTPUT-RO computation ----------

def normalize_all_certs(old_cert_rows, new_certs):
    """Combine old (raw workbook tuples) and new (RER export dicts) certs into one
    common structure for per-station aggregation."""
    all_certs = []
    for row in old_cert_rows:
        # 0-idx: 4=acc,6=capacity,11=output_period,12=numcerts,15=mwhpercert,16=issuedate,17=status
        all_certs.append({
            'acc': row[4], 'capacity': row[6],
            'output_period': to_date(row[11]),
            'num_certs': row[12], 'mwh_per_cert': row[15],
            'issue_date': to_date(row[16]),
            'status': (row[17] or '').strip(),
        })
    for row in new_certs:
        all_certs.append({
            'acc': row['Accreditation reference'],
            'capacity': parse_num(row['Station TIC in kW']),
            'output_period': parse_output_period(row['Output period']),
            'num_certs': parse_num(row['Number of certificates']),
            'mwh_per_cert': parse_num(row['MWh per certificate']),
            'issue_date': parse_uk_date(row['Issue Date']),
            'status': (row['Cert status'] or '').strip(),
        })
    return all_certs


def compute_output_rows(old_stations, appended_stations, by_acc, by_acc_cap, cutoff_date):

    def station_capacity_mw(acc, station_capacity_kw):
        # Use old-certs-only pool so newly-appended certs with updated TIC values
        # don't retroactively inflate capacity for existing stations.
        # cutoff_date is a frozen baseline guarding the same thing: RER restates
        # Station TIC on every cert row, so trusting all vintages would let a
        # later re-registration raise a station's historic capacity. Don't
        # advance it per cycle. station_mw below is NOT cutoff-filtered, so real
        # uprates still come through via the max().
        certs = by_acc_cap.get(acc, [])
        cert_caps = [c['capacity'] for c in certs
                     if c['issue_date'] and c['issue_date'] <= cutoff_date and c['capacity']]
        cert_max = max(cert_caps) / 1000 if cert_caps else 0
        station_mw = (station_capacity_kw or 0) / 1000
        return max(cert_max, station_mw)

    def date_last_first_cert(acc):
        certs = by_acc.get(acc, [])
        issue_dates = [c['issue_date'] for c in certs if c['issue_date']]
        if not issue_dates:
            return None, None
        return max(issue_dates), min(issue_dates)

    def historical_avg_generation(acc, date_last_cert):
        if not date_last_cert:
            return 0
        y = date_last_cert.year
        certs = by_acc.get(acc, [])
        yearly_sums = {}
        for k in range(1, 6):
            target_year = y - k
            total = 0
            for c in certs:
                op = c['output_period']
                if op and op.year == target_year:
                    if c['status'] in ('Redeemed', 'Issued') and c['num_certs'] and c['mwh_per_cert']:
                        total += c['num_certs'] * c['mwh_per_cert']
            yearly_sums[target_year] = total
        nonzero = [v for v in yearly_sums.values() if v != 0]
        if not nonzero:
            return 0
        return sum(nonzero) / len(nonzero)

    def build_row(acc, status, gen_station, scheme, capacity_kw, country, tech,
                   contract, acc_date, comm_date, org):
        cap_mw = station_capacity_mw(acc, capacity_kw)
        dl, df = date_last_first_cert(acc)
        hist_avg = historical_avg_generation(acc, dl)
        return {
            'Accreditation Number': acc, 'Status': status, 'Generating Station': gen_station,
            'Scheme': scheme, 'Capacity': capacity_kw, 'Country': country, 'Technology': tech,
            'Contract Type': contract, 'Accreditation Date': acc_date, 'Commission Date': comm_date,
            'Organisation': org, 'Capacity (MW)': round(cap_mw, 4),
            'Date Last Certificate': dl, 'Date First Certificate': df,
            'Historical avg generation (MWh/year)': round(hist_avg, 2),
        }

    output_rows = []
    for acc, v in old_stations.items():
        output_rows.append(build_row(acc, v[1], v[2], v[3], v[4], v[5], v[6], v[7], v[8], v[9], v[10]))
    for s in appended_stations:
        output_rows.append(build_row(
            s['Accreditation Number'], s['Status'], s['Generating Station'], s['Scheme'],
            s['Capacity'], s['Country'], s['Technology'], s['Contract Type'],
            s['Accreditation Date'], s['Commission Date'], s['Organisation']))
    return output_rows


# ---------- CSV writers ----------

def write_new_stations_csv(path, appended_stations):
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['Accreditation Number', 'Status', 'Generating Station', 'Scheme', 'Capacity',
                    'Country', 'Technology', 'Contract Type', 'Accreditation Date',
                    'Commission Date', 'Organisation'])
        for s in appended_stations:
            w.writerow([s['Accreditation Number'], s['Status'], s['Generating Station'], s['Scheme'],
                        s['Capacity'], s['Country'], s['Technology'], s['Contract Type'],
                        fmt_date(s['Accreditation Date']), fmt_date(s['Commission Date']),
                        s['Organisation']])


def write_new_certs_csv(path, appended_certs):
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['Year', 'Valid Certificates Banding (MWh)', '', 'Accreditation Number',
                    'Generating Station', 'Capacity', 'Scheme', 'Country', 'Technology Group',
                    'Generation Type', 'Output period', 'No. of Certificates',
                    'Start Certificate No.', 'End Certificate No.', 'MWh per certificate',
                    'Issue Date', 'Certificate Status', 'Status Date',
                    'Current Holder Organisation Name', 'Company Registration No.'])
        for c in appended_certs:
            op = parse_output_period(c['Output period'])
            issue_date = parse_uk_date(c['Issue Date'])
            status_date = parse_uk_date(c['Status date'])
            num_certs = parse_num(c['Number of certificates'])
            mwh_per_cert = parse_num(c['MWh per certificate'])
            status = (c['Cert status'] or '').strip()
            year_val = date(op.year, 12, 31) if op else None
            valid_banding = (num_certs * mwh_per_cert) if (
                status in ('Redeemed', 'Issued') and num_certs and mwh_per_cert) else 0
            w.writerow([
                fmt_date(year_val), valid_banding, '',
                c['Accreditation reference'], c['Station name'], parse_num(c['Station TIC in kW']),
                c['Scheme'], c['Country'], c['Technology'], c['Generation Type'],
                c['Output period'], num_certs, c['Start certificate number'], c['End certificate number'],
                mwh_per_cert, fmt_date(issue_date), status, fmt_date(status_date),
                c['Current holder org name'], c['Current holder company reg number'],
            ])


def write_output_ro_csv(path, output_rows):
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['Accreditation Number', 'Status', 'Generating Station', 'Scheme', 'Capacity',
                    'Country', 'Technology', 'Contract Type', 'Accreditation Date', 'Commission Date',
                    'Organisation', 'Capacity (MW)', 'Date Last Certificate', 'Date First Certificate',
                    'Historical avg generation (MWh/year)'])
        for r in output_rows:
            w.writerow([
                r['Accreditation Number'], r['Status'], r['Generating Station'], r['Scheme'],
                r['Capacity'], r['Country'], r['Technology'], r['Contract Type'],
                fmt_date(r['Accreditation Date']), fmt_date(r['Commission Date']), r['Organisation'],
                r['Capacity (MW)'], fmt_date(r['Date Last Certificate']),
                fmt_date(r['Date First Certificate']), r['Historical avg generation (MWh/year)'],
            ])


# ---------- main ----------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--old-workbook', required=True, help='Path to previous cycle RO_Pre-loading.xlsx')
    ap.add_argument('--new-stations', required=True, help='Path to freshly downloaded Accredited Stations CSV')
    ap.add_argument('--new-certificates', required=True, help='Path to freshly downloaded Certificates CSV')
    ap.add_argument('--capacity-cutoff-date', default=None,
                     help='YYYY-MM-DD. Defaults to reusing the cutoff already in the old workbook '
                          '(OUTPUT-RO!C5) if it can be read; otherwise falls back to 2023-12-31. '
                          'This is a frozen baseline, NOT a per-cycle setting — advancing it '
                          're-bases historic capacity retroactively. Only pass it for a '
                          'deliberate re-baselining.')
    ap.add_argument('--out-dir', default='.', help='Directory to write output CSVs into')
    args = ap.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    print('Loading previous cycle workbook...')
    old_stations = load_old_stations(args.old_workbook)
    old_cert_rows = load_old_certs(args.old_workbook)
    old_cert_keys = {row[13] for row in old_cert_rows}
    print(f'  {len(old_stations)} existing stations, {len(old_cert_rows)} existing certificate rows')

    if args.capacity_cutoff_date:
        cutoff = datetime.strptime(args.capacity_cutoff_date, '%Y-%m-%d').date()
    else:
        wb = openpyxl.load_workbook(args.old_workbook, read_only=True, data_only=True)
        cell_val = wb['OUTPUT-RO']['C5'].value
        wb.close()
        cutoff = to_date(cell_val) or date(2023, 12, 31)
    print(f'  Using capacity cutoff date: {cutoff}')

    print('Loading new downloads...')
    new_stations = load_new_stations(args.new_stations)
    new_certs = load_new_certs(args.new_certificates)
    print(f'  {len(new_stations)} station rows, {len(new_certs)} certificate rows in new downloads')

    appended_stations = build_appended_stations(old_stations, new_stations)
    appended_certs = build_appended_certs(old_cert_keys, new_certs)
    print(f'  -> {len(appended_stations)} genuinely new stations to append (ex-NI)')
    print(f'  -> {len(appended_certs)} genuinely new certificate rows to append (ex-NI)')

    all_certs = normalize_all_certs(old_cert_rows, appended_certs)
    by_acc = defaultdict(list)
    for c in all_certs:
        if c['acc']:
            by_acc[c['acc']].append(c)

    # Separate pool containing only old-workbook certs, used exclusively for
    # capacity computation so newly-appended certs with revised TIC values
    # cannot retroactively inflate existing stations' capacities.
    old_certs_only = normalize_all_certs(old_cert_rows, [])
    by_acc_cap = defaultdict(list)
    for c in old_certs_only:
        if c['acc']:
            by_acc_cap[c['acc']].append(c)

    output_rows = compute_output_rows(old_stations, appended_stations, by_acc, by_acc_cap, cutoff)
    print(f'  Computed OUTPUT-RO for {len(output_rows)} total stations')

    write_new_stations_csv(out_dir / 'RO_new_stations_to_append.csv', appended_stations)
    write_new_certs_csv(out_dir / 'RO_new_certificates_to_append.csv', appended_certs)
    write_output_ro_csv(out_dir / 'RES_DB_paste_INPUT-RO_U8.csv', output_rows)

    print(f'\nDone. Files written to {out_dir}:')
    print('  RO_new_stations_to_append.csv       -> paste into RO_Pre-loading.xlsx  INPUT-RO-Stations (append below last row)')
    print('  RO_new_certificates_to_append.csv   -> paste into RO_Pre-loading.xlsx  INPUT-RO-Certificates (append below last row)')
    print('  RES_DB_paste_INPUT-RO_U8.csv        -> paste as VALUES into RES DB .xlsb  \'INPUT-RO\'!U8')
    print('\nOr hand all three to local_excel_update.py to write them automatically.')


if __name__ == '__main__':
    main()
