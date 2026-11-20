#!/usr/bin/env python3
"""
local_excel_update.py — writes the RES Database update INTO the real Excel files.

This script is meant to run LOCALLY on your own machine (not in a cloud sandbox),
driving a real installed copy of Excel via COM automation (pywin32). That's what
lets it write both the large pre-loading .xlsx files AND the RES DB .xlsb — neither
of which can be written from a memory-constrained sandbox (.xlsb has no writable
Python library at all; the .xlsx files are large enough that loading + resaving
them with openpyxl reliably runs out of memory in that environment).

Prerequisites (one-time):
    pip install pywin32

What it does, in order:
  1. Opens RO_Pre-loading.xlsx, appends new rows to INPUT-RO-Stations and
     INPUT-RO-Certificates from the CSVs produced by process_ro.py, saves as
     "RO_Pre-loading - updated.xlsx".
  2. Opens Fit_Pre-loading.xlsx, clears the Installation tab from row 6 and pastes
     in the full replacement CSV from process_fit.py, saves as
     "Fit_Pre-loading - updated.xlsx".
  3. Opens the RES DB .xlsb:
     a. Pastes the OUTPUT-RO CSV as VALUES into 'INPUT-RO'!U8 (cols U:AI), then
        drags the sheet's existing formulas in cols A:S and AK:LR down to match
        the new bottom row.
     b. Pastes the FiT OUTPUT CSV as VALUES into 'INPUT-FiT'!K8 (cols K:P), then
        drags formulas in cols A:J and R:KY down to match.
     c. Reads the raw CfD register download (no pre-loading file, no other
        preprocessing) and full-replaces 'INPUT-CfD'!R8 (cols R:BD). Matches
        and reorders the download's columns by header name against the
        sheet's row 7 headers first (LCCC's column order isn't stable
        between exports), aborting only if a header is genuinely missing.
        Then drags formulas in cols A:P and BF:MM down to match.
     Saves into --out-dir as "<original name> - updated.xlsb" (or to
     --res-db-output if given).

Dragging formulas down means literally what it sounds like: take the last row
that already had a formula, and AutoFill it down through the new last data
row, so every new row gets the same (relatively-adjusted) formula a manual
fill-handle drag would produce. If this cycle's data is shorter than last
cycle's, the excess formula rows below the new bottom row are cleared instead.

Usage:
    python local_excel_update.py ^
        --ro-pre-loading "C:\\prmf-process\\input_from_prev_cycle\\RO_Pre-loading.xlsx" ^
        --fit-pre-loading "C:\\prmf-process\\input_from_prev_cycle\\Fit_Pre-loading.xlsx" ^
        --res-db "C:\\prmf-process\\input_from_prev_cycle\\RES_Database_Central-8z-Q2_2026.xlsb" ^
        --ro-new-stations RO_new_stations_to_append.csv ^
        --ro-new-certificates RO_new_certificates_to_append.csv ^
        --ro-output-csv RES_DB_paste_INPUT-RO_U8.csv ^
        --fit-installation-csv FiT_Installation_tab_full_replace.csv ^
        --fit-output-csv RES_DB_paste_INPUT-FiT_K8.csv ^
        --cfd-file "C:\\prmf-process\\res-db-update\\working\\cfd_register_download.xlsx" ^
        --out-dir "C:\\prmf-process\\output"

Notes:
  - Excel must be installed and closable/controllable via COM (this is normal for
    a standard local install; no special configuration usually needed).
  - The script makes Excel visible=False and disables alerts, but if something
    goes wrong it's often easiest to debug by setting SHOW_EXCEL=True below and
    watching it work.
  - Any step can be skipped with its --skip-* flag if you only need to do part
    of the update (e.g. --skip-res-db to just refresh the pre-loading files).
"""
import argparse
import csv
import os
import re
import sys
from datetime import datetime as _dt
from pathlib import Path

try:
    import win32com.client as win32
except ImportError:
    print("This script requires pywin32. Install it with:  pip install pywin32", file=sys.stderr)
    sys.exit(1)

SHOW_EXCEL = False  # flip to True for visual debugging

_DD_MM_YYYY = re.compile(r'^(\d{2})/(\d{2})/(\d{4})$')

# Column layout on the RES DB .xlsb's INPUT-* tabs. Each tab has a paste range
# (raw values overwritten each cycle) and one or more formula ranges (existing
# formulas that must be dragged down to cover the new bottom row, exactly like
# manually pulling the fill handle). There is always a deliberate one-column
# gap between a formula range and the adjacent paste range (e.g. INPUT-RO's
# col T between S and U) - leave those alone, they're outside both ranges on
# purpose and this script never touches them.
#
# Column letters -> (first_col, last_col) as 1-indexed numbers.
RO_PASTE_COLS = (21, 35)                    # U:AI
RO_FORMULA_RANGES = [(1, 19), (37, 330)]    # A:S, AK:LR

FIT_PASTE_COLS = (11, 16)                   # K:P
FIT_FORMULA_RANGES = [(1, 10), (18, 311)]   # A:J, R:KY

CFD_PASTE_COLS = (18, 56)                   # R:BD
CFD_ALLOW_MISSING = False                   # set by --cfd-allow-missing: paste missing headers as blanks
CFD_FORMULA_RANGES = [(1, 16), (58, 351)]   # A:P, BF:MM

_EXCEL_EPOCH = _dt(1899, 12, 30)  # Excel day 0; day 1 == 1900-01-01


def _coerce_dates(rows):
    """Convert 'DD/MM/YYYY' strings to Excel date serial numbers so COM writes
    real Excel dates rather than locale-dependent text strings. Excel would
    otherwise mis-parse days > 12 as text (since it can't treat them as
    MM/DD/YYYY months) and silently mis-parse days <= 12 by swapping day and
    month.

    Serials rather than datetime objects, because passing naive datetimes inside
    a 2-D array to Range.Value goes through pywin32's VARIANT conversion, which:
      * raises OSError EINVAL for any date before 1970 (Windows CRT limit) —
        the RO stations feed carries commission dates back to the 1950s; and
      * treats the naive value as UTC, so a BST date comes back as the previous
        day at 23:00 (2026-07-21 -> 2026-07-20 23:00).
    A plain float has neither problem and round-trips exactly. NumberFormat is
    deliberately left untouched so pasted cells display the same way as the rows
    already above them, matching a manual paste-as-values."""
    out = []
    for row in rows:
        new_row = []
        for v in row:
            if isinstance(v, str):
                m = _DD_MM_YYYY.match(v)
                if m:
                    try:
                        v = float((_dt(int(m.group(3)), int(m.group(2)),
                                       int(m.group(1))) - _EXCEL_EPOCH).days)
                    except ValueError:
                        pass
            new_row.append(v)
        out.append(new_row)
    return out


def start_excel():
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = SHOW_EXCEL
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.ScreenUpdating = False
    return excel


def read_csv_rows(path):
    with open(path, newline='', encoding='utf-8-sig') as f:
        return list(csv.reader(f))


def paste_rows_at(sheet, rows, start_row, start_col=1):
    """Write a list of row-lists as values starting at (start_row, start_col)."""
    if not rows:
        return
    n_rows = len(rows)
    n_cols = len(rows[0])
    target = sheet.Range(
        sheet.Cells(start_row, start_col),
        sheet.Cells(start_row + n_rows - 1, start_col + n_cols - 1),
    )
    target.Value = rows


_XL_FILL_DEFAULT = 0  # xlFillDefault


def fill_formulas_down(sheet, formula_ranges, from_row, to_row):
    """Drag existing formulas in each (start_col, end_col) range down from
    from_row through to_row - the same operation as selecting the last
    formula row and pulling the fill handle to the new bottom row. Excel
    adjusts relative references per row exactly as a manual drag would.

    If to_row < from_row (this cycle's data is shorter than before), there's
    nothing to drag down to - instead clear the now-unused formula rows so
    stale formulas don't hang below the new data.

    from_row must already contain valid formulas (i.e. be a real previous
    cycle's last row) - if a tab has never been populated before, there's
    nothing to copy from and the caller should skip this entirely.
    """
    for start_col, end_col in formula_ranges:
        if to_row > from_row:
            src = sheet.Range(sheet.Cells(from_row, start_col), sheet.Cells(from_row, end_col))
            dst = sheet.Range(sheet.Cells(from_row, start_col), sheet.Cells(to_row, end_col))
            src.AutoFill(dst, _XL_FILL_DEFAULT)
        elif to_row < from_row:
            sheet.Range(
                sheet.Cells(to_row + 1, start_col), sheet.Cells(from_row, end_col)
            ).ClearContents()
        # to_row == from_row: row count unchanged, nothing to do


def update_ro_pre_loading(excel, path, new_stations_csv, new_certs_csv, out_dir):
    print(f'Opening {path} ...')
    wb = excel.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0)
    try:
        stations_rows = _coerce_dates(read_csv_rows(new_stations_csv)[1:])
        certs_rows = _coerce_dates(read_csv_rows(new_certs_csv)[1:])

        ws_stations = wb.Sheets('INPUT-RO-Stations')
        last_row = ws_stations.Cells(ws_stations.Rows.Count, 3).End(-4162).Row  # xlUp = -4162, col C
        print(f'  INPUT-RO-Stations: appending {len(stations_rows)} rows after row {last_row}')
        paste_rows_at(ws_stations, stations_rows, last_row + 1, start_col=3)  # starts at col C

        ws_certs = wb.Sheets('INPUT-RO-Certificates')
        last_row = ws_certs.Cells(ws_certs.Rows.Count, 5).End(-4162).Row  # col E = Accreditation Number
        print(f'  INPUT-RO-Certificates: appending {len(certs_rows)} rows after row {last_row}')
        paste_rows_at(ws_certs, certs_rows, last_row + 1, start_col=2)  # starts at col B (Year)

        out_path = Path(out_dir) / 'RO_Pre-loading - updated.xlsx'
        wb.SaveAs(str(out_path.resolve()), FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
        print(f'  saved {out_path}')
    finally:
        wb.Close(SaveChanges=False)


def update_fit_pre_loading(excel, path, installation_csv, out_dir):
    print(f'Opening {path} ...')
    wb = excel.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0)
    try:
        excel.Calculation = -4135  # xlCalculationManual — prevent recalc during paste
        rows = _coerce_dates(read_csv_rows(installation_csv)[1:])

        ws = wb.Sheets('Installation')
        last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        if last_row >= 6:
            print(f'  clearing existing Installation data (rows 6:{last_row})')
            ws.Range(ws.Cells(6, 1), ws.Cells(last_row, 30)).ClearContents()

        CHUNK = 10000
        print(f'  writing {len(rows)} rows from row 6 (chunks of {CHUNK})')
        for i in range(0, len(rows), CHUNK):
            chunk = rows[i:i + CHUNK]
            paste_rows_at(ws, chunk, start_row=6 + i, start_col=1)
            if i % 100000 == 0 and i > 0:
                print(f'    {i}/{len(rows)} rows written...')

        out_path = Path(out_dir) / 'Fit_Pre-loading - updated.xlsx'
        wb.SaveAs(str(out_path.resolve()), FileFormat=51)
        print(f'  saved {out_path}')
    finally:
        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception:
            pass
        wb.Close(SaveChanges=False)


def read_source_workbook_rows(excel, path):
    """Open a raw downloaded workbook (e.g. the CfD register export) read-only,
    pull its UsedRange as values, and close it again. Returns (header_row,
    data_rows) where header_row is a list and data_rows is a list of lists."""
    wb = excel.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0, ReadOnly=True)
    try:
        values = wb.Sheets(1).UsedRange.Value  # tuple of tuples
    finally:
        wb.Close(SaveChanges=False)
    rows = [list(r) for r in values]
    return rows[0], rows[1:]


def update_res_db(excel, path, ro_output_csv, fit_output_csv, cfd_file, out_dir,
                   res_db_output, skip_cfd=False, skip_ro=False, skip_fit=False):
    print(f'Opening {path} ...')
    wb = excel.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0)
    try:
        if not skip_ro:
            ro_rows = _coerce_dates(read_csv_rows(ro_output_csv)[1:])
            ws_ro = wb.Sheets('INPUT-RO')
            print(f'  clearing + pasting {len(ro_rows)} rows into INPUT-RO!U8')
            old_last_row = ws_ro.Cells(ws_ro.Rows.Count, RO_PASTE_COLS[0]).End(-4162).Row  # col U = 21
            if old_last_row >= 8:
                ws_ro.Range(ws_ro.Cells(8, RO_PASTE_COLS[0]),
                            ws_ro.Cells(max(old_last_row, 8), RO_PASTE_COLS[1])).ClearContents()  # U:AI
            paste_rows_at(ws_ro, ro_rows, start_row=8, start_col=RO_PASTE_COLS[0])
            new_last_row = (8 + len(ro_rows) - 1) if ro_rows else old_last_row
            if old_last_row >= 8:
                print(f'  dragging INPUT-RO formulas (A:S, AK:LR) from row {old_last_row} to {new_last_row}')
                fill_formulas_down(ws_ro, RO_FORMULA_RANGES, old_last_row, new_last_row)
            else:
                print('  INPUT-RO has no previous data to drag formulas from - skipping formula fill; '
                      'formulas may need to be seeded manually for a first-ever run.')
        else:
            print('  --skip-res-db-ro set - skipping INPUT-RO update')

        if not skip_fit:
            fit_rows = _coerce_dates(read_csv_rows(fit_output_csv)[1:])
            ws_fit = wb.Sheets('INPUT-FiT')
            print(f'  clearing + pasting {len(fit_rows)} rows into INPUT-FiT!K8')
            old_last_row = ws_fit.Cells(ws_fit.Rows.Count, FIT_PASTE_COLS[0]).End(-4162).Row  # col K = 11
            if old_last_row >= 8:
                ws_fit.Range(ws_fit.Cells(8, FIT_PASTE_COLS[0]),
                             ws_fit.Cells(max(old_last_row, 8), FIT_PASTE_COLS[1])).ClearContents()  # K:P
            paste_rows_at(ws_fit, fit_rows, start_row=8, start_col=FIT_PASTE_COLS[0])
            new_last_row = (8 + len(fit_rows) - 1) if fit_rows else old_last_row
            if old_last_row >= 8:
                print(f'  dragging INPUT-FiT formulas (A:J, R:KY) from row {old_last_row} to {new_last_row}')
                fill_formulas_down(ws_fit, FIT_FORMULA_RANGES, old_last_row, new_last_row)
            else:
                print('  INPUT-FiT has no previous data to drag formulas from - skipping formula fill; '
                      'formulas may need to be seeded manually for a first-ever run.')
        else:
            print('  --skip-res-db-fit set - skipping INPUT-FiT update')

        if not skip_cfd and cfd_file:
            ws_cfd = wb.Sheets('INPUT-CfD')
            print(f'  reading CfD source file {cfd_file} ...')
            header_row, cfd_rows = read_source_workbook_rows(excel, cfd_file)

            target_header_range = ws_cfd.Range(
                ws_cfd.Cells(7, CFD_PASTE_COLS[0]), ws_cfd.Cells(7, CFD_PASTE_COLS[1])
            )
            target_headers = list(target_header_range.Value[0])

            # The register download's column order isn't guaranteed to match
            # INPUT-CfD's column order even when the set of headers is the
            # same, so match by name and reorder rather than assuming
            # positional alignment.
            header_index = {}
            for i, h in enumerate(header_row):
                header_index.setdefault(h, i)  # first occurrence wins if a header repeats
            missing = [h for h in target_headers if h not in header_index]
            if missing and CFD_ALLOW_MISSING:
                print(f'  WARNING: CfD download lacks {sorted(set(missing))}; pasting these as BLANK columns '
                      '(--cfd-allow-missing)')
            elif missing:
                raise SystemExit(
                    "CfD download is missing column(s) that INPUT-CfD row 7 (columns R:BD) "
                    "expects - aborting before paste to avoid misaligned data.\n"
                    f"  Missing: {missing}\n"
                    f"  Downloaded headers: {header_row}\n"
                    f"  Expected headers:   {target_headers}\n"
                    "Check whether the LCCC register export changed its column layout, "
                    "and update column_mappings.md / the CFD_PASTE_COLS range if so."
                )
            extra = [h for h in header_row if h not in target_headers]
            if extra:
                print(f'  note: CfD download has {len(extra)} extra column(s) not on INPUT-CfD, '
                      f'ignoring: {extra}')

            col_order = [header_index.get(h) for h in target_headers]
            cfd_rows = [[(row[i] if i is not None else None) for i in col_order] for row in cfd_rows]

            print(f'  clearing + pasting {len(cfd_rows)} rows into INPUT-CfD!R8 (full replace)')
            old_last_row = ws_cfd.Cells(ws_cfd.Rows.Count, CFD_PASTE_COLS[0]).End(-4162).Row  # col R = 18
            if old_last_row >= 8:
                ws_cfd.Range(ws_cfd.Cells(8, CFD_PASTE_COLS[0]),
                             ws_cfd.Cells(max(old_last_row, 8), CFD_PASTE_COLS[1])).ClearContents()  # R:BD
            paste_rows_at(ws_cfd, cfd_rows, start_row=8, start_col=CFD_PASTE_COLS[0])
            new_last_row = (8 + len(cfd_rows) - 1) if cfd_rows else old_last_row
            if old_last_row >= 8:
                print(f'  dragging INPUT-CfD formulas (A:P, BF:MM) from row {old_last_row} to {new_last_row}')
                fill_formulas_down(ws_cfd, CFD_FORMULA_RANGES, old_last_row, new_last_row)
            else:
                print('  INPUT-CfD has no previous data to drag formulas from - skipping formula fill; '
                      'formulas may need to be seeded manually for a first-ever run.')
        elif not cfd_file:
            print('  no --cfd-file given - skipping INPUT-CfD update')
        else:
            print('  --skip-cfd set - skipping INPUT-CfD update')

        if res_db_output:
            out_path = Path(res_db_output)
        else:
            # Match the ' - updated' convention used by the two pre-loading
            # files, keeping the cycle tag from the original name.
            src = Path(path)
            out_path = Path(out_dir) / f'{src.stem} - updated{src.suffix}'
        # xlsb format code is 50 (xlExcel12)
        wb.SaveAs(str(out_path.resolve()), FileFormat=50)
        print(f'  saved {out_path}')
    finally:
        wb.Close(SaveChanges=False)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--ro-pre-loading', help='Path to previous cycle RO_Pre-loading.xlsx')
    ap.add_argument('--fit-pre-loading', help='Path to previous cycle Fit_Pre-loading.xlsx')
    ap.add_argument('--res-db', help='Path to the RES Database .xlsb')

    ap.add_argument('--ro-new-stations', help='CSV from process_ro.py: RO_new_stations_to_append.csv')
    ap.add_argument('--ro-new-certificates', help='CSV from process_ro.py: RO_new_certificates_to_append.csv')
    ap.add_argument('--ro-output-csv', help='CSV from process_ro.py: RES_DB_paste_INPUT-RO_U8.csv')

    ap.add_argument('--fit-installation-csv', help='CSV from process_fit.py: FiT_Installation_tab_full_replace.csv')
    ap.add_argument('--fit-output-csv', help='CSV from process_fit.py: RES_DB_paste_INPUT-FiT_K8.csv')

    ap.add_argument('--cfd-file', help='Path to the raw CfD register download (.xlsx) from '
                                        'register.lowcarboncontracts.uk/downloadfiles - pasted straight '
                                        'into INPUT-CfD!R8 with no preprocessing (columns matched by '
                                        'header name and reordered as needed; same header set required)')

    ap.add_argument('--out-dir', default='.', help='Directory to save updated files into')
    ap.add_argument('--res-db-output', help='Optional explicit output path for the updated .xlsb '
                                             '(defaults to overwriting in --out-dir with the same filename)')

    ap.add_argument('--skip-ro', action='store_true', help='Skip the RO_Pre-loading.xlsx update')
    ap.add_argument('--skip-fit', action='store_true', help='Skip the Fit_Pre-loading.xlsx update')
    ap.add_argument('--skip-res-db', action='store_true', help='Skip the RES DB .xlsb update entirely '
                                                                 '(INPUT-RO, INPUT-FiT, and INPUT-CfD)')
    ap.add_argument('--skip-res-db-ro', action='store_true', help='Within the RES DB .xlsb update, skip '
                                                                    'just the INPUT-RO part')
    ap.add_argument('--skip-res-db-fit', action='store_true', help='Within the RES DB .xlsb update, skip '
                                                                     'just the INPUT-FiT part')
    ap.add_argument('--skip-cfd', action='store_true', help='Within the RES DB .xlsb update, skip just the '
                                                              'INPUT-CfD part')
    ap.add_argument('--cfd-allow-missing', action='store_true', help='Paste INPUT-CfD columns missing from the '
                                                                       'download as blanks instead of aborting')
    args = ap.parse_args()
    global CFD_ALLOW_MISSING
    CFD_ALLOW_MISSING = args.cfd_allow_missing

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    excel = start_excel()
    try:
        if not args.skip_ro:
            update_ro_pre_loading(excel, args.ro_pre_loading, args.ro_new_stations,
                                   args.ro_new_certificates, out_dir)
        if not args.skip_fit:
            update_fit_pre_loading(excel, args.fit_pre_loading, args.fit_installation_csv, out_dir)
        if not args.skip_res_db:
            update_res_db(excel, args.res_db, args.ro_output_csv, args.fit_output_csv,
                           args.cfd_file, out_dir, args.res_db_output, skip_cfd=args.skip_cfd,
                           skip_ro=args.skip_res_db_ro, skip_fit=args.skip_res_db_fit)
        print('\nAll done.')
    finally:
        excel.Quit()


if __name__ == '__main__':
    main()
