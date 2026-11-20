# Download instructions

## FiT Installation Report (public, no login)

1. Go to: https://www.ofgem.gov.uk/feed-tariffs-fit/contacts-guidance-and-resources/public-reports-and-data-fit/installation-reports
2. Find the most recent quarter's report (published quarterly).
3. Download all parts (currently split into 3 `.xlsx` files due to size, each
   ~30-36MB). File names look like:
   `feed_in_tariff_installation_report_part_1_<month>_<year>....xlsx`, etc.
4. Save all 3 parts to `C:\prmf-process\res-db-update\working\`.

Put the files on disk rather than attaching them to the chat — at ~30-36MB each
they'd consume an enormous amount of context, and the processing scripts read
them straight off disk anyway.

This report is a full cumulative snapshot each quarter, not incremental — the
processing script rebuilds the `Installation` tab from scratch each cycle,
matching the pre-loading workbook's own `Intro` tab instructions.

## RO Certificates & Accredited Stations (requires RER login)

1. Log into https://rer.ofgem.gov.uk/ with your Aurora RER account.
2. Click **"Renewable Energy Register"**.
3. Click **"View Public Reports"**.
4. Download the **"Certificates (RO)"** report CSV.
5. Download the **"Accredited Stations (RO)"** report CSV.
6. Save both files to `C:\prmf-process\res-db-update\working\`.

Note `working\` (fresh downloads), not `files_to_update\` — the latter holds
only the previous cycle's three workbooks, which the pipeline reads and never
writes.

### Important known quirks in this export (as of Aug 2026)

- **Certificates only go back to ~April 2015**, not full history. This is why
  the RO side of the pipeline is strictly append-only — a full rebuild would
  destroy pre-2015 history that only exists in the previous cycle's workbook.
- **Accredited Stations now includes Northern Ireland and micro-scale
  (<=50kW) generation** that the database has never tracked. `process_ro.py`
  filters these out automatically (see `column_mappings.md` for the exact
  logic and the investigation that led to this decision).
- Column names differ from the old pre-loading workbook's expectations (e.g.
  `Accreditation reference` vs `Accreditation Number`). See
  `column_mappings.md` for the full mapping — this is handled automatically by
  `process_ro.py`, documented here in case Ofgem changes the export again.

If Ofgem changes the export format again, the fix is almost always in
`process_ro.py`'s field-name constants near the top of `load_new_stations`/
`build_appended_stations`/`build_appended_certs` — check actual column headers
against what the script expects before assuming something deeper is wrong.

## CfD Register (public, no login)

1. Go to: https://register.lowcarboncontracts.uk/downloadfiles
2. This lists every scheme's daily-ish exports in a table with **Date**,
   **Uploaded At**, **Scheme**, and **Action** columns, newest first. Find the
   most recent row whose **Scheme** column says **"CfD"** (not "LCHA" or
   anything else — the list interleaves several schemes) and click its
   **Download** link.
3. Save the file to `C:\prmf-process\res-db-update\working\`.

Unlike RO/FiT, there's no separate pre-loading workbook for CfD and no
computation step — the downloaded file's columns already match the
`INPUT-CfD` sheet's paste columns (R:BD) directly, so it's a straight
full-replace paste. `local_excel_update.py` verifies the downloaded file's
header row against `INPUT-CfD` row 7 before pasting anything, and aborts with
a clear diff if they don't match — if LCCC ever changes their export's column
layout, that's your signal to check `column_mappings.md` and update the
`CFD_PASTE_COLS` range in `local_excel_update.py` accordingly.
