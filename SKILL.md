---
name: res-db-update
description: Update the RES Database (a large Excel workbook tracking GB renewable generation assets, including RO, FiT, and CfD schemes) each cycle from fresh RO, FiT, and CfD source downloads. Trigger this whenever the user asks to "update the RES database", "run the RES DB update", "refresh the RES database", mentions RO Pre-loading.xlsx / Fit Pre-loading.xlsx / RES Database Central*.xlsb / INPUT-CfD by name, or asks to process a new Ofgem RER/FiT installation report or LCCC CfD register download for the renewables database. Also use this if the user has the raw RER, FiT, or CfD downloads and asks what to do with them for the database.
---

# RES Database update

Updates the RES Database (GB renewable generation assets, tracked across an
`.xlsb` workbook plus two `.xlsx` "pre-loading" files) each cycle from fresh
Ofgem downloads.

**Run the whole pipeline end to end yourself.** This runs on the user's own
Windows machine with a real Excel install, so there is nothing to hand off —
download the sources, process them, and write the updated workbooks. The only
thing to ask the user for is the previous cycle's files if they aren't already
on disk (see Step 0), or a decision on something genuinely ambiguous. Don't
stop after producing the intermediate CSVs and ask them to run the last step;
that's your job.

All the actual computation lives in `scripts/`. This file tells you which
script to run, in what order, and how to interpret the output. Don't re-derive
the column mappings, filters, or formula logic yourself — they're already
implemented, and the reasoning behind them is in `references/`.

## Architecture

```
res-db-update/
├── SKILL.md                       <- you are here
├── scripts/
│   ├── process_ro.py              <- RO CSVs -> paste-ready CSVs
│   ├── process_fit.py             <- FiT report parts -> paste-ready CSVs
│   └── local_excel_update.py      <- writes the .xlsx and .xlsb files via Excel COM
└── references/
    ├── download_instructions.md   <- manual fallback if the automated download fails
    └── column_mappings.md         <- full mapping/filter/formula documentation
```

**Why the final step goes through Excel COM rather than a Python library:** the
pre-loading `.xlsx` files are large (the FiT one is ~270 MB, ~470k rows) and
loading + resaving them with openpyxl exhausts memory long before it finishes.
The `.xlsb` can't be written by any available Python library at all. Driving a
real Excel process via `pywin32` sidesteps both problems. This is why the
pipeline needs a local Windows + Excel machine, and why it works fine here.

**CfD has no processing script and no pre-loading file.** Unlike RO and FiT,
the CfD register download's columns already match the RES DB's `INPUT-CfD`
paste range directly, so there's nothing to compute in Python — the raw
download is handed straight to `local_excel_update.py`. Don't go looking for
a `process_cfd.py`; it doesn't exist by design.

## Directory layout

Everything lives under `C:\prmf-process\res-db-update\`, split three ways so
it's always obvious what's an input, what's a deliverable, and what's scratch:

| Directory | Contents |
|---|---|
| `files_to_update\` | **Inputs.** The previous cycle's three workbooks. Read, never written. |
| `updated_files\` | **Deliverables.** The three updated workbooks this run produces. |
| `working\` | **Everything else.** Fresh Ofgem downloads, intermediate CSVs, logs. |

Create any of these that don't exist (`New-Item -ItemType Directory -Force`)
before doing anything else. Use these paths for the script arguments throughout
— don't ask the user where files are unless something expected is missing.

The three files in `files_to_update\` are:
- `RO Pre-loading.xlsx` — RO pre-loading workbook
- `Fit Pre-loading.xlsx` — FiT pre-loading workbook
- `RES Database Central*.xlsb` — the main RES database. The filename carries a
  cycle tag (e.g. `RES Database Central-8z-Q2_2026.xlsb`), so glob for it
  rather than hardcoding.

Keeping inputs read-only and writing to `updated_files\` matters: if a run goes
wrong halfway through, the previous cycle's files are still intact and the run
is just repeatable. Never write into `files_to_update\`.

## Step 0: check inputs

First, check `files_to_update\` for all three workbooks. If any are missing,
stop and ask the user to paste the previous cycle's copies in, naming exactly
what's absent and the full path — e.g. "I need the previous cycle's `Fit
Pre-loading.xlsx` in `C:\prmf-process\res-db-update\files_to_update\` before I
can run the FiT side. Let me know once it's there." Don't try to work around a
missing input; every step depends on the accumulated history in these files.

Then check for the fresh Ofgem/LCCC downloads in `working\`:
- 3 FiT Installation Report parts (`.xlsx`) — public Ofgem download
- RO Certificates export (`.csv`) — requires RER login
- RO Accredited Stations export (`.csv`) — requires RER login
- CfD register export (`.xlsx`) — public LCCC download, no pre-loading file

If they're all there, skip to Step 2. Otherwise run Step 1.

**Source files belong on disk, not in the chat.** The FiT parts are ~30–36 MB
each. Attaching them to the conversation burns an enormous amount of context
for data the scripts are going to parse off disk anyway. If the user offers to
upload them, redirect: ask them to drop the files in `working\` instead, and
read only the script summaries.

## Step 1: download the Ofgem source files

Save everything into `working\`. Browsers download to
`C:\Users\<username>\Downloads\` by default, so move the files across once
they land and confirm the exact filenames (they include the quarter and year).

Try the automated route first. If anything fails — no login, page structure
changed, download times out — fall back immediately rather than retrying
blindly: read `references/download_instructions.md` and give the user clear
manual steps for the part that failed. Never leave them without a path forward.

### FiT Installation Report (public — use the in-app browser)

1. Load the in-app browser tools if not already loaded:
   `ToolSearch select:mcp__Claude_Browser__navigate,mcp__Claude_Browser__read_page,mcp__Claude_Browser__computer`
2. Navigate to:
   `https://www.ofgem.gov.uk/feed-tariffs-fit/contacts-guidance-and-resources/public-reports-and-data-fit/installation-reports`
3. Read the page to find the most recent quarter's report — 3 `.xlsx` files,
   each ~30–36 MB, labelled "part 1", "part 2", "part 3".
4. Click each download link in turn, then move the files into `working\`.

### RO Certificates & Accredited Stations (requires RER login — use real Chrome)

The RER needs the user's existing logged-in session, so this one goes through
their real Chrome rather than the in-app browser.

1. Load the tools:
   `ToolSearch select:mcp__claude-in-chrome__tabs_context_mcp,mcp__claude-in-chrome__navigate,mcp__claude-in-chrome__computer,mcp__claude-in-chrome__read_page`
2. Navigate to `https://rer.ofgem.gov.uk/` and read the page to confirm the
   user is logged in. If they aren't, don't wait indefinitely — hand them the
   manual steps from `references/download_instructions.md` and ask them to put
   the exports in `working\` when done.
3. Go to **Reports**, run the **"Certificates (RO)"** report:
   - Set the "To" date to cover the current cycle. The previous cutoff is in
     the existing workbook's `OUTPUT-RO!C5` — advance by one quarter unless the
     user says otherwise.
   - The export takes 5–10 minutes. Wait for it to finish before downloading.
4. Run the **"Accredited Stations (RO)"** report and download it.
5. Move both exports into `working\`.

### CfD Register export (public — use the in-app browser)

1. Load the in-app browser tools if not already loaded:
   `ToolSearch select:mcp__Claude_Browser__navigate,mcp__Claude_Browser__read_page,mcp__Claude_Browser__computer`
2. Navigate to: `https://register.lowcarboncontracts.uk/downloadfiles`
3. Read the page to find the table of downloads (Date, Uploaded At, Scheme,
   Action columns, newest first). Find the most recent row whose **Scheme**
   column reads **"CfD"** — the list interleaves other schemes (e.g. "LCHA")
   so don't just take the first row.
4. Click that row's **Download** link, then move the file into `working\`.

There's no login and no pre-loading file for CfD — the download goes straight
into `local_excel_update.py` in Step 3, with no Python processing step in
between.

## Step 2: process the sources

Both scripts are independent and can run in either order. Write their output
to `working\` — these CSVs are intermediates, not deliverables.

**CfD has no equivalent script here** — its raw download needs no
preprocessing (see Step 3) and is handed straight to `local_excel_update.py`.
Don't wait on a `process_cfd.py` that doesn't exist.

```bash
python scripts/process_ro.py \
    --old-workbook "C:\prmf-process\res-db-update\files_to_update\RO Pre-loading.xlsx" \
    --new-stations "<working\...Accredited Stations csv>" \
    --new-certificates "<working\...Certificates csv>" \
    --out-dir "C:\prmf-process\res-db-update\working"

python scripts/process_fit.py \
    --old-workbook "C:\prmf-process\res-db-update\files_to_update\Fit Pre-loading.xlsx" \
    --report-parts "<working\part1.xlsx>" "<working\part2.xlsx>" "<working\part3.xlsx>" \
    --out-dir "C:\prmf-process\res-db-update\working"
```

This produces 5 CSVs in `working\`:
`RO_new_stations_to_append.csv`, `RO_new_certificates_to_append.csv`,
`RES_DB_paste_INPUT-RO_U8.csv`, `FiT_Installation_tab_full_replace.csv`,
`RES_DB_paste_INPUT-FiT_K8.csv`.

Leave them in `working\` — don't present them to the user. They're a useful
audit trail if something later looks wrong, but they aren't the deliverable and
surfacing them just adds noise to a run that's still in progress.

Each script prints a summary as it goes (row counts, filters applied, any
capacity-cutoff assumption in use). Read it. You'll fold the key numbers into
your final report in Step 4, and it's also your first chance to catch a source
format change — a sudden collapse in new-row counts usually means Ofgem renamed
a column, not that nothing happened this quarter.

## Step 3: write the updated workbooks

Run this yourself. It needs `pywin32` and a local Excel — both are normally
already present, but if the import fails, `pip install pywin32` and retry.

```bash
python scripts/local_excel_update.py \
    --ro-pre-loading "C:\prmf-process\res-db-update\files_to_update\RO Pre-loading.xlsx" \
    --fit-pre-loading "C:\prmf-process\res-db-update\files_to_update\Fit Pre-loading.xlsx" \
    --res-db "<glob files_to_update\RES Database Central*.xlsb>" \
    --ro-new-stations "C:\prmf-process\res-db-update\working\RO_new_stations_to_append.csv" \
    --ro-new-certificates "C:\prmf-process\res-db-update\working\RO_new_certificates_to_append.csv" \
    --ro-output-csv "C:\prmf-process\res-db-update\working\RES_DB_paste_INPUT-RO_U8.csv" \
    --fit-installation-csv "C:\prmf-process\res-db-update\working\FiT_Installation_tab_full_replace.csv" \
    --fit-output-csv "C:\prmf-process\res-db-update\working\RES_DB_paste_INPUT-FiT_K8.csv" \
    --cfd-file "<working\...CfD register download.xlsx>" \
    --out-dir "C:\prmf-process\res-db-update\updated_files"
```

The FiT step writes ~470k rows through COM and is the slow part — expect
several minutes and don't assume a long-running step has hung. Give it a
generous timeout rather than killing and restarting it.

Three files land in `updated_files\`: `RO Pre-loading - updated.xlsx`,
`Fit Pre-loading - updated.xlsx`, and the `.xlsb` with ` - updated` appended to
its original name (e.g. `RES Database Central-8z-Q2_2026 - updated.xlsb`), so
the cycle tag is preserved.

**Beyond pasting values, this step also drags formulas down** on `INPUT-RO`,
`INPUT-FiT`, and `INPUT-CfD` within the `.xlsb` — the columns that aren't the
paste range (e.g. `A:S`/`AK:LR` on `INPUT-RO`) hold formulas that need to
cover every row through the new bottom row, exactly like manually pulling the
fill handle down after adding rows. This happens automatically as part of the
same script; see `references/column_mappings.md` for the exact column ranges
per sheet. If a formula's logic ever needs to change, edit it by hand in what
will become the new "last row" *before* running this script that cycle — the
drag just propagates whatever's already there.

**CfD is a full replace, not an append**, and unlike RO/FiT there's no
pre-loading file feeding it — `--cfd-file` points straight at the raw LCCC
register download from Step 1. LCCC's column order isn't always stable
between exports, so the script matches and reorders columns by header name
rather than assuming positional alignment — a shuffled column order is
handled automatically. Only a genuinely *missing* header (one that doesn't
exist anywhere in the download, not just moved) aborts the write, with a
clear list of what's missing; if that happens, don't retry blindly — check
whether LCCC changed the export's actual field set and update
`column_mappings.md` / `CFD_PASTE_COLS` in `local_excel_update.py` if so. Use
`--skip-cfd` to run the RO/FiT parts of this step without touching
`INPUT-CfD` (e.g. if the CfD download isn't ready yet).

`--skip-ro` and `--skip-fit` skip the corresponding pre-loading `.xlsx`
update entirely; `--skip-res-db` skips the whole `.xlsb` update; and within
the `.xlsb` update, `--skip-res-db-ro`, `--skip-res-db-fit`, and `--skip-cfd`
let you rerun just one sheet's part of it. That's the right tool if one stage
fails: fix the cause and rerun that stage alone rather than redoing
everything.

**If it fails**, the most likely cause is a sheet name or cell anchor having
moved since this was written. Check `references/column_mappings.md` for the
expected anchors (`INPUT-RO-Stations` col C, `INPUT-RO-Certificates` col B,
`INPUT-RO!U8`, `INPUT-FiT!K8`, `INPUT-CfD!R8`, `Installation` tab from row 6)
against the actual workbook before assuming something deeper is wrong. A
stale Excel process left over from a crashed run can also lock the files —
check for a lingering `EXCEL.EXE` if a workbook won't open.

## Step 4: report back

Tell the user what actually changed, not just that it finished. Cover:

- Where the updated files are, and which cycle's inputs they came from.
- How many new RO stations and certificates were appended, and the total FiT
  capacity.
- Whether the CfD update ran (or was skipped), and how many CfD contract rows
  were pasted this cycle.
- Anything worth their attention from the list below.

## Things worth flagging (don't skip silently)

- **RO Stations filter**: `process_ro.py` excludes Northern Ireland stations by
  design — RER's export now includes NI plus micro-scale generation the
  database has never tracked (`column_mappings.md` has the full
  investigation). If a cycle looks like it's dropping something it shouldn't,
  check this filter first.
- **Capacity cutoff date**: don't flag this as a per-cycle decision. It is a
  deliberately frozen baseline, not something that advances with the calendar —
  it pins which vintage of RER `Station TIC in kW` restatements is allowed into
  the certificate-derived capacity basis, so moving it re-bases historic
  capacity retroactively. `process_ro.py` reuses whatever's in the workbook's
  `OUTPUT-RO!C5` and that's the correct default; leave it alone unless the user
  explicitly asks to re-baseline. Current registered capacity is unaffected
  either way — it comes from the stations export and is not cutoff-filtered.
- **FiT total capacity moving cycle-to-cycle**: sub-0.1% changes are normal
  (Ofgem data corrections), not a bug. The workbook's own `Comparing to
  previous` tab exists for exactly this check.
- **Row counts that barely move, or move enormously**: the clearest signal of a
  changed source format. Worth a sanity check against the previous cycle before
  the user relies on the output.
- **FiT `Tech & Cap` in the `INPUT-FiT` pivot is always `'Utility-scale'`**:
  by design, `process_fit.py`'s OUTPUT pivot no longer groups by `Tech & Cap`
  — every row in `RES_DB_paste_INPUT-FiT_K8.csv` reports `'Utility-scale'`
  for this column, overwriting whatever the per-installation value would
  otherwise be. This only affects that pivoted output; the `Installation`
  tab's own column C (`Tech & Cap`) is still the genuine per-row value
  computed from `TariffCode`. See `references/column_mappings.md` for how to
  revert this if a future cycle needs the old behavior back.
- **Formula-dragging on `INPUT-RO`/`INPUT-FiT`/`INPUT-CfD` assumes the
  previous cycle's formulas are correct and consistent all the way to the old
  last row.** `local_excel_update.py` copies whatever's sitting in that row
  and fills it down — it doesn't validate the formula itself. If a formula
  was wrong or inconsistent before this cycle, dragging it down just
  propagates the same problem further. Worth a spot-check of a few dragged
  rows against the pattern above them, especially the first time this runs
  after any manual edit to the workbook.
- **CfD column order is matched by header name, not position** — LCCC's
  export doesn't always keep the same column order between cycles, so
  `local_excel_update.py` reorders the download to match `INPUT-CfD` before
  pasting rather than assuming positional alignment. A genuinely missing
  header still aborts the whole `.xlsb` write (not just the CfD part), with
  the missing header(s) listed alongside both full header lists — that's a
  real schema change, not just reordering, and isn't retried automatically.
  Don't retry blindly; check `column_mappings.md`, update `CFD_PASTE_COLS` in
  `local_excel_update.py` if the layout genuinely changed, and rerun with
  `--skip-res-db-ro --skip-res-db-fit` to redo just the CfD part once fixed.
