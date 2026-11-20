# Column mappings and filter logic

This documents the reasoning behind `process_ro.py` and `process_fit.py` so
future changes (e.g. Ofgem changing an export format) can be made with full
context rather than re-deriving everything from scratch.

## RO: Stations

| RO_Pre-loading expects | RER export gives          |
|---|---|
| Accreditation Number    | Accreditation reference   |
| Generating Station      | Station name               |
| Capacity (kW)           | Station DNC in kW          |
| Country                 | Country                    |
| Technology              | Technology                 |
| Contract Type           | Contract type               |
| Accreditation Date      | Effective date               |
| Commission Date         | Commissioning date            |
| Organisation            | Organisation name             |
| Status                  | *(not present in current export - left blank)* |
| Scheme                  | *(not present - hardcoded to `'RO'`, since this file is RO-only)* |

Validated by cross-checking ~2,788 overlapping accreditation numbers between
an old cycle's data and a fresh RER download — every field matched exactly
except the two not present in the new export.

### Filter: exclude Northern Ireland

The RER's public Accredited Stations export was found to include ~23,790 rows
that the database has never tracked:
- All Northern Ireland stations (never present under the old export).
- ~22,715 of those NI rows carry a "...with a DNC of 50kW or less" technology
  label (micro-scale PV/wind) — this category turned out to be **entirely
  contained within the NI set** (22,715 of 22,715 overlap), so excluding NI
  alone removes almost all of the scope creep.
- The remaining single non-NI row (across a real test run) was a genuine new
  large-scale station (a fuelled/CHP site), which should NOT be filtered out.

Decision: **filter by `Country == 'Northern Ireland'` only.** Do not
additionally filter by technology label — NI exclusion already captures
essentially all of the unwanted scope, and filtering by technology risks
dropping legitimate future non-NI micro-generation if Ofgem ever reports it.

Confirmed via commission-date analysis that these are not upcoming/new-build
plants — they're historical (mostly 2012-2017 commissioning), just newly
appearing in the export.

## RO: Certificates

| RO_Pre-loading expects        | RER export gives              |
|---|---|
| Accreditation Number            | Accreditation reference         |
| Generating Station               | Station name                     |
| Capacity                          | Station TIC in kW                 |
| Technology Group                   | Technology                          |
| Generation Type                     | Generation Type                       |
| Output period                        | Output period (e.g. "Sep 2022")         |
| No. of Certificates                   | Number of certificates                    |
| Start/End Certificate No.              | Start/End certificate number                |
| MWh per certificate                     | MWh per certificate                           |
| Issue Date                                | Issue Date                                      |
| Certificate Status                          | Cert status                                       |
| Status Date                                   | Status date                                         |
| Current Holder Organisation Name                | Current holder org name                               |
| Company Registration No.                          | Current holder company reg number                       |

**Important:** `Station TIC in kW` / `Capacity` is a **current snapshot as of
the download date, not a historical value at the time of the certificate.**
Cross-checking the same certificate (same `Start certificate number`) across
two different cycles' downloads showed the capacity value can change (e.g.
1887 kW -> 2935 kW for the same cert) because the underlying station's
registered capacity changed since. This is expected and does not indicate a
data error. It's not a practical problem for the append-only approach since
existing rows are never touched, only new ones added.

### Append-only, keyed by `Start certificate number`

RER's Certificates export only goes back to ~April 2015. A full rebuild each
cycle would destroy pre-2015 history that exists only in the accumulated
workbook. So: never delete/rebuild, only add rows whose `Start certificate
number` isn't already present.

### Filter: exclude Northern Ireland (by the cert row's own Country field)

Unlike Stations, certs are filtered by their own `Country` field directly,
NOT by whether their accreditation number appears in the (also NI-filtered)
Stations list. This matters because ~20 certs reference accreditation numbers
for stations no longer present in the current Accredited Stations snapshot
(likely revoked/closed stations) — these are legitimate certificate history
worth keeping, since certs and stations have always been separate tables in
this workbook (the certs table has always been a superset).

### Derived fields computed in Python (not written as Excel formulas)

The original workbook computes two fields via formula for every certificate
row:
- `Year` = `IFERROR(DATE(YEAR(Output period),12,31), <fallback using Issue
  Date and last-4-chars of Output period text>)`. In practice this resolves
  to: parse Output period (e.g. "Sep 2022") and use 31 Dec of that year.
- `Valid Certificates Banding (MWh)` = `No. of Certificates * MWh per
  certificate` if Certificate Status is "Redeemed" or "Issued", else 0.

And per-station output (`OUTPUT-RO`, pasted into the RES DB's `INPUT-RO!U8`)
computes, per accreditation number, over ALL certs (old + newly appended):
- `Capacity (MW)` = `max(station's registered capacity, max cert capacity
  among certs issued on/before a cutoff date) / 1000`. The cutoff date
  defaults to whatever is already in the workbook's `OUTPUT-RO!C5` assumption
  cell (2023-12-31 as of this writing) unless overridden with
  `--capacity-cutoff-date`.

  **The cutoff is a frozen baseline and should not be advanced each cycle.**
  RER restates `Station TIC in kW` on *every* certificate row, so a station
  that re-registers at a higher TIC has the new figure stamped on all its
  subsequent certs. Taking `max()` over all certs would therefore inflate that
  station's *historic* capacity. The cutoff pins which vintage of TIC
  restatements is trusted; advancing it doesn't add information about when a
  change occurred, it just admits later restatements and re-bases the series
  retroactively — the exact effect the cutoff exists to prevent.

  Two consequences worth knowing:
  - The cert pool used here (`by_acc_cap`) is built from old-workbook certs
    only, so this cycle's appended certs are already excluded independently of
    the cutoff. The two guards overlap.
  - For stations accredited after the cutoff, `cert_max` is always 0 and
    capacity comes purely from the registered TIC in the stations export, which
    is *not* cutoff-filtered. So genuine uprates still flow through via the
    `max()`, and a static cutoff does not go stale.
- `Date Last/First Certificate` = max/min Issue Date across all certs for
  that accreditation number.
- `Historical avg generation (MWh/year)` = average of the yearly sum of
  "Valid Certificates Banding" over up to 5 years ending the year before Date
  Last Certificate, excluding zero-sum years, 0 if no data. This replicates
  the workbook's original `LET`/`MAXIFS` array formula, which was computed in
  Python rather than as an Excel formula because `LET` doesn't reliably
  recalculate outside genuine Excel at this row count.

**Why CSVs, not a rewritten .xlsx:** loading + resaving the ~470k-row
pre-loading workbook with openpyxl reliably exhausts memory in a constrained
sandbox (tested: killed by OOM after ~2 minutes). `local_excel_update.py`
solves this by writing through real Excel (via COM) on a normal local
machine instead, which has no such ceiling.

## FiT: Installation Report -> Installation tab

The raw report columns map 1:1 onto the `Installation` tab's paste columns
I:AD (`Extension (Y/N)` through `FIT ID`) with headers at row 5 of the
downloaded report and data from row 6.

Derived columns (A-H), computed in Python from the raw columns:

| Col | Name | Logic |
|---|---|---|
| A | Tariff Code | part of `TariffCode` after the 2nd `/` |
| B | PV Letter | `'H'`/`'M'`/`'L'` — whichever appears in Tariff Code, else `''` |
| C | Tech & Cap | part of `TariffCode` before the 2nd `/`, + PV Letter |
| D | Capacity (MW) | `Installed capacity` (kW) / 1000 |
| E | Start Month | first-of-month(Commissioning date, else Application date), floored at 2010-04-01 |
| F | Clean postcode | postcode with spaces removed |
| G | Zone | lookup of clean postcode against `REGION.PC` (PV/Wind only) |
| H | Region | lookup of postcode area, or GOR+LocalAuthority fallback, against `REGION.PC` (PV/Wind only) |

### Zone/Region lookup — faithful replication of the original IFERROR chain

This is the trickiest part and worth getting exactly right. The original
Excel formulas are pure `IFERROR(primary_lookup, fallback_lookup)` chains —
**not** truthiness checks on the result. A few subtleties that caused bugs
during development, kept here so they aren't reintroduced:

- **A successful lookup that returns `0` is a legitimate answer, not a
  signal to fall back further.** `REGION.PC` genuinely maps the
  `"Unknown-Unknown"` combined fallback key to `0` for both Photovoltaic and
  Wind. Filtering out `0` as "falsy" (an earlier draft of this script did
  this) is WRONG — it deviates from what real Excel would compute.
- **A genuinely blank postcode should give Zone = blank, not "Missing
  postcode".** The Zone formula checks `F6=""` *before* attempting any
  lookup — "Missing postcode" is the `IFERROR` fallback text used only when a
  postcode IS present but doesn't match anything in the zone table. Don't
  convert missing postcodes to the string `'nan'` (a pandas/Python artifact
  of `.astype(str)` on `NaN`) before checking for blank — `fillna('')` first,
  or the blank-check will silently fail and everything falls through to
  "Missing postcode" incorrectly.
- As of the last processed report, ~594 of 869,815 installations (0.07%)
  have no location data at all (`PostCode`, GOR, and Local Authority all
  blank/"Unknown") — these correctly resolve to Region=0, Zone=blank per the
  logic above. This is not a bug; verify against the `Unknown-Unknown` key in
  `REGION.PC` if it ever looks wrong again.

### OUTPUT (pasted into the RES DB's `INPUT-FiT!K8`)

A `groupby(Technology, Region, Zone, Start Month)` summing Capacity (MW).
This is a full rebuild each cycle (matching the raw Installation tab being a
full rebuild too), not an append.

**`Tech & Cap` is forced to `'Utility-scale'` in this OUTPUT, not grouped
on.** Unlike the Installation tab's own column C (which keeps the genuine
per-row `C_TechCap` derived from `TariffCode`), the pivoted OUTPUT that feeds
`INPUT-FiT!K8` always reports `Tech & Cap = 'Utility-scale'` regardless of
what column C says. Because `Tech & Cap` is no longer a distinguishing
groupby key here, rows that would previously have been split out by
`Tech & Cap` (within the same Technology/Region/Zone/Start Month) are now
summed together into a single row. If a future cycle needs the old
Tech & Cap-aware pivot back, restore `C_TechCap` as a groupby key in
`build_output_pivot()` (`scripts/process_fit.py`) instead of inserting the
constant.

**Why CSVs, not a rewritten .xlsx:** same memory constraint as RO, worse —
~870k rows across 3 report parts. `local_excel_update.py` handles the actual
write via real Excel locally.

## CfD: Register download -> INPUT-CfD (RES DB .xlsb)

Unlike RO and FiT, CfD has **no separate pre-loading workbook** and **no
Python preprocessing step** — the raw file downloaded from
`register.lowcarboncontracts.uk/downloadfiles` (Scheme = "CfD") has the same
set of column headers as the `INPUT-CfD` sheet's paste range, so
`local_excel_update.py` reads it and pastes it across with no other
transformation. This is a genuinely different shape from RO/FiT, so don't
look for a `process_cfd.py` — there isn't one by design.

Column layout on `INPUT-CfD` (all within the RES DB `.xlsb`, not a separate
file):

| Range | Purpose |
|---|---|
| `A:P` | Formulas — dragged down each cycle, never pasted into |
| `Q` | Deliberate gap column — left alone |
| `R:BD` | **Paste range.** Full-replace from the CfD register download each cycle, starting row 8 (header at row 7) |
| `BE` | Deliberate gap column — left alone |
| `BF:MM` | Formulas — dragged down each cycle, never pasted into |

**Column order in the download isn't guaranteed to match `INPUT-CfD`'s order,
even though the same headers appear.** `local_excel_update.py` matches
columns by header name, not position: it builds a name -> index map from the
downloaded file's header row, then reorders every row so it lines up with
`INPUT-CfD` row 7's header order before pasting. This is the fix for LCCC
occasionally shuffling column order between exports.

**A genuinely missing header still aborts the write.** If a header that
`INPUT-CfD` row 7 expects isn't present anywhere in the download (not just in
a different position), that's a real schema change, not a reordering — the
script raises an error listing the missing header(s) plus both full header
lists, and nothing gets pasted. This is the only guard against a silent
misalignment for this tab — don't remove it or make it lenient. Any download
column *not* found in `INPUT-CfD`'s headers is logged and ignored (not an
error) — it's assumed to be new/unused data, not a break.

**Full replace, like FiT** — each cycle's download is a fresh cumulative
snapshot of live CfD contracts, so the entire `R:BD` range is cleared and
re-pasted rather than appended to.

**Formula-dragging**: after the R:BD paste, `local_excel_update.py` drags
whatever formulas already exist in `A:P` and `BF:MM` down from the previous
last row to the new last row (`AutoFill`, equivalent to a manual fill-handle
drag). If this cycle has fewer rows than last cycle, the excess formula rows
below the new last row are cleared instead of extended.

## Formula-dragging on INPUT-RO and INPUT-FiT (RES DB .xlsb)

The same drag-down behavior also applies to the RES DB's `INPUT-RO` and
`INPUT-FiT` sheets, in the columns surrounding their respective paste ranges:

| Sheet | Paste range (values, cleared+repasted each cycle) | Formula ranges (dragged down each cycle) |
|---|---|---|
| `INPUT-RO` | `U:AI` (from `RES_DB_paste_INPUT-RO_U8.csv`) | `A:S`, `AK:LR` |
| `INPUT-FiT` | `K:P` (from `RES_DB_paste_INPUT-FiT_K8.csv`) | `A:J`, `R:KY` |
| `INPUT-CfD` | `R:BD` (from the raw CfD register download) | `A:P`, `BF:MM` |

Each sheet has a one-column gap between its formula range(s) and paste
range(s) (e.g. `T` on `INPUT-RO`, `Q` on `INPUT-FiT`) that is neither pasted
into nor dragged — leave it alone.

`local_excel_update.py`'s `fill_formulas_down()` implements this generically:
it takes the previous last row (captured before the paste-range clear) and
the new last row (after the paste), and for each formula range calls Excel's
`AutoFill` from the old last row down through the new one — exactly what a
manual fill-handle drag produces, including Excel adjusting each row's
relative references. If a cycle's row count decreases, it clears the excess
formula rows below the new last row instead.

If a formula ever needs to change, edit it once in the workbook by hand on
what will become the new "last row" *before* running the pipeline that
cycle — `AutoFill` propagates whatever formula is already sitting in the
source row, so a hand-edit there is exactly how you'd update the pattern
going forward.
