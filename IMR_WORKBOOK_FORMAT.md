# IMR Workbook Format Standard

This document defines the exact Excel workbook structure the IMR dashboard
requires. Follow it and your uploads will parse cleanly with no corrections.
Deviate and the dashboard will silently misplace data or drop readings.

---

## 1. Filename

The filename is the only source of the report date and plant group label.
It must contain **[Month] [Year]** and use `.xlsx` or `.xls`.

**Accepted patterns**

| Style | Example |
|---|---|
| Full month + 4-digit year | `IMR April-2025 Ankleshwar.xlsx` |
| Abbreviated month + 2-digit year | `master sheet IMR Dec.-25.xlsx` |
| Full month + 4-digit year (multi-group) | `IMR January-2025 Ank,JH,Panoli.xlsx` |

**Rules**
- Month name can be full (`April`) or abbreviated (`Apr`, `Dec.`) — the dot after an abbreviation is fine.
- Year can be 4-digit (`2025`) or 2-digit (`25`). Two-digit years are treated as 20xx.
- The separator between month and year can be a dash (`-`), dot (`.`), slash (`/`), or a space.
- The group/location name can appear before OR after the date — both are stripped correctly.
- Do NOT include special characters other than commas, spaces, dashes, dots, and parentheses.
- Do NOT start the filename with `~$` (Excel locks temporary files with this prefix and the parser skips them).

**What breaks**
- No month/year in the name at all → report date defaults to today's date (shown as a warning at upload).
- Typo in the group label (e.g. `Panol` vs `Panoli`) → creates a separate group in the dashboard. Use the same spelling every month.

---

## 2. Workbook structure (sheets)

A workbook contains:
- **One sheet per RO plant** (see Section 3)
- **One `mis` sheet** — the plant register (see Section 4)

The `mis` sheet is detected by name (case-insensitive). All other sheets are
treated as plant reading sheets.

---

## 3. Plant reading sheets

### 3a. Sheet name

Name each sheet after the plant, ending with its Plant Serial Number in
parentheses:

```
Lupin Ank RO1 (1639)
Bharat Rasayan RO 6 (3523)
BEIL PTHP (3550)
```

The number in parentheses is how the dashboard links a reading to its zone
in the MIS register. Without it, the plant shows as **"Unknown"** zone in all
fleet views.

- The number must match the `PLANT SR NO` in the `mis` sheet exactly.
- Trailing spaces or a stray dot after the closing parenthesis are fine — e.g. `(3523) .` still works.
- If there are no parentheses, the parser tries a trailing 3+ digit number as a fallback (`BEIL PTHP 3550`), but the parentheses form is preferred and unambiguous.

### 3b. Layout — wide block format (preferred)

Place one "block" of columns per RO stage. Each block has exactly **four**
columns in this order:

| Column | Accepted header text (any capitalisation) |
|---|---|
| Module number | `Mo No`, `Mod No`, `Module No`, `Module Number` |
| Installation date | Any header containing both "inst" and "date" |
| Flow rate | Any header containing "flow" AND one of: "total", "rate", "liter", "lph" |
| Conductivity | Any header containing "cond", "conductivity", or "us cm" |

For multiple stages side-by-side, repeat the four-column block. Put the
**stage label** (e.g. `Stage I`, `Stage II`, `RO-1`) in the row(s) **above**
the header row — the parser scans up to 5 rows above the header for the label.

**Example header structure (two stages)**

```
Row N-1:  |        Stage I         |        Stage II        |
Row N:    | Mo No | Inst Date | Total flow liter/hr. | Cond. us/cm | Mo No | Inst Date | Total flow liter/hr. | Cond. us/cm |
Row N+1:  |  1    | 01-Jan-24 | 12000                | 340         |  1    | 01-Jan-24 | 8500                 | 620         |
Row N+2:  |  2    | 03-Mar-23 | 11800                | 360         |  2    | 03-Mar-23 | 8400                 | 630         |
```

**Data rows**
- Module No: integer (rows without a number are skipped entirely).
- Flow: number in L/hr. Blank is acceptable if conductivity is present.
- Conductivity: number in µS/cm. Values above 50,000 are treated as data-entry errors and are blanked automatically (shown as a warning at upload).
- Bypass: write `BY PASS` or `Bypass` anywhere in the module's row — the parser will record it as a bypassed module (the strongest replacement signal). Do NOT leave the row blank and expect it to count.

### 3c. Layout — structured table format (alternative)

If the plant data is already a tidy one-row-per-reading table (e.g. exported
from another system), use these column headers:

| Field | Accepted header |
|---|---|
| Module number | `Mo No`, `Module No` (required) |
| Flow | Any header with "flow" + "total"/"rate"/"liter"/"lph" (required) |
| Conductivity | Any header with "cond"/"conductivity"/"us cm" (required) |
| Installation date | Any header with "inst" + "date" |
| Reading date | `Date`, `Reading Date`, `Report Date`, `Sample Date` |
| Stage | `Stage`, `RO Stage` |
| Site name | `Plant`, `Site`, `Site Name`, `Plant Name` |

Required columns are marked. All others are optional. Rows with no module
number are skipped.

### 3d. Operating parameters (optional)

Place instrument readings (pressure, temperature, feed conductivity) anywhere
on the sheet as **(tag, value, unit)** triples — three consecutive cells on
the same row, with the **unit in the rightmost cell**:

| Accepted unit text | Type |
|---|---|
| `bar` | Pressure |
| `deg c` | Temperature |
| `us cm` | Conductivity |
| `lit hrs` | Flow |
| `m3 hr` | Flow |

Example: `PI 1601 | 56 | bar` → tag `PI 1601`, value 56, type pressure.

The feed-array pressure tag `PI 1601` is used as the operating pressure context
for each reading. Other tags are stored but not yet surfaced on the dashboard.

---

## 4. MIS register sheet

Name this sheet exactly **`mis`** (case-insensitive). It must appear in the
same workbook as the reading sheets it describes.

### 4a. Header row

The header row must contain at least **`PLANT SR NO`** and **`Site Name`** (case-insensitive, extra spaces OK). The parser searches the first 8 rows for this header.

### 4b. Columns (all column names case-insensitive)

| Column | Notes |
|---|---|
| `ZONE` | The geographic/administrative zone. Can be a merged cell spanning all plants in that zone — blank continuation rows are forward-filled automatically. |
| `ZM NAME` | Zone Manager name. Same merged-cell behaviour as ZONE. |
| `PLANT SR NO` | **Required.** Integer plant serial number. Rows without this are skipped. Must match the number in the plant sheet name (see 3a). |
| `Site Name` | Human-readable plant name. |
| `STATUS` | Operating status string (e.g. `Active`, `Shutdown`). |
| `Membrane Required` | Integer count of membranes flagged for replacement. |
| `REMARKS` | Free text. |

### 4c. Example layout

```
Row 1:  ZONE     | ZM NAME    | PLANT SR NO | Site Name              | STATUS | Membrane Required | REMARKS
Row 2:  Ankleshwar | Ravi Patel | 1639        | Lupin Ank RO1          | Active |                   |
Row 3:            |            | 1640        | Lupin Ank RO2          | Active | 2                 | Stage II degraded
Row 4:  Panoli    | Amit Shah  | 3523        | Bharat Rasayan RO 6    | Active |                   |
```

Rows 2 and 3 share the same zone/ZM (merged cell); row 3 has a blank ZONE
cell and inherits "Ankleshwar" automatically.

---

## 5. Checklist before uploading

- [ ] Filename contains month name and year (`April-2025` / `Dec.-25`)
- [ ] Filename spells the group/location the same way as previous months
- [ ] Every plant sheet name ends with `(PLANT_SR_NO)` matching the MIS
- [ ] Header row contains `Mo No` (or variant), `Inst Date`, flow, and conductivity columns
- [ ] Bypassed modules are marked `BY PASS` / `Bypass` — not just left blank
- [ ] The `mis` sheet is named exactly `mis`
- [ ] `PLANT SR NO` column exists in the `mis` sheet header (within the first 8 rows)
- [ ] No merged cells in the module/flow/conductivity data rows (merged cells in ZONE/ZM NAME are fine)
- [ ] Conductivity values are in µS/cm (not mS/cm — a value like `0.34` will be treated as near-zero; it should be `340`)
- [ ] File is saved as `.xlsx` or `.xls`, not `.csv` or `.ods`

---

## 6. What the upload gate will warn you about

When you upload a file through the dashboard, the system shows a per-file
quality report before anything is written to the database. Here is what each
message means:

| Level | Code | Meaning |
|---|---|---|
| ERROR | `no_readings` | No data rows were extracted. The file cannot be committed. Check that the header row exists and columns are named correctly. |
| WARN | `date_fallback` | No month/year found in the filename. The report date has been guessed from the file's last-modified time — this is almost always wrong. Rename the file. |
| WARN | `raw_group` | The plant group label looks like a raw filename (contains digit runs, underscores, or too many tokens). Check the filename convention. |
| WARN | `no_sr_no` | One or more plant sheet names have no `(NNNN)` serial number. Those plants will show as "Unknown" zone everywhere. Fix the sheet names. |
| WARN | `low_zone_match` | Fewer than 50% of the readings could be matched to a MIS zone entry. Check that `PLANT SR NO` values in the reading sheets match those in `mis`. |
| INFO | `cond_blanked` | Some conductivity values exceeded 50,000 µS/cm and were treated as data-entry errors (blanked). Check those cells. |
| INFO | `will_replace` | A previous version of this file is already in the database. Committing will replace it. |

An ERROR blocks the commit. WARNs and INFOs allow commit but indicate data
quality issues you should investigate.
