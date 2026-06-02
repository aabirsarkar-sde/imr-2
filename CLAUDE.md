# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-file Streamlit dashboard (`app.py`) that monitors reverse-osmosis (RO) plant
membrane health. It ingests monthly Excel/CSV inspection reports, extracts per-module
flow rate and conductivity readings, fits a linear degradation trend, and projects a
predicted failure date for each module.

## Commands

```bash
pip install -r requirements.txt
streamlit run app.py          # launches the dashboard at http://localhost:8501
```

There is no test suite, linter config, or build step — the entire app is `app.py`.

## Data flow and architecture

The app is structured as a pipeline. Understanding it requires reading these stages in order:

1. **Discovery** — `report_paths()` globs `*.xlsx/*.xls/*.csv` from the app directory
   and the `uploaded_reports/` subfolder. It skips `manual_readings.csv` and Excel temp
   files (`~$`).

2. **Metadata parsing** — `parse_report_metadata()` derives `plant_group` and a
   `report_date` from the *filename* (e.g. `01- IMR January - 2024 Ank,JH,Panoli.xlsx`),
   matching a month-name + year pattern. Report date defaults to the last day of that
   month, or file mtime if no match.

3. **Extraction** — `read_report()` dispatches per file. Two extraction strategies exist
   because real reports are messy spreadsheets, not clean tables:
   - `extract_rows_from_sheet()` — handles **wide, multi-block sheets** where each module
     is a horizontal block of columns. `find_header_row()` scores rows to locate the
     header, and `find_module_blocks()` splits columns into per-module blocks keyed off
     repeated "Mo No" headers. This is the primary path for the Excel reports.
   - `extract_rows_from_structured_table()` — fallback for already-tidy tables (one row
     per reading), used for CSVs and sheets where the block parser finds nothing.
   - Header detection is fuzzy: the `is_*_header()` predicates run on `normalize_text()`
     (lowercased, alphanumeric-only) so column-name variations still match. Excel sheets
     named "mis" are skipped.

4. **Normalization** — `load_reports()` (cached via `@st.cache_data`) concatenates all
   records plus `read_manual_readings()`, coerces types, drops empty/duplicate rows, and
   returns a single tidy DataFrame with one row per (file, plant, module, date) reading.

5. **Analysis** — per selected module/metric: `aggregate_series()` averages by report
   date, `fit_linear_trend()` does a `np.polyfit` degree-1 fit (the "Predictive ML
   Trendline"), `predicted_failure_date()` extrapolates the line to the threshold
   crossing, and `system_status()` classifies Healthy/Warning/Critical.

6. **Rendering** — `render_dashboard()` draws metric cards, a Plotly chart with trend +
   failure marker, and a cleaned-readings table. `sidebar_filters()` and
   `render_add_data_controls()` build the sidebar (filters + upload/manual-entry forms).

## Key conventions

- **The `METRICS` dict (top of file) is the source of truth** for the two tracked metrics.
  Each entry defines the DataFrame column, raw spreadsheet label, `bad_direction`
  (`"down"` for flow — degradation is a *decrease*; `"up"` for conductivity), and a
  `threshold_factor` applied to the first reading to set the default failure threshold.
  Add or change tracked metrics here.

- **`MANUAL_COLUMNS`** defines the canonical record schema. Both the manual-entry CSV and
  every extracted record conform to it. Keep extraction outputs and `manual_readings.csv`
  in sync with this list.

- **Two write targets, both in the app directory**: uploaded report files land in
  `uploaded_reports/`; manual readings are appended to `manual_readings.csv`. After either
  write the code calls `load_reports.clear()` to bust the cache, then `rerun_app()`.

- When changing extraction logic, account for *both* extractors — a fix in the block
  parser often needs a mirror in the structured-table parser.
