# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-file Streamlit app (`app.py`) that monitors reverse-osmosis (RO) plant
membrane health. It ingests monthly Excel inspection reports ("IMR" workbooks), parses
each one **once** into a **PostgreSQL** database, and serves a fleet-wide and per-plant
view. The dashboard never re-parses Excel to render — it queries the database.

## Commands

```bash
pip install -r requirements.txt
streamlit run app.py          # http://localhost:8501
```

The database connection comes from a `DATABASE_URL` env var **or** `.streamlit/secrets.toml`:

```toml
[postgres]
url = "postgresql://USER:PASSWORD@HOST:5432/DBNAME?sslmode=require"
```

There is no test suite or linter. Verify changes with a **scratch SQLite ingest** (see
"Verifying changes") — the schema uses SQLAlchemy generic column types, so the same code
runs on Postgres (prod) and SQLite (tests).

## Pages (st.navigation, in `main()`)

1. **Portfolio** (default landing) — fleet-wide across ALL plants, no plant selection.
   KPI cards, plant ranking, fleet trend, zone rollup, worst-25 modules, age profile.
2. **Dashboard** — per-plant/module flow & conductivity history (pick plant in sidebar).
3. **Replacement Candidates** — per-plant, per-stage outlier/absolute-limit flagging.

## Data flow (ingest → DB → render)

1. **Discovery** — `report_paths()` globs `*.xlsx/*.xls/*.csv` from the app directory and
   `uploaded_reports/`, skipping `manual_readings.csv` (legacy) and Excel temp files (`~$`).

2. **Metadata** — `parse_report_metadata()` derives `plant_group` + `report_date` from the
   *filename* (month-name + year; date defaults to last day of that month).

3. **Extraction** — `read_report()` dispatches per sheet. Sheets named `mis` go to the MIS
   parser; all others to the reading parsers:
   - `extract_rows_from_sheet()` — wide, multi-block sheets (each module is a horizontal
     block of columns; `find_module_blocks()` keys off repeated "Mo No" headers).
   - `extract_rows_from_structured_table()` — already-tidy one-row-per-reading tables.
   - Header detection is fuzzy via `normalize_text()` + the `is_*_header()` predicates.
   - **Bypass**: `is_bypass_marker()` flags a module whose block reads "BY PASS"/"Bypass";
     it has a module number but no flow/conductivity, and is kept with `status="bypass"`
     (it is the strongest replacement signal). Otherwise `status="active"`.
   - **MIS** (`extract_mis_rows()`): the per-workbook plant register — ZONE, ZM NAME,
     PLANT SR NO, Site Name, STATUS, Membrane Required, REMARKS. ZONE/ZM NAME are merged
     cells, **forward-filled** down continuation rows; rows without a PLANT SR NO are skipped.

4. **Normalization** — `normalize_reading_records()` / `normalize_parameter_records()` /
   `normalize_mis_records()` coerce types and dedup. Reading guards: keep a row if it has a
   module number AND (`status=="bypass"` OR flow/conductivity present); conductivity above
   `MAX_PLAUSIBLE_CONDUCTIVITY_US_CM` (50,000) is a data-entry error → NaN.

5. **Persistence** — `ingest_reports()` is **idempotent and parse-once**: it hashes each
   file (SHA-256) and skips any whose hash is unchanged (recorded in `ingested_files`). A
   new/changed file is parsed and written via DELETE-then-insert by `source_file` into
   `readings`, `parameters`, `mis`, with per-file counts (`n_readings/n_parameters/n_mis`).
   `main()` runs `init_db` + `ingest_reports` on startup, then busts the cached `load_*()`.

6. **Render** — `load_readings()/load_parameters()/load_mis()` (each `@st.cache_data`) query
   the DB. `compute_fleet_status()` (cached) annotates every reading as bypass/degraded/need.

## Persistence, schema, and migrations

- **Tables** are SQLAlchemy `Table`s on `DB_METADATA` (`readings`, `parameters`, `mis`,
  `ingested_files`). `READINGS_COLUMNS` / `MIS_COLUMNS` derive from the table definitions,
  and the normalizers `reindex(columns=...)` to them — **add a column to the `Table` and it
  flows through automatically.**
- **`init_db()` calls `create_all()`, which creates MISSING tables but does NOT ALTER
  existing ones.** So adding a column to an existing table needs a one-time migration
  against the live DB. The established pattern (used for `status`, `plant_sr_no`, `n_mis`):

  ```sql
  ALTER TABLE readings       ADD COLUMN IF NOT EXISTS <col> <type>;
  ALTER TABLE ingested_files ADD COLUMN IF NOT EXISTS n_<x> integer;
  DELETE FROM ingested_files;   -- clear hashes to force a full re-parse
  ```
  Then run `ingest_reports(engine, APP_DIR)` so every file re-parses and back-fills the new
  columns. Because ingest is DELETE-then-insert by `source_file`, re-ingest never duplicates.
  **Always run the ALTERs before re-ingest** — `to_sql(..., if_exists="append")` fails if a
  DataFrame column has no matching table column.

## Key conventions

- **`METRICS` (top of file)** is the source of truth for the two tracked metrics (flow,
  conductivity): DataFrame column, label, `bad_direction`, `unit`.
- **Zones are data-driven** — read from the MIS `zone` column (never hardcoded). Readings
  inherit a zone by joining `readings.plant_sr_no` (the trailing `(NNNN)` parsed from the
  plant sheet name, via `plant_sr_no_from_name()`) to `mis.plant_sr_no`; missing → "Unknown".
- **Degraded definition** is shared: `evaluate_stage_readings()` (absolute-limit method).
  The Portfolio uses stage-aware cutoffs `STAGE_CONDUCTIVITY_CUTOFFS` (I>1000, II>1500,
  III>2000 µS/cm via `stage_cutoff()`); the Replacement page lets the user pick the method.
- **When changing extraction logic, account for *both* reading extractors** — a fix in the
  block parser usually needs a mirror in the structured-table parser.
- Styling: reuse `metric_card()` and the `plotly_white` Plotly theme.

## Verifying changes

No test suite. For data/extraction/schema changes, write a throwaway script that builds a
**SQLite** engine, ingests the two `.xlsx` files in the repo, and asserts on the result —
then delete the script (don't commit it). This exercises the real ingest pipeline without
touching production. Boot the app (`DATABASE_URL=sqlite:////tmp/x.db streamlit run app.py`)
to confirm pages render (HTTP 200).
