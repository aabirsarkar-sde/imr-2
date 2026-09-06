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
   KPI cards, stage-wise membrane requirement, plant ranking, fleet trend, zone rollup,
   worst-25 modules, age profile.
2. **Dashboard** — per-plant/module flow & conductivity history (pick plant in sidebar).
3. **Replacement Candidates** — per-plant, per-stage outlier/absolute-limit flagging.
4. **IMR Tracker** — who has submitted each month, by zone; hosts the "not running this
   month" marks.

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
- **Two tables hold facts no workbook can carry**, edited in-app rather than parsed, so
  ingest never touches them and a re-parse never clears them:
  - `plant_downtime` (plant_sr_no, month "YYYY-MM", not_running, remarks) — a plant that
    wasn't running that month. A plant that didn't run sends nothing, which is
    indistinguishable from one that just didn't file, so the IMR Tracker subtracts these
    from "expected" instead of chasing them. Set on the Tracker; `save_downtime()` is
    scoped to the plants shown, so unticking clears a mark and hidden plants are untouched.
    A plant that DID submit overrides a stale mark.
  - `bypass_notes` (plant_key, stage_label, module_label, month, reason, remarks) — WHO
    bypassed a module. The sheet only ever says "BY PASS"; `BYPASS_REASON_OPTIONS` is
    ROCHEM vs Client, and blank ("Not recorded") is a real state, never defaulted.
    `attach_bypass_notes()` joins them onto the fleet-status frame in the Portfolio;
    clearing both fields deletes the note.
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
- **Exception: tables listed in `MIGRATED_TABLES`** (`plants`, `parameters`) migrate
  themselves. `init_db()` calls `migrate_added_columns()`, which diffs the `Table`
  definition against the live table and ALTERs in whatever is missing — needed because
  `plants` is maintained in-app, so no re-ingest would ever back-fill it. Add a column to
  `PLANTS_TABLE` and it appears on the next boot; add its default back-fill next to the
  `plant_type` one if blank isn't an acceptable starting value.
  For an **ingest-built** table in that list, the ALTER can't fill the column on its own,
  so pair it with a `DELETE FROM ingested_files` in the same `if` (see the
  `("parameters", "plant_sr_no")` case) — that makes the startup `ingest_reports` treat
  every report as new and back-fill by re-parsing. `init_db()` runs before both ingest
  passes in `main()`, so the ALTER always lands first.

## Plant identity

**A plant IS its `plant_sr_no`, never its display name.** Sheet names collide in both
directions: several sites run two RO trains under one name ("Aarti Jhagadia" is both 1957
and 2708), and one plant is renamed or misspelled across months ("Glenmark Ank RO 1
(1748)" is now "Alivus Life science LTD"). So **every per-plant groupby, dedup, join or
picker keys on `plant_identity_key()`** (the SR, falling back to the name only when no SR
could be resolved) — grouping on `plant` merges two physically separate plants and
mislabels their Plant SR No. `readings` and `parameters` both carry `plant_sr_no`;
`plant_level_frame()` attaches it (plus `plant_key` and `zone`) to a parameter frame, and
`current_plant_name()` picks the name to display for a set of rows.

## Key conventions

- **`METRICS` (top of file)** is the source of truth for the two tracked metrics (flow,
  conductivity): DataFrame column, label, `bad_direction`, `unit`.
- **Zones are data-driven** — read from the MIS `zone` column (never hardcoded). Readings
  inherit a zone by joining `readings.plant_sr_no` (the trailing `(NNNN)` parsed from the
  plant sheet name, via `plant_sr_no_from_name()`) to `mis.plant_sr_no`; missing → "Unknown".
- **The register carries three per-plant facts, each with its own field and its own
  default.** `plant_type` (PLANT_TYPE_OPTIONS, default PT) is the system installed;
  `client` (CLIENT_OPTIONS — ROCHEM / ROSERVE / RENT, default ROCHEM) is who it runs
  under commercially; `status` is the state below. Each has a `canonical_*()` that maps
  a raw value to the option list or None, and every write (`save_plants`, `add_plant`,
  `seed_plants_if_empty`, `load_plants`) falls back to the default rather than storing
  a blank. `plants` is in `MIGRATED_TABLES`, so a new column ALTERs itself in at boot —
  pair it with a back-fill in `_add_column` when blank isn't an acceptable start.
- **A plant's status is a state; its type is a kind.** `PLANT_STATUS_OPTIONS` answers
  whether the plant is running for us (Active / Inactive / Stand By RO / Not in Rochem
  Scope / Plant Shutdown); `PLANT_TYPE_OPTIONS` answers what system is installed.
  STPT RO / SPRO / UF RO were statuses once, which made one field answer both questions
  (a shut-down SPRO had to pick one) — `migrate_status_types()` moves any row still
  stored under them onto `plant_type`, but only where no type was chosen, and sets the
  status Active. **The IMR Tracker's roster asks both questions**:
  `build_submission_roster()` keeps a plant only if `status_is_reporting()` AND
  `type_is_reporting()` — `NON_IMR_PLANT_TYPES` (SP / UF / STPT) don't file a per-module
  IMR at all, so a new non-module system belongs in that set, not in the status list.
- **Degraded definition** is shared: `evaluate_stage_readings()` (peer-IQR or
  absolute-limit); the Replacement page lets the user pick the method.
- **Two degradation signals, and the Portfolio can drop one.** `compute_fleet_status()`
  sets `degraded_iqr` (peer outlier in its stage) and `degraded_mom` (month-over-month
  conductivity jump) and ORs them into `degraded`/`need`. The Portfolio's **Peer outliers
  only** toggle re-derives those two verdict columns via `apply_degradation_signals()`,
  applied *before* any zone/stage filtering so every count on the page moves together.
  It never clears `degraded_iqr`/`degraded_mom` — those stay as the evidence behind the
  Reason column, so a peer outlier that also jumped still says so. `mom_only_count()` is
  what the toggle adds or withholds. **Anything new that counts modules must read
  `degraded`/`need` off the frame, never re-OR the two raw signals** — that would ignore
  the toggle.
- **Stages are canonicalized, never compared raw.** Sheets spell the stage every way there
  is ("I STAGE", "1st Stage", "Stage-2"), so `canonical_stage()` folds a `stage_label` to
  `"I"/"II"/"III"` (or None → `UNSTAGED_LABEL`) and `compute_fleet_status()` stamps it as
  `stage_key`. **Every stage-wise rollup groups on `stage_key`** — the membrane requirement
  is raised per stage, so grouping on the raw label would split one stage into three.
  `stage_display()` renders it as "1st/2nd/3rd Stage"; `stage_pills()` is the button row
  (falls back to a multiselect on a Streamlit without `st.pills`).
- The Portfolio's stage buttons filter the whole page, but `render_stage_requirement()` is
  deliberately fed the month's **unfiltered** snapshot — the breakdown must keep showing
  every stage no matter which button is pressed.
- **When changing extraction logic, account for *both* reading extractors** — a fix in the
  block parser usually needs a mirror in the structured-table parser.
- Styling: reuse `metric_card()` and the `plotly_white` Plotly theme.

## Verifying changes

No test suite, and no sample workbooks are committed (they get auto-ingested into prod, so
they're kept out of the repo). For data/extraction/schema changes, drop a real IMR workbook
somewhere temporary, then write a throwaway script that builds a **SQLite** engine, ingests
that file via `ingest_reports()`, and asserts on the result — then delete the script (don't
commit it). This exercises the real ingest pipeline without touching production. Boot the app
(`DATABASE_URL=sqlite:////tmp/x.db streamlit run app.py`) to confirm pages render (HTTP 200).
