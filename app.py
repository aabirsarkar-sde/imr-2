from __future__ import annotations

import calendar
import hashlib
import os
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from sqlalchemy import (
    Column,
    Date,
    DateTime,
    Float,
    Integer,
    MetaData,
    Table,
    Text,
    create_engine,
    text,
)
from sqlalchemy.engine import Engine


APP_DIR = Path(__file__).resolve().parent
UPLOAD_DIR_NAME = "uploaded_reports"
MANUAL_DATA_FILE = "manual_readings.csv"
REPORT_GLOBS = ("*.xlsx", "*.xls", "*.csv")
MONTHS = {
    month.lower(): index
    for index, month in enumerate(calendar.month_name)
    if month
}

METRICS = {
    "Flow Rate": {
        "column": "flow_lph",
        "raw_label": "Total flow liter/hr.",
        "axis_label": "Flow Rate (liter/hr.)",
        "unit": "L/hr",
        "bad_direction": "down",
        "threshold_factor": 0.70,
    },
    "Conductivity": {
        "column": "conductivity_us_cm",
        "raw_label": "Cond. us/cm",
        "axis_label": "Conductivity (uS/cm)",
        "unit": "uS/cm",
        "bad_direction": "up",
        "threshold_factor": 1.30,
    },
}

# Units that mark an operating-parameter cell in the lower/right section of a
# plant sheet. Each reading is laid out as a (tag, value, unit) triple, e.g.
# "PI 1601" | 56 | "bar". The unit normalizes (see normalize_text) to one of
# these keys, which also tells us what kind of parameter it is.
OP_PARAM_KINDS = {
    "bar": "pressure",
    "deg c": "temperature",
    "us cm": "conductivity",
    "lit hrs": "flow",
    "m3 hr": "flow",
}

PARAM_COLUMNS = [
    "source_file",
    "plant_group",
    "plant",
    "report_date",
    "tag",
    "kind",
    "value",
    "unit",
]

# Feed pressure to the high-pressure membrane array. The standard instrument
# scheme on these reports tags the array feed as PI 1601. We surface this as the
# operating pressure at which each flow/conductivity reading was taken. Other
# pressure tags are ignored for now.
FEED_TAG = "1601"


@dataclass(frozen=True)
class ReportMeta:
    plant_group: str
    report_date: pd.Timestamp


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\n", " ").strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def clean_label(value: object) -> str:
    text = "" if pd.isna(value) else str(value)
    text = re.sub(r"\s+", " ", text).strip()
    return text or "Unknown"


def clean_sheet_name(sheet_name: str) -> str:
    return re.sub(r"\s+", " ", sheet_name).strip()


# A plant group is the set of regions a workbook covers, taken verbatim from the
# filename (e.g. "Ank,JH,Panoli"). A spelling/spacing/case slip in one month's
# filename — "Panol" vs "Panoli" — would otherwise fragment one logical group
# into two on the Portfolio/Dashboard pickers. canonicalize_plant_group() folds
# known variants to one name. To merge a future typo, add its normalized key
# here; do NOT special-case it in the parser.
PLANT_GROUP_ALIASES = {
    "ank,jh,panol": "Ank,JH,Panoli",
    "ank,jh,panoli": "Ank,JH,Panoli",
}


def canonicalize_plant_group(value: object) -> str:
    """Map a raw plant-group label to its canonical name. The lookup key is the
    label with whitespace stripped, comma-spacing collapsed, and lower-cased, so
    "Ank, JH, Panol" and "ank,jh,panoli" both resolve to the same entry. Unknown
    labels pass through cleaned (whitespace collapsed) but otherwise unchanged."""
    text = "" if pd.isna(value) else re.sub(r"\s+", " ", str(value)).strip()
    if not text:
        return text
    key = re.sub(r"\s*,\s*", ",", text).lower()
    return PLANT_GROUP_ALIASES.get(key, text)


def plant_sr_no_from_name(name: object) -> int | None:
    """Pull the trailing "(NNNN)" plant id from a sheet name, e.g.
    "Lupin Ank RO1 (1639)" -> 1639. This joins a reading to its MIS row."""
    match = re.search(r"\((\d+)\)\s*$", str(name))
    return int(match.group(1)) if match else None


def parse_report_metadata(path: Path) -> ReportMeta:
    stem = path.stem
    month_regex = "|".join(MONTHS)
    match = re.search(
        rf"\b({month_regex})\b\s*-?\s*(20\d{{2}}|19\d{{2}})",
        stem,
        flags=re.IGNORECASE,
    )

    if match:
        month_number = MONTHS[match.group(1).lower()]
        year = int(match.group(2))
        day = calendar.monthrange(year, month_number)[1]
        report_date = pd.Timestamp(year=year, month=month_number, day=day)
        plant_group = stem[match.end() :].strip(" -_")
    else:
        report_date = pd.Timestamp(path.stat().st_mtime, unit="s").normalize()
        plant_group = stem

    if not plant_group:
        plant_group = re.sub(r"^\d+\s*-\s*IMR\s*", "", stem, flags=re.IGNORECASE)
        plant_group = plant_group.strip(" -_") or stem

    return ReportMeta(
        plant_group=canonicalize_plant_group(plant_group), report_date=report_date
    )


def report_paths(root: Path) -> list[Path]:
    source_roots = [root, root / UPLOAD_DIR_NAME]
    paths: list[Path] = []

    for source_root in source_roots:
        if not source_root.exists():
            continue
        paths.extend(
            path
            for pattern in REPORT_GLOBS
            for path in source_root.glob(pattern)
            if path.name != MANUAL_DATA_FILE and not path.name.startswith("~$")
        )

    return sorted(paths)


def is_module_header(value: object) -> bool:
    text = normalize_text(value)
    return text in {"mo no", "mod no", "module no", "module number"} or bool(
        re.fullmatch(r"(mo|mod|module)\s*no", text)
    )


def is_install_date_header(value: object) -> bool:
    text = normalize_text(value)
    return "inst" in text and "date" in text


def is_flow_header(value: object) -> bool:
    text = normalize_text(value)
    return "flow" in text and (
        "total" in text or "rate" in text or "liter" in text or "lph" in text
    )


def is_conductivity_header(value: object) -> bool:
    text = normalize_text(value)
    return "cond" in text or "conductivity" in text or "us cm" in text


def is_plant_group_header(value: object) -> bool:
    text = normalize_text(value)
    return text in {"plant group", "site group", "area group"} or "plant group" in text


def is_plant_header(value: object) -> bool:
    text = normalize_text(value)
    return text in {"plant", "site", "site name", "plant name"} or "site name" in text


def is_stage_header(value: object) -> bool:
    text = normalize_text(value)
    return text in {"stage", "ro stage"}


def is_reading_date_header(value: object) -> bool:
    text = normalize_text(value)
    return (
        text in {"date", "reading date", "report date", "sample date"}
        or "reading date" in text
        or "report date" in text
    )


def find_header_row(raw: pd.DataFrame) -> int | None:
    max_rows = min(len(raw), 80)
    best_row = None
    best_score = 0

    for row_index in range(max_rows):
        values = raw.iloc[row_index].tolist()
        module_count = sum(is_module_header(value) for value in values)
        has_install_date = any(is_install_date_header(value) for value in values)
        has_flow = any(is_flow_header(value) for value in values)
        has_conductivity = any(is_conductivity_header(value) for value in values)

        score = module_count * 3 + has_install_date + has_flow + has_conductivity
        if module_count and has_install_date and has_flow and has_conductivity and score > best_score:
            best_row = row_index
            best_score = score

    return best_row


def find_stage_label(raw: pd.DataFrame, header_row: int, start_col: int, end_col: int) -> str:
    if header_row <= 0:
        return ""

    for row_index in range(header_row - 1, max(-1, header_row - 5), -1):
        values = raw.iloc[row_index, start_col:end_col].dropna().tolist()
        labels = [clean_label(value) for value in values if clean_label(value).lower() != "nan"]
        if labels:
            return labels[0]

    return ""


def find_module_blocks(raw: pd.DataFrame, header_row: int) -> list[dict[str, int | str]]:
    headers = raw.iloc[header_row].tolist()
    module_cols = [col for col, value in enumerate(headers) if is_module_header(value)]
    blocks: list[dict[str, int | str]] = []

    for index, module_col in enumerate(module_cols):
        next_module_col = module_cols[index + 1] if index + 1 < len(module_cols) else len(headers)
        span = range(module_col, next_module_col)
        block: dict[str, int | str] = {"module_col": module_col}

        for col in span:
            header = headers[col]
            if is_install_date_header(header):
                block["install_date_col"] = col
            elif is_flow_header(header):
                block["flow_col"] = col
            elif is_conductivity_header(header):
                block["conductivity_col"] = col

        required = {"module_col", "install_date_col", "flow_col", "conductivity_col"}
        if required.issubset(block):
            block["stage"] = find_stage_label(raw, header_row, module_col, next_module_col)
            block["end_col"] = next_module_col
            blocks.append(block)

    return blocks


def value_at(raw: pd.DataFrame, row: int, col: int | str) -> object:
    if not isinstance(col, int) or row >= len(raw) or col >= raw.shape[1]:
        return np.nan
    return raw.iat[row, col]


def is_bypass_marker(values: list[object]) -> bool:
    """A module is bypassed when any cell in its block reads BY PASS / Bypass.

    A bypassed membrane has a module number but empty flow/conductivity cells,
    so it would otherwise be dropped — yet it is the strongest replacement
    signal we have, so we keep it as data (status="bypass")."""
    for value in values:
        normalized = normalize_text(value)
        if "by pass" in normalized or "bypass" in normalized:
            return True
    return False


def extract_rows_from_sheet(
    raw: pd.DataFrame,
    *,
    report_path: Path,
    sheet_name: str,
    metadata: ReportMeta,
) -> list[dict[str, object]]:
    header_row = find_header_row(raw)
    if header_row is None:
        return []

    blocks = find_module_blocks(raw, header_row)
    if not blocks:
        return []

    records: list[dict[str, object]] = []
    plant = clean_sheet_name(sheet_name)
    plant_sr_no = plant_sr_no_from_name(sheet_name)

    for row_index in range(header_row + 1, len(raw)):
        for block in blocks:
            module_raw = value_at(raw, row_index, block["module_col"])
            module_number = pd.to_numeric(module_raw, errors="coerce")
            if pd.isna(module_number):
                continue

            block_cells = [
                value_at(raw, row_index, col)
                for col in range(int(block["module_col"]), int(block["end_col"]))
            ]
            bypass = is_bypass_marker(block_cells)

            if bypass:
                flow = np.nan
                conductivity = np.nan
            else:
                flow = pd.to_numeric(value_at(raw, row_index, block["flow_col"]), errors="coerce")
                conductivity = pd.to_numeric(
                    value_at(raw, row_index, block["conductivity_col"]),
                    errors="coerce",
                )
                if pd.isna(flow) and pd.isna(conductivity):
                    continue

            records.append(
                {
                    "source_file": report_path.name,
                    "plant_group": metadata.plant_group,
                    "plant": plant,
                    "plant_sr_no": plant_sr_no,
                    "stage": block.get("stage", ""),
                    "module_number": float(module_number),
                    "install_date": value_at(raw, row_index, block["install_date_col"]),
                    "report_date": metadata.report_date,
                    "flow_lph": float(flow) if pd.notna(flow) else np.nan,
                    "conductivity_us_cm": float(conductivity) if pd.notna(conductivity) else np.nan,
                    "status": "bypass" if bypass else "active",
                }
            )

    return records


def find_column(columns: pd.Index, predicate) -> str | None:
    for column in columns:
        if predicate(column):
            return str(column)
    return None


def extract_rows_from_structured_table(
    table: pd.DataFrame,
    *,
    report_path: Path,
    sheet_name: str,
    metadata: ReportMeta,
) -> list[dict[str, object]]:
    if table.empty:
        return []

    table = table.rename(columns={column: clean_label(column) for column in table.columns})
    columns = table.columns
    module_col = find_column(columns, is_module_header)
    flow_col = find_column(columns, is_flow_header)
    conductivity_col = find_column(columns, is_conductivity_header)
    reading_date_col = find_column(columns, is_reading_date_header)
    install_date_col = find_column(columns, is_install_date_header)
    plant_group_col = find_column(columns, is_plant_group_header)
    plant_col = find_column(columns, is_plant_header)
    stage_col = find_column(columns, is_stage_header)

    if not all([module_col, flow_col, conductivity_col]):
        return []

    records: list[dict[str, object]] = []
    default_plant = clean_sheet_name(sheet_name) if sheet_name else metadata.plant_group

    for _, row in table.iterrows():
        module_number = pd.to_numeric(row.get(module_col), errors="coerce")
        if pd.isna(module_number):
            continue

        bypass = is_bypass_marker(list(row.values))
        if bypass:
            flow = np.nan
            conductivity = np.nan
        else:
            flow = pd.to_numeric(row.get(flow_col), errors="coerce")
            conductivity = pd.to_numeric(row.get(conductivity_col), errors="coerce")
            if pd.isna(flow) and pd.isna(conductivity):
                continue

        reading_date = metadata.report_date
        if reading_date_col is not None and pd.notna(row.get(reading_date_col)):
            reading_date = row.get(reading_date_col)
        elif install_date_col is not None and pd.notna(row.get(install_date_col)):
            reading_date = row.get(install_date_col)

        install_date = reading_date
        if install_date_col is not None and pd.notna(row.get(install_date_col)):
            install_date = row.get(install_date_col)

        plant_group = metadata.plant_group
        if plant_group_col is not None and pd.notna(row.get(plant_group_col)):
            plant_group = canonicalize_plant_group(row.get(plant_group_col))

        plant = default_plant
        if plant_col is not None and pd.notna(row.get(plant_col)):
            plant = clean_label(row.get(plant_col))

        stage = ""
        if stage_col is not None and pd.notna(row.get(stage_col)):
            stage = clean_label(row.get(stage_col))

        records.append(
            {
                "source_file": report_path.name,
                "plant_group": plant_group,
                "plant": plant,
                "plant_sr_no": plant_sr_no_from_name(plant),
                "stage": stage,
                "module_number": float(module_number),
                "install_date": install_date,
                "report_date": reading_date,
                "flow_lph": float(flow) if pd.notna(flow) else np.nan,
                "conductivity_us_cm": float(conductivity) if pd.notna(conductivity) else np.nan,
                "status": "bypass" if bypass else "active",
            }
        )

    return records


def read_report(path: Path) -> list[dict[str, object]]:
    metadata = parse_report_metadata(path)
    records: list[dict[str, object]] = []

    if path.suffix.lower() == ".csv":
        try:
            structured = pd.read_csv(path)
        except Exception:  # noqa: BLE001
            structured = pd.DataFrame()

        if not structured.empty:
            records.extend(
                extract_rows_from_structured_table(
                    structured,
                    report_path=path,
                    sheet_name=metadata.plant_group,
                    metadata=metadata,
                )
            )
        if records:
            return records

        raw = pd.read_csv(path, header=None)
        records.extend(
            extract_rows_from_sheet(
                raw,
                report_path=path,
                sheet_name=metadata.plant_group,
                metadata=metadata,
            )
        )
        return records

    workbook = pd.ExcelFile(path)
    for sheet_name in workbook.sheet_names:
        if clean_sheet_name(sheet_name).lower() == "mis":
            continue
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
        sheet_records = extract_rows_from_sheet(
            raw,
            report_path=path,
            sheet_name=sheet_name,
            metadata=metadata,
        )
        if not sheet_records:
            structured = pd.read_excel(path, sheet_name=sheet_name)
            sheet_records = extract_rows_from_structured_table(
                structured,
                report_path=path,
                sheet_name=sheet_name,
                metadata=metadata,
            )
        records.extend(sheet_records)

    return records


def display_tag(normalized_tag: str) -> str:
    return normalized_tag.upper()


def extract_operating_parameters(
    raw: pd.DataFrame,
    *,
    report_path: Path,
    sheet_name: str,
    metadata: ReportMeta,
) -> list[dict[str, object]]:
    """Scan a raw sheet for (tag, value, unit) instrument readings.

    Operating parameters appear in two layouts: a key-value block below the
    module table, or extra columns to the right of it. Both place the unit
    immediately right of the value, with the instrument tag one cell further
    left, so a single scan keyed on the unit cell handles both. The last
    occurrence of a tag wins.
    """
    plant = clean_sheet_name(sheet_name)
    found: dict[str, dict[str, object]] = {}
    rows, cols = raw.shape

    for row_index in range(rows):
        for col_index in range(cols):
            unit = normalize_text(raw.iat[row_index, col_index])
            kind = OP_PARAM_KINDS.get(unit)
            if kind is None or col_index < 2:
                continue

            value = pd.to_numeric(raw.iat[row_index, col_index - 1], errors="coerce")
            tag = normalize_text(raw.iat[row_index, col_index - 2])
            if not tag or pd.isna(value):
                continue

            found[tag] = {
                "source_file": report_path.name,
                "plant_group": metadata.plant_group,
                "plant": plant,
                "report_date": metadata.report_date,
                "tag": display_tag(tag),
                "kind": kind,
                "value": float(value),
                "unit": unit,
            }

    return list(found.values())


def read_report_parameters(path: Path) -> list[dict[str, object]]:
    if path.suffix.lower() == ".csv":
        return []

    metadata = parse_report_metadata(path)
    records: list[dict[str, object]] = []
    workbook = pd.ExcelFile(path)
    for sheet_name in workbook.sheet_names:
        if clean_sheet_name(sheet_name).lower() == "mis":
            continue
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
        records.extend(
            extract_operating_parameters(
                raw,
                report_path=path,
                sheet_name=sheet_name,
                metadata=metadata,
            )
        )
    return records


def find_mis_header_row(raw: pd.DataFrame) -> int | None:
    for row_index in range(min(len(raw), 8)):
        joined = " ".join(normalize_text(value) for value in raw.iloc[row_index].tolist())
        if "plant sr no" in joined and "site name" in joined:
            return row_index
    return None


def map_mis_columns(header_cells: list[object]) -> dict[str, int]:
    """Map MIS field name -> column index by matching the normalized header.
    'plant sr no' is checked before the plain 'sr no' so the two don't collide."""
    mapping: dict[str, int] = {}
    for col, cell in enumerate(header_cells):
        normalized = normalize_text(cell)
        if not normalized:
            continue
        if "plant sr no" in normalized:
            mapping["plant_sr_no"] = col
        elif normalized == "zone":
            mapping["zone"] = col
        elif "zm name" in normalized:
            mapping["zm_name"] = col
        elif "site name" in normalized:
            mapping["site_name"] = col
        elif normalized == "status":
            mapping["status"] = col
        elif "membrane required" in normalized:
            mapping["membrane_required"] = col
        elif "remark" in normalized:
            mapping["remarks"] = col
    return mapping


def extract_mis_rows(
    raw: pd.DataFrame, *, report_path: Path, metadata: ReportMeta
) -> list[dict[str, object]]:
    """Parse the MIS register sheet (one per workbook).

    ZONE and ZM NAME are merged cells that are blank on continuation rows, so we
    forward-fill them down. Rows without a PLANT SR NO are skipped.
    """
    header_row = find_mis_header_row(raw)
    if header_row is None:
        return []

    mapping = map_mis_columns(raw.iloc[header_row].tolist())
    if "plant_sr_no" not in mapping:
        return []

    def cell_text(row_index: int, field: str) -> str | None:
        value = value_at(raw, row_index, mapping.get(field))
        if pd.isna(value):
            return None
        return str(value).strip() or None

    records: list[dict[str, object]] = []
    last_zone: str | None = None
    last_zm: str | None = None

    for row_index in range(header_row + 1, len(raw)):
        zone = cell_text(row_index, "zone")
        if zone:  # merged cell: only set on the group's first row
            last_zone = zone
        zm_name = cell_text(row_index, "zm_name")
        if zm_name:
            last_zm = zm_name

        plant_sr_no = pd.to_numeric(value_at(raw, row_index, mapping["plant_sr_no"]), errors="coerce")
        if pd.isna(plant_sr_no):
            continue

        records.append(
            {
                "source_file": report_path.name,
                "report_date": metadata.report_date,
                "zone": last_zone,
                "zm_name": last_zm,
                "plant_sr_no": int(plant_sr_no),
                "site_name": cell_text(row_index, "site_name"),
                "status": cell_text(row_index, "status"),
                "membrane_required": pd.to_numeric(
                    value_at(raw, row_index, mapping.get("membrane_required")), errors="coerce"
                ),
                "remarks": cell_text(row_index, "remarks"),
            }
        )

    return records


def read_report_mis(path: Path) -> list[dict[str, object]]:
    if path.suffix.lower() == ".csv":
        return []

    metadata = parse_report_metadata(path)
    records: list[dict[str, object]] = []
    workbook = pd.ExcelFile(path)
    for sheet_name in workbook.sheet_names:
        if clean_sheet_name(sheet_name).lower() != "mis":
            continue
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
        records.extend(extract_mis_rows(raw, report_path=path, metadata=metadata))
    return records


def safe_upload_name(filename: str) -> str:
    name = Path(filename).name
    name = re.sub(r"[^A-Za-z0-9., _()\-]+", "_", name).strip()
    return name or "uploaded_report.xlsx"


def save_uploaded_reports(files: list[object], root: Path) -> tuple[list[str], list[str]]:
    upload_dir = root / UPLOAD_DIR_NAME
    upload_dir.mkdir(exist_ok=True)

    saved: list[str] = []
    skipped: list[str] = []
    allowed_suffixes = {".xlsx", ".xls"}

    for uploaded_file in files:
        filename = safe_upload_name(getattr(uploaded_file, "name", "uploaded_report.xlsx"))
        suffix = Path(filename).suffix.lower()
        if suffix not in allowed_suffixes:
            skipped.append(f"{filename}: unsupported file type")
            continue
        if (root / filename).exists():
            skipped.append(f"{filename}: already exists in the main report folder")
            continue

        destination = upload_dir / filename
        destination.write_bytes(uploaded_file.getbuffer())
        saved.append(filename)

    return saved, skipped


# --------------------------------------------------------------------------- #
# Persistence (PostgreSQL)
#
# Reports are parsed exactly once, at ingest. Each report's normalized rows are
# written to the `readings` and `parameters` tables and the file is recorded in
# `ingested_files` with a content hash. On every startup we re-scan the report
# folder, but only files whose hash is new or changed are parsed again — so the
# dashboard never re-parses Excel to render; it just queries the database.
#
# The schema uses SQLAlchemy's generic column types, so the same code runs on
# Postgres (production) and SQLite (tests) without change.
# --------------------------------------------------------------------------- #

DB_METADATA = MetaData()

READINGS_TABLE = Table(
    "readings",
    DB_METADATA,
    Column("source_file", Text),
    Column("plant_group", Text),
    Column("plant", Text),
    # Trailing "(NNNN)" id parsed from the plant sheet name; joins to mis.plant_sr_no.
    Column("plant_sr_no", Integer),
    Column("stage", Text),
    Column("module_number", Float),
    Column("module_label", Text),
    Column("install_date", Date),
    Column("report_date", Date),
    Column("flow_lph", Float),
    Column("conductivity_us_cm", Float),
    Column("status", Text),
)

PARAMETERS_TABLE = Table(
    "parameters",
    DB_METADATA,
    Column("source_file", Text),
    Column("plant_group", Text),
    Column("plant", Text),
    Column("report_date", Date),
    Column("tag", Text),
    Column("kind", Text),
    Column("value", Float),
    Column("unit", Text),
)

# The MIS sheet (one per workbook) is the plant register: real ZONE, owner
# (ZM NAME), per-plant STATUS, membrane-required quantity, and PLANT SR NO,
# which joins to readings.plant_sr_no.
MIS_TABLE = Table(
    "mis",
    DB_METADATA,
    Column("source_file", Text),
    Column("report_date", Date),
    Column("zone", Text),
    Column("zm_name", Text),
    Column("plant_sr_no", Integer),
    Column("site_name", Text),
    Column("status", Text),
    Column("membrane_required", Float),
    Column("remarks", Text),
)

INGESTED_FILES_TABLE = Table(
    "ingested_files",
    DB_METADATA,
    Column("filename", Text, primary_key=True),
    Column("content_hash", Text, nullable=False),
    Column("ingested_at", DateTime),
    Column("n_readings", Integer),
    Column("n_parameters", Integer),
    Column("n_mis", Integer),
)

READINGS_COLUMNS = [c.name for c in READINGS_TABLE.columns]
MIS_COLUMNS = [c.name for c in MIS_TABLE.columns]


def database_url() -> str | None:
    """Resolve the DB connection string from DATABASE_URL or st.secrets, and
    normalize it to the psycopg2 driver SQLAlchemy expects."""
    url = os.environ.get("DATABASE_URL")
    if not url:
        try:
            url = st.secrets["postgres"]["url"]
        except Exception:  # noqa: BLE001 - secrets file may be absent
            url = None
    if not url:
        return None
    if url.startswith("postgres://"):
        url = "postgresql+psycopg2://" + url[len("postgres://") :]
    elif url.startswith("postgresql://"):
        url = "postgresql+psycopg2://" + url[len("postgresql://") :]
    return url


@st.cache_resource
def get_engine() -> Engine | None:
    url = database_url()
    if not url:
        return None
    return create_engine(url, pool_pre_ping=True)


def init_db(engine: Engine) -> None:
    DB_METADATA.create_all(engine)


def file_fingerprint(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


# Conductivity readings above this are data-entry errors (the raw sheets contain
# values like 2,984,050 uS/cm); blank them rather than let them skew the analysis.
MAX_PLAUSIBLE_CONDUCTIVITY_US_CM = 50000.0


def normalize_reading_records(records: list[dict[str, object]]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame(columns=READINGS_COLUMNS)

    df = pd.DataFrame(records)
    df["install_date"] = pd.to_datetime(df["install_date"], errors="coerce")
    df["report_date"] = pd.to_datetime(df["report_date"], errors="coerce")
    df["flow_lph"] = pd.to_numeric(df["flow_lph"], errors="coerce")
    df["conductivity_us_cm"] = pd.to_numeric(df["conductivity_us_cm"], errors="coerce")
    df["module_number"] = pd.to_numeric(df["module_number"], errors="coerce")
    df["module_label"] = df["module_number"].map(format_module_number)
    if "plant_sr_no" not in df.columns:
        df["plant_sr_no"] = pd.NA
    df["plant_sr_no"] = pd.to_numeric(df["plant_sr_no"], errors="coerce").astype("Int64")
    if "status" not in df.columns:
        df["status"] = "active"
    df["status"] = df["status"].fillna("active")
    # Implausibly high conductivity is a data-entry error, not a reading.
    df.loc[df["conductivity_us_cm"] > MAX_PLAUSIBLE_CONDUCTIVITY_US_CM, "conductivity_us_cm"] = np.nan
    df = df.dropna(subset=["report_date", "module_number"])
    # Keep a row if it has a module number AND (it is bypassed OR has at least
    # one of flow/conductivity). Bypassed membranes carry no flow/conductivity
    # but are the strongest replacement signal, so they must be retained.
    keep = (
        (df["status"] == "bypass")
        | df["flow_lph"].notna()
        | df["conductivity_us_cm"].notna()
    )
    df = df[keep]
    # One reading per (file, plant, stage, module, month); a re-ingest of the
    # same file replaces its rows wholesale, so file-local dedup is enough.
    df = df.drop_duplicates(
        subset=["source_file", "plant", "stage", "module_number", "report_date"],
        keep="last",
    )
    return df.reindex(columns=READINGS_COLUMNS)


def normalize_parameter_records(records: list[dict[str, object]]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame(columns=PARAM_COLUMNS)

    df = pd.DataFrame(records)
    df["report_date"] = pd.to_datetime(df["report_date"], errors="coerce")
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df = df.dropna(subset=["report_date", "value"])
    df = df.drop_duplicates(
        subset=["source_file", "plant", "report_date", "tag"], keep="last"
    )
    return df.reindex(columns=PARAM_COLUMNS)


def normalize_mis_records(records: list[dict[str, object]]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame(columns=MIS_COLUMNS)

    df = pd.DataFrame(records)
    df["report_date"] = pd.to_datetime(df["report_date"], errors="coerce")
    df["plant_sr_no"] = pd.to_numeric(df["plant_sr_no"], errors="coerce")
    df["membrane_required"] = pd.to_numeric(df["membrane_required"], errors="coerce")
    df = df.dropna(subset=["plant_sr_no"])
    df["plant_sr_no"] = df["plant_sr_no"].astype(int)
    df = df.drop_duplicates(subset=["source_file", "plant_sr_no"], keep="last")
    return df.reindex(columns=MIS_COLUMNS)


def ingest_reports(engine: Engine, data_dir: str) -> dict[str, list[str]]:
    """Parse and store any report whose content hash is new or changed.

    Returns a summary of what was ingested / skipped (already up to date) /
    failed. Already-ingested files are never re-parsed.
    """
    root = Path(data_dir)
    summary: dict[str, list[str]] = {"ingested": [], "skipped": [], "failed": []}

    with engine.connect() as conn:
        known = dict(
            conn.execute(text("SELECT filename, content_hash FROM ingested_files")).all()
        )

    for path in report_paths(root):
        try:
            fingerprint = file_fingerprint(path)
        except OSError as exc:
            summary["failed"].append(f"{path.name}: {exc}")
            continue

        if known.get(path.name) == fingerprint:
            summary["skipped"].append(path.name)
            continue

        try:
            readings = normalize_reading_records(read_report(path))
            parameters = normalize_parameter_records(read_report_parameters(path))
            mis = normalize_mis_records(read_report_mis(path))
        except Exception as exc:  # noqa: BLE001 - one bad report shouldn't sink the rest
            summary["failed"].append(f"{path.name}: {exc}")
            continue

        if readings.empty:
            summary["failed"].append(f"{path.name}: no module table found")
            continue

        with engine.begin() as conn:
            conn.execute(
                text("DELETE FROM readings WHERE source_file = :f"), {"f": path.name}
            )
            conn.execute(
                text("DELETE FROM parameters WHERE source_file = :f"), {"f": path.name}
            )
            conn.execute(
                text("DELETE FROM mis WHERE source_file = :f"), {"f": path.name}
            )
            readings.to_sql(
                "readings", conn, if_exists="append", index=False,
                method="multi", chunksize=1000,
            )
            if not parameters.empty:
                parameters.to_sql(
                    "parameters", conn, if_exists="append", index=False,
                    method="multi", chunksize=1000,
                )
            if not mis.empty:
                mis.to_sql(
                    "mis", conn, if_exists="append", index=False,
                    method="multi", chunksize=1000,
                )
            conn.execute(
                text(
                    """
                    INSERT INTO ingested_files
                        (filename, content_hash, ingested_at, n_readings, n_parameters, n_mis)
                    VALUES (:f, :h, :ts, :nr, :np, :nm)
                    ON CONFLICT (filename) DO UPDATE SET
                        content_hash = excluded.content_hash,
                        ingested_at = excluded.ingested_at,
                        n_readings = excluded.n_readings,
                        n_parameters = excluded.n_parameters,
                        n_mis = excluded.n_mis
                    """
                ),
                {
                    "f": path.name,
                    "h": fingerprint,
                    "ts": datetime.now(timezone.utc),
                    "nr": int(len(readings)),
                    "np": int(len(parameters)),
                    "nm": int(len(mis)),
                },
            )
        summary["ingested"].append(path.name)

    return summary


@st.cache_data(show_spinner="Loading readings from the database...")
def load_readings() -> pd.DataFrame:
    engine = get_engine()
    if engine is None:
        return pd.DataFrame(columns=READINGS_COLUMNS)
    df = pd.read_sql(
        "SELECT * FROM readings",
        engine,
        parse_dates=["install_date", "report_date"],
    )
    if df.empty:
        return df
    df["module_label"] = df["module_label"].astype("string")
    # Fold filename-typo group variants ("Panol" vs "Panoli") at read time too, so
    # rows ingested before canonicalize_plant_group() existed are merged without a
    # forced re-parse.
    df["plant_group"] = df["plant_group"].map(canonicalize_plant_group)
    return df.sort_values(["plant_group", "plant", "module_number", "report_date"]).reset_index(
        drop=True
    )


@st.cache_data(show_spinner="Loading operating parameters from the database...")
def load_parameters() -> pd.DataFrame:
    engine = get_engine()
    if engine is None:
        return pd.DataFrame(columns=PARAM_COLUMNS)
    df = pd.read_sql("SELECT * FROM parameters", engine, parse_dates=["report_date"])
    if df.empty:
        return pd.DataFrame(columns=PARAM_COLUMNS)
    df["plant_group"] = df["plant_group"].map(canonicalize_plant_group)
    return df.sort_values(["plant_group", "plant", "tag", "report_date"]).reset_index(drop=True)


@st.cache_data(show_spinner="Loading the plant register (MIS) from the database...")
def load_mis() -> pd.DataFrame:
    engine = get_engine()
    if engine is None:
        return pd.DataFrame(columns=MIS_COLUMNS)
    df = pd.read_sql("SELECT * FROM mis", engine, parse_dates=["report_date"])
    if df.empty:
        return pd.DataFrame(columns=MIS_COLUMNS)
    return df.sort_values(["zone", "plant_sr_no", "report_date"]).reset_index(drop=True)


def feed_pressure_series(
    params: pd.DataFrame, plant_group: str, plant: str
) -> pd.DataFrame:
    pressures = params[
        (params["plant_group"] == plant_group)
        & (params["plant"] == plant)
        & (params["kind"] == "pressure")
    ]
    if pressures.empty:
        return pd.DataFrame(columns=["report_date", "value"])

    rows: list[dict[str, object]] = []
    for report_date, group in pressures.groupby("report_date"):
        feed = group[group["tag"].str.contains(FEED_TAG, regex=False)]["value"]
        if feed.empty:
            continue
        rows.append({"report_date": report_date, "value": float(feed.iloc[0])})

    series = pd.DataFrame(rows)
    if series.empty:
        return pd.DataFrame(columns=["report_date", "value"])
    return series.sort_values("report_date").reset_index(drop=True)


def format_module_number(value: object) -> str:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return "Unknown"
    if float(number).is_integer():
        return str(int(number))
    return f"{number:g}"


def aggregate_series(df: pd.DataFrame, metric_col: str) -> pd.DataFrame:
    metric_df = df.dropna(subset=[metric_col]).copy()
    if metric_df.empty:
        return metric_df

    grouped = (
        metric_df.groupby("report_date", as_index=False)
        .agg(
            value=(metric_col, "mean"),
            install_date=("install_date", "first"),
            source_file=("source_file", "first"),
            stage=("stage", "first"),
        )
        .sort_values("report_date")
    )
    return grouped


# Teal, distinct from the blue reading line, so the operating-pressure context
# reads clearly on its own right-hand axis.
PRESSURE_LINE_COLOR = "#0d9488"


def add_feed_pressure_axis(fig: go.Figure, pressure: pd.DataFrame | None) -> None:
    """Overlay PI 1601 feed pressure as a dashed line on a secondary right-hand axis.

    Shared by the single-module and comparison charts so the operating pressure
    each reading was taken at is visible at a glance, not just on hover. Pressure
    is a plant-level series, so in the comparison view this is one shared line
    behind all the module lines.
    """
    if pressure is None or pressure.empty:
        return
    ordered = pressure.dropna(subset=["value"]).sort_values("report_date")
    if ordered.empty:
        return
    fig.add_trace(
        go.Scatter(
            x=ordered["report_date"],
            y=ordered["value"],
            mode="lines+markers",
            name="Feed pressure (PI 1601)",
            yaxis="y2",
            line={"color": PRESSURE_LINE_COLOR, "width": 2, "dash": "dot"},
            marker={"size": 6, "symbol": "diamond"},
            hovertemplate="Feed pressure (PI 1601)<br>%{x|%d %b %Y}<br>%{y:,.2f} bar<extra></extra>",
        )
    )
    fig.update_layout(
        yaxis2={
            "title": "Feed Pressure (bar, PI 1601)",
            "overlaying": "y",
            "side": "right",
            "showgrid": False,
            "automargin": True,
        }
    )


def make_chart(
    series: pd.DataFrame,
    *,
    metric_name: str,
    metric_config: dict[str, object],
    pressure: pd.DataFrame | None = None,
) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=series["report_date"],
            y=series["value"],
            mode="lines+markers",
            name="Historical Reading",
            line={"color": "#2563eb", "width": 3},
            marker={"size": 8},
            hovertemplate="%{x|%d %b %Y}<br>%{y:,.2f}<extra></extra>",
        )
    )

    add_feed_pressure_axis(fig, pressure)

    fig.update_layout(
        title=f"{metric_name} History",
        xaxis_title="Report Date",
        yaxis_title=str(metric_config["axis_label"]),
        hovermode="x unified",
        margin={"l": 20, "r": 20, "t": 60, "b": 20},
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "right", "x": 1},
        template="plotly_white",
        height=470,
    )
    return fig


def metric_card(title: str, value: str, subtitle: str = "", color: str = "#0f172a") -> None:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">{title}</div>
            <div class="metric-value" style="color:{color};">{value}</div>
            <div class="metric-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def rerun_app() -> None:
    if hasattr(st, "rerun"):
        st.rerun()
    else:
        st.experimental_rerun()


def sidebar_filters(df: pd.DataFrame) -> tuple[pd.DataFrame, str, str, str, str]:
    st.sidebar.header("Filters")

    plant_groups = sorted(df["plant_group"].dropna().unique())
    if len(plant_groups) > 1:
        selected_group = st.sidebar.selectbox("Plant Group", plant_groups)
    else:
        selected_group = plant_groups[0]
        st.sidebar.caption(f"Plant group: {selected_group}")

    group_df = df[df["plant_group"] == selected_group]
    plants = sorted(group_df["plant"].dropna().unique())
    selected_plant = st.sidebar.selectbox("Plant", plants)

    plant_df = group_df[group_df["plant"] == selected_plant]
    module_labels = sorted(
        plant_df["module_label"].dropna().unique(),
        key=lambda item: pd.to_numeric(item, errors="coerce"),
    )
    selected_module = st.sidebar.selectbox("Module Number", module_labels)

    metric_name = st.sidebar.radio("View Metric", list(METRICS), horizontal=False)

    filtered = plant_df[plant_df["module_label"] == selected_module].copy()
    return filtered, selected_group, selected_plant, selected_module, metric_name


def render_add_data_controls() -> None:
    st.sidebar.markdown("---")
    with st.sidebar.expander("Add New Data"):
        uploaded_files = st.file_uploader(
            "Upload monthly Excel report",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="Saved files are loaded from the uploaded_reports folder on the next refresh.",
        )
        if st.button("Save Uploaded Report(s)", disabled=not uploaded_files):
            saved, skipped = save_uploaded_reports(uploaded_files or [], APP_DIR)
            if saved:
                st.success(f"Saved {len(saved)} report(s). Ingesting on refresh…")
            if skipped:
                st.warning("; ".join(skipped))
            # Saved files are parsed into the database by ingest on the rerun.
            rerun_app()


def module_series_map(
    plant_df: pd.DataFrame, module_labels: list[str], metric_col: str
) -> dict[str, pd.DataFrame]:
    series_map: dict[str, pd.DataFrame] = {}
    for module_label in module_labels:
        module_df = plant_df[plant_df["module_label"] == module_label]
        series = aggregate_series(module_df, metric_col)
        if not series.empty:
            series_map[module_label] = series
    return series_map


def make_comparison_chart(
    series_map: dict[str, pd.DataFrame],
    *,
    metric_config: dict[str, object],
    average: pd.DataFrame | None,
    pressure: pd.DataFrame | None = None,
) -> go.Figure:
    fig = go.Figure()
    for module_label, series in series_map.items():
        ordered = series.sort_values("report_date")
        fig.add_trace(
            go.Scatter(
                x=ordered["report_date"],
                y=ordered["value"],
                mode="lines+markers",
                name=f"Module {module_label}",
                marker={"size": 7},
                hovertemplate=f"Module {module_label}<br>%{{x|%d %b %Y}}<br>%{{y:,.2f}}<extra></extra>",
            )
        )

    if average is not None and not average.empty:
        ordered = average.sort_values("report_date")
        fig.add_trace(
            go.Scatter(
                x=ordered["report_date"],
                y=ordered["value"],
                mode="lines",
                name="Plant average",
                line={"color": "#475569", "width": 2, "dash": "dash"},
                hovertemplate="Plant average<br>%{x|%d %b %Y}<br>%{y:,.2f}<extra></extra>",
            )
        )

    add_feed_pressure_axis(fig, pressure)

    fig.update_layout(
        title="Module Comparison",
        xaxis_title="Report Date",
        yaxis_title=str(metric_config["axis_label"]),
        hovermode="x unified",
        margin={"l": 20, "r": 20, "t": 60, "b": 20},
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "right", "x": 1},
        template="plotly_white",
        height=470,
    )
    return fig


def build_comparison_table(
    series_map: dict[str, pd.DataFrame], metric_config: dict[str, object]
) -> pd.DataFrame:
    unit = str(metric_config["unit"])
    rows: list[dict[str, object]] = []

    for module_label, series in series_map.items():
        ordered = series.sort_values("report_date")
        latest_row = ordered.iloc[-1]
        latest = float(latest_row["value"])
        rows.append(
            {
                "Module": module_label,
                "Latest": f"{latest:,.2f} {unit}",
                "Latest Date": pd.Timestamp(latest_row["report_date"]).strftime("%d %b %Y"),
                "_sort": pd.to_numeric(module_label, errors="coerce"),
            }
        )

    table = pd.DataFrame(rows)
    if table.empty:
        return table
    return table.sort_values("_sort").drop(columns="_sort").reset_index(drop=True)


def render_comparison_section(
    plant_df: pd.DataFrame,
    selected_module: str,
    metric_name: str,
    metric_config: dict[str, object],
    pressure: pd.DataFrame | None = None,
) -> None:
    module_labels = sorted(
        plant_df["module_label"].dropna().unique(),
        key=lambda item: pd.to_numeric(item, errors="coerce"),
    )
    if len(module_labels) < 2:
        return

    st.markdown("---")
    st.subheader("Compare Modules")
    st.caption(
        f"Overlay multiple modules for {metric_name.lower()} to spot outliers against their peers. "
        "The dashed teal line (right axis) is the shared plant feed pressure (PI 1601) "
        "all modules were operating at."
    )

    default_selection = [selected_module] if selected_module in module_labels else module_labels[:1]
    chosen = st.multiselect(
        "Modules to compare",
        options=module_labels,
        default=default_selection,
        key="compare_modules",
    )
    show_average = st.checkbox("Show plant average", value=True, key="compare_average")

    if not chosen:
        st.info("Select at least one module to compare.")
        return

    metric_col = str(metric_config["column"])
    series_map = module_series_map(plant_df, chosen, metric_col)
    if not series_map:
        st.warning("No valid readings for the selected modules and metric.")
        return

    average = aggregate_series(plant_df, metric_col) if show_average else None

    st.plotly_chart(
        make_comparison_chart(
            series_map, metric_config=metric_config, average=average, pressure=pressure
        ),
        width="stretch",
    )

    table = build_comparison_table(series_map, metric_config)
    if not table.empty:
        st.dataframe(table, width="stretch", hide_index=True)


# The replacement page offers two ways to flag a membrane. "Peer outlier" is
# relative — who stands out above the rest of this stage. "Absolute limit" is a
# fixed operator-set cutoff — who is over an acceptable conductivity outright,
# which catches a whole stage that is uniformly degraded (where a relative test
# finds nothing, because the peer baseline itself is elevated).
PEER_METHOD = "Peer outlier (IQR)"
ABSOLUTE_METHOD = "Absolute limit (uS/cm)"
OUTLIER_METHODS = (PEER_METHOD, ABSOLUTE_METHOD)
DEFAULT_ABSOLUTE_LIMIT = 500.0

# Fleet-wide degradation cutoffs rise per RO stage: the reject side naturally
# runs at higher TDS, so a higher bar avoids false positives. Used by the
# Portfolio page via evaluate_stage_readings (absolute-limit method).
STAGE_CONDUCTIVITY_CUTOFFS = {"i": 1000.0, "ii": 1500.0, "iii": 2000.0}
DEFAULT_STAGE_CUTOFF = 1000.0


def stage_cutoff(stage_label: str) -> float:
    normalized = normalize_text(stage_label)
    for roman in ("iii", "ii", "i"):
        if re.search(rf"\b{roman}\b", normalized):
            return STAGE_CONDUCTIVITY_CUTOFFS[roman]
    return DEFAULT_STAGE_CUTOFF


def iqr_outlier_stats(values: pd.Series, sensitivity: float) -> dict[str, object]:
    """High-side IQR (Tukey) cutoff plus a per-module 'IQRs above Q3' score.

    Conductivity degradation is a *rise*, so only the upper tail matters. The
    median and upper fence are returned in uS/cm (for the chart line and the flag
    test). IQR is robust: the very outliers we are hunting don't inflate the
    cutoff the way a mean + std-dev would. Below 4 readings the quartiles are too
    unstable, so the fence is left NaN and the caller flags nothing.
    """
    clean = pd.to_numeric(values, errors="coerce").dropna()
    n = int(clean.shape[0])
    stats: dict[str, object] = {
        "n": n,
        "median": float(clean.median()) if n else float("nan"),
        "upper_fence": float("nan"),
        "scores": pd.Series(np.zeros(len(values)), index=values.index),
    }
    if n < 4:
        return stats

    q1 = float(clean.quantile(0.25))
    q3 = float(clean.quantile(0.75))
    iqr = q3 - q1
    stats["upper_fence"] = q3 + sensitivity * iqr
    if iqr > 0:
        stats["scores"] = (values - q3) / iqr
    return stats


def make_outlier_chart(
    readings: pd.DataFrame,
    *,
    median: float,
    fence: float,
    unit: str,
    title: str,
    fence_label: str = "Cutoff",
) -> go.Figure:
    colors = ["#dc2626" if flag else "#2563eb" for flag in readings["flag"]]
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=readings["module_label"].astype(str),
            y=readings["conductivity"],
            marker_color=colors,
            customdata=readings["pct_vs_median"],
            hovertemplate=(
                "Module %{x}<br>%{y:,.0f} " + unit
                + "<br>%{customdata:+.0f}% vs stage median<extra></extra>"
            ),
        )
    )
    if pd.notna(median):
        fig.add_hline(
            y=median,
            line_dash="dot",
            line_color="#64748b",
            annotation_text="Stage median",
            annotation_position="top left",
        )
    if pd.notna(fence):
        fig.add_hline(
            y=fence,
            line_dash="dash",
            line_color="#dc2626",
            annotation_text=fence_label,
            annotation_position="top right",
        )
    fig.update_layout(
        title=title,
        xaxis_title="Module",
        yaxis_title=f"Conductivity ({unit})",
        margin={"l": 20, "r": 20, "t": 60, "b": 20},
        template="plotly_white",
        height=420,
        showlegend=False,
    )
    fig.update_xaxes(
        type="category",
        categoryorder="array",
        categoryarray=readings["module_label"].astype(str).tolist(),
    )
    return fig


def evaluate_stage_readings(
    readings: pd.DataFrame, *, method: str, sensitivity: float, limit: float
) -> tuple[pd.DataFrame, float, float, bool]:
    """Annotate one stage's per-module readings with flag / score / pct_vs_median.

    Returns (readings, stage_median, cutoff_fence, too_few). The peer baseline is
    per stage, so this is called once per stage when consolidating a whole plant.
    """
    readings = readings.copy()
    median = float(pd.to_numeric(readings["conductivity"], errors="coerce").median())
    n_modules = int(readings["conductivity"].notna().sum())
    too_few = False

    if method == PEER_METHOD:
        stats = iqr_outlier_stats(readings["conductivity"], sensitivity)
        fence = float(stats["upper_fence"])
        readings["score"] = pd.Series(stats["scores"]).to_numpy()
        too_few = n_modules < 4
    else:  # ABSOLUTE_METHOD
        fence = float(limit)
        readings["score"] = readings["conductivity"] / fence if fence else np.nan

    readings["pct_vs_median"] = (
        (readings["conductivity"] / median - 1.0) * 100.0 if median else np.nan
    )
    readings["flag"] = readings["conductivity"] > fence if pd.notna(fence) else False
    readings = readings.sort_values("conductivity", ascending=False).reset_index(drop=True)
    return readings, median, fence, too_few


def render_outlier_section(plant_df: pd.DataFrame) -> None:
    metric_config = METRICS["Conductivity"]
    col = str(metric_config["column"])
    unit = str(metric_config["unit"])

    data = plant_df.dropna(subset=[col]).copy()
    if data.empty:
        st.info("No conductivity readings for this plant.")
        return

    data["stage_label"] = data["stage"].fillna("").map(
        lambda value: str(value).strip() if str(value).strip() else "Unspecified"
    )

    date_options = sorted(data["report_date"].dropna().unique(), reverse=True)
    if not date_options:
        st.info("No dated conductivity readings for this plant.")
        return

    controls = st.columns([2, 3])
    with controls[0]:
        report_date = st.selectbox(
            "Report month",
            date_options,
            format_func=lambda value: pd.Timestamp(value).strftime("%b %Y"),
            key="outlier_date",
        )
    with controls[1]:
        method = st.selectbox("Detection method", OUTLIER_METHODS, key="outlier_method")

    if method == PEER_METHOD:
        sensitivity = st.slider(
            "Sensitivity (k)",
            min_value=1.0,
            max_value=5.0,
            value=1.5,
            step=0.5,
            help="Stricter = fewer flags. IQR cutoff = Q3 + k·IQR (1.5 standard, 3.0 = extreme).",
            key="outlier_k",
        )
        limit = DEFAULT_ABSOLUTE_LIMIT
        score_label = "IQR above Q3"
        score_fmt = lambda value: f"{value:,.1f}"
        cutoff_label = "IQR Fence"
    else:  # ABSOLUTE_METHOD
        sensitivity = 1.5
        limit = st.number_input(
            f"Replace above ({unit})",
            min_value=0.0,
            value=DEFAULT_ABSOLUTE_LIMIT,
            step=50.0,
            help="Flag every module over this fixed conductivity, regardless of its peers. "
            "Use when a whole stage is degraded and the peer test finds nothing.",
            key="outlier_limit",
        )
        score_label = "x Limit"
        score_fmt = lambda value: f"{value:,.2f}x"
        cutoff_label = "Replacement Limit"

    snapshot = data[data["report_date"] == report_date]
    month_text = pd.Timestamp(report_date).strftime("%b %Y")
    stages_present = sorted(snapshot["stage_label"].unique())

    # Evaluate every stage (each against its own peers) and collect the flags.
    per_stage: dict[str, tuple[pd.DataFrame, float, float, bool]] = {}
    flagged_records: list[dict[str, object]] = []
    total_modules = 0
    too_few_stages: list[str] = []

    for stage in stages_present:
        stage_readings = snapshot[snapshot["stage_label"] == stage].groupby(
            "module_label", as_index=False
        ).agg(
            conductivity=(col, "mean"),
            flow=("flow_lph", "mean"),
            install_date=("install_date", "first"),
        )
        if stage_readings.empty:
            continue
        evaluated, median, fence, too_few = evaluate_stage_readings(
            stage_readings, method=method, sensitivity=sensitivity, limit=limit
        )
        per_stage[stage] = (evaluated, median, fence, too_few)
        total_modules += int(evaluated["conductivity"].notna().sum())
        if too_few:
            too_few_stages.append(stage)
        for _, row in evaluated[evaluated["flag"]].iterrows():
            flagged_records.append(
                {
                    "Stage": stage,
                    "Module": row["module_label"],
                    "conductivity": float(row["conductivity"]),
                    "pct_vs_median": float(row["pct_vs_median"]),
                    "score": float(row["score"]),
                    "install_date": row["install_date"],
                }
            )

    # ----- Consolidated headline + table (the whole plant, this month) -----
    n_flagged = len(flagged_records)
    card1, card2, card3 = st.columns(3)
    with card1:
        flag_color = "#dc2626" if n_flagged else "#16a34a"
        metric_card("Modules Needing Replacement", str(n_flagged), month_text, flag_color)
    with card2:
        metric_card("Modules Evaluated", str(total_modules), f"{len(stages_present)} stage(s)")
    with card3:
        cutoff_text = f"{limit:,.0f} {unit}" if method == ABSOLUTE_METHOD else f"k = {sensitivity:g}"
        metric_card("Method", method.split(" (")[0], cutoff_text)

    if too_few_stages:
        st.info(
            "Too few modules (<4) for a reliable peer test in: "
            + ", ".join(too_few_stages)
            + ". Use an absolute limit to flag those stages."
        )

    if flagged_records:
        consolidated = pd.DataFrame(flagged_records).sort_values(
            ["Stage", "conductivity"], ascending=[True, False]
        )
        display_table = pd.DataFrame(
            {
                "Stage": consolidated["Stage"],
                "Module": consolidated["Module"],
                f"Conductivity ({unit})": consolidated["conductivity"].map(lambda v: f"{v:,.0f}"),
                "vs Stage Median": consolidated["pct_vs_median"].map(lambda v: f"{v:+,.0f}%"),
                score_label: consolidated["score"].map(score_fmt),
                "Install Date": consolidated["install_date"].map(
                    lambda v: pd.Timestamp(v).strftime("%d %b %Y") if pd.notna(v) else "Unknown"
                ),
            }
        )
        st.markdown(
            f"**{n_flagged} module(s) need replacement** across "
            f"{consolidated['Stage'].nunique()} stage(s) — {month_text}:"
        )
        st.dataframe(display_table, width="stretch", hide_index=True)
        st.download_button(
            "Download replacement list (CSV)",
            display_table.to_csv(index=False).encode("utf-8"),
            file_name=f"replacement_candidates_{month_text.replace(' ', '_')}.csv",
            mime="text/csv",
        )
    else:
        st.success(f"No modules need replacement in any stage for {month_text}.")

    # ----- Per-stage drill-down chart -----
    if per_stage:
        st.markdown("---")
        st.markdown("**Stage detail**")
        stage = st.selectbox("Stage to chart", list(per_stage), key="outlier_stage")
        evaluated, median, fence, _ = per_stage[stage]
        st.plotly_chart(
            make_outlier_chart(
                evaluated,
                median=median,
                fence=fence,
                unit=unit,
                title=f"{stage} conductivity by module — {month_text}",
                fence_label=cutoff_label,
            ),
            width="stretch",
        )


def select_group_plant(df: pd.DataFrame) -> tuple[str, str]:
    """Sidebar plant picker for the standalone Replacement page (mirrors the
    dashboard's group/plant selectors, with its own widget keys)."""
    st.sidebar.header("Plant")
    groups = sorted(df["plant_group"].dropna().unique())
    if len(groups) > 1:
        group = st.sidebar.selectbox("Plant Group", groups, key="rep_group")
    else:
        group = groups[0]
        st.sidebar.caption(f"Plant group: {group}")
    plants = sorted(df[df["plant_group"] == group]["plant"].dropna().unique())
    plant = st.sidebar.selectbox("Plant", plants, key="rep_plant")
    return group, plant


def render_replacement_page(df: pd.DataFrame) -> None:
    st.title("Membrane Replacement Candidates")
    st.caption(
        "For the selected plant and month, every stage is checked and the flagged "
        "membranes are consolidated into one list with a total count — flagged either "
        "for standing out above their stage peers (conductivity outliers) or for being "
        "over a fixed acceptable limit. Each stage is judged on its own readings, so a "
        "naturally high-TDS stage isn't penalised against a cleaner one."
    )
    group, plant = select_group_plant(df)
    st.caption(f"{group} | {plant}")
    plant_df = df[(df["plant_group"] == group) & (df["plant"] == plant)]
    render_outlier_section(plant_df)


@st.cache_data(show_spinner="Building the fleet view...")
def compute_fleet_status(readings: pd.DataFrame, mis: pd.DataFrame) -> pd.DataFrame:
    """Annotate every reading with bypass / degraded / need, for all months.

    Degradation reuses evaluate_stage_readings (absolute-limit method) once per
    (plant, stage, month) with a stage-aware cutoff, so the definition matches
    the Replacement page. A module "needs attention" if it is bypassed or
    degraded. Returns one row per (plant, stage, module, month).
    """
    col = str(METRICS["Conductivity"]["column"])
    columns = [
        "plant_group", "plant", "plant_sr_no", "zone", "stage_label",
        "module_label", "report_date", "conductivity", "install_date",
        "status", "degraded", "need", "cutoff",
    ]
    if readings.empty:
        return pd.DataFrame(columns=columns)

    data = readings.copy()
    data["stage_label"] = data["stage"].fillna("").map(
        lambda value: str(value).strip() if str(value).strip() else "Unspecified"
    )

    # plant_sr_no -> zone, from the most recent MIS row for that plant.
    zone_by_sr: dict[object, object] = {}
    if not mis.empty:
        latest_mis = mis.sort_values("report_date").drop_duplicates("plant_sr_no", keep="last")
        for _, row in latest_mis.iterrows():
            zone = row.get("zone")
            zone_by_sr[row["plant_sr_no"]] = zone if (pd.notna(zone) and str(zone).strip()) else None

    plant_meta: dict[str, tuple[object, object, str]] = {}
    for plant, pdf in data.groupby("plant"):
        plant_group = pdf["plant_group"].iloc[0]
        srs = pdf["plant_sr_no"].dropna()
        plant_sr = int(srs.iloc[0]) if not srs.empty else None
        zone = zone_by_sr.get(plant_sr) if plant_sr is not None else None
        plant_meta[plant] = (plant_group, plant_sr, zone if zone else "Unknown")

    records: list[dict[str, object]] = []
    for (plant, stage_label, report_date), group in data.groupby(
        ["plant", "stage_label", "report_date"]
    ):
        agg = group.groupby("module_label", as_index=False).agg(
            conductivity=(col, "mean"),
            install_date=("install_date", "first"),
            is_bypass=("status", lambda s: bool((s == "bypass").any())),
        )
        active = agg[~agg["is_bypass"]]
        cutoff = stage_cutoff(stage_label)
        degraded_map: dict[object, bool] = {}
        if not active.empty:
            evaluated, _, _, _ = evaluate_stage_readings(
                active, method=ABSOLUTE_METHOD, sensitivity=1.5, limit=cutoff
            )
            degraded_map = dict(zip(evaluated["module_label"], evaluated["flag"]))

        plant_group, plant_sr, zone = plant_meta[plant]
        for _, row in agg.iterrows():
            bypass = bool(row["is_bypass"])
            degraded = (not bypass) and bool(degraded_map.get(row["module_label"], False))
            records.append(
                {
                    "plant_group": plant_group,
                    "plant": plant,
                    "plant_sr_no": plant_sr,
                    "zone": zone,
                    "stage_label": stage_label,
                    "module_label": row["module_label"],
                    "report_date": report_date,
                    "conductivity": row["conductivity"],
                    "install_date": row["install_date"],
                    "status": "bypass" if bypass else "active",
                    "degraded": degraded,
                    "need": bypass or degraded,
                    "cutoff": cutoff,
                }
            )

    return pd.DataFrame(records, columns=columns)


def build_plant_ranking(snapshot: pd.DataFrame, latest: pd.Timestamp) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for plant, pdf in snapshot.groupby("plant"):
        module_count = len(pdf)
        bypassed = int((pdf["status"] == "bypass").sum())
        degraded = int(pdf["degraded"].sum())
        need = int(pdf["need"].sum())
        avg_cond = pdf.loc[pdf["status"] == "active", "conductivity"].mean()
        ages = (
            pd.Timestamp(latest) - pd.to_datetime(pdf["install_date"], errors="coerce")
        ).dt.days / 365.25
        rows.append(
            {
                "Plant": plant,
                "Zone": pdf["zone"].iloc[0],
                "Modules": module_count,
                "Bypassed": bypassed,
                "Degraded": degraded,
                "Need": need,
                "Need %": round(need / module_count * 100, 1) if module_count else 0.0,
                "Avg Cond (uS/cm)": round(float(avg_cond), 0) if pd.notna(avg_cond) else np.nan,
                "Median Age (yrs)": round(float(ages.median()), 1) if ages.notna().any() else np.nan,
            }
        )
    ranking = pd.DataFrame(rows)
    if ranking.empty:
        return ranking
    return ranking.sort_values("Need", ascending=False).reset_index(drop=True)


def make_fleet_trend_chart(status: pd.DataFrame) -> go.Figure:
    months = sorted(status["report_date"].dropna().unique())
    degraded_pct: list[float] = []
    avg_cond: list[float] = []
    for month in months:
        sub = status[status["report_date"] == month]
        total = len(sub)
        degraded_pct.append(float(sub["degraded"].sum()) / total * 100 if total else np.nan)
        avg_cond.append(sub.loc[sub["status"] == "active", "conductivity"].mean())

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=months, y=degraded_pct, name="Degraded %", mode="lines+markers",
            line={"color": "#dc2626", "width": 3},
            hovertemplate="%{x|%b %Y}<br>%{y:.1f}% degraded<extra></extra>",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=months, y=avg_cond, name="Avg conductivity (uS/cm)", mode="lines+markers",
            yaxis="y2", line={"color": "#2563eb", "width": 2, "dash": "dash"},
            hovertemplate="%{x|%b %Y}<br>%{y:,.0f} uS/cm avg<extra></extra>",
        )
    )
    fig.update_layout(
        title="Fleet deterioration over time",
        xaxis_title="Report month",
        yaxis_title="Degraded %",
        yaxis2={"title": "Avg conductivity (uS/cm)", "overlaying": "y", "side": "right", "showgrid": False},
        hovermode="x unified",
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "right", "x": 1},
        margin={"l": 20, "r": 20, "t": 60, "b": 20},
        template="plotly_white",
        height=420,
    )
    return fig


def make_zone_rollup_chart(snapshot: pd.DataFrame) -> go.Figure:
    rollup = (
        snapshot.groupby("zone")
        .agg(
            bypassed=("status", lambda s: int((s == "bypass").sum())),
            degraded=("degraded", "sum"),
        )
        .reset_index()
    )
    rollup["total"] = rollup["bypassed"] + rollup["degraded"]
    rollup = rollup.sort_values("total", ascending=False)

    fig = go.Figure()
    fig.add_trace(go.Bar(x=rollup["zone"], y=rollup["degraded"], name="Degraded", marker_color="#d97706"))
    fig.add_trace(go.Bar(x=rollup["zone"], y=rollup["bypassed"], name="Bypassed", marker_color="#dc2626"))
    fig.update_layout(
        barmode="stack",
        title="Modules needing attention by zone",
        xaxis_title="Zone",
        yaxis_title="Modules",
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "right", "x": 1},
        margin={"l": 20, "r": 20, "t": 60, "b": 20},
        template="plotly_white",
        height=380,
    )
    return fig


def make_age_profile_chart(snapshot: pd.DataFrame, latest: pd.Timestamp) -> go.Figure | None:
    years = pd.to_datetime(snapshot["install_date"], errors="coerce").dt.year.dropna()
    if years.empty:
        return None
    counts = years.astype(int).value_counts().sort_index()
    latest_year = pd.Timestamp(latest).year
    colors = ["#dc2626" if (latest_year - year) > 5 else "#2563eb" for year in counts.index]

    fig = go.Figure(
        go.Bar(
            x=[str(year) for year in counts.index],
            y=counts.values,
            marker_color=colors,
            hovertemplate="%{x}<br>%{y} modules<extra></extra>",
        )
    )
    fig.update_layout(
        title="Membrane age profile — install year of fitted modules (>5 yrs old in red)",
        xaxis_title="Install year",
        yaxis_title="Modules",
        margin={"l": 20, "r": 20, "t": 60, "b": 20},
        template="plotly_white",
        height=380,
        showlegend=False,
    )
    fig.update_xaxes(type="category")
    return fig


def render_portfolio_page(df: pd.DataFrame, mis: pd.DataFrame) -> None:
    st.title("Fleet Portfolio")
    st.caption(
        "Fleet-wide membrane health across every plant — no plant selection needed. "
        "A module needs attention if it is bypassed, or degraded (active and over its "
        "stage's conductivity limit: Stage I > 1,000, II > 1,500, III > 2,000 uS/cm)."
    )

    status = compute_fleet_status(df, mis)
    if status.empty:
        st.info("No readings are available to build the fleet view.")
        return

    latest = status["report_date"].max()
    latest_text = pd.Timestamp(latest).strftime("%b %Y")
    snapshot = status[status["report_date"] == latest]

    total_plants = int(snapshot["plant"].nunique())
    total_modules = len(snapshot)
    active_modules = int((snapshot["status"] == "active").sum())
    degraded = int(snapshot["degraded"].sum())
    bypassed = int((snapshot["status"] == "bypass").sum())
    need = int(snapshot["need"].sum())
    need_pct = need / total_modules * 100 if total_modules else 0.0

    # ----- 1. KPI cards -----
    st.subheader(f"Fleet snapshot — {latest_text}")
    kpis = st.columns(5)
    with kpis[0]:
        metric_card("Plants", f"{total_plants:,}", "in fleet")
    with kpis[1]:
        metric_card("Active Modules", f"{active_modules:,}", f"{total_modules:,} total this month")
    with kpis[2]:
        metric_card("Degraded", f"{degraded:,}", "active, over stage limit", "#d97706")
    with kpis[3]:
        metric_card("Bypassed", f"{bypassed:,}", "offline membranes", "#dc2626")
    with kpis[4]:
        attention_color = "#dc2626" if need_pct >= 10 else "#16a34a"
        metric_card("Needing Attention", f"{need_pct:.1f}%", f"{need:,} of {total_modules:,} modules", attention_color)

    # ----- 2. Plant ranking (centerpiece) -----
    st.markdown("---")
    st.subheader("Plant ranking")
    st.caption("One row per plant, sorted by modules needing attention. Click a column header to re-sort.")
    ranking = build_plant_ranking(snapshot, latest)
    st.dataframe(ranking, width="stretch", hide_index=True)
    st.download_button(
        "Download plant ranking (CSV)",
        ranking.to_csv(index=False).encode("utf-8"),
        file_name=f"fleet_ranking_{latest_text.replace(' ', '_')}.csv",
        mime="text/csv",
    )

    # ----- 3. Fleet trend -----
    st.markdown("---")
    st.subheader("Fleet trend")
    st.caption("Month-over-month deterioration across the whole fleet.")
    st.plotly_chart(make_fleet_trend_chart(status), width="stretch")

    # ----- 4 & 5. Zone rollup + worst modules -----
    left, right = st.columns(2)
    with left:
        st.subheader("By zone")
        st.plotly_chart(make_zone_rollup_chart(snapshot), width="stretch")
    with right:
        st.subheader(f"Worst 25 modules — {latest_text}")
        active_snap = snapshot[snapshot["status"] == "active"].dropna(subset=["conductivity"])
        worst = active_snap.sort_values("conductivity", ascending=False).head(25)
        worst_table = pd.DataFrame(
            {
                "Plant": worst["plant"],
                "Stage": worst["stage_label"],
                "Module": worst["module_label"],
                "Conductivity (uS/cm)": worst["conductivity"].map(lambda v: f"{v:,.0f}"),
                "Install Date": worst["install_date"].map(
                    lambda v: pd.Timestamp(v).strftime("%d %b %Y") if pd.notna(v) else "Unknown"
                ),
            }
        )
        st.dataframe(worst_table, width="stretch", hide_index=True, height=380)

    # ----- 6. Membrane age profile -----
    st.markdown("---")
    st.subheader("Membrane age profile")
    age_chart = make_age_profile_chart(snapshot, latest)
    if age_chart is None:
        st.info("No install dates available to build the age profile.")
    else:
        st.plotly_chart(age_chart, width="stretch")


def render_dashboard(df: pd.DataFrame, params: pd.DataFrame) -> None:
    filtered, selected_group, selected_plant, selected_module, metric_name = sidebar_filters(df)
    render_add_data_controls()
    metric_config = METRICS[metric_name]
    metric_col = str(metric_config["column"])
    series = aggregate_series(filtered, metric_col)

    st.title("RO Plant Membrane Health Dashboard")
    st.caption(f"{selected_group} | {selected_plant}")

    if series.empty:
        st.warning("No valid numeric readings are available for this module and metric.")
        return

    # PI 1601 feed pressure is the operating pressure each reading was taken at.
    # It is plant-level (shared across modules), overlaid on a secondary axis.
    pressure_by_date = feed_pressure_series(params, selected_group, selected_plant)

    ordered = series.sort_values("report_date")
    latest_row = ordered.iloc[-1]
    first_row = ordered.iloc[0]
    latest = float(latest_row["value"])
    first = float(first_row["value"])
    unit = str(metric_config["unit"])
    latest_date = pd.Timestamp(latest_row["report_date"]).strftime("%d %b %Y")

    change = latest - first
    if first:
        change_subtitle = f"{change / first * 100:+,.1f}% since {pd.Timestamp(first_row['report_date']).strftime('%b %Y')}"
    else:
        change_subtitle = f"since {pd.Timestamp(first_row['report_date']).strftime('%b %Y')}"

    col1, col2, col3 = st.columns(3)
    with col1:
        metric_card("Latest Reading", f"{latest:,.2f} {unit}", latest_date)
    with col2:
        metric_card("Change Since First", f"{change:+,.2f} {unit}", change_subtitle)
    with col3:
        metric_card("Readings on Record", str(len(ordered)), "report months")

    st.plotly_chart(
        make_chart(
            series,
            metric_name=metric_name,
            metric_config=metric_config,
            pressure=pressure_by_date,
        ),
        width="stretch",
    )

    with st.expander("Cleaned readings"):
        display = filtered[
            [
                "source_file",
                "report_date",
                "plant",
                "stage",
                "module_label",
                "install_date",
                "flow_lph",
                "conductivity_us_cm",
            ]
        ].rename(
            columns={
                "source_file": "Source File",
                "report_date": "Report Date",
                "plant": "Plant",
                "stage": "Stage",
                "module_label": "Module Number",
                "install_date": "Install Date",
                "flow_lph": "Total flow liter/hr.",
                "conductivity_us_cm": "Cond. us/cm",
            }
        )
        st.dataframe(display, width="stretch", hide_index=True)

    plant_df = df[(df["plant_group"] == selected_group) & (df["plant"] == selected_plant)]
    render_comparison_section(
        plant_df, selected_module, metric_name, metric_config, pressure_by_date
    )


def main() -> None:
    st.set_page_config(
        page_title="RO Membrane Health Dashboard",
        page_icon=":bar_chart:",
        layout="wide",
    )
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.8rem; }
        .metric-card {
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            background: #ffffff;
            padding: 16px 18px;
            min-height: 118px;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.06);
        }
        .metric-title {
            color: #64748b;
            font-size: 0.82rem;
            font-weight: 700;
            letter-spacing: 0;
            text-transform: uppercase;
        }
        .metric-value {
            color: #0f172a;
            font-size: clamp(1.35rem, 2vw, 2rem);
            line-height: 1.15;
            font-weight: 800;
            margin-top: 0.42rem;
            overflow-wrap: anywhere;
        }
        .metric-subtitle {
            color: #64748b;
            font-size: 0.86rem;
            margin-top: 0.45rem;
            min-height: 1.1rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    engine = get_engine()
    if engine is None:
        st.error(
            "No database is configured. Set a `DATABASE_URL` environment variable, "
            "or add a `.streamlit/secrets.toml` with:\n\n"
            "```toml\n[postgres]\nurl = \"postgresql://USER:PASSWORD@HOST:5432/DBNAME\"\n```"
        )
        return

    try:
        init_db(engine)
        summary = ingest_reports(engine, str(APP_DIR))
    except Exception as exc:  # noqa: BLE001
        st.error(f"Could not reach the database: {exc}")
        return

    # New/changed reports were parsed this run — drop the cached query results so
    # the fresh rows show up.
    if summary["ingested"]:
        load_readings.clear()
        load_parameters.clear()
        load_mis.clear()
        compute_fleet_status.clear()

    df = load_readings()
    if df.empty:
        st.error("No RO module readings are in the database yet.")
        if summary["failed"]:
            st.write(summary["failed"])
        return

    if summary["failed"]:
        with st.sidebar.expander("Skipped inputs"):
            for item in summary["failed"]:
                st.write(item)

    params = load_parameters()
    mis = load_mis()

    pages = st.navigation(
        [
            st.Page(
                lambda: render_portfolio_page(df, mis),
                title="Portfolio",
                icon="🌐",
                url_path="portfolio",
                default=True,
            ),
            st.Page(
                lambda: render_dashboard(df, params),
                title="Dashboard",
                icon="📊",
                url_path="dashboard",
            ),
            st.Page(
                lambda: render_replacement_page(df),
                title="Replacement Candidates",
                icon="🛠️",
                url_path="replacement",
            ),
        ]
    )
    pages.run()


if __name__ == "__main__":
    main()
