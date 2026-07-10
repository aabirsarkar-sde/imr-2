from __future__ import annotations

import calendar
import hashlib
import io
import json
import os
import re
import shutil
import tempfile
import time
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
    LargeBinary,
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
# Month name -> number, including abbreviations ("jan", "sept", "dec") so
# filenames like "master sheet IMR Dec.-25" parse, not just full month names.
MONTHS = {}
for _month_index in range(1, 13):
    MONTHS[calendar.month_name[_month_index].lower()] = _month_index
    MONTHS[calendar.month_abbr[_month_index].lower()] = _month_index
MONTHS["sept"] = 9

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
    # False when no Month-Year was found in the filename and report_date had to
    # be guessed from the file's mtime — the validation gate surfaces this as a
    # warning instead of letting a wrong date land silently.
    date_from_filename: bool = True


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
    # The monthly "master sheet IMR <month>-<yy>" workbooks all cover the Dahej
    # zone; fold their filename-derived label ("master sheet") to that zone name.
    "master sheet": "Dahej",
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
    """Pull the plant id from a sheet name. Prefer the last "(NNNN)" group, e.g.
    "Lupin Ank RO1 (1639)" -> 1639 and "Bharat rasayan RO 6 (3523) ." -> 3523
    (trailing junk after the parens). If there are no parentheses, fall back to a
    trailing 3+ digit run ("BEIL PTHP 3550" -> 3550). The 3-digit floor avoids
    grabbing a stray "(2" from a truncated name. This joins a reading to its MIS row."""
    text = str(name)
    parens = re.findall(r"\((\d+)\)", text)
    if parens:
        return int(parens[-1])
    match = re.search(r"(\d{3,})\s*\.?\s*$", text)
    return int(match.group(1)) if match else None


def parse_report_metadata(path: Path) -> ReportMeta:
    stem = path.stem
    # Longest names first so "march" wins over "mar", "sept" over "sep", etc.
    month_regex = "|".join(sorted(MONTHS, key=len, reverse=True))
    # Month (+ optional abbreviation dot), a spaces/dot/dash separator, then a
    # 4- *or* 2-digit year. Handles "April-2025", "Jan-26", "Dec.-25", "June -25".
    match = re.search(
        rf"\b({month_regex})\b\.?\s*[-./]?\s*((?:19|20)\d{{2}}|\d{{2}})",
        stem,
        flags=re.IGNORECASE,
    )

    if match:
        month_number = MONTHS[match.group(1).lower()]
        year = int(match.group(2))
        if year < 100:  # 2-digit year -> 20xx
            year += 2000
        day = calendar.monthrange(year, month_number)[1]
        report_date = pd.Timestamp(year=year, month=month_number, day=day)
        date_from_filename = True
        # The location label is whatever is NOT the date. The old filenames put it
        # after the date ("IMR <date> <loc>"); the master-sheet naming puts it
        # before ("<loc> IMR <date>"). Take both sides so either convention works.
        remainder = stem[: match.start()] + " " + stem[match.end() :]
    else:
        report_date = pd.Timestamp(path.stat().st_mtime, unit="s").normalize()
        date_from_filename = False
        remainder = stem

    # Strip the "IMR" report-type noise and any leading "NN - " index, then tidy.
    remainder = re.sub(r"\bIMR\b", " ", remainder, flags=re.IGNORECASE)
    remainder = re.sub(r"^\s*\d+\s*-\s*", "", remainder)
    plant_group = re.sub(r"\s+", " ", remainder).strip(" -_,") or stem

    return ReportMeta(
        plant_group=canonicalize_plant_group(plant_group),
        report_date=report_date,
        date_from_filename=date_from_filename,
    )


# A hardened per-plant report carries its own identity in a COVER block at the top
# of each sheet, so plant id / date / zone come from the sheet itself instead of
# the filename + MIS join. The label text (normalized) maps to a field; the value
# is the first non-empty cell to its right. Old master workbooks have no COVER, so
# read_cover_block() returns {} for them and every caller falls back to the
# filename-derived metadata + sheet-name id — i.e. unchanged behavior.
COVER_LABELS = {
    "plant_sr_no": "plant sr no",
    "report_date": "report date",
    "zone": "zone",
    "site_name": "site name",
    "plant_name": "plant name",
    "plant_capacity": "plant capacity",
}


def _value_right(raw: pd.DataFrame, row: int, col: int, span: int = 8) -> object:
    """First non-empty cell to the right of (row, col), within `span` columns.
    Skips the blanks that merged label/value cells leave behind."""
    last = min(col + 1 + span, raw.shape[1])
    for c in range(col + 1, last):
        value = raw.iat[row, c]
        if not pd.isna(value) and str(value).strip() != "":
            return value
    return None


def read_cover_block(raw: pd.DataFrame) -> dict[str, object]:
    """Read a sheet's COVER identity block, if present. Returns {} when none of the
    labels are found (old-format sheets), so callers can fall back cleanly."""
    raw_hits: dict[str, object] = {}
    max_row = min(len(raw), 14)
    max_col = min(raw.shape[1], 12)
    for r in range(max_row):
        for c in range(max_col):
            label = normalize_text(raw.iat[r, c])
            if not label:
                continue
            for field, key in COVER_LABELS.items():
                if field not in raw_hits and key in label:
                    value = _value_right(raw, r, c)
                    if value is not None:
                        raw_hits[field] = value

    cover: dict[str, object] = {}
    if "plant_sr_no" in raw_hits:
        sr = pd.to_numeric(raw_hits["plant_sr_no"], errors="coerce")
        if pd.notna(sr):
            cover["plant_sr_no"] = int(sr)
    if "report_date" in raw_hits:
        value = raw_hits["report_date"]
        try:
            # Real Excel dates arrive as datetime — unambiguous, no warning.
            date = pd.Timestamp(value)
        except (ValueError, TypeError):
            # Free-typed text: assume dd-mm (the Indian convention on these forms).
            date = pd.to_datetime(str(value), errors="coerce", dayfirst=True)
        if pd.notna(date):
            cover["report_date"] = date
    for field in ("zone", "site_name", "plant_name", "plant_capacity"):
        if field in raw_hits:
            text = str(raw_hits[field]).strip()
            if text:
                cover[field] = text
    return cover


def sheet_context(
    raw: pd.DataFrame, sheet_name: str, metadata: ReportMeta
) -> tuple[str, int | None, ReportMeta, dict[str, object]]:
    """Resolve a sheet's plant id, sr-no and effective metadata, preferring the
    COVER block and falling back to the sheet name + filename metadata."""
    cover = read_cover_block(raw)
    # A real COVER block is marked by an identity/date cell (sr-no, zone or date).
    # Only then do we take the plant's display name from the sheet's name cell
    # (site_name); old-format sheets fall back to the tab name, even if they
    # happen to carry a stray "SITE NAME" label.
    has_cover = bool(
        cover.get("plant_sr_no") or cover.get("zone") or cover.get("report_date")
    )
    identity = cover.get("plant_name") or cover.get("site_name")
    plant = str(identity) if (has_cover and identity) else clean_sheet_name(sheet_name)
    sr_value = cover.get("plant_sr_no")
    plant_sr_no = int(sr_value) if sr_value is not None else plant_sr_no_from_name(sheet_name)
    report_date = cover.get("report_date") or metadata.report_date
    zone = cover.get("zone")
    plant_group = canonicalize_plant_group(str(zone)) if zone else metadata.plant_group
    effective = ReportMeta(
        plant_group=plant_group,
        report_date=report_date,
        date_from_filename=metadata.date_from_filename if not cover.get("report_date") else True,
    )
    return plant, plant_sr_no, effective, cover


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


def is_time_header(value: object) -> bool:
    """The 'Time for ___ ml' column — the raw stopwatch reading (seconds) from
    which flow is computed as (volume_ml / 1000) / (time / 3600) = volume*3.6/time."""
    text = normalize_text(value)
    return "time for" in text or ("time" in text and ("ml" in text or "ltr" in text))


def parse_volume_ml(value: object) -> float | None:
    """Millilitres named in a 'Time for X ml' header. Accepts '500 ml', '1000ml',
    '1 ltr', '2 ltr', '1Ltr' (litres → ×1000). Returns None when no volume is
    printed (the blank template header 'Time for ______ ml')."""
    text = normalize_text(value)
    match = re.search(r"(\d+(?:\.\d+)?)\s*(ml|litre|liter|ltr|l)\b", text)
    if not match:
        return None
    amount = float(match.group(1))
    return amount if match.group(2) == "ml" else amount * 1000.0


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
            elif is_time_header(header):
                block["time_col"] = col
                volume = parse_volume_ml(header)
                if volume is not None:
                    block["volume_ml"] = volume

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
    plant_override: str | None = None,
    plant_sr_no_override: int | None = None,
) -> list[dict[str, object]]:
    header_row = find_header_row(raw)
    if header_row is None:
        return []

    blocks = find_module_blocks(raw, header_row)
    if not blocks:
        return []

    records: list[dict[str, object]] = []
    plant = plant_override or clean_sheet_name(sheet_name)
    plant_sr_no = (
        plant_sr_no_override
        if plant_sr_no_override is not None
        else plant_sr_no_from_name(sheet_name)
    )

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
                # A live '=<volume*3.6>/<time>' formula written by openpyxl carries
                # no cached result, so pandas reads the flow cell as NaN. Recompute
                # it from the timed reading exactly as the sheet's formula would.
                if pd.isna(flow) and "time_col" in block and block.get("volume_ml"):
                    seconds = pd.to_numeric(value_at(raw, row_index, block["time_col"]), errors="coerce")
                    if pd.notna(seconds) and seconds != 0:
                        flow = float(block["volume_ml"]) * 3.6 / float(seconds)
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
    time_col = find_column(columns, is_time_header)
    time_volume_ml = parse_volume_ml(time_col) if time_col else None
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
            if pd.isna(flow) and time_col is not None and time_volume_ml:
                seconds = pd.to_numeric(row.get(time_col), errors="coerce")
                if pd.notna(seconds) and seconds != 0:
                    flow = time_volume_ml * 3.6 / float(seconds)
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
        plant, plant_sr_no, eff_meta, _cover = sheet_context(raw, sheet_name, metadata)
        sheet_records = extract_rows_from_sheet(
            raw,
            report_path=path,
            sheet_name=sheet_name,
            metadata=eff_meta,
            plant_override=plant,
            plant_sr_no_override=plant_sr_no,
        )
        if not sheet_records:
            structured = pd.read_excel(path, sheet_name=sheet_name)
            sheet_records = extract_rows_from_structured_table(
                structured,
                report_path=path,
                sheet_name=sheet_name,
                metadata=eff_meta,
            )
        records.extend(sheet_records)

    return records


def extract_operating_parameters(
    raw: pd.DataFrame,
    *,
    report_path: Path,
    sheet_name: str,
    metadata: ReportMeta,
    plant_override: str | None = None,
) -> list[dict[str, object]]:
    """Scan a raw sheet for (tag, value, unit) instrument readings.

    Operating parameters appear in two layouts: a key-value block below the
    module table, or extra columns to the right of it. Both place the unit
    immediately right of the value, with the instrument tag one cell further
    left, so a single scan keyed on the unit cell handles both. The last
    occurrence of a tag wins.
    """
    plant = plant_override or clean_sheet_name(sheet_name)
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
            # Some sheets leave a blank spacer column between the tag and its
            # value, e.g. the feed/permeate conductivity block reads
            # "CIS 151" | <blank> | 8200 | "us/cm". When the immediate left cell
            # is empty, look one cell further left for the tag (requiring a
            # letter so a stray number doesn't masquerade as an instrument tag).
            if not tag and col_index >= 3:
                candidate = normalize_text(raw.iat[row_index, col_index - 3])
                if re.search(r"[a-z]", candidate):
                    tag = candidate
            if not tag or pd.isna(value):
                continue

            found[tag] = {
                "source_file": report_path.name,
                "plant_group": metadata.plant_group,
                "plant": plant,
                "report_date": metadata.report_date,
                "tag": tag.upper(),
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
        plant, _sr, eff_meta, _cover = sheet_context(raw, sheet_name, metadata)
        records.extend(
            extract_operating_parameters(
                raw,
                plant_override=plant,
                report_path=path,
                sheet_name=sheet_name,
                metadata=eff_meta,
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
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
        if clean_sheet_name(sheet_name).lower() == "mis":
            records.extend(extract_mis_rows(raw, report_path=path, metadata=metadata))
            continue
        # Per-plant template: no MIS sheet, but each sheet's COVER block carries
        # its zone/site/sr-no. Synthesize an MIS row so zone resolution (which
        # joins readings.plant_sr_no -> mis.zone) works without a register sheet.
        cover = read_cover_block(raw)
        sr_value = cover.get("plant_sr_no")
        zone = cover.get("zone")
        if sr_value is not None and zone:
            records.append(
                {
                    "source_file": path.name,
                    "report_date": cover.get("report_date") or metadata.report_date,
                    "zone": str(zone),
                    "zm_name": None,
                    "plant_sr_no": int(sr_value),
                    "site_name": cover.get("site_name"),
                    "status": None,
                    "membrane_required": None,
                    "remarks": None,
                }
            )
    return records


def safe_upload_name(filename: str) -> str:
    name = Path(filename).name
    name = re.sub(r"[^A-Za-z0-9., _()\-]+", "_", name).strip()
    return name or "uploaded_report.xlsx"


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

# Durable copy of each uploaded workbook's raw bytes, so a report can be
# re-parsed later (after a parser improvement) without depending on the
# ephemeral Cloud disk. New table -> create_all makes it; no migration needed.
REPORT_FILES_TABLE = Table(
    "report_files",
    DB_METADATA,
    Column("filename", Text, primary_key=True),
    Column("content_hash", Text, nullable=False),
    Column("file_bytes", LargeBinary, nullable=False),
    Column("uploaded_at", DateTime),
    Column("status", Text),
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


@dataclass
class ParseResult:
    """The normalized output of parsing one report, before it is committed.
    Carries the file's metadata plus a quality signal the gate surfaces."""
    meta: ReportMeta
    readings: pd.DataFrame
    parameters: pd.DataFrame
    mis: pd.DataFrame
    n_cond_blanked: int  # raw conductivity values dropped as out-of-range


def parse_report_file(path: Path) -> ParseResult:
    """Parse a report on disk into normalized DataFrames WITHOUT committing.
    The same parse the old ingest did inline, now reusable by the upload gate."""
    meta = parse_report_metadata(path)
    raw_readings = read_report(path)
    n_blanked = 0
    for record in raw_readings:
        value = pd.to_numeric(record.get("conductivity_us_cm"), errors="coerce")
        if pd.notna(value) and value > MAX_PLAUSIBLE_CONDUCTIVITY_US_CM:
            n_blanked += 1
    return ParseResult(
        meta=meta,
        readings=normalize_reading_records(raw_readings),
        parameters=normalize_parameter_records(read_report_parameters(path)),
        mis=normalize_mis_records(read_report_mis(path)),
        n_cond_blanked=n_blanked,
    )


def parse_report_bytes(filename: str, data: bytes) -> ParseResult:
    """Parse uploaded bytes without a durable disk write. The bytes are written
    to a transient temp dir under their ORIGINAL (sanitized) filename — group and
    report_date are derived from the name — then parsed via parse_report_file and
    the temp dir removed. Use safe_upload_name(filename) as the canonical key for
    staging / commit / byte storage so it matches the source_file stamped in rows."""
    tmpdir = Path(tempfile.mkdtemp(prefix="imr_stage_"))
    try:
        target = tmpdir / safe_upload_name(filename)
        target.write_bytes(data)
        return parse_report_file(target)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def commit_parse_result(
    engine: Engine, filename: str, content_hash: str, result: ParseResult
) -> dict[str, int]:
    """Write a parsed result to the live tables (DELETE-then-insert by
    source_file, idempotent) and record it in ingested_files. Returns counts."""
    readings, parameters, mis = result.readings, result.parameters, result.mis
    with engine.begin() as conn:
        conn.execute(text("DELETE FROM readings WHERE source_file = :f"), {"f": filename})
        conn.execute(text("DELETE FROM parameters WHERE source_file = :f"), {"f": filename})
        conn.execute(text("DELETE FROM mis WHERE source_file = :f"), {"f": filename})
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
                "f": filename,
                "h": content_hash,
                "ts": datetime.now(timezone.utc),
                "nr": int(len(readings)),
                "np": int(len(parameters)),
                "nm": int(len(mis)),
            },
        )
    return {
        "readings": int(len(readings)),
        "parameters": int(len(parameters)),
        "mis": int(len(mis)),
    }


def store_report_bytes(
    engine: Engine, filename: str, content_hash: str, data: bytes, status: str = "committed"
) -> None:
    """Persist an uploaded workbook's raw bytes so it can be re-parsed later
    without the original file on disk. Bytes are bound as a parameter (not via
    to_sql) — psycopg2 adapts to bytea, sqlite3 to BLOB."""
    with engine.begin() as conn:
        conn.execute(
            text(
                """
                INSERT INTO report_files (filename, content_hash, file_bytes, uploaded_at, status)
                VALUES (:f, :h, :b, :ts, :s)
                ON CONFLICT (filename) DO UPDATE SET
                    content_hash = excluded.content_hash,
                    file_bytes = excluded.file_bytes,
                    uploaded_at = excluded.uploaded_at,
                    status = excluded.status
                """
            ),
            {
                "f": filename,
                "h": content_hash,
                "b": data,
                "ts": datetime.now(timezone.utc),
                "s": status,
            },
        )


def load_report_bytes(engine: Engine, filename: str) -> bytes | None:
    """Return the stored raw bytes for a report, or None if not stored."""
    with engine.connect() as conn:
        row = conn.execute(
            text("SELECT file_bytes FROM report_files WHERE filename = :f"), {"f": filename}
        ).first()
    if row is None or row[0] is None:
        return None
    return bytes(row[0])  # psycopg2 returns a memoryview; normalize to bytes


def report_files_index(engine: Engine) -> list[str]:
    """Filenames that have durable bytes stored (eligible for re-parse)."""
    with engine.connect() as conn:
        return [
            r[0]
            for r in conn.execute(
                text("SELECT filename FROM report_files ORDER BY filename")
            ).all()
        ]


def ingest_reports(engine: Engine, data_dir: str) -> dict[str, list[str]]:
    """Parse and store any report on disk whose content hash is new or changed.

    Returns a summary of what was ingested / skipped (already up to date) /
    failed. Already-ingested files are never re-parsed. This is the startup path
    for repo-committed files; uploads go through the validation gate instead.
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
            result = parse_report_file(path)
        except Exception as exc:  # noqa: BLE001 - one bad report shouldn't sink the rest
            summary["failed"].append(f"{path.name}: {exc}")
            continue

        if result.readings.empty:
            summary["failed"].append(f"{path.name}: no module table found")
            continue

        commit_parse_result(engine, path.name, fingerprint, result)
        summary["ingested"].append(path.name)

    return summary


def ingest_report_files(engine: Engine) -> dict[str, list[str]]:
    """Startup safety net: re-commit any durably-stored report whose content_hash
    is not already recorded in ingested_files — e.g. the derived tables were wiped
    but report_files survived a Cloud restart. Normally a no-op (hashes match)."""
    summary: dict[str, list[str]] = {"ingested": [], "skipped": [], "failed": []}
    with engine.connect() as conn:
        known = dict(
            conn.execute(text("SELECT filename, content_hash FROM ingested_files")).all()
        )
        stored = conn.execute(
            text("SELECT filename, content_hash FROM report_files")
        ).all()

    for filename, content_hash in stored:
        if known.get(filename) == content_hash:
            summary["skipped"].append(filename)
            continue
        data = load_report_bytes(engine, filename)
        if data is None:
            summary["failed"].append(f"{filename}: no stored bytes")
            continue
        try:
            result = parse_report_bytes(filename, data)
        except Exception as exc:  # noqa: BLE001
            summary["failed"].append(f"{filename}: {exc}")
            continue
        if result.readings.empty:
            summary["failed"].append(f"{filename}: no module table found")
            continue
        commit_parse_result(engine, filename, content_hash, result)
        summary["ingested"].append(filename)

    return summary


@dataclass
class Issue:
    """One line in a report's data-quality report. level is ERROR (blocks the
    commit), WARN (commit allowed, but something is off), or INFO (FYI)."""
    level: str
    code: str
    message: str


def looks_like_raw_filename(plant_group: str) -> bool:
    """Heuristic: a plant_group that still looks like an undigested filename
    (underscores, long digit runs, or many tokens) — a sign the date/label split
    didn't cleanly isolate the location."""
    if "_" in plant_group:
        return True
    if re.search(r"\d{4,}", plant_group):
        return True
    return len(plant_group.split()) > 4


def validate_parse(engine: Engine, filename: str, result: ParseResult) -> list[Issue]:
    """Produce a data-quality report for a parsed-but-not-yet-committed report.
    The upload gate shows these and blocks Confirm if any ERROR is present."""
    issues: list[Issue] = []
    readings = result.readings

    if readings.empty:
        issues.append(Issue("ERROR", "no_readings",
            "No module table found — there is nothing to commit."))
        return issues

    if not result.meta.date_from_filename:
        issues.append(Issue("WARN", "date_fallback",
            f"No month/year in the filename; report_date was guessed as "
            f"{result.meta.report_date:%d %b %Y}. Rename like '… April-2025 …' to fix."))

    if looks_like_raw_filename(result.meta.plant_group):
        issues.append(Issue("WARN", "raw_group",
            f"Plant group '{result.meta.plant_group}' looks like a raw filename, "
            f"not a clean location label."))

    missing = sorted(
        str(p) for p in
        readings.loc[readings["plant_sr_no"].isna(), "plant"].dropna().unique()
    )
    if missing:
        shown = ", ".join(missing[:8]) + (" …" if len(missing) > 8 else "")
        issues.append(Issue("WARN", "no_sr_no",
            f"{len(missing)} sheet(s) have no plant id (NNNN) — their readings get "
            f"zone 'Unknown': {shown}"))

    # Zone match: of the distinct plant_sr_no in this file, how many resolve to a
    # zone via this file's MIS or the existing DB register.
    sr_nos = {int(x) for x in readings["plant_sr_no"].dropna().unique()}
    if sr_nos:
        known_sr = {int(x) for x in result.mis["plant_sr_no"].dropna().unique()}
        with engine.connect() as conn:
            known_sr |= {
                int(r[0]) for r in
                conn.execute(text("SELECT DISTINCT plant_sr_no FROM mis")).all()
                if r[0] is not None
            }
        matched = len(sr_nos & known_sr)
        if matched / len(sr_nos) < 0.5:
            issues.append(Issue("WARN", "low_zone_match",
                f"Only {matched}/{len(sr_nos)} plants match a zone in the MIS register; "
                f"most readings will show zone 'Unknown'."))

    if result.n_cond_blanked:
        issues.append(Issue("INFO", "cond_blanked",
            f"{result.n_cond_blanked} conductivity value(s) above "
            f"{int(MAX_PLAUSIBLE_CONDUCTIVITY_US_CM):,} µS/cm dropped as data-entry errors."))

    with engine.connect() as conn:
        already = conn.execute(
            text("SELECT 1 FROM ingested_files WHERE filename = :f"), {"f": filename}
        ).first()
    if already:
        issues.append(Issue("INFO", "will_replace",
            f"'{filename}' is already in the database — committing replaces it."))

    return issues


def ingested_report_summary(engine: Engine) -> pd.DataFrame:
    """List every ingested report with its stored row counts, for the
    manage-data UI. One row per file in `ingested_files`."""
    with engine.connect() as conn:
        rows = conn.execute(
            text(
                "SELECT filename, n_readings, n_parameters, n_mis, ingested_at "
                "FROM ingested_files ORDER BY filename"
            )
        ).all()
    return pd.DataFrame(
        rows,
        columns=["filename", "n_readings", "n_parameters", "n_mis", "ingested_at"],
    )


def remove_report(engine: Engine, root: Path, filename: str) -> dict[str, int]:
    """Completely remove one report from the system.

    Deletes its rows from `readings`/`parameters`/`mis`, drops its `ingested_files`
    record, and deletes the file on disk. The on-disk delete is essential: if the
    file survived, the next `ingest_reports()` run would re-parse it (its hash is
    no longer known) and restore everything we just deleted. Returns per-table
    delete counts.
    """
    counts: dict[str, int] = {}
    with engine.begin() as conn:
        for table in ("readings", "parameters", "mis"):
            result = conn.execute(
                text(f"DELETE FROM {table} WHERE source_file = :f"), {"f": filename}
            )
            counts[table] = result.rowcount or 0
        conn.execute(
            text("DELETE FROM ingested_files WHERE filename = :f"), {"f": filename}
        )
        # Drop the durable byte copy too, else a re-parse could resurrect it.
        conn.execute(
            text("DELETE FROM report_files WHERE filename = :f"), {"f": filename}
        )

    # Delete the source file wherever it lives so it is not re-ingested on refresh.
    for source_root in (root, root / UPLOAD_DIR_NAME):
        candidate = source_root / filename
        if candidate.exists():
            candidate.unlink()

    return counts


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


def hero_card(title: str, value: str, subtitle: str = "", color: str = "#dc2626") -> None:
    """Oversized single-number card for the page's headline metric."""
    st.markdown(
        f"""
        <div class="hero-card">
            <div class="hero-title">{title}</div>
            <div class="hero-value" style="color:{color};">{value}</div>
            <div class="hero-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def jump_nav(items: list[tuple[str, str, str]]) -> None:
    """Render a row of clickable cards that scroll to in-page section anchors.

    Each item is (label, anchor, subtitle); `anchor` must match the `anchor=`
    set on the target section's subheader. Pure anchor links — no rerun.
    """
    cards = "".join(
        f'<a class="nav-card" href="#{anchor}">'
        f'<span class="nav-label">{label}</span>'
        f'<span class="nav-sub">{subtitle}</span></a>'
        for label, anchor, subtitle in items
    )
    st.markdown(f'<div class="nav-grid">{cards}</div>', unsafe_allow_html=True)


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


def stage_uploaded_files(engine: Engine, files: list[object]) -> None:
    """Parse + validate each uploaded file into st.session_state['staged']
    WITHOUT committing. Keyed by sanitized filename so a re-upload overwrites."""
    staged = st.session_state.setdefault("staged", {})
    for uploaded in files:
        data = uploaded.getvalue()
        name = safe_upload_name(getattr(uploaded, "name", "uploaded_report.xlsx"))
        try:
            result = parse_report_bytes(name, data)
            issues = validate_parse(engine, name, result)
        except Exception as exc:  # noqa: BLE001 - surface parse failures in the gate
            staged[name] = {"error": str(exc)}
            continue
        staged[name] = {
            "result": result,
            "issues": issues,
            "content_hash": hashlib.sha256(data).hexdigest(),
            "data": data,
        }


# Map an Issue level to the Streamlit callout that renders it.
ISSUE_RENDERERS = {"ERROR": st.error, "WARN": st.warning, "INFO": st.info}


def render_staged_reports(engine: Engine) -> None:
    """Render the pending-review queue: per-file stats + quality issues + a
    Confirm/Discard pair. Confirm commits to the live tables; nothing is written
    until then."""
    staged = st.session_state.get("staged", {})
    if not staged:
        return
    st.markdown("**Pending review**")
    for name in list(staged):
        entry = staged[name]
        with st.container(border=True):
            st.markdown(f"**{name}**")
            if "error" in entry:
                st.error(f"Could not parse: {entry['error']}")
                if st.button("Discard", key=f"discard_{name}"):
                    staged.pop(name, None)
                    rerun_app()
                continue

            result: ParseResult = entry["result"]
            issues: list[Issue] = entry["issues"]
            readings = result.readings
            bypass = int((readings["status"] == "bypass").sum()) if "status" in readings else 0
            dates = readings["report_date"].dropna()
            span = ""
            if not dates.empty:
                lo, hi = dates.min(), dates.max()
                span = f" · {lo:%b %Y}" + (f"–{hi:%b %Y}" if hi != lo else "")
            st.caption(
                f"{result.meta.plant_group} · {result.meta.report_date:%d %b %Y} · "
                f"{readings['plant'].nunique()} plants · {len(readings)} readings · "
                f"{bypass} bypass{span}"
            )
            for issue in issues:
                ISSUE_RENDERERS.get(issue.level, st.info)(issue.message)

            blocked = any(i.level == "ERROR" for i in issues)
            ok_col, no_col = st.columns(2)
            if ok_col.button(
                "Confirm & commit", key=f"commit_{name}", type="primary", disabled=blocked
            ):
                commit_parse_result(engine, name, entry["content_hash"], result)
                store_report_bytes(engine, name, entry["content_hash"], entry["data"])
                load_readings.clear()
                load_parameters.clear()
                load_mis.clear()
                compute_fleet_status.clear()
                staged.pop(name, None)
                st.success(f"Committed {name}.")
                rerun_app()
            if no_col.button("Discard", key=f"discard_{name}"):
                staged.pop(name, None)
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

# Month-over-month escape hatch (Portfolio degraded flag). The IQR peer test is
# blind to a stage that degrades in lockstep — when every module rises together,
# none stands out as an outlier. So a module the IQR clears is ALSO flagged if
# its conductivity jumps unusually versus its OWN prior-month reading: at least
# +MOM_JUMP_RATIO relative AND +MOM_MIN_DELTA absolute (the floor stops trivial
# rises on near-zero permeate readings, e.g. 60 -> 100 uS/cm, from tripping it).
MOM_JUMP_RATIO = 1.5
MOM_MIN_DELTA = 150.0

# Whole-plant conductivity salt passage. The plant-level instrument block records
# feed conductivity as CIS 151 and combined permeate conductivity as CIS 180
# (us/cm), once per plant sheet. Salt passage % = permeate / feed * 100 — the
# share of feed salinity that leaks through the entire RO train. A healthy train
# passes only a few percent, so a plant above this fraction is surfaced for review.
FEED_COND_TAG = "151"
PERMEATE_COND_TAG = "180"
SALT_PASSAGE_FLAG_PCT = 10.0

# Whole-plant permeate (product) flow, logged once per plant sheet as FIS 180 in
# either m3/hr or litres/hr. A sustained month-over-month fall in permeate flow is
# a classic fouling/membrane-loss signal, so the Portfolio surfaces the biggest
# drops. PERMEATE_FLOW_M3HR_PER_LPH converts litres/hr -> m3/hr so plants logged
# in either unit rank on one scale.
PERMEATE_FLOW_TAG = "180"
PERMEATE_FLOW_M3HR_PER_LPH = 0.001


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

    A module is degraded if EITHER signal fires:
    - **Peer outlier (IQR):** its conductivity exceeds the stage's Tukey upper
      fence (Q3 + 1.5·IQR), computed against peers in the same stage that month
      via evaluate_stage_readings (k=1.5). Robust down to ~4 modules per stage.
    - **Month-over-month jump:** its conductivity rose unusually versus the SAME
      module's prior reading (>= MOM_JUMP_RATIO and >= MOM_MIN_DELTA). This is a
      per-module time-series test, so it catches a stage degrading in lockstep —
      the case the within-month IQR peer test is structurally blind to.

    A module "needs attention" if it is bypassed or degraded. Returns one row per
    (plant, stage, module, month), with degraded_iqr / degraded_mom kept split so
    callers can show *why* a module was flagged.
    """
    col = str(METRICS["Conductivity"]["column"])
    flow_col = str(METRICS["Flow Rate"]["column"])
    columns = [
        "plant_group", "plant", "plant_sr_no", "zone", "stage_label",
        "module_label", "report_date", "conductivity", "prev_conductivity",
        "flow", "prev_flow", "install_date", "status", "degraded_iqr",
        "degraded_mom", "degraded", "need", "cutoff",
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
            flow=(flow_col, "mean"),
            install_date=("install_date", "first"),
            is_bypass=("status", lambda s: bool((s == "bypass").any())),
        )
        active = agg[~agg["is_bypass"]]
        cutoff = float("nan")
        iqr_map: dict[object, bool] = {}
        if not active.empty:
            evaluated, _, cutoff, _ = evaluate_stage_readings(
                active, method=PEER_METHOD, sensitivity=1.5, limit=0.0
            )
            iqr_map = dict(zip(evaluated["module_label"], evaluated["flag"]))

        plant_group, plant_sr, zone = plant_meta[plant]
        for _, row in agg.iterrows():
            bypass = bool(row["is_bypass"])
            degraded_iqr = (not bypass) and bool(iqr_map.get(row["module_label"], False))
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
                    "flow": row["flow"],
                    "install_date": row["install_date"],
                    "status": "bypass" if bypass else "active",
                    "degraded_iqr": degraded_iqr,
                    "cutoff": cutoff,
                }
            )

    result = pd.DataFrame(records)
    if result.empty:
        return pd.DataFrame(columns=columns)

    # Second signal: an unusually high month-over-month jump vs the SAME module's
    # prior reading. Computed per (plant, stage, module) time series — blind to
    # within-stage peers, so it catches a stage that degrades in lockstep, the
    # case IQR misses. shift() compares each reading to that module's previous
    # available report (no flag on a module's first-ever month).
    result = result.sort_values(["plant", "stage_label", "module_label", "report_date"])
    grouped = result.groupby(["plant", "stage_label", "module_label"])
    prev = grouped["conductivity"].shift()
    result["prev_conductivity"] = prev
    result["prev_flow"] = grouped["flow"].shift()
    result["degraded_mom"] = (
        (result["status"] == "active")
        & prev.notna()
        & (result["conductivity"] >= prev * MOM_JUMP_RATIO)
        & (result["conductivity"] - prev >= MOM_MIN_DELTA)
    )
    result["degraded"] = result["degraded_iqr"] | result["degraded_mom"]
    result["need"] = (result["status"] == "bypass") | result["degraded"]

    return result.reindex(columns=columns).reset_index(drop=True)


def build_zone_by_sr(mis: pd.DataFrame) -> dict[object, object]:
    """plant_sr_no -> zone, from the most recent MIS row for each plant."""
    zone_by_sr: dict[object, object] = {}
    if mis is not None and not mis.empty:
        latest_mis = mis.sort_values("report_date").drop_duplicates("plant_sr_no", keep="last")
        for _, row in latest_mis.iterrows():
            zone = row.get("zone")
            zone_by_sr[row["plant_sr_no"]] = zone if (pd.notna(zone) and str(zone).strip()) else None
    return zone_by_sr


def compute_salt_passage(parameters: pd.DataFrame, mis: pd.DataFrame) -> pd.DataFrame:
    """Per-plant per-month conductivity salt passage from the CIS feed/permeate tags.

    Feed conductivity (CIS 151) and combined permeate conductivity (CIS 180) are
    recorded once per plant sheet. Salt passage % = permeate / feed * 100 — the
    share of feed salinity that leaks through the whole RO train. Returns one row
    per (plant, month) that carries BOTH readings, with the plant's zone joined in
    (via plant_sr_no -> latest MIS row, mirroring compute_fleet_status), sorted by
    passage descending.
    """
    columns = [
        "plant_group", "plant", "plant_sr_no", "zone",
        "report_date", "feed", "permeate", "passage_pct",
    ]
    if parameters is None or parameters.empty:
        return pd.DataFrame(columns=columns)

    cond = parameters[parameters["kind"] == "conductivity"].copy()
    if cond.empty:
        return pd.DataFrame(columns=columns)

    tag = cond["tag"].astype(str)
    cond["role"] = pd.NA
    cond.loc[tag.str.contains(rf"\b{FEED_COND_TAG}\b"), "role"] = "feed"
    cond.loc[tag.str.contains(rf"\b{PERMEATE_COND_TAG}\b"), "role"] = "permeate"
    cond = cond.dropna(subset=["role"])
    if cond.empty:
        return pd.DataFrame(columns=columns)

    # One value per (plant, month, role); last write wins on any duplicate.
    cond = cond.drop_duplicates(["plant", "report_date", "role"], keep="last")
    wide = cond.pivot_table(
        index=["plant_group", "plant", "report_date"],
        columns="role", values="value", aggfunc="last",
    ).reset_index()
    if "feed" not in wide.columns or "permeate" not in wide.columns:
        return pd.DataFrame(columns=columns)
    wide = wide.dropna(subset=["feed", "permeate"])
    wide = wide[wide["feed"] > 0]
    if wide.empty:
        return pd.DataFrame(columns=columns)
    wide["passage_pct"] = wide["permeate"] / wide["feed"] * 100.0

    zone_by_sr = build_zone_by_sr(mis)
    wide["plant_sr_no"] = wide["plant"].map(plant_sr_no_from_name)
    wide["zone"] = wide["plant_sr_no"].map(lambda sr: zone_by_sr.get(sr) or "Unknown")

    return (
        wide.reindex(columns=columns)
        .sort_values("passage_pct", ascending=False)
        .reset_index(drop=True)
    )


def compute_permeate_flow(parameters: pd.DataFrame, mis: pd.DataFrame) -> pd.DataFrame:
    """Per-plant per-month permeate (product) flow from the FIS 180 tag, in m3/hr.

    Permeate flow is logged once per plant sheet as FIS 180, in either m3/hr or
    litres/hr; litres/hr is converted to m3/hr so every plant ranks on one scale.
    Returns one row per (plant, month) with the month-over-month change vs the
    plant's OWN prior reading (`fall_m3hr` positive = a drop), zone joined in.
    """
    columns = [
        "plant_group", "plant", "plant_sr_no", "zone", "report_date",
        "flow_m3hr", "prev_flow_m3hr", "fall_m3hr", "change_pct",
    ]
    if parameters is None or parameters.empty:
        return pd.DataFrame(columns=columns)

    flow = parameters[parameters["kind"] == "flow"].copy()
    if flow.empty:
        return pd.DataFrame(columns=columns)
    flow = flow[flow["tag"].astype(str).str.contains(rf"\b{PERMEATE_FLOW_TAG}\b")]
    if flow.empty:
        return pd.DataFrame(columns=columns)

    # Normalize litres/hr -> m3/hr (m3/hr passes through unchanged).
    unit = flow["unit"].astype(str)
    factor = pd.Series(1.0, index=flow.index)
    factor[unit.str.contains("lit")] = PERMEATE_FLOW_M3HR_PER_LPH
    flow["flow_m3hr"] = pd.to_numeric(flow["value"], errors="coerce") * factor
    flow = flow.dropna(subset=["flow_m3hr"])
    flow = flow[flow["flow_m3hr"] > 0]
    if flow.empty:
        return pd.DataFrame(columns=columns)

    # One value per (plant, month); last write wins, then a per-plant time series.
    flow = flow.drop_duplicates(["plant", "report_date"], keep="last")
    flow = flow.sort_values(["plant", "report_date"])
    prev = flow.groupby("plant")["flow_m3hr"].shift()
    flow["prev_flow_m3hr"] = prev
    flow["fall_m3hr"] = prev - flow["flow_m3hr"]
    flow["change_pct"] = (flow["flow_m3hr"] / prev - 1.0) * 100.0

    zone_by_sr = build_zone_by_sr(mis)
    flow["plant_sr_no"] = flow["plant"].map(plant_sr_no_from_name)
    flow["zone"] = flow["plant_sr_no"].map(lambda sr: zone_by_sr.get(sr) or "Unknown")

    return flow.reindex(columns=columns).reset_index(drop=True)


def compute_permeate_conductivity(parameters: pd.DataFrame, mis: pd.DataFrame) -> pd.DataFrame:
    """Per-plant per-month permeate conductivity (CIS 180) with its MoM rise.

    Permeate (product) conductivity is logged once per plant sheet as CIS 180.
    Returns one row per (plant, month) with the month-over-month change vs the
    plant's OWN prior reading (`rise` positive = product water got saltier — a
    plant-wide membrane / feed signal), zone joined in.
    """
    columns = [
        "plant_group", "plant", "plant_sr_no", "zone", "report_date",
        "permeate", "prev_permeate", "rise", "change_pct",
    ]
    if parameters is None or parameters.empty:
        return pd.DataFrame(columns=columns)

    cond = parameters[parameters["kind"] == "conductivity"].copy()
    if cond.empty:
        return pd.DataFrame(columns=columns)
    cond = cond[cond["tag"].astype(str).str.contains(rf"\b{PERMEATE_COND_TAG}\b")]
    if cond.empty:
        return pd.DataFrame(columns=columns)

    cond["permeate"] = pd.to_numeric(cond["value"], errors="coerce")
    cond = cond.dropna(subset=["permeate"])
    cond = cond[cond["permeate"] > 0]
    if cond.empty:
        return pd.DataFrame(columns=columns)

    # One value per (plant, month); last write wins, then a per-plant time series.
    cond = cond.drop_duplicates(["plant", "report_date"], keep="last")
    cond = cond.sort_values(["plant", "report_date"])
    prev = cond.groupby("plant")["permeate"].shift()
    cond["prev_permeate"] = prev
    cond["rise"] = cond["permeate"] - prev
    cond["change_pct"] = (cond["permeate"] / prev - 1.0) * 100.0

    zone_by_sr = build_zone_by_sr(mis)
    cond["plant_sr_no"] = cond["plant"].map(plant_sr_no_from_name)
    cond["zone"] = cond["plant_sr_no"].map(lambda sr: zone_by_sr.get(sr) or "Unknown")

    return cond.reindex(columns=columns).reset_index(drop=True)


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


def render_plant_flagged_modules(
    snapshot: pd.DataFrame, plant: str, month_text: str
) -> None:
    """List every module flagged (degraded or bypassed) for one plant this month.

    `snapshot` is the already-zone/month-filtered fleet status, so degraded/need
    and the per-stage IQR cutoff are read straight off it — no recomputation.
    """
    plant_rows = snapshot[snapshot["plant"] == plant]
    flagged = plant_rows[plant_rows["need"]].copy()

    st.markdown(f"#### {plant} — flagged modules · {month_text}")
    if flagged.empty:
        st.success(f"No modules flagged for {plant} in {month_text}.")
        return

    n_degraded = int(flagged["degraded"].sum())
    n_bypass = int((flagged["status"] == "bypass").sum())
    n_iqr = int(flagged["degraded_iqr"].sum())
    n_mom = int(flagged["degraded_mom"].sum())
    st.caption(
        f"{len(flagged):,} module(s) need attention — {n_degraded:,} degraded "
        f"({n_iqr:,} peer outlier · {n_mom:,} month-over-month jump) "
        f"+ {n_bypass:,} bypassed."
    )

    def _reason(row: pd.Series) -> str:
        if row["status"] == "bypass":
            return "Bypassed"
        tags = []
        if row["degraded_iqr"]:
            tags.append("Peer outlier (IQR)")
        if row["degraded_mom"]:
            tags.append("MoM jump")
        return " + ".join(tags) if tags else "Degraded"

    flagged = flagged.sort_values(
        ["status", "conductivity"], ascending=[True, False]
    )
    detail = pd.DataFrame(
        {
            "Stage": flagged["stage_label"],
            "Module": flagged["module_label"],
            "Reason": flagged.apply(_reason, axis=1),
            "Conductivity (uS/cm)": flagged["conductivity"].map(
                lambda v: f"{v:,.0f}" if pd.notna(v) else "—"
            ),
            "Prev Month (uS/cm)": flagged["prev_conductivity"].map(
                lambda v: f"{v:,.0f}" if pd.notna(v) else "—"
            ),
            "Stage Fence (uS/cm)": flagged["cutoff"].map(
                lambda v: f"{v:,.0f}" if pd.notna(v) else "—"
            ),
            "Install Date": flagged["install_date"].map(
                lambda v: pd.Timestamp(v).strftime("%d %b %Y") if pd.notna(v) else "Unknown"
            ),
        }
    )
    st.dataframe(detail, width="stretch", hide_index=True)
    st.download_button(
        "Download flagged modules (CSV)",
        detail.to_csv(index=False).encode("utf-8"),
        file_name=f"{plant.replace(' ', '_')}_flagged_{month_text.replace(' ', '_')}.csv",
        mime="text/csv",
        key="portfolio_plant_flagged_csv",
    )


def render_portfolio_page(df: pd.DataFrame, params: pd.DataFrame, mis: pd.DataFrame) -> None:
    st.title("Fleet Portfolio")
    st.caption(
        "Fleet-wide membrane health across every plant — no plant selection needed. "
        "A module needs attention if it is bypassed, or degraded — an active module "
        "that is a peer outlier in its stage (Q3 + 1.5·IQR around the stage median) "
        "or shows an unusually high month-over-month jump in conductivity."
    )

    status = compute_fleet_status(df, mis)
    if status.empty:
        st.info("No readings are available to build the fleet view.")
        return

    # Month + zone slicers — the whole page renders for the chosen month and the
    # chosen zones, defaulting to the most recent month and all zones. Every
    # downstream section keys off `snapshot`/`selected`/`status_zoned`.
    months = sorted(status["report_date"].dropna().unique(), reverse=True)
    month_labels = [pd.Timestamp(m).strftime("%b %Y") for m in months]
    zones = sorted(status["zone"].dropna().astype(str).unique().tolist())
    picker_col, zone_col = st.columns([1, 2])
    with picker_col:
        picked = st.selectbox("Report month", month_labels, index=0)
    with zone_col:
        selected_zones = st.multiselect(
            "Zones",
            zones,
            default=zones,
            help="Filter the whole page to one or more zones. Clear the box to show all zones.",
        )
    selected = months[month_labels.index(picked)]
    month_text = picked
    active_zones = selected_zones or zones  # empty selection = all zones
    status_zoned = status[status["zone"].isin(active_zones)]
    snapshot = status_zoned[status_zoned["report_date"] == selected]

    total_plants = int(snapshot["plant"].nunique())
    total_modules = len(snapshot)
    active_modules = int((snapshot["status"] == "active").sum())
    degraded = int(snapshot["degraded"].sum())
    bypassed = int((snapshot["status"] == "bypass").sum())
    need = int(snapshot["need"].sum())
    need_pct = need / total_modules * 100 if total_modules else 0.0

    # Last month's demand (modules to replace) for the same zones, and the delta.
    # The previous month is the next entry in `months` (sorted newest-first).
    picked_index = month_labels.index(picked)
    prev_month = months[picked_index + 1] if picked_index + 1 < len(months) else None
    if prev_month is not None:
        prev_need = int(
            status_zoned.loc[status_zoned["report_date"] == prev_month, "need"].sum()
        )
        prev_month_text = pd.Timestamp(prev_month).strftime("%b %Y")
        demand_delta = need - prev_need
    else:
        prev_need = None
        prev_month_text = "—"
        demand_delta = None

    # ----- 1. Headline KPI: modules to replace -----
    hero_card(
        "Modules to Replace",
        f"{need:,}",
        f"{month_text} — {degraded:,} degraded + {bypassed:,} bypassed "
        f"across {total_plants:,} plants ({need_pct:.1f}% of {total_modules:,} modules)",
    )

    # ----- 2. Supporting KPI cards -----
    st.markdown("")
    kpis = st.columns(6)
    with kpis[0]:
        metric_card("Plants", f"{total_plants:,}", "in fleet")
    with kpis[1]:
        metric_card("Active Modules", f"{active_modules:,}", f"{total_modules:,} total this month")
    with kpis[2]:
        metric_card("Degraded", f"{degraded:,}", "active, IQR or MoM jump", "#d97706")
    with kpis[3]:
        metric_card("Bypassed", f"{bypassed:,}", "offline modules", "#dc2626")
    with kpis[4]:
        metric_card(
            "Last Month Demand",
            f"{prev_need:,}" if prev_need is not None else "—",
            f"modules to replace · {prev_month_text}",
        )
    with kpis[5]:
        if demand_delta is None:
            metric_card("Δ vs Last Month", "—", "no prior month")
        else:
            # More modules to replace than last month is bad (red); fewer is good (green).
            delta_color = (
                "#dc2626" if demand_delta > 0 else "#16a34a" if demand_delta < 0 else "#0f172a"
            )
            metric_card(
                "Δ vs Last Month",
                f"{demand_delta:+,}",
                f"vs {prev_need:,} in {prev_month_text}",
                delta_color,
            )

    # ----- Jump navigation: clickable cards that scroll to each section -----
    st.markdown("")
    jump_nav(
        [
            ("Plant ranking", "plant-ranking", "per-plant attention"),
            ("Worst 50 modules", "worst-modules", "highest conductivity"),
            ("Conductivity rises", "cond-rises", "module MoM"),
            ("Flow drops", "flow-drops", "module MoM"),
            ("Salt passage", "salt-passage", "sites > 10%"),
            ("Permeate cond. rises", "permeate-cond-rises", "site MoM"),
            ("Permeate flow drops", "permeate-flow-drops", "site MoM"),
            ("Age profile", "age-profile", "install years"),
        ]
    )

    # ----- 2. Plant ranking (centerpiece) -----
    st.markdown("---")
    st.subheader("Plant ranking", anchor="plant-ranking")
    st.caption(
        "One row per plant, sorted by modules needing attention. Click a column header "
        "to re-sort, or select a row to see that plant's flagged modules below."
    )
    ranking = build_plant_ranking(snapshot, selected)
    ranking_event = st.dataframe(
        ranking,
        width="stretch",
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key="portfolio_ranking",
    )
    st.download_button(
        "Download plant ranking (CSV)",
        ranking.to_csv(index=False).encode("utf-8"),
        file_name=f"fleet_ranking_{month_text.replace(' ', '_')}.csv",
        mime="text/csv",
    )

    # Drill-down: clicking a plant row lists every module flagged for it this month.
    selected_rows = ranking_event.selection.get("rows", []) if ranking_event else []
    if selected_rows and not ranking.empty:
        chosen_plant = str(ranking.iloc[selected_rows[0]]["Plant"])
        render_plant_flagged_modules(snapshot, chosen_plant, month_text)

    # ----- 3. Fleet trend -----
    st.markdown("---")
    st.subheader("Fleet trend")
    st.caption("Month-over-month deterioration across the selected zones.")
    st.plotly_chart(make_fleet_trend_chart(status_zoned), width="stretch")

    # ----- 4 & 5. Zone rollup + worst modules -----
    left, right = st.columns(2)
    with left:
        st.subheader("By zone")
        st.plotly_chart(make_zone_rollup_chart(snapshot), width="stretch")
    with right:
        st.subheader(f"Worst 50 modules — {month_text}", anchor="worst-modules")
        active_snap = snapshot[snapshot["status"] == "active"].dropna(subset=["conductivity"])
        worst = active_snap.sort_values("conductivity", ascending=False).head(50)
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
        st.dataframe(worst_table, width="stretch", hide_index=True, height=600)

    # ----- 5b. Biggest month-over-month conductivity rises -----
    st.markdown("---")
    st.subheader(f"Top 50 conductivity rises — {month_text}", anchor="cond-rises")
    st.caption(
        "Active modules whose conductivity rose the most versus their own prior "
        "reading. A jump shared across a whole plant/stage usually points to a feed "
        "or cleaning issue rather than individual membranes."
    )
    jumps = active_snap.dropna(subset=["prev_conductivity"]).copy()
    jumps["delta"] = jumps["conductivity"] - jumps["prev_conductivity"]
    jumps = jumps[jumps["delta"] > 0].sort_values("delta", ascending=False).head(50)
    if jumps.empty:
        st.info("No month-over-month increases to show (no prior-month readings yet).")
    else:
        jump_table = pd.DataFrame(
            {
                "Plant": jumps["plant"],
                "Stage": jumps["stage_label"],
                "Module": jumps["module_label"],
                "Prev (uS/cm)": jumps["prev_conductivity"].map(lambda v: f"{v:,.0f}"),
                "Now (uS/cm)": jumps["conductivity"].map(lambda v: f"{v:,.0f}"),
                "Change (uS/cm)": jumps["delta"].map(lambda v: f"+{v:,.0f}"),
                "Change (x)": (jumps["conductivity"] / jumps["prev_conductivity"]).map(
                    lambda v: f"{v:,.1f}x"
                ),
            }
        )
        st.dataframe(jump_table, width="stretch", hide_index=True, height=600)
        st.download_button(
            "Download conductivity rises (CSV)",
            jump_table.to_csv(index=False).encode("utf-8"),
            file_name=f"mom_cond_rises_{month_text.replace(' ', '_')}.csv",
            mime="text/csv",
        )

    # ----- 5c. Biggest month-over-month flow drops -----
    st.markdown("---")
    st.subheader(f"Top 50 flow drops — {month_text}", anchor="flow-drops")
    st.caption(
        "Active modules whose flow fell the most versus their own prior reading. "
        "A falling permeate flow is a classic fouling / membrane-loss signal."
    )
    drops = active_snap.dropna(subset=["flow", "prev_flow"]).copy()
    drops["delta"] = drops["prev_flow"] - drops["flow"]
    drops = drops[drops["delta"] > 0].sort_values("delta", ascending=False).head(50)
    if drops.empty:
        st.info("No month-over-month flow drops to show (no prior-month readings yet).")
    else:
        drop_table = pd.DataFrame(
            {
                "Plant": drops["plant"],
                "Stage": drops["stage_label"],
                "Module": drops["module_label"],
                "Prev (L/hr)": drops["prev_flow"].map(lambda v: f"{v:,.0f}"),
                "Now (L/hr)": drops["flow"].map(lambda v: f"{v:,.0f}"),
                "Change (L/hr)": drops["delta"].map(lambda v: f"-{v:,.0f}"),
                "Change (%)": (drops["flow"] / drops["prev_flow"] - 1.0).map(
                    lambda v: f"{v * 100:,.0f}%"
                ),
            }
        )
        st.dataframe(drop_table, width="stretch", hide_index=True, height=600)
        st.download_button(
            "Download flow drops (CSV)",
            drop_table.to_csv(index=False).encode("utf-8"),
            file_name=f"mom_flow_drops_{month_text.replace(' ', '_')}.csv",
            mime="text/csv",
        )

    # ----- 5d. Highest conductivity salt passage (feed vs permeate) -----
    st.markdown("---")
    st.subheader(
        f"Top 10 sites — salt passage > {SALT_PASSAGE_FLAG_PCT:.0f}% — {month_text}",
        anchor="salt-passage",
    )
    st.caption(
        "Conductivity salt passage = permeate conductivity (CIS 180) ÷ feed "
        "conductivity (CIS 151) × 100 — the share of feed salinity leaking through "
        "the whole RO train. A healthy train passes only a few percent; a high "
        "figure means the membranes are letting salts through plant-wide."
    )
    passage = compute_salt_passage(params, mis)
    passage = passage[passage["zone"].isin(active_zones)]
    passage = passage[passage["report_date"] == selected]
    passage = passage[passage["passage_pct"] > SALT_PASSAGE_FLAG_PCT]
    passage = passage.sort_values("passage_pct", ascending=False).head(10)
    if passage.empty:
        st.info(
            f"No site exceeds {SALT_PASSAGE_FLAG_PCT:.0f}% salt passage this month "
            "(or feed/permeate conductivity wasn't recorded in these workbooks)."
        )
    else:
        passage_table = pd.DataFrame(
            {
                "Plant": passage["plant"],
                "Zone": passage["zone"],
                "Feed (uS/cm)": passage["feed"].map(lambda v: f"{v:,.0f}"),
                "Permeate (uS/cm)": passage["permeate"].map(lambda v: f"{v:,.0f}"),
                "Salt passage": passage["passage_pct"].map(lambda v: f"{v:.1f}%"),
            }
        )
        st.dataframe(passage_table, width="stretch", hide_index=True)
        st.download_button(
            "Download salt passage (CSV)",
            passage_table.to_csv(index=False).encode("utf-8"),
            file_name=f"salt_passage_{month_text.replace(' ', '_')}.csv",
            mime="text/csv",
        )

    # ----- 5e. Highest month-over-month rise in permeate conductivity -----
    st.markdown("---")
    st.subheader(
        f"Top 10 sites — permeate conductivity rise — {month_text}",
        anchor="permeate-cond-rises",
    )
    st.caption(
        "Whole-plant permeate (product) conductivity from CIS 180, versus the "
        "plant's own prior month. A site-wide rise means the product water got "
        "saltier across the train — a feed, cleaning, or membrane-integrity signal."
    )
    perm_rise = compute_permeate_conductivity(params, mis)
    perm_rise = perm_rise[perm_rise["zone"].isin(active_zones)]
    perm_rise = perm_rise[perm_rise["report_date"] == selected]
    perm_rise = perm_rise.dropna(subset=["prev_permeate"])
    perm_rise = perm_rise[perm_rise["rise"] > 0]
    perm_rise = perm_rise.sort_values("rise", ascending=False).head(10)
    if perm_rise.empty:
        st.info(
            "No site shows a permeate-conductivity rise this month "
            "(or CIS 180 wasn't recorded with a prior month to compare)."
        )
    else:
        perm_table = pd.DataFrame(
            {
                "Plant": perm_rise["plant"],
                "Zone": perm_rise["zone"],
                "Prev (uS/cm)": perm_rise["prev_permeate"].map(lambda v: f"{v:,.0f}"),
                "Now (uS/cm)": perm_rise["permeate"].map(lambda v: f"{v:,.0f}"),
                "Rise (uS/cm)": perm_rise["rise"].map(lambda v: f"+{v:,.0f}"),
                "Change": perm_rise["change_pct"].map(lambda v: f"{v:,.0f}%"),
            }
        )
        st.dataframe(perm_table, width="stretch", hide_index=True)
        st.download_button(
            "Download permeate conductivity rises (CSV)",
            perm_table.to_csv(index=False).encode("utf-8"),
            file_name=f"permeate_cond_rises_{month_text.replace(' ', '_')}.csv",
            mime="text/csv",
        )

    # ----- 5f. Biggest fall in permeate flow (month over month) -----
    st.markdown("---")
    st.subheader(
        f"Top 10 sites — biggest permeate-flow drop — {month_text}",
        anchor="permeate-flow-drops",
    )
    st.caption(
        "Whole-plant permeate (product) flow from FIS 180, normalized to m3/hr, "
        "versus the plant's own prior month. A sustained fall in permeate flow is "
        "a classic fouling / membrane-loss signal."
    )
    flow_mom = compute_permeate_flow(params, mis)
    flow_mom = flow_mom[flow_mom["zone"].isin(active_zones)]
    flow_mom = flow_mom[flow_mom["report_date"] == selected]
    flow_mom = flow_mom.dropna(subset=["prev_flow_m3hr"])
    flow_mom = flow_mom[flow_mom["fall_m3hr"] > 0]
    flow_mom = flow_mom.sort_values("fall_m3hr", ascending=False).head(10)
    if flow_mom.empty:
        st.info(
            "No site shows a permeate-flow drop this month "
            "(or FIS 180 flow wasn't recorded with a prior month to compare)."
        )
    else:
        flow_table = pd.DataFrame(
            {
                "Plant": flow_mom["plant"],
                "Zone": flow_mom["zone"],
                "Prev (m3/hr)": flow_mom["prev_flow_m3hr"].map(lambda v: f"{v:,.2f}"),
                "Now (m3/hr)": flow_mom["flow_m3hr"].map(lambda v: f"{v:,.2f}"),
                "Fall (m3/hr)": flow_mom["fall_m3hr"].map(lambda v: f"-{v:,.2f}"),
                "Change": flow_mom["change_pct"].map(lambda v: f"{v:,.1f}%"),
            }
        )
        st.dataframe(flow_table, width="stretch", hide_index=True)
        st.download_button(
            "Download permeate-flow drops (CSV)",
            flow_table.to_csv(index=False).encode("utf-8"),
            file_name=f"permeate_flow_drops_{month_text.replace(' ', '_')}.csv",
            mime="text/csv",
        )

    # ----- 6. Membrane age profile (click a bar to list its modules) -----
    st.markdown("---")
    st.subheader("Membrane age profile", anchor="age-profile")
    age_chart = make_age_profile_chart(snapshot, selected)
    if age_chart is None:
        st.info("No install dates available to build the age profile.")
    else:
        st.caption("Click a bar to list the modules installed that year.")
        age_event = st.plotly_chart(
            age_chart,
            width="stretch",
            on_select="rerun",
            selection_mode="points",
            key="portfolio_age_profile",
        )
        picked_points = age_event.selection.get("points", []) if age_event else []
        picked_years = {str(p.get("x")) for p in picked_points if p.get("x") is not None}
        if picked_years:
            install_year = pd.to_datetime(
                snapshot["install_date"], errors="coerce"
            ).dt.year
            cohort = snapshot[install_year.astype("Int64").astype(str).isin(picked_years)]
            year_text = ", ".join(sorted(picked_years))
            st.markdown(f"**Modules installed in {year_text}** — {len(cohort):,} fitted")
            cohort_table = pd.DataFrame(
                {
                    "Plant": cohort["plant"],
                    "Zone": cohort["zone"],
                    "Stage": cohort["stage_label"],
                    "Module": cohort["module_label"],
                    "Install Date": pd.to_datetime(
                        cohort["install_date"], errors="coerce"
                    ).dt.strftime("%d %b %Y"),
                    "Status": cohort["status"],
                }
            ).sort_values(["Plant", "Stage", "Module"])
            st.dataframe(cohort_table, width="stretch", hide_index=True, height=360)
            st.download_button(
                "Download these modules (CSV)",
                cohort_table.to_csv(index=False).encode("utf-8"),
                file_name=f"modules_installed_{year_text.replace(', ', '_')}.csv",
                mime="text/csv",
            )


def render_dashboard(df: pd.DataFrame, params: pd.DataFrame) -> None:
    filtered, selected_group, selected_plant, selected_module, metric_name = sidebar_filters(df)
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


def render_data_manager(engine: Engine) -> None:
    """Sidebar admin panel: upload reports through a validation gate, review the
    pending queue, and permanently remove a bad report. Shown on every page."""
    with st.sidebar.expander("🗂 Manage Data"):
        # --- Upload + validate (no commit until Confirm) ---
        uploaded_files = st.file_uploader(
            "Upload monthly Excel report",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="Files are parsed and checked; nothing is saved until you Confirm.",
            key="data_uploader",
        )
        if st.button("Validate & preview", disabled=not uploaded_files):
            stage_uploaded_files(engine, uploaded_files or [])
            rerun_app()

        render_staged_reports(engine)

        # --- Re-parse a stored report through the same gate ---
        stored = report_files_index(engine)
        if stored:
            st.markdown("---")
            reparse_choice = st.selectbox(
                "Re-parse a stored report",
                stored,
                index=None,
                placeholder="Pick a file to re-parse…",
                help="Re-runs the current parser on the stored bytes, back into the review queue.",
                key="reparse_choice",
            )
            if reparse_choice and st.button("Re-parse", key="reparse_btn"):
                data = load_report_bytes(engine, reparse_choice)
                if data is None:
                    st.warning("No stored bytes for that file.")
                else:
                    result = parse_report_bytes(reparse_choice, data)
                    st.session_state.setdefault("staged", {})[reparse_choice] = {
                        "result": result,
                        "issues": validate_parse(engine, reparse_choice, result),
                        "content_hash": hashlib.sha256(data).hexdigest(),
                        "data": data,
                    }
                    rerun_app()

        st.markdown("---")

        # --- Remove an already-committed report ---
        reports = ingested_report_summary(engine)
        if reports.empty:
            st.caption("No reports have been committed yet.")
            return

        st.caption(f"{len(reports)} report(s) in the database.")
        choice = st.selectbox(
            "Remove a report",
            reports["filename"].tolist(),
            index=None,
            placeholder="Select a file to remove…",
            key="remove_report_choice",
        )
        if not choice:
            return

        row = reports[reports["filename"] == choice].iloc[0]
        st.caption(
            f"{int(row['n_readings'] or 0)} readings · "
            f"{int(row['n_parameters'] or 0)} parameters · "
            f"{int(row['n_mis'] or 0)} MIS rows"
        )
        confirm = st.checkbox(
            f"Confirm permanent removal of **{choice}**",
            key="remove_report_confirm",
        )
        if st.button(
            "Delete report data", type="primary", disabled=not confirm
        ):
            counts = remove_report(engine, APP_DIR, choice)
            load_readings.clear()
            load_parameters.clear()
            load_mis.clear()
            compute_fleet_status.clear()
            st.success(
                f"Removed {choice}: {counts['readings']} readings, "
                f"{counts['parameters']} parameters, {counts['mis']} MIS rows."
            )
            rerun_app()


# --------------------------------------------------------------------------- #
# Scan IMR: handwritten printout photo -> Gemini Vision -> filled template.
# Zero-cost path: a free Google AI Studio key + gemini-2.5-flash. The rate limit
# (~a handful per minute) is fine for one-IMR-at-a-time entry.
# --------------------------------------------------------------------------- #
SCAN_ZONES = ["Ahmedabad", "Vadodara", "Ankleshwar", "Panoli", "Jhagadia", "Dahej", "Vapi"]
SCAN_MODEL = "gemini-2.5-flash"
TEMPLATE_PATH = APP_DIR / "templates" / "IMR_template.xlsx"

# Where each value lands in templates/IMR_template.xlsx (verified cell map).
SCAN_STAGE_BLOCKS = {  # leading roman numeral -> (mo, inst, flow, cond) columns
    "I": (1, 2, 5, 6),       # A B E F
    "II": (8, 9, 12, 13),    # H I L M
    "III": (15, 16, 19, 20), # O P S T
}
SCAN_DATA_TOP, SCAN_DATA_BOTTOM = 10, 32

EXTRACTION_PROMPT = """You are reading a handwritten "Individual Module Report (IMR)" for a
reverse-osmosis plant. Return ONLY JSON matching exactly this shape:

{
  "plant_sr_no": <integer or null>,
  "report_date": "<YYYY-MM-DD or null>",
  "zone": "<one of: Ahmedabad, Vadodara, Ankleshwar, Panoli, Jhagadia, Dahej, Vapi, or null>",
  "site_name": "<string or null>",
  "plant_capacity": "<string or null>",
  "stages": [
    {"stage_label": "<e.g. I STAGE, II STAGE, III STAGE>",
     "volume_ml": <number or null, the volume printed in the "Time for ___ ml" header, in millilitres; 1 ltr = 1000>,
     "modules": [
       {"mo_no": <int>, "inst_date": "<YYYY-MM-DD or null>",
        "time_sec": <number or null>,
        "flow": <number or null>, "cond": <number or null>,
        "remark": "<string or null, e.g. BY PASS>"}
     ]}
  ],
  "parameters": [ {"tag": "<e.g. CIS - 151 -, FIS - 180 -, PI - 1601 ->", "value": <number>} ]
}

Rules:
- Transcribe digits exactly as written; do not invent or 'correct' values.
- If a cell is blank or unreadable, use null (do not guess).
- 'time_sec' is the "Time for ___ ml" column (the stopwatch reading in seconds);
  'volume_ml' is the millilitre volume named in that column's header (e.g. 500, 1000).
- 'flow' is the Total flow (liter/hr) column; it is often blank because the form
  computes it from the time — read it only if a number is actually written.
- 'cond' is the Cond. us/cm column.
- Keep stage labels and parameter tags verbatim as printed on the form.
Return only the JSON object, no prose."""


def _secret_key(env_var: str, section: str) -> str | None:
    key = os.environ.get(env_var)
    if not key:
        try:
            key = st.secrets[section]["api_key"]
        except Exception:  # noqa: BLE001 - secrets file / section may be absent
            key = None
    return key or None


def get_gemini_key() -> str | None:
    return _secret_key("GEMINI_API_KEY", "gemini")


def get_groq_key() -> str | None:
    return _secret_key("GROQ_API_KEY", "groq")


class GeminiRateLimited(Exception):
    """Free-tier quota / rate limit hit (429) — retrying soon won't help."""


class GeminiUnavailable(Exception):
    """Model is transiently overloaded (503/500/UNAVAILABLE) after retries."""


# When the primary model is overloaded, fall back to a second model alias whose
# pool is usually separate. Same family, still free.
SCAN_FALLBACK_MODEL = "gemini-flash-latest"


def _coerce_extracted(obj: object) -> dict:
    """Models sometimes wrap the JSON object in an array (e.g. `[{...}]`) despite
    the prompt. Normalize any list to the first dict element so callers always get
    the single object they expect."""
    if isinstance(obj, dict):
        return obj
    if isinstance(obj, list):
        for item in obj:
            if isinstance(item, dict):
                return item
    return {}


def extract_imr_from_images(
    images: list[tuple[bytes, str]], model: str, api_key: str
) -> dict:
    """Send one or more form photos/PDFs to Gemini and return the parsed JSON dict.

    Resilient to the transient overloads (503 UNAVAILABLE) Gemini throws under load:
    retries the primary model with backoff, then tries a fallback model. Raises
    GeminiRateLimited on a 429 (wait for quota) and GeminiUnavailable if every
    attempt is overloaded."""
    from google import genai  # lazy: only needed on this page
    from google.genai import types

    client = genai.Client(api_key=api_key)
    parts: list[object] = [
        types.Part.from_bytes(data=data, mime_type=mime) for data, mime in images
    ]
    parts.append(EXTRACTION_PROMPT)
    config = types.GenerateContentConfig(response_mime_type="application/json", temperature=0)

    # (model, seconds-to-wait-before-this-attempt): try primary, retry primary
    # after a pause, then fall back to a second model.
    attempts = [(model, 0.0), (model, 2.0), (SCAN_FALLBACK_MODEL, 3.0)]
    last_error: Exception | None = None
    for mdl, wait in attempts:
        if wait:
            time.sleep(wait)
        try:
            response = client.models.generate_content(model=mdl, contents=parts, config=config)
            return _coerce_extracted(json.loads(response.text or "{}"))
        except Exception as exc:  # noqa: BLE001
            text = str(exc).lower()
            if "429" in text or "resource_exhausted" in text or "quota" in text:
                raise GeminiRateLimited(str(exc)) from exc
            if any(s in text for s in ("503", "500", "unavailable", "overload", "internal")):
                last_error = exc
                continue  # transient — try again / fall back
            raise  # a real error (bad request, auth, etc.) — surface it
    raise GeminiUnavailable(str(last_error) if last_error else "Gemini unavailable")


# Groq backup: separate infrastructure, so it survives a Gemini overload. Llama 4
# Scout is multimodal; its vision input is IMAGES ONLY, so PDFs are rasterized to
# page images first. A notch below Gemini on handwriting — fine for a backup.
GROQ_MODEL = "meta-llama/llama-4-scout-17b-16e-instruct"


def _to_images(files: list[tuple[bytes, str]]) -> list[tuple[bytes, str]]:
    """Expand any PDFs into per-page PNG images; pass real images through. Used
    for providers (Groq) whose vision input does not accept PDF directly."""
    out: list[tuple[bytes, str]] = []
    for data, mime in files:
        if "pdf" in (mime or "").lower():
            import fitz  # PyMuPDF, lazy

            doc = fitz.open(stream=data, filetype="pdf")
            try:
                for page in doc:
                    pix = page.get_pixmap(dpi=150)
                    out.append((pix.tobytes("png"), "image/png"))
            finally:
                doc.close()
        else:
            out.append((data, mime or "image/jpeg"))
    return out


def extract_imr_via_groq(files: list[tuple[bytes, str]], api_key: str) -> dict:
    """Backup extractor: Groq Llama 4 Scout vision -> parsed JSON dict."""
    import base64

    from groq import Groq  # lazy: only needed on fallback

    images = _to_images(files)
    content: list[dict] = [{"type": "text", "text": EXTRACTION_PROMPT}]
    for data, mime in images:
        b64 = base64.b64encode(data).decode("ascii")
        content.append({"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}})

    client = Groq(api_key=api_key)
    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": content}],
        response_format={"type": "json_object"},
        temperature=0,
    )
    return _coerce_extracted(json.loads(response.choices[0].message.content or "{}"))


class ExtractionFailed(Exception):
    """Every configured provider failed; message lists what happened."""


def extract_imr(
    files: list[tuple[bytes, str]], gemini_key: str | None, groq_key: str | None
) -> tuple[dict, str]:
    """Extract with Gemini (primary); fall back to Groq on any Gemini failure.
    Returns (data, provider_label). Raises ExtractionFailed if all providers fail."""
    notes: list[str] = []
    if gemini_key:
        try:
            return extract_imr_from_images(files, SCAN_MODEL, gemini_key), "Gemini (2.5-flash)"
        except GeminiRateLimited:
            notes.append("Gemini hit its free-tier limit (429)")
        except GeminiUnavailable:
            notes.append("Gemini was overloaded (503)")
        except Exception as exc:  # noqa: BLE001
            notes.append(f"Gemini error: {str(exc)[:120]}")
    if groq_key:
        try:
            return extract_imr_via_groq(files, groq_key), "Groq (Llama 4 Scout)"
        except Exception as exc:  # noqa: BLE001
            notes.append(f"Groq error: {str(exc)[:120]}")
    if not gemini_key and not groq_key:
        notes.append("no API key configured")
    raise ExtractionFailed(" → ".join(notes))


def _stage_block(stage_label: str) -> tuple[int, int, int, int] | None:
    """Map a stage label (e.g. 'III STAGE (PT)') to its block columns by the
    leading roman numeral."""
    token = str(stage_label).strip().upper().split()[0] if str(stage_label).strip() else ""
    return SCAN_STAGE_BLOCKS.get(token)


def fill_imr_template(extracted: dict) -> bytes:
    """Write an extracted/corrected IMR dict into a copy of the finalized template
    and return the workbook bytes."""
    import openpyxl
    from openpyxl.utils import get_column_letter

    def clean(v: object) -> object:
        """Blank cells from the editor arrive as NaN/''; keep them empty so the
        sheet never shows a literal 'nan'."""
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        if isinstance(v, str) and not v.strip():
            return None
        return v

    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb["IMR"]

    # cover
    ws["C4"] = clean(extracted.get("plant_sr_no"))
    ws["J4"] = clean(extracted.get("report_date"))
    ws["C5"] = clean(extracted.get("site_name"))
    ws["J5"] = clean(extracted.get("zone"))
    ws["C6"] = clean(extracted.get("plant_capacity"))

    # stage grids
    for stage in extracted.get("stages") or []:
        cols = _stage_block(stage.get("stage_label", ""))
        if not cols:
            continue
        mo_c, inst_c, flow_c, cond_c = cols
        time_c = mo_c + 2  # the "Time for ___ ml" column sits two right of Mo no.
        time_letter = get_column_letter(time_c)
        # keep the printed stage label verbatim on the band (row 8 = block start)
        ws.cell(row=8, column=mo_c, value=stage.get("stage_label"))
        # Fill the timed volume into the header so the live flow formula is
        # self-documenting; numerator = volume_ml / 1000 * 3600 = volume_ml * 3.6.
        volume_ml = pd.to_numeric(stage.get("volume_ml"), errors="coerce")
        numerator = None
        if pd.notna(volume_ml) and volume_ml > 0:
            ws.cell(row=9, column=time_c, value=f"Time for {volume_ml:g} ml")
            numerator = volume_ml * 3.6
        modules = (stage.get("modules") or [])[: SCAN_DATA_BOTTOM - SCAN_DATA_TOP + 1]
        for i, m in enumerate(modules):
            mo = clean(m.get("mo_no"))
            if mo is None:
                continue  # skip blank editor rows
            r = SCAN_DATA_TOP + i
            ws.cell(row=r, column=mo_c, value=mo)
            ws.cell(row=r, column=inst_c, value=clean(m.get("inst_date")))
            remark = str(clean(m.get("remark")) or "")
            if "by pass" in remark.lower() or "bypass" in remark.lower():
                ws.cell(row=r, column=cond_c, value="BY PASS")
            else:
                seconds = pd.to_numeric(m.get("time_sec"), errors="coerce")
                ws.cell(row=r, column=time_c, value=float(seconds) if pd.notna(seconds) else None)
                # With both the timed volume and the reading, write the same live
                # formula the paper form uses so editing the time updates the flow;
                # otherwise fall back to the transcribed flow number.
                if numerator is not None and pd.notna(seconds) and seconds != 0:
                    ws.cell(row=r, column=flow_c, value=f"={numerator:g}/{time_letter}{r}")
                else:
                    ws.cell(row=r, column=flow_c, value=clean(m.get("flow")))
                ws.cell(row=r, column=cond_c, value=clean(m.get("cond")))

    # operating parameters: map EVERY label cell in the param region to the value
    # cell on its right, then write each extracted param to its matching label.
    # Matching on normalized text means text-only tags ("Reject cond", "Feed Flow",
    # "Permeat Flow") work too, and "CIS - 151 -" vs "CIS 151" both resolve.
    value_cell_by_label: dict[str, object] = {}
    for row in ws.iter_rows(min_row=35, max_row=46, min_col=1, max_col=10):
        for cell in row:
            norm = normalize_text(cell.value)
            if norm:
                value_cell_by_label.setdefault(norm, ws.cell(row=cell.row, column=cell.column + 1))
    for p in extracted.get("parameters") or []:
        target = value_cell_by_label.get(normalize_text(p.get("tag")))
        val = clean(p.get("value"))
        if target is not None and val is not None:
            target.value = val

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def render_scan_page(engine: Engine) -> None:
    st.title("Scan IMR — photo to Excel")
    st.caption(
        "Photograph or scan (PDF) a handwritten IMR printout; it's read into the "
        "finalized Excel format. Review and fix anything, then download it or commit "
        "it straight to the database. Gemini is primary, Groq is the backup if Gemini "
        "is overloaded. No more retyping."
    )

    gemini_key = get_gemini_key()
    groq_key = get_groq_key()
    if not gemini_key and not groq_key:
        st.info(
            "Add a **free** API key to enable scanning (`.streamlit/secrets.toml`):\n\n"
            "```toml\n[gemini]\napi_key = \"YOUR_FREE_GEMINI_KEY\"   # primary\n\n"
            "[groq]\napi_key = \"YOUR_FREE_GROQ_KEY\"       # backup\n```\n\n"
            "Gemini: https://aistudio.google.com/apikey · Groq: https://console.groq.com/keys "
            "(both free)."
        )
        return

    # ----- capture -----
    up_tab, cam_tab = st.tabs(["Upload photo(s)", "Use camera"])
    images: list[tuple[bytes, str]] = []
    with up_tab:
        files = st.file_uploader(
            "Photo(s) or a scanned PDF of the same form",
            type=["jpg", "jpeg", "png", "webp", "pdf"],
            accept_multiple_files=True,
        )
        for f in files or []:
            mime = f.type or (
                "application/pdf" if str(f.name).lower().endswith(".pdf") else "image/jpeg"
            )
            images.append((f.getvalue(), mime))
    with cam_tab:
        shot = st.camera_input("Snap the form")
        if shot is not None:
            images.append((shot.getvalue(), shot.type or "image/jpeg"))

    if st.button("Extract data", type="primary", disabled=not images):
        with st.spinner(f"Reading {len(images)} file(s)…"):
            try:
                extracted, provider = extract_imr(images, gemini_key, groq_key)
            except ExtractionFailed as exc:
                st.error(
                    f"Couldn't extract — every provider failed: {exc}. "
                    "If both are overloaded, wait a minute and retry."
                )
                return
        st.session_state["scan_extracted"] = extracted
        st.session_state["scan_provider"] = provider

    extracted = st.session_state.get("scan_extracted")
    if not extracted:
        return
    provider = st.session_state.get("scan_provider", "")
    if provider:
        st.success(f"Extracted via {provider}. Review and fix below.")

    # ----- review & correct -----
    st.markdown("---")
    st.subheader("Review & fix")
    st.caption("Correct anything the OCR misread before you download or commit.")

    c = st.columns(5)
    with c[0]:
        sr = st.number_input("Plant SR No", value=int(extracted.get("plant_sr_no") or 0), step=1, min_value=0)
    with c[1]:
        date_str = st.text_input("Report date (YYYY-MM-DD)", value=str(extracted.get("report_date") or ""))
    with c[2]:
        zone_val = extracted.get("zone") if extracted.get("zone") in SCAN_ZONES else None
        zone = st.selectbox("Zone", SCAN_ZONES, index=SCAN_ZONES.index(zone_val) if zone_val else 0)
    with c[3]:
        site = st.text_input("Site name", value=str(extracted.get("site_name") or ""))
    with c[4]:
        cap = st.text_input("Plant capacity", value=str(extracted.get("plant_capacity") or ""))

    corrected_stages = []
    for idx, stage in enumerate(extracted.get("stages") or []):
        st.markdown(f"**{stage.get('stage_label', f'Stage {idx+1}')}**")
        vol_default = pd.to_numeric(stage.get("volume_ml"), errors="coerce")
        volume_ml = st.number_input(
            "Time for ___ ml (volume timed per reading)",
            value=float(vol_default) if pd.notna(vol_default) else 0.0,
            step=100.0, min_value=0.0, key=f"scan_vol_{idx}",
            help="Flow is computed as volume × 3.6 ÷ time. Leave 0 if the form has "
                 "no timed volume — flow is then taken from the Flow column as-is.",
        )
        mod_df = pd.DataFrame(stage.get("modules") or [],
                              columns=["mo_no", "inst_date", "time_sec", "flow", "cond", "remark"])
        edited = st.data_editor(
            mod_df, num_rows="dynamic", width="stretch", key=f"scan_stage_{idx}",
            column_config={
                "mo_no": "Mo no.", "inst_date": "Inst date", "time_sec": "Time (sec)",
                "flow": "Flow (L/hr)", "cond": "Cond (uS/cm)", "remark": "Remark",
            },
        )
        corrected_stages.append({
            "stage_label": stage.get("stage_label"),
            "volume_ml": volume_ml or None,
            "modules": edited.to_dict("records"),
        })

    params = extracted.get("parameters") or []
    if params:
        st.markdown("**Operating parameters**")
        par_df = pd.DataFrame(params, columns=["tag", "value"])
        par_edit = st.data_editor(par_df, num_rows="dynamic", width="stretch", key="scan_params")
        corrected_params = par_edit.to_dict("records")
    else:
        corrected_params = []

    corrected = {
        "plant_sr_no": int(sr) or None,
        "report_date": date_str.strip() or None,
        "zone": zone,
        "site_name": site.strip() or None,
        "plant_capacity": cap.strip() or None,
        "stages": corrected_stages,
        "parameters": corrected_params,
    }

    xlsx_bytes = fill_imr_template(corrected)
    base = f"{corrected['plant_sr_no'] or 'imr'}_{corrected['report_date'] or 'report'}"
    filename = safe_upload_name(f"{base}.xlsx")

    # ----- download / commit -----
    st.markdown("---")
    dl, cm = st.columns(2)
    with dl:
        st.download_button(
            "⬇ Download Excel", xlsx_bytes, file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
    with cm:
        commit = st.button("✓ Commit to database", type="primary", width="stretch")

    if commit:
        content_hash = hashlib.sha256(xlsx_bytes).hexdigest()
        result = parse_report_bytes(filename, xlsx_bytes)
        issues = validate_parse(engine, filename, result)
        blockers = [i for i in issues if i.level == "ERROR"]
        for issue in issues:
            ISSUE_RENDERERS.get(issue.level, st.info)(issue.message)
        if blockers:
            st.error("Fix the blocking issues above (or correct the data) before committing.")
        elif result.readings.empty:
            st.error("No readings were parsed from the filled form — check the module rows.")
        else:
            store_report_bytes(engine, filename, content_hash, xlsx_bytes, status="committed")
            counts = commit_parse_result(engine, filename, content_hash, result)
            load_readings.clear()
            load_parameters.clear()
            load_mis.clear()
            st.success(
                f"Committed {filename}: {counts['readings']} readings, "
                f"{counts['parameters']} parameters, {counts['mis']} MIS rows."
            )
            st.session_state.pop("scan_extracted", None)


def build_submission_roster(mis: pd.DataFrame) -> pd.DataFrame:
    """The roster of plants expected to report: the latest MIS row per
    plant_sr_no, with its zone and site name. This is 'who should send an IMR'."""
    cols = ["plant_sr_no", "zone", "site_name"]
    if mis is None or mis.empty:
        return pd.DataFrame(columns=cols)
    latest = (
        mis.dropna(subset=["plant_sr_no"])
        .sort_values("report_date")
        .drop_duplicates("plant_sr_no", keep="last")
        .copy()
    )
    latest["plant_sr_no"] = latest["plant_sr_no"].astype(int)
    latest["zone"] = latest["zone"].fillna("").astype(str).str.strip().replace("", "Unknown")
    latest["site_name"] = latest["site_name"].fillna("").astype(str).str.strip()
    return latest[cols].reset_index(drop=True)


def submission_matrix(df: pd.DataFrame) -> pd.DataFrame:
    """Boolean grid (index plant_sr_no x columns 'YYYY-MM'): True if that plant
    had at least one reading that month."""
    sub = df.dropna(subset=["plant_sr_no"]).copy()
    if sub.empty:
        return pd.DataFrame()
    sub["plant_sr_no"] = sub["plant_sr_no"].astype(int)
    sub["month"] = sub["report_date"].dt.to_period("M").astype(str)
    return sub.groupby(["plant_sr_no", "month"]).size().unstack(fill_value=0) > 0


def render_input_tracker(df: pd.DataFrame, mis: pd.DataFrame) -> None:
    st.title("IMR Input Tracker")
    st.caption(
        "Who has sent their IMR each month, by zone — the plant register (MIS) "
        "checked against the readings actually received. Use it to chase the sites "
        "that haven't reported yet."
    )

    roster = build_submission_roster(mis)
    grid = submission_matrix(df)
    if roster.empty or grid.empty:
        st.info(
            "Need both a plant register (MIS) and some received readings to track "
            "submissions. Ingest at least one report with plant SR numbers."
        )
        return

    months = sorted(grid.columns.tolist(), reverse=True)
    month_labels = [pd.Period(m).strftime("%b %Y") for m in months]
    picked_label = st.selectbox("Month", month_labels, index=0)
    picked = months[month_labels.index(picked_label)]

    received_sr = set(grid.index[grid[picked]]) if picked in grid.columns else set()
    roster = roster.copy()
    roster["received"] = roster["plant_sr_no"].isin(received_sr)

    # Display name: prefer the readings sheet name, else the MIS site name.
    name_by_sr = (
        df.dropna(subset=["plant_sr_no"])
        .assign(plant_sr_no=lambda d: d["plant_sr_no"].astype(int))
        .sort_values("report_date")
        .drop_duplicates("plant_sr_no", keep="last")
        .set_index("plant_sr_no")["plant"]
        .to_dict()
    )
    roster["site"] = roster.apply(
        lambda r: str(name_by_sr.get(r["plant_sr_no"]) or r["site_name"] or f"SR {r['plant_sr_no']}"),
        axis=1,
    )

    # Last month each plant submitted (for chasing the chronic non-reporters).
    def last_submitted(sr: int) -> str:
        if sr in grid.index:
            hits = [m for m in months if grid.loc[sr, m]]
            if hits:
                return pd.Period(max(hits)).strftime("%b %Y")
        return "Never"
    roster["last_submitted"] = roster["plant_sr_no"].map(last_submitted)

    expected = len(roster)
    got = int(roster["received"].sum())
    missing = expected - got
    pct = got / expected * 100 if expected else 0.0
    zone_grp = roster.groupby("zone")
    zones_total = roster["zone"].nunique()
    zones_done = int((zone_grp["received"].mean() == 1).sum())

    # ----- 1. Headline + KPIs -----
    hero_card(
        "Missing IMRs",
        f"{missing:,}",
        f"{picked_label} — {got}/{expected} sites received ({pct:.0f}%)",
        color="#16a34a" if missing == 0 else "#dc2626",
    )
    st.markdown("")
    k = st.columns(4)
    with k[0]:
        metric_card("Sites expected", f"{expected:,}", "in the register")
    with k[1]:
        metric_card("Received", f"{got:,}", f"{pct:.0f}% this month", "#16a34a")
    with k[2]:
        metric_card("Missing", f"{missing:,}", "not yet received", "#dc2626")
    with k[3]:
        metric_card("Zones complete", f"{zones_done}/{zones_total}", "all sites in")

    # ----- 2. By zone -----
    st.markdown("---")
    st.subheader(f"By zone — {picked_label}")
    zs = (
        zone_grp.agg(expected=("plant_sr_no", "count"), received=("received", "sum"))
        .reset_index()
    )
    zs["missing"] = zs["expected"] - zs["received"]
    zs["completion"] = (zs["received"] / zs["expected"] * 100).round(0)
    zs = zs.sort_values(["missing", "zone"], ascending=[False, True])
    fig = go.Figure(
        go.Bar(
            x=zs["completion"],
            y=zs["zone"],
            orientation="h",
            marker_color=[
                "#16a34a" if c == 100 else "#f59e0b" if c >= 50 else "#dc2626"
                for c in zs["completion"]
            ],
            text=[f"{r}/{e}" for r, e in zip(zs["received"], zs["expected"])],
            textposition="auto",
            hovertemplate="%{y}: %{x:.0f}% received (%{text})<extra></extra>",
        )
    )
    fig.update_layout(
        template="plotly_white",
        height=max(220, 42 * len(zs)),
        xaxis_title="Completion %",
        xaxis_range=[0, 100],
        margin={"l": 10, "r": 10, "t": 10, "b": 10},
    )
    st.plotly_chart(fig, width="stretch")
    ztab = pd.DataFrame(
        {
            "Zone": zs["zone"],
            "Received": zs["received"],
            "Expected": zs["expected"],
            "Missing": zs["missing"],
            "Completion": zs["completion"].map(lambda v: f"{v:.0f}%"),
        }
    )
    st.dataframe(ztab, width="stretch", hide_index=True)

    # ----- 3. Missing this month (the actionable list) -----
    st.markdown("---")
    st.subheader(f"Missing this month — {picked_label}")
    miss = roster[~roster["received"]].sort_values(["zone", "site"])
    if miss.empty:
        st.success("Every registered site has submitted this month. 🎉")
    else:
        for zone, grp in miss.groupby("zone"):
            with st.expander(f"{zone} — {len(grp)} missing", expanded=True):
                st.dataframe(
                    pd.DataFrame(
                        {
                            "Site": grp["site"],
                            "SR No": grp["plant_sr_no"],
                            "Last submitted": grp["last_submitted"],
                        }
                    ),
                    width="stretch",
                    hide_index=True,
                )
        miss_out = pd.DataFrame(
            {
                "Zone": miss["zone"],
                "Site": miss["site"],
                "SR No": miss["plant_sr_no"],
                "Last submitted": miss["last_submitted"],
            }
        )
        st.download_button(
            "Download missing list (CSV)",
            miss_out.to_csv(index=False).encode("utf-8"),
            file_name=f"missing_imr_{picked}.csv",
            mime="text/csv",
        )

    # ----- 4. Submission history matrix -----
    st.markdown("---")
    st.subheader("Submission history")
    st.caption("✓ received · ✗ missing, by zone. Most recent 8 months.")
    zones = sorted(roster["zone"].unique())
    sel_zones = st.multiselect("Zones", zones, default=zones, key="tracker_zones")
    recent = months[:8][::-1]  # oldest -> newest for left-to-right reading
    month_cols = [pd.Period(m).strftime("%b %y") for m in recent]
    view = roster[roster["zone"].isin(sel_zones or zones)].sort_values(["zone", "site"])
    rows = []
    for _, r in view.iterrows():
        sr = r["plant_sr_no"]
        row = {"Zone": r["zone"], "Site": r["site"]}
        for m, label in zip(recent, month_cols):
            received_here = bool(grid.loc[sr, m]) if (sr in grid.index and m in grid.columns) else False
            row[label] = "✓" if received_here else "✗"
        rows.append(row)
    mat = pd.DataFrame(rows)
    if mat.empty:
        st.info("No sites in the selected zones.")
    else:
        def color_cell(v: str) -> str:
            if v == "✓":
                return "color:#16a34a;font-weight:700"
            if v == "✗":
                return "color:#dc2626;font-weight:700"
            return ""
        styler = mat.style
        styler = (styler.map if hasattr(styler, "map") else styler.applymap)(
            color_cell, subset=month_cols
        )
        st.dataframe(
            styler,
            width="stretch",
            hide_index=True,
            height=min(640, 60 + 28 * len(mat)),
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
        .hero-card {
            border: 1px solid #fecaca;
            border-radius: 12px;
            background: linear-gradient(180deg, #fff5f5 0%, #ffffff 70%);
            padding: 22px 26px;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08);
        }
        .hero-title {
            color: #64748b;
            font-size: 0.95rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.02em;
        }
        .hero-value {
            font-size: clamp(3rem, 7vw, 5rem);
            line-height: 1.05;
            font-weight: 800;
            margin-top: 0.2rem;
        }
        .hero-subtitle {
            color: #64748b;
            font-size: 1rem;
            margin-top: 0.35rem;
        }
        .nav-grid {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin: 6px 0 2px;
        }
        .nav-card {
            flex: 1 1 150px;
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            background: #f8fafc;
            padding: 10px 14px;
            text-decoration: none;
            display: flex;
            flex-direction: column;
            gap: 2px;
            transition: background .12s ease, border-color .12s ease, box-shadow .12s ease;
        }
        .nav-card:hover {
            background: #eff6ff;
            border-color: #bfdbfe;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08);
        }
        .nav-card .nav-label {
            color: #0f172a;
            font-weight: 700;
            font-size: 0.9rem;
        }
        .nav-card .nav-sub {
            color: #64748b;
            font-size: 0.78rem;
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
        # Rebuild from durable bytes anything whose derived rows went missing
        # (e.g. Cloud restart wiped disk but report_files survived). Usually a no-op.
        recovered = ingest_report_files(engine)
        summary["ingested"].extend(recovered["ingested"])
        summary["failed"].extend(recovered["failed"])
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

    render_data_manager(engine)

    pages = st.navigation(
        [
            st.Page(
                lambda: render_portfolio_page(df, params, mis),
                title="Portfolio",
                icon="🌐",
                url_path="portfolio",
                default=True,
            ),
            st.Page(
                lambda: render_input_tracker(df, mis),
                title="IMR Tracker",
                icon="📥",
                url_path="tracker",
            ),
            st.Page(
                lambda: render_scan_page(engine),
                title="Scan IMR",
                icon="📷",
                url_path="scan",
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
