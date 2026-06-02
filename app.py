from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
UPLOAD_DIR_NAME = "uploaded_reports"
MANUAL_DATA_FILE = "manual_readings.csv"
REPORT_GLOBS = ("*.xlsx", "*.xls", "*.csv")
MANUAL_COLUMNS = [
    "source_file",
    "plant_group",
    "plant",
    "stage",
    "module_number",
    "install_date",
    "report_date",
    "flow_lph",
    "conductivity_us_cm",
]
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

# Differential pressure across the high-pressure membrane array. The standard
# instrument scheme on these reports tags the array feed as PI 1601 and the
# array reject/concentrate as PI 1602; dP = feed - reject. Rising dP over time
# is the classic membrane fouling / scaling signal.
DP_FEED_TAG = "1601"
DP_REJECT_TAG = "1602"

DP_METRIC = {
    "axis_label": "Differential Pressure (bar)",
    "unit": "bar",
    "bad_direction": "up",
    "threshold_factor": 1.30,
}


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

    return ReportMeta(plant_group=plant_group, report_date=report_date)


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
            blocks.append(block)

    return blocks


def value_at(raw: pd.DataFrame, row: int, col: int | str) -> object:
    if not isinstance(col, int) or row >= len(raw) or col >= raw.shape[1]:
        return np.nan
    return raw.iat[row, col]


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

    for row_index in range(header_row + 1, len(raw)):
        for block in blocks:
            module_raw = value_at(raw, row_index, block["module_col"])
            module_number = pd.to_numeric(module_raw, errors="coerce")
            if pd.isna(module_number):
                continue

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
                    "stage": block.get("stage", ""),
                    "module_number": float(module_number),
                    "install_date": value_at(raw, row_index, block["install_date_col"]),
                    "report_date": metadata.report_date,
                    "flow_lph": float(flow) if pd.notna(flow) else np.nan,
                    "conductivity_us_cm": float(conductivity) if pd.notna(conductivity) else np.nan,
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
            plant_group = clean_label(row.get(plant_group_col))

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
                "stage": stage,
                "module_number": float(module_number),
                "install_date": install_date,
                "report_date": reading_date,
                "flow_lph": float(flow) if pd.notna(flow) else np.nan,
                "conductivity_us_cm": float(conductivity) if pd.notna(conductivity) else np.nan,
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


def read_manual_readings(root: Path) -> list[dict[str, object]]:
    manual_path = root / MANUAL_DATA_FILE
    if not manual_path.exists():
        return []

    table = pd.read_csv(manual_path)
    for column in MANUAL_COLUMNS:
        if column not in table.columns:
            table[column] = np.nan

    return table[MANUAL_COLUMNS].to_dict("records")


def append_manual_reading(root: Path, row: dict[str, object]) -> None:
    manual_path = root / MANUAL_DATA_FILE
    if manual_path.exists():
        table = pd.read_csv(manual_path)
    else:
        table = pd.DataFrame(columns=MANUAL_COLUMNS)

    table = pd.concat([table, pd.DataFrame([row], columns=MANUAL_COLUMNS)], ignore_index=True)
    table.to_csv(manual_path, index=False)


def safe_upload_name(filename: str) -> str:
    name = Path(filename).name
    name = re.sub(r"[^A-Za-z0-9., _()\-]+", "_", name).strip()
    return name or "uploaded_report.xlsx"


def save_uploaded_reports(files: list[object], root: Path) -> tuple[list[str], list[str]]:
    upload_dir = root / UPLOAD_DIR_NAME
    upload_dir.mkdir(exist_ok=True)

    saved: list[str] = []
    skipped: list[str] = []
    allowed_suffixes = {".xlsx", ".xls", ".csv"}

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


@st.cache_data(show_spinner="Loading and cleaning monthly RO reports...")
def load_reports(data_dir: str) -> tuple[pd.DataFrame, list[str]]:
    root = Path(data_dir)
    paths = report_paths(root)

    skipped: list[str] = []
    records: list[dict[str, object]] = []

    for path in paths:
        try:
            report_records = read_report(path)
        except Exception as exc:  # noqa: BLE001
            skipped.append(f"{path.name}: {exc}")
            continue

        if report_records:
            records.extend(report_records)
        else:
            skipped.append(f"{path.name}: no module table found")

    records.extend(read_manual_readings(root))

    if not records:
        return pd.DataFrame(), skipped

    df = pd.DataFrame(records)
    df["install_date"] = pd.to_datetime(df["install_date"], errors="coerce")
    df["report_date"] = pd.to_datetime(df["report_date"], errors="coerce")
    df["flow_lph"] = pd.to_numeric(df["flow_lph"], errors="coerce")
    df["conductivity_us_cm"] = pd.to_numeric(df["conductivity_us_cm"], errors="coerce")
    df["module_number"] = pd.to_numeric(df["module_number"], errors="coerce")
    df["module_label"] = df["module_number"].map(format_module_number)
    df = df.dropna(subset=["report_date", "module_number"])
    df = df.dropna(subset=["flow_lph", "conductivity_us_cm"], how="all")
    df = df.drop_duplicates(
        subset=[
            "source_file",
            "plant_group",
            "plant",
            "stage",
            "module_number",
            "report_date",
            "flow_lph",
            "conductivity_us_cm",
        ]
    )
    df = df.sort_values(["plant_group", "plant", "module_number", "report_date"])
    return df.reset_index(drop=True), skipped


@st.cache_data(show_spinner="Reading plant operating parameters...")
def load_parameters(data_dir: str) -> pd.DataFrame:
    root = Path(data_dir)
    records: list[dict[str, object]] = []

    for path in report_paths(root):
        try:
            records.extend(read_report_parameters(path))
        except Exception:  # noqa: BLE001
            continue

    if not records:
        return pd.DataFrame(columns=PARAM_COLUMNS)

    df = pd.DataFrame(records)
    df["report_date"] = pd.to_datetime(df["report_date"], errors="coerce")
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df = df.dropna(subset=["report_date", "value"])
    df = df.drop_duplicates(
        subset=["plant_group", "plant", "report_date", "tag"], keep="last"
    )
    return df.sort_values(["plant_group", "plant", "tag", "report_date"]).reset_index(drop=True)


def differential_pressure_series(
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
        feed = group[group["tag"].str.contains(DP_FEED_TAG, regex=False)]["value"]
        reject = group[group["tag"].str.contains(DP_REJECT_TAG, regex=False)]["value"]
        if feed.empty or reject.empty:
            continue
        rows.append({"report_date": report_date, "value": float(feed.iloc[0]) - float(reject.iloc[0])})

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


def fit_linear_trend(series: pd.DataFrame) -> dict[str, object] | None:
    clean = series.dropna(subset=["report_date", "value"]).sort_values("report_date")
    if len(clean) < 2:
        return None

    x = (clean["report_date"] - clean["report_date"].min()).dt.days.to_numpy(dtype=float)
    y = clean["value"].to_numpy(dtype=float)
    if np.isclose(x.max(), x.min()):
        return None

    slope, intercept = np.polyfit(x, y, 1)
    fitted = intercept + slope * x
    ss_res = float(np.sum((y - fitted) ** 2))
    ss_tot = float(np.sum((y - y.mean()) ** 2))
    r_squared = 1 - ss_res / ss_tot if ss_tot else 1.0

    return {
        "slope_per_day": float(slope),
        "intercept": float(intercept),
        "start_date": clean["report_date"].min(),
        "last_date": clean["report_date"].max(),
        "r_squared": float(r_squared),
    }


def default_threshold(series: pd.DataFrame, metric_config: dict[str, object]) -> float:
    values = series["value"].dropna()
    if values.empty:
        return 0.0
    baseline = float(values.iloc[0])
    factor = float(metric_config["threshold_factor"])
    return max(baseline * factor, 0.0)


def system_status(latest: float, threshold: float, bad_direction: str) -> tuple[str, str]:
    if threshold <= 0 or pd.isna(latest):
        return "Unknown", "#64748b"

    if bad_direction == "down":
        if latest <= threshold:
            return "Critical", "#dc2626"
        if latest <= threshold * 1.15:
            return "Warning", "#d97706"
        return "Healthy", "#059669"

    if latest >= threshold:
        return "Critical", "#dc2626"
    if latest >= threshold * 0.85:
        return "Warning", "#d97706"
    return "Healthy", "#059669"


def predicted_failure_date(
    trend: dict[str, object] | None,
    threshold: float,
    bad_direction: str,
) -> pd.Timestamp | None:
    if trend is None:
        return None

    slope = float(trend["slope_per_day"])
    intercept = float(trend["intercept"])
    start_date = pd.Timestamp(trend["start_date"])
    last_date = pd.Timestamp(trend["last_date"])

    if np.isclose(slope, 0.0):
        return None
    if bad_direction == "down" and slope >= 0:
        return None
    if bad_direction == "up" and slope <= 0:
        return None

    crossing_day = (threshold - intercept) / slope
    crossing_date = start_date + pd.to_timedelta(crossing_day, unit="D")
    if crossing_date < last_date:
        return last_date
    if crossing_date > last_date + pd.Timedelta(days=3650):
        return None
    return crossing_date.normalize()


def is_at_or_beyond_threshold(latest: float, threshold: float, bad_direction: str) -> bool:
    if threshold <= 0 or pd.isna(latest):
        return False
    if bad_direction == "down":
        return latest <= threshold
    return latest >= threshold


def make_chart(
    series: pd.DataFrame,
    *,
    metric_name: str,
    metric_config: dict[str, object],
    trend: dict[str, object] | None,
    threshold: float,
    show_trend: bool,
    failure_date: pd.Timestamp | None,
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

    fig.add_hline(
        y=threshold,
        line_dash="dot",
        line_color="#dc2626",
        annotation_text="Failure threshold",
        annotation_position="top left",
    )

    if show_trend and trend is not None:
        last_date = pd.Timestamp(trend["last_date"])
        start_date = pd.Timestamp(trend["start_date"])
        forecast_end = last_date + pd.Timedelta(days=365)
        if failure_date is not None:
            forecast_end = max(forecast_end, failure_date + pd.Timedelta(days=30))
            forecast_end = min(forecast_end, last_date + pd.Timedelta(days=1095))

        dates = pd.date_range(start=start_date, end=forecast_end, periods=80)
        days = (dates - start_date).days.to_numpy(dtype=float)
        values = float(trend["intercept"]) + float(trend["slope_per_day"]) * days

        fig.add_trace(
            go.Scatter(
                x=dates,
                y=values,
                mode="lines",
                name="Predictive ML Trendline",
                line={"color": "#7c3aed", "width": 2, "dash": "dash"},
                hovertemplate="%{x|%d %b %Y}<br>%{y:,.2f}<extra></extra>",
            )
        )

        if failure_date is not None:
            failure_marker = pd.Timestamp(failure_date).to_pydatetime()
            fig.add_shape(
                type="line",
                x0=failure_marker,
                x1=failure_marker,
                y0=0,
                y1=1,
                xref="x",
                yref="paper",
                line={"color": "#dc2626", "dash": "dash", "width": 2},
            )
            fig.add_annotation(
                x=failure_marker,
                y=1,
                xref="x",
                yref="paper",
                text="Predicted failure",
                showarrow=False,
                xanchor="left",
                yanchor="bottom",
                bgcolor="rgba(255,255,255,0.85)",
                font={"color": "#dc2626", "size": 12},
            )

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


def sidebar_filters(df: pd.DataFrame) -> tuple[pd.DataFrame, str, str, str, str, bool]:
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
    show_trend = st.sidebar.checkbox("Show Predictive ML Trendline", value=True)

    filtered = plant_df[plant_df["module_label"] == selected_module].copy()
    return filtered, selected_group, selected_plant, selected_module, metric_name, show_trend


def render_add_data_controls(
    df: pd.DataFrame,
    selected_group: str,
    selected_plant: str,
    selected_module: str,
) -> None:
    st.sidebar.markdown("---")
    with st.sidebar.expander("Add New Data"):
        uploaded_files = st.file_uploader(
            "Upload monthly Excel/CSV report",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            help="Saved files are loaded from the uploaded_reports folder on the next refresh.",
        )
        if st.button("Save Uploaded Report(s)", disabled=not uploaded_files):
            saved, skipped = save_uploaded_reports(uploaded_files or [], APP_DIR)
            load_reports.clear()
            load_parameters.clear()
            if saved:
                st.success(f"Saved {len(saved)} report(s).")
            if skipped:
                st.warning("; ".join(skipped))
            rerun_app()

        st.markdown("---")
        default_stage = ""
        selected_rows = df[
            (df["plant_group"] == selected_group)
            & (df["plant"] == selected_plant)
            & (df["module_label"] == selected_module)
        ]
        if not selected_rows.empty:
            stages = selected_rows["stage"].dropna()
            if not stages.empty:
                default_stage = clean_label(stages.iloc[-1])
                if default_stage == "Unknown":
                    default_stage = ""

        with st.form("manual_reading_form"):
            st.caption("Manual module reading")
            plant_group = st.text_input("Plant Group", value=selected_group)
            plant = st.text_input("Plant", value=selected_plant)
            module_number = st.text_input("Module Number", value=selected_module)
            stage = st.text_input("Stage", value=default_stage)
            reading_date = st.date_input("Reading Date", value=pd.Timestamp.today().date())
            install_date = st.date_input("Install Date", value=pd.Timestamp.today().date())
            flow_lph = st.number_input(
                "Total flow liter/hr.",
                min_value=0.0,
                value=0.0,
                step=1.0,
                format="%.2f",
            )
            conductivity = st.number_input(
                "Cond. us/cm",
                min_value=0.0,
                value=0.0,
                step=1.0,
                format="%.2f",
            )
            submitted = st.form_submit_button("Add Reading")

        if submitted:
            module_value = pd.to_numeric(module_number, errors="coerce")
            if pd.isna(module_value):
                st.error("Module Number must be numeric.")
                return
            if not plant_group.strip() or not plant.strip():
                st.error("Plant Group and Plant are required.")
                return

            append_manual_reading(
                APP_DIR,
                {
                    "source_file": "Manual Entry",
                    "plant_group": plant_group.strip(),
                    "plant": plant.strip(),
                    "stage": stage.strip(),
                    "module_number": float(module_value),
                    "install_date": pd.Timestamp(install_date).date().isoformat(),
                    "report_date": pd.Timestamp(reading_date).date().isoformat(),
                    "flow_lph": float(flow_lph),
                    "conductivity_us_cm": float(conductivity),
                },
            )
            load_reports.clear()
            st.success("Manual reading added.")
            rerun_app()


def make_pressure_overview_chart(pressures: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    for tag, group in pressures.groupby("tag"):
        ordered = group.sort_values("report_date")
        fig.add_trace(
            go.Scatter(
                x=ordered["report_date"],
                y=ordered["value"],
                mode="lines+markers",
                name=str(tag),
                hovertemplate=f"{tag}<br>%{{x|%d %b %Y}}<br>%{{y:,.2f}} bar<extra></extra>",
            )
        )
    fig.update_layout(
        title="All Pressure Tags",
        xaxis_title="Report Date",
        yaxis_title="Pressure (bar)",
        hovermode="x unified",
        margin={"l": 20, "r": 20, "t": 60, "b": 20},
        template="plotly_white",
        height=420,
    )
    return fig


def render_pressure_section(
    params: pd.DataFrame,
    selected_group: str,
    selected_plant: str,
    show_trend: bool,
) -> None:
    if params.empty:
        return

    pressures = params[
        (params["plant_group"] == selected_group)
        & (params["plant"] == selected_plant)
        & (params["kind"] == "pressure")
    ]

    st.markdown("---")
    st.subheader("Membrane Differential Pressure")
    st.caption(
        "Differential pressure across the high-pressure array "
        f"(dP = PI {DP_FEED_TAG} array feed − PI {DP_REJECT_TAG} array reject). "
        "A rising dP is the classic membrane fouling / scaling signal."
    )

    if pressures.empty:
        st.info("No pressure (bar) readings were found for this plant.")
        return

    dp = differential_pressure_series(params, selected_group, selected_plant)
    unit = str(DP_METRIC["unit"])
    bad_direction = str(DP_METRIC["bad_direction"])

    if dp.empty:
        st.info(
            "Pressure readings exist, but this plant is missing one side of the "
            f"PI {DP_FEED_TAG}/PI {DP_REJECT_TAG} pair needed to compute differential pressure. "
            "Raw pressures are shown below."
        )
    else:
        trend = fit_linear_trend(dp)
        threshold_default = default_threshold(dp, DP_METRIC)
        threshold_step = max(threshold_default * 0.05, 0.1)
        threshold = st.number_input(
            "Differential Pressure Failure Threshold (bar)",
            min_value=0.0,
            value=float(round(threshold_default, 2)),
            step=float(round(threshold_step, 2)),
            help="Rising dP crossing this value flags fouling. Used for status and failure-date prediction.",
            key="dp_threshold",
        )

        latest_row = dp.sort_values("report_date").iloc[-1]
        latest = float(latest_row["value"])
        status, status_color = system_status(latest, threshold, bad_direction)
        slope_30 = float(trend["slope_per_day"]) * 30 if trend else np.nan
        failure_date = predicted_failure_date(trend, threshold, bad_direction)
        already_critical = is_at_or_beyond_threshold(latest, threshold, bad_direction)
        failure_marker_date = None if already_critical else failure_date
        latest_date = pd.Timestamp(latest_row["report_date"]).strftime("%d %b %Y")

        if trend:
            slope_text = f"{slope_30:+,.2f} {unit}/30d"
            slope_subtitle = f"R2 {float(trend['r_squared']):.2f}"
        else:
            slope_text = "Need more data"
            slope_subtitle = "At least 2 report months required"

        if already_critical:
            failure_text = "Already Critical"
            failure_subtitle = f"dP crossed threshold on {latest_date}"
        elif failure_date is not None:
            failure_text = failure_date.strftime("%d %b %Y")
            failure_subtitle = "Linear trend threshold crossing"
        else:
            failure_text = "Not projected"
            failure_subtitle = "Trend does not cross threshold"

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            metric_card("Latest dP", f"{latest:,.2f} {unit}", latest_date)
        with col2:
            metric_card("30-Day dP Rate", slope_text, slope_subtitle)
        with col3:
            metric_card("Fouling Status", status, f"Threshold: {threshold:,.2f} {unit}", status_color)
        with col4:
            metric_card("Predicted Failure", failure_text, failure_subtitle)

        st.plotly_chart(
            make_chart(
                dp,
                metric_name="Membrane Differential Pressure",
                metric_config=DP_METRIC,
                trend=trend,
                threshold=threshold,
                show_trend=show_trend,
                failure_date=failure_marker_date,
            ),
            width="stretch",
        )

    st.plotly_chart(make_pressure_overview_chart(pressures), width="stretch")

    with st.expander("Pressure readings"):
        display = (
            pressures[["report_date", "tag", "value", "unit"]]
            .sort_values(["report_date", "tag"])
            .rename(
                columns={
                    "report_date": "Report Date",
                    "tag": "Tag",
                    "value": "Value",
                    "unit": "Unit",
                }
            )
        )
        st.dataframe(display, width="stretch", hide_index=True)


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
    bad_direction = str(metric_config["bad_direction"])
    rows: list[dict[str, object]] = []

    for module_label, series in series_map.items():
        ordered = series.sort_values("report_date")
        latest_row = ordered.iloc[-1]
        latest = float(latest_row["value"])
        trend = fit_linear_trend(series)
        threshold = default_threshold(series, metric_config)
        status, _ = system_status(latest, threshold, bad_direction)
        failure_date = predicted_failure_date(trend, threshold, bad_direction)
        already_critical = is_at_or_beyond_threshold(latest, threshold, bad_direction)

        if trend:
            rate = float(trend["slope_per_day"]) * 30
            rate_text = f"{rate:+,.2f} {unit}/30d"
            r2_text = f"{float(trend['r_squared']):.2f}"
        else:
            rate_text = "n/a"
            r2_text = "n/a"

        if already_critical:
            failure_text = "Already critical"
        elif failure_date is not None:
            failure_text = failure_date.strftime("%d %b %Y")
        else:
            failure_text = "Not projected"

        rows.append(
            {
                "Module": module_label,
                "Latest": f"{latest:,.2f} {unit}",
                "Latest Date": pd.Timestamp(latest_row["report_date"]).strftime("%d %b %Y"),
                "30-Day Rate": rate_text,
                "R2": r2_text,
                "Status": status,
                "Predicted Failure": failure_text,
                "_sort": failure_date if failure_date is not None else pd.Timestamp.max,
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
        "Status and predicted failure use each module's own baseline-derived threshold."
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
        make_comparison_chart(series_map, metric_config=metric_config, average=average),
        width="stretch",
    )

    table = build_comparison_table(series_map, metric_config)
    if not table.empty:
        st.caption("Sorted by predicted failure date (soonest first).")
        st.dataframe(table, width="stretch", hide_index=True)


def render_dashboard(df: pd.DataFrame, params: pd.DataFrame) -> None:
    filtered, selected_group, selected_plant, selected_module, metric_name, show_trend = sidebar_filters(df)
    render_add_data_controls(df, selected_group, selected_plant, selected_module)
    metric_config = METRICS[metric_name]
    metric_col = str(metric_config["column"])
    series = aggregate_series(filtered, metric_col)

    st.title("RO Plant Membrane Health Dashboard")
    st.caption(f"{selected_group} | {selected_plant}")

    if series.empty:
        st.warning("No valid numeric readings are available for this module and metric.")
        return

    trend = fit_linear_trend(series)
    threshold_default = default_threshold(series, metric_config)
    threshold_step = max(threshold_default * 0.05, 1.0)
    threshold = st.sidebar.number_input(
        f"{metric_name} Failure Threshold",
        min_value=0.0,
        value=float(round(threshold_default, 2)),
        step=float(round(threshold_step, 2)),
        help="Used for status and failure-date prediction.",
    )

    latest_row = series.sort_values("report_date").iloc[-1]
    latest = float(latest_row["value"])
    unit = str(metric_config["unit"])
    bad_direction = str(metric_config["bad_direction"])
    status, status_color = system_status(latest, threshold, bad_direction)
    slope_30 = float(trend["slope_per_day"]) * 30 if trend else np.nan
    failure_date = predicted_failure_date(trend, threshold, bad_direction)
    already_critical = is_at_or_beyond_threshold(latest, threshold, bad_direction)
    failure_marker_date = None if already_critical else failure_date

    latest_date = pd.Timestamp(latest_row["report_date"]).strftime("%d %b %Y")
    if trend:
        slope_text = f"{slope_30:+,.2f} {unit}/30d"
        slope_subtitle = f"R2 {float(trend['r_squared']):.2f}"
    else:
        slope_text = "Need more data"
        slope_subtitle = "At least 2 report months required"

    if already_critical:
        failure_text = "Already Critical"
        failure_subtitle = f"Latest reading crossed threshold on {latest_date}"
    elif failure_date is not None:
        failure_text = failure_date.strftime("%d %b %Y")
        failure_subtitle = "Linear trend threshold crossing"
    else:
        failure_text = "Not projected"
        failure_subtitle = "Trend does not cross threshold"

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        metric_card("Latest Reading", f"{latest:,.2f} {unit}", latest_date)
    with col2:
        metric_card("30-Day Degradation Rate", slope_text, slope_subtitle)
    with col3:
        metric_card("System Status", status, f"Threshold: {threshold:,.2f} {unit}", status_color)
    with col4:
        metric_card("Predicted Failure", failure_text, failure_subtitle)

    st.plotly_chart(
        make_chart(
            series,
            metric_name=metric_name,
            metric_config=metric_config,
            trend=trend,
            threshold=threshold,
            show_trend=show_trend,
            failure_date=failure_marker_date,
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
    render_comparison_section(plant_df, selected_module, metric_name, metric_config)

    render_pressure_section(params, selected_group, selected_plant, show_trend)


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

    df, skipped = load_reports(str(APP_DIR))
    if df.empty:
        st.error("No RO module readings were found in the current folder.")
        if skipped:
            st.write(skipped)
        return

    if skipped:
        with st.sidebar.expander("Skipped inputs"):
            for item in skipped:
                st.write(item)

    params = load_parameters(str(APP_DIR))
    render_dashboard(df, params)


if __name__ == "__main__":
    main()
