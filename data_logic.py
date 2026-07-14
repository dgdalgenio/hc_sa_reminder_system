"""
data_logic.py
--------------
All data loading, cleaning, merging, and filtering logic for the
"Visit Views per SA" app lives here.

Expected inputs
----------------
1. addresses_data   -> file with (at least) columns: 'Contract', 'City',
                        'Barangay', 'Address', 'Landmark'
2. contracts_data    -> file with (at least) columns: 'CONTRACT  #', 'SUPERIOR',
                        'AGENT', 'CLIENT NAME', 'INSTALMENT', 'INST. DATE',
                        'DAYS TO INST', 'CONTACT PHONE'
3. form_data         -> file with (at least) columns:
                        'Input contract number (Copy it out from the form to reduce errors)',
                        'When did you visit?'

Any of the three can be .csv or .xlsx/.xls. CSV encoding is auto-detected
so files exported with odd encodings (Windows-1252, latin-1, etc.) still load.
"""

from __future__ import annotations

import io
from dataclasses import dataclass
from datetime import date
from typing import Optional, Union

import chardet
import pandas as pd

# ---------------------------------------------------------------------------
# Column name constants
# ---------------------------------------------------------------------------

CONTRACT_COL = "CONTRACT  #"  # NOTE: double space, matches source system export
ADDR_CONTRACT_COL = "Contract"
SUPERIOR_COL = "SUPERIOR"
AGENT_COL = "AGENT"
DATE_COL = "INST. DATE"
DAYS_TO_INST_COL = "DAYS TO INST"

FORM_CONTRACT_COL = "Input contract number (Copy it out from the form to reduce errors)"
FORM_VISIT_DATE_COL = "When did you visit?"

# Columns required to exist (after cleaning) in each source file
REQUIRED_ADDRESS_COLS = {ADDR_CONTRACT_COL, "City", "Barangay", "Address", "Landmark"}
REQUIRED_CONTRACT_COLS = {
    CONTRACT_COL,
    SUPERIOR_COL,
    AGENT_COL,
    "CLIENT NAME",
    "INSTALMENT",
    DATE_COL,
    DAYS_TO_INST_COL,
    "CONTACT PHONE",
}
REQUIRED_FORM_COLS = {FORM_CONTRACT_COL, FORM_VISIT_DATE_COL}

# Final columns requested for the output table, and the "clean" labels
# they should be renamed to for display.
OUTPUT_COLUMN_ORDER = [
    AGENT_COL,
    SUPERIOR_COL,
    CONTRACT_COL,
    "CLIENT NAME",
    "INSTALMENT",
    DATE_COL,
    DAYS_TO_INST_COL,
    "Visited?",
    "Date of visit",
    "Visit Status",
    "CONTACT PHONE",
    "City",
    "Barangay",
    "Address",
    "Landmark",
]

CLEAN_COLUMN_NAMES = {
    AGENT_COL: "Agent",
    SUPERIOR_COL: "Superior",
    CONTRACT_COL: "Contract #",
    "CLIENT NAME": "Client Name",
    "INSTALMENT": "Instalment",
    DATE_COL: "Instalment Date",
    DAYS_TO_INST_COL: "Days to Instalment",
    "Visited?": "Visited?",
    "Date of visit": "Date of Visit",
    "Visit Status": "Visit Status",
    "CONTACT PHONE": "Contact Phone",
    "City": "City",
    "Barangay": "Barangay",
    "Address": "Address",
    "Landmark": "Landmark",
}


# Clean (post-rename) column names used repeatedly by the UI layer
CLEAN_DAYS_TO_INST_COL = CLEAN_COLUMN_NAMES[DAYS_TO_INST_COL]
CLEAN_VISITED_COL = CLEAN_COLUMN_NAMES["Visited?"]
CLEAN_INSTALMENT_COL = CLEAN_COLUMN_NAMES["INSTALMENT"]
CLEAN_CONTACT_PHONE_COL = CLEAN_COLUMN_NAMES["CONTACT PHONE"]
CLEAN_SUPERIOR_COL = CLEAN_COLUMN_NAMES[SUPERIOR_COL]
CLEAN_AGENT_COL = CLEAN_COLUMN_NAMES[AGENT_COL]

# Row-highlight thresholds (in days to instalment due)
DUE_SOON_RED_RANGE = (0, 2)   # light red
DUE_SOON_YELLOW_RANGE = (3, 5)  # light yellow
COLOR_RED = "FFCCCC"
COLOR_YELLOW = "FFF3B0"


class DataLoadError(Exception):
    """Raised when an uploaded file can't be parsed or is missing required columns."""


# ---------------------------------------------------------------------------
# File loading helpers
# ---------------------------------------------------------------------------

def _detect_encoding(raw_bytes: bytes) -> str:
    """Best-effort encoding detection, falling back to utf-8."""
    try:
        result = chardet.detect(raw_bytes)
        encoding = result.get("encoding")
        return encoding or "utf-8"
    except Exception:
        return "utf-8"


def read_any_table(file: Union[str, "io.BytesIO", "io.BufferedReader"], label: str = "file") -> pd.DataFrame:
    """
    Robustly read a CSV or Excel file into a DataFrame.

    Accepts:
      - a file path (str)
      - a Streamlit UploadedFile / file-like object (has .name and .read())

    Handles:
      - .csv / .txt with unknown/odd encodings (auto-detected via chardet)
      - .xlsx / .xls / .xlsm via openpyxl
    """
    # Figure out the filename to inspect the extension
    name = getattr(file, "name", None) or (file if isinstance(file, str) else "")
    name_lower = str(name).lower()

    try:
        if name_lower.endswith((".xlsx", ".xls", ".xlsm")):
            return pd.read_excel(file)

        if name_lower.endswith((".csv", ".txt")) or name_lower == "":
            # Need raw bytes to sniff encoding
            if isinstance(file, str):
                with open(file, "rb") as f:
                    raw = f.read()
                encoding = _detect_encoding(raw)
                return pd.read_csv(file, encoding=encoding)
            else:
                raw = file.read()
                file.seek(0)
                encoding = _detect_encoding(raw)
                return pd.read_csv(io.BytesIO(raw), encoding=encoding)

        # Unknown extension: try excel first, then csv as a fallback
        try:
            return pd.read_excel(file)
        except Exception:
            if hasattr(file, "seek"):
                file.seek(0)
            raw = file.read() if hasattr(file, "read") else open(file, "rb").read()
            encoding = _detect_encoding(raw)
            return pd.read_csv(io.BytesIO(raw), encoding=encoding)

    except Exception as exc:  # pragma: no cover - defensive
        raise DataLoadError(f"Could not read {label}: {exc}") from exc


def _strip_whitespace_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip leading/trailing whitespace from column names (but preserve internal spacing,
    since e.g. 'CONTRACT  #' intentionally has a double space in the source system)."""
    df = df.copy()
    df.columns = [str(c).strip() if str(c).strip() != str(c) else c for c in df.columns]
    return df


def _validate_columns(df: pd.DataFrame, required: set, label: str) -> None:
    missing = required - set(df.columns)
    if missing:
        raise DataLoadError(
            f"{label} is missing expected column(s): {sorted(missing)}. "
            f"Found columns: {list(df.columns)}"
        )


# ---------------------------------------------------------------------------
# Loading + cleaning each source
# ---------------------------------------------------------------------------

def load_addresses(file) -> pd.DataFrame:
    df = read_any_table(file, label="addresses data")
    df = _strip_whitespace_columns(df)
    _validate_columns(df, REQUIRED_ADDRESS_COLS, "Addresses data")
    df[ADDR_CONTRACT_COL] = df[ADDR_CONTRACT_COL].astype(str).str.strip()
    return df


def load_contracts(file) -> pd.DataFrame:
    df = read_any_table(file, label="contracts data")
    df = _strip_whitespace_columns(df)
    _validate_columns(df, REQUIRED_CONTRACT_COLS, "Contracts data")

    # Drop rows without a SUPERIOR (mirrors notebook's cleaning step)
    df = df[~df[SUPERIOR_COL].isna()].copy()

    df[CONTRACT_COL] = df[CONTRACT_COL].astype(str).str.strip()
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    df[DAYS_TO_INST_COL] = pd.to_numeric(df[DAYS_TO_INST_COL], errors="coerce")
    return df


def load_form_data(file) -> pd.DataFrame:
    df = read_any_table(file, label="form data")
    df = _strip_whitespace_columns(df)
    _validate_columns(df, REQUIRED_FORM_COLS, "Form data")
    df[FORM_CONTRACT_COL] = df[FORM_CONTRACT_COL].astype(str).str.strip()
    return df


# ---------------------------------------------------------------------------
# Core pipeline
# ---------------------------------------------------------------------------

def build_reference_data(addresses_data: pd.DataFrame, contracts_data: pd.DataFrame) -> pd.DataFrame:
    """Merge contracts data with addresses data on contract number."""
    reference_data = contracts_data.merge(
        addresses_data,
        left_on=CONTRACT_COL,
        right_on=ADDR_CONTRACT_COL,
        how="inner",
    )
    return reference_data


def get_superior_options(reference_data: pd.DataFrame) -> list:
    return sorted(reference_data[SUPERIOR_COL].dropna().unique().tolist())


def get_agent_options(reference_data: pd.DataFrame, superior: Optional[str]) -> list:
    if superior is None:
        subset = reference_data
    else:
        subset = reference_data[reference_data[SUPERIOR_COL] == superior]
    return sorted(subset[AGENT_COL].dropna().unique().tolist())


@dataclass
class DateWindow:
    selection_date: date
    history_days: int
    forward_days: int

    @property
    def start_date(self) -> pd.Timestamp:
        return pd.to_datetime(self.selection_date) - pd.Timedelta(days=self.history_days)

    @property
    def end_date(self) -> pd.Timestamp:
        return pd.to_datetime(self.selection_date) + pd.Timedelta(days=self.forward_days)


def filter_reference_data(
    reference_data: pd.DataFrame,
    dss_filter: Optional[str],
    agents_filter: Optional[str],
    window: DateWindow,
) -> pd.DataFrame:
    """Filter reference_data by date window, superior (DSS), and agent."""
    if reference_data.empty:
        return reference_data.copy()

    mask = (reference_data[DATE_COL] >= window.start_date) & (
        reference_data[DATE_COL] <= window.end_date
    )

    if dss_filter and dss_filter != "All":
        # mask &= ((reference_data[SUPERIOR_COL] == dss_filter) or (reference_data[AGENT_COL] == dss_filter))
        mask &= (reference_data[SUPERIOR_COL] == dss_filter) | (reference_data[AGENT_COL] == dss_filter)

    if agents_filter and agents_filter != "All":
        mask &= reference_data[AGENT_COL] == agents_filter

    filtered = reference_data.loc[mask].copy()
    filtered = filtered.sort_values(by=DAYS_TO_INST_COL)
    return filtered


def attach_visit_status(filtered_data: pd.DataFrame, form_data: pd.DataFrame) -> pd.DataFrame:
    """
    Select the output columns, join against the visit-report form data, and
    derive 'Visited?' / 'Date of visit' / 'Visit Status' columns.
    """
    base_cols = [
        AGENT_COL,
        SUPERIOR_COL,
        CONTRACT_COL,
        "CLIENT NAME",
        "INSTALMENT",
        DATE_COL,
        DAYS_TO_INST_COL,
        "CONTACT PHONE",
        "City",
        "Barangay",
        "Address",
        "Landmark",
    ]
    data_to_show = filtered_data[base_cols].copy()

    columns_to_pull = [FORM_CONTRACT_COL, FORM_VISIT_DATE_COL]
    merged_data = data_to_show.merge(
        form_data[columns_to_pull],
        left_on=CONTRACT_COL,
        right_on=FORM_CONTRACT_COL,
        how="left",
    )

    merged_data["Visit Status"] = merged_data[FORM_CONTRACT_COL].apply(
        lambda x: "visited" if pd.notna(x) else "not visited"
    )
    merged_data["Visited?"] = merged_data["Visit Status"].map(
        {"visited": "Yes", "not visited": "No"}
    )
    merged_data["Date of visit"] = merged_data[FORM_VISIT_DATE_COL]

    merged_data = merged_data.drop(columns=[FORM_CONTRACT_COL, FORM_VISIT_DATE_COL])

    # Guarantee full, requested column order
    merged_data = merged_data[OUTPUT_COLUMN_ORDER]
    return merged_data


def clean_output_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename columns to display-friendly labels for the final output table."""
    return df.rename(columns=CLEAN_COLUMN_NAMES)


def filter_by_visit_status(df: pd.DataFrame, include: str) -> pd.DataFrame:
    """
    Filter the (already clean-column-named) output table by visit status.

    include: one of "All", "Visited only", "Not visited only".
    """
    if df.empty or include == "All":
        return df
    if include == "Visited only":
        return df[df[CLEAN_VISITED_COL] == "Yes"].copy()
    if include == "Not visited only":
        return df[df[CLEAN_VISITED_COL] == "No"].copy()
    raise ValueError(f"Unknown visit-status filter: {include!r}")


def sort_output(df: pd.DataFrame, sort_col: Optional[str], ascending: bool) -> pd.DataFrame:
    """Sort the output table by a column, tolerating missing/None column choice."""
    if df.empty or not sort_col or sort_col not in df.columns:
        return df
    return df.sort_values(by=sort_col, ascending=ascending, na_position="last").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Row highlighting (0-2 days => light red, 3-5 days => light yellow)
# ---------------------------------------------------------------------------

def _days_to_inst_bucket(value) -> Optional[str]:
    """Return 'red', 'yellow', or None for a given days-to-instalment value."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return None
    if pd.isna(v):
        return None
    lo, hi = DUE_SOON_RED_RANGE
    if lo <= v <= hi:
        return "red"
    lo, hi = DUE_SOON_YELLOW_RANGE
    if lo <= v <= hi:
        return "yellow"
    return None


def coerce_display_dtypes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize dtypes for display/export:
      - 'Days to Instalment' -> nullable integer (no decimal point)
      - 'Contact Phone'      -> nullable integer (no decimal point)
    'Instalment' is left as numeric (float) since it's *formatted* as pesos
    for display rather than converted to a string, so it stays sortable.
    """
    df = df.copy()
    if CLEAN_DAYS_TO_INST_COL in df.columns:
        df[CLEAN_DAYS_TO_INST_COL] = pd.to_numeric(
            df[CLEAN_DAYS_TO_INST_COL], errors="coerce"
        ).round().astype("Int64")
    if CLEAN_CONTACT_PHONE_COL in df.columns:
        df[CLEAN_CONTACT_PHONE_COL] = pd.to_numeric(
            df[CLEAN_CONTACT_PHONE_COL], errors="coerce"
        ).round().astype("Int64")
    return df

def toggle_columns(df: pd.DataFrame, include_superior: bool = False, include_agent: bool=True) -> pd.DataFrame:
    """
    'Superior' or 'Agent' is excluded from the output table by default; pass
    include_superior=True to keep it (e.g. via a sidebar/table-view toggle).
    include_agent=True to keep it
    """
    if not include_superior and CLEAN_SUPERIOR_COL in df.columns:
        df = df.drop(columns=[CLEAN_SUPERIOR_COL])
    if not include_agent and CLEAN_AGENT_COL in df.columns:
        df = df.drop(columns=[CLEAN_AGENT_COL])
    return df
    
def style_output_table(df: pd.DataFrame):
    """
    Return a pandas Styler that:
      - highlights each row light red (0-2 days to due) or light yellow
        (3-5 days to due), based on CLEAN_DAYS_TO_INST_COL
      - formats Instalment as pesos, and Days to Instalment / Contact Phone
        as plain integers (no decimal point)
    Safe to call on an empty DataFrame.
    """
    if df.empty:
        return df.style

    def _row_style(row):
        bucket = _days_to_inst_bucket(row.get(CLEAN_DAYS_TO_INST_COL))
        if bucket == "red":
            return [f"background-color: #{COLOR_RED}"] * len(row)
        if bucket == "yellow":
            return [f"background-color: #{COLOR_YELLOW}"] * len(row)
        return [""] * len(row)

    styler = df.style.apply(_row_style, axis=1)

    fmt = {}
    if CLEAN_INSTALMENT_COL in df.columns:
        fmt[CLEAN_INSTALMENT_COL] = lambda v: f"\u20b1{v:,.0f}" if pd.notna(v) else ""
    if CLEAN_DAYS_TO_INST_COL in df.columns:
        fmt[CLEAN_DAYS_TO_INST_COL] = lambda v: f"{int(v)}" if pd.notna(v) else ""
    if CLEAN_CONTACT_PHONE_COL in df.columns:
        fmt[CLEAN_CONTACT_PHONE_COL] = lambda v: f"{int(v)}" if pd.notna(v) else ""
    if fmt:
        styler = styler.format(fmt)

    df[CLEAN_COLUMN_NAMES[DATE_COL]] = df[CLEAN_COLUMN_NAMES[DATE_COL]].dt.strftime("%Y-%m-%d")

    return styler


def run_pipeline(
    addresses_file,
    contracts_file,
    form_file,
    dss_filter: Optional[str],
    agents_filter: Optional[str],
    window: DateWindow,
) -> pd.DataFrame:
    """
    End-to-end pipeline: load all three sources, merge, filter, attach visit
    status, and return the final, display-ready DataFrame.
    """
    addresses_data = load_addresses(addresses_file)
    contracts_data = load_contracts(contracts_file)
    form_data = load_form_data(form_file)

    reference_data = build_reference_data(addresses_data, contracts_data)
    if reference_data.empty:
        raise DataLoadError(
            "No rows matched between addresses data and contracts data. "
            "Check that contract numbers line up between the two files."
        )

    filtered_data = filter_reference_data(reference_data, dss_filter, agents_filter, window)
    result = attach_visit_status(filtered_data, form_data)
    return clean_output_columns(result)