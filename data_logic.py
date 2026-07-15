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
PAID_COL = "PAID"

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
    PAID_COL,
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
    PAID_COL,
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
    PAID_COL: "Paid",
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
CLEAN_PAID_COL = CLEAN_COLUMN_NAMES[PAID_COL]

# Row-highlight thresholds (in days to instalment due)
DUE_SOON_RED_RANGE = (0, 2)   # light red
DUE_SOON_YELLOW_RANGE = (3, 5)  # light yellow
COLOR_RED = "FFCCCC"
COLOR_YELLOW = "FFF3B0"

# ---------------------------------------------------------------------------
# Visit-status-per-DSS ("visit stats") constants
# ---------------------------------------------------------------------------
VISIT_STATUS_COL = "Visit Status"  # produced by attach_visit_status()
VISIT_STATUS_NOT_VISITED = "not visited"

STATS_AGENT_COL = "Agent"
STATS_TOTAL_DUE_COL = "Total Due Dates"
STATS_MISSED_COL = "Missed Visits"
STATS_COVERAGE_COL = "% Visits Covered"

WEIGHTED_AVERAGE_LABEL = "\u2696\ufe0f Weighted Average"
DEFAULT_OVERDUE_LOOKBACK_DAYS = 20


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


PAID_YES_VALUES = {"1", "1.0", "yes", "y", "true", "paid"}
PAID_NO_VALUES = {"0", "0.0", "no", "n", "false", "unpaid"}


def _coerce_paid_to_yes_no(value):
    """
    Normalize a binary/varied 'PAID' cell (0/1, True/False, 'Yes'/'No', ...)
    into a plain 'Yes' / 'No' string. Unrecognized non-null values are left
    as-is rather than silently dropped, so bad source data stays visible.
    """
    if pd.isna(value):
        return pd.NA
    text = str(value).strip().lower()
    if text in PAID_YES_VALUES:
        return "Yes"
    if text in PAID_NO_VALUES:
        return "No"
    return value


def load_contracts(file) -> pd.DataFrame:
    df = read_any_table(file, label="contracts data")
    df = _strip_whitespace_columns(df)
    _validate_columns(df, REQUIRED_CONTRACT_COLS, "Contracts data")

    # Drop rows without a SUPERIOR (mirrors notebook's cleaning step)
    df = df[~df[SUPERIOR_COL].isna()].copy()

    df[CONTRACT_COL] = df[CONTRACT_COL].astype(str).str.strip()
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")
    df[DAYS_TO_INST_COL] = pd.to_numeric(df[DAYS_TO_INST_COL], errors="coerce")
    df[PAID_COL] = df[PAID_COL].apply(_coerce_paid_to_yes_no)
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


def filter_reference_data_by_scope(
    reference_data: pd.DataFrame,
    dss_filter: Optional[str],
    agents_filter: Optional[str],
) -> pd.DataFrame:
    """
    Same DSS/agent scoping as filter_reference_data(), but WITHOUT the date
    window mask. Used by the visit-stats tab, which looks back from each
    row's own 'DAYS TO INST' rather than from the sidebar's date window.
    """
    if reference_data.empty:
        return reference_data.copy()

    mask = pd.Series(True, index=reference_data.index)

    if dss_filter and dss_filter != "All":
        mask &= (reference_data[SUPERIOR_COL] == dss_filter) | (reference_data[AGENT_COL] == dss_filter)

    if agents_filter and agents_filter != "All":
        mask &= reference_data[AGENT_COL] == agents_filter

    return reference_data.loc[mask].copy()


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
        PAID_COL,
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


# ---------------------------------------------------------------------------
# Visit-status-per-DSS ("visit stats") pipeline
#
# Ports the logic from visit_stats_peragent.ipynb: for accounts overdue by
# at least `lookback_days` (DAYS TO INST <= -lookback_days), what fraction
# of them has each agent actually visited? The notebook's approach checked
# out, so this mirrors it — reusing attach_visit_status() for the visit join
# rather than re-deriving it, and adding a weighted-average rollup.
# ---------------------------------------------------------------------------

def compute_agent_visit_stats(
    scoped_reference_data: pd.DataFrame,
    form_data: pd.DataFrame,
    lookback_days: int,
) -> pd.DataFrame:
    """
    For each agent in scoped_reference_data, compute visit coverage among
    that agent's accounts overdue by at least `lookback_days`
    (DAYS TO INST <= -lookback_days).

    Returns one row per agent that has at least one such overdue account,
    with columns: STATS_AGENT_COL, STATS_TOTAL_DUE_COL, STATS_MISSED_COL,
    STATS_COVERAGE_COL (fraction, 0-1). Agents with zero overdue accounts
    are omitted here — the caller can diff against the full agent roster to
    report those separately. Rows with a missing/NaN agent are dropped.
    """
    empty_result = pd.DataFrame(
        columns=[STATS_AGENT_COL, STATS_TOTAL_DUE_COL, STATS_MISSED_COL, STATS_COVERAGE_COL]
    )
    if scoped_reference_data.empty:
        return empty_result

    merged = attach_visit_status(scoped_reference_data, form_data)
    merged = merged[merged[AGENT_COL].notna()]

    overdue = merged[merged[DAYS_TO_INST_COL] <= -lookback_days]
    if overdue.empty:
        return empty_result

    grouped = overdue.groupby(AGENT_COL, dropna=True)
    records = []
    for agent, group in grouped:
        total_due = len(group)
        missed = int((group[VISIT_STATUS_COL] == VISIT_STATUS_NOT_VISITED).sum())
        coverage = 1 - (missed / total_due)
        records.append(
            {
                STATS_AGENT_COL: agent,
                STATS_TOTAL_DUE_COL: total_due,
                STATS_MISSED_COL: missed,
                STATS_COVERAGE_COL: coverage,
            }
        )

    result = pd.DataFrame(records).sort_values(by=STATS_AGENT_COL).reset_index(drop=True)
    return result


def compute_weighted_average_coverage(stats_df: pd.DataFrame) -> Optional[float]:
    """
    Weighted average visit-coverage across all agents in stats_df, weighted
    by each agent's Total Due Dates (not a plain average of percentages).
    Returns None if there are no due dates to weight by.
    """
    if stats_df.empty:
        return None
    total_due = stats_df[STATS_TOTAL_DUE_COL].sum()
    if total_due == 0:
        return None
    total_missed = stats_df[STATS_MISSED_COL].sum()
    return 1 - (total_missed / total_due)


def append_weighted_average_row(stats_df: pd.DataFrame) -> pd.DataFrame:
    """Append a bottom 'Weighted Average' summary row to a per-agent stats table."""
    if stats_df.empty:
        return stats_df

    total_due = int(stats_df[STATS_TOTAL_DUE_COL].sum())
    total_missed = int(stats_df[STATS_MISSED_COL].sum())
    weighted_avg = compute_weighted_average_coverage(stats_df)

    summary_row = pd.DataFrame(
        [
            {
                STATS_AGENT_COL: WEIGHTED_AVERAGE_LABEL,
                STATS_TOTAL_DUE_COL: total_due,
                STATS_MISSED_COL: total_missed,
                STATS_COVERAGE_COL: weighted_avg,
            }
        ]
    )
    return pd.concat([stats_df, summary_row], ignore_index=True)


@dataclass
class DssVisitStats:
    """Visit-stats results for a single DSS/Superior group."""
    superior: str
    stats_df: pd.DataFrame           # per-agent rows, no weighted-average row
    agents_with_no_due_dates: list
    weighted_average_coverage: Optional[float]


def compute_visit_stats_by_dss(
    reference_data: pd.DataFrame,
    form_data: pd.DataFrame,
    dss_filter: Optional[str],
    agents_filter: Optional[str],
    lookback_days: int,
) -> list:
    """
    Build one DssVisitStats per DSS/Superior in scope. If dss_filter == "All",
    this produces one entry per superior found in reference_data; otherwise
    a single-entry list for that DSS. agents_filter narrows every group down
    to that one agent (mirrors the sidebar's Agent filter).
    """
    superiors = get_superior_options(reference_data) if (not dss_filter or dss_filter == "All") else [dss_filter]

    results = []
    for superior in superiors:
        scoped = filter_reference_data_by_scope(reference_data, superior, agents_filter)
        scoped_agents = sorted(scoped[AGENT_COL].dropna().unique().tolist())

        stats_df = compute_agent_visit_stats(scoped, form_data, lookback_days)
        agents_with_due = set(stats_df[STATS_AGENT_COL]) if not stats_df.empty else set()
        agents_with_no_due_dates = [a for a in scoped_agents if a not in agents_with_due]

        results.append(
            DssVisitStats(
                superior=str(superior),
                stats_df=stats_df,
                agents_with_no_due_dates=agents_with_no_due_dates,
                weighted_average_coverage=compute_weighted_average_coverage(stats_df),
            )
        )
    return results


def style_visit_stats_table(df: pd.DataFrame):
    """
    Return a pandas Styler for a per-agent visit-stats table: percentage
    formatting for the coverage column, plain integers for the count
    columns, and a bolded/highlighted bottom row for the weighted average
    (if present, identified by WEIGHTED_AVERAGE_LABEL).
    """
    if df.empty:
        return df.style

    def _summary_row_style(row):
        if row.get(STATS_AGENT_COL) == WEIGHTED_AVERAGE_LABEL:
            return ["font-weight: bold; background-color: #EEECE1"] * len(row)
        return [""] * len(row)

    styler = df.style.apply(_summary_row_style, axis=1)

    fmt = {
        STATS_TOTAL_DUE_COL: lambda v: f"{int(v)}" if pd.notna(v) else "",
        STATS_MISSED_COL: lambda v: f"{int(v)}" if pd.notna(v) else "",
        STATS_COVERAGE_COL: lambda v: f"{v:.1%}" if pd.notna(v) else "\u2014",
    }
    return styler.format(fmt)


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