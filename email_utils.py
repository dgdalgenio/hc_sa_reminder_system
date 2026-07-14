"""
email_utils.py
---------------
Bulk email prep helper.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

from data_logic import (
    AGENT_COL,
    SUPERIOR_COL,
    DateWindow,
    attach_visit_status,
    clean_output_columns,
    coerce_display_dtypes,
    filter_by_visit_status,
    filter_reference_data,
)

from report_builder import build_visit_report_workbook

DEFAULT_SUBJECT = "PAALALA: Mga Due Date ng Client Repayment + Store Visit Check-in Survey"

PERSON_COL_CANDIDATES = ["Person", "Name", "Agent", "AGENT", "agent"]
EMAIL_COL_CANDIDATES = ["Email", "EMAIL", "email", "Email Address", "email address"]


class EmailMappingError(Exception):
    """Raised when the person->email mapping file can't be parsed."""


def load_person_email_map(file) -> dict:
    """
    Read an uploaded csv/xlsx file with (at least) a Person column and an
    Email column and return {person_name: email_address}. "Person" covers
    both agents and superiors — one flat mapping resolves both the agent's
    own send-to address and their superior's CC address.
    """
    from data_logic import read_any_table, _strip_whitespace_columns

    df = read_any_table(file, label="person email mapping")
    df = _strip_whitespace_columns(df)

    person_col = next((c for c in PERSON_COL_CANDIDATES if c in df.columns), None)
    email_col = next((c for c in EMAIL_COL_CANDIDATES if c in df.columns), None)
    if person_col is None or email_col is None:
        raise EmailMappingError(
            f"Person email mapping file must have a 'Person' column and an "
            f"'Email' column. Found columns: {list(df.columns)}"
        )

    df[person_col] = df[person_col].astype(str).str.strip()
    df[email_col] = df[email_col].astype(str).str.strip()
    df = df[(df[person_col] != "") & (df[email_col] != "")]

    return dict(zip(df[person_col], df[email_col]))


def _normalize_name(name) -> str:
    """Lowercase + whitespace-collapse a person's name for robust matching
    (handles stray leading/trailing spaces, double spaces, and case
    differences between the mapping file and the source data)."""
    return " ".join(str(name).split()).lower()


def _build_normalized_map(person_email_map: dict) -> dict:
    return {_normalize_name(k): v for k, v in person_email_map.items() if k}


@dataclass
class AgentEmailJob:
    """Everything needed to represent one agent's email row in the tracker."""
    superior: str
    agent: str
    email: Optional[str]
    cc_emails: list = field(default_factory=list)
    row_count: int = 0
    display_df: pd.DataFrame = field(repr=False, default=None)
    html_body: str = field(default="", repr=False)
    subject: str = ""
    status: str = "Ready"          # Missing email / Ready


def _build_html_body(job_display_df: pd.DataFrame) -> str:
    """
    Build just the styled table HTML (no greeting/template wrapper) — the
    Power Automate flow supplies its own email body/greeting and only needs
    the table itself dropped in.
    """
    _, _, report_html = build_visit_report_workbook(job_display_df)
    return report_html


def build_tracker(
    reference_data: pd.DataFrame,
    form_data: pd.DataFrame,
    dss_filter: Optional[str],
    window: DateWindow,
    visit_inclusion: str,
    person_email_map: dict,
    subject: str = DEFAULT_SUBJECT,
    default_cc_list: Optional[list] = None,
    report_signature: str = "MSME Team",  # unused now — HTMLBody is table-only, no greeting
) -> list:
    """
    For every agent that falls under the current DSS/Superior filter (across
    ALL superiors if dss_filter == "All"), build the same output table that
    would appear in the Table View tab, and package it into an AgentEmailJob
    with the styled table HTML (no greeting) ready for the Power Automate
    export. `report_signature` is accepted for call-site compatibility but
    no longer used, since the flow's own template supplies the greeting.

    Each job's CC list is: [the agent's own superior's email (if found in
    the mapping)] + default_cc_list (deduped, agent's own To address
    excluded from CC).

    Agents with zero rows after filtering are skipped entirely (no job is
    created for them).
    """
    jobs: list[AgentEmailJob] = []
    default_cc_list = default_cc_list or []
    normalized_map = _build_normalized_map(person_email_map)

    if reference_data.empty:
        return jobs

    # Scope down to the selected DSS/Superior (or everyone, if "All"),
    # within the selected date window, but don't narrow by agent yet.
    scoped = filter_reference_data(reference_data, dss_filter, "All", window)
    if scoped.empty:
        return jobs

    agents = sorted(scoped[AGENT_COL].dropna().unique().tolist())

    for agent in agents:
        agent_rows = scoped[scoped[AGENT_COL] == agent]
        superior = agent_rows[SUPERIOR_COL].dropna().iloc[0] if not agent_rows[SUPERIOR_COL].dropna().empty else ""

        result = attach_visit_status(agent_rows, form_data)
        result_clean = clean_output_columns(result)
        result_clean = filter_by_visit_status(result_clean, visit_inclusion)
        result_clean = coerce_display_dtypes(result_clean)

        if result_clean.empty:
            # Only include agents with a nonzero number of rows.
            continue

        email = normalized_map.get(_normalize_name(agent))
        superior_email = normalized_map.get(_normalize_name(superior)) if superior else None

        cc_emails = [superior_email] if superior_email else []
        for cc in default_cc_list:
            if cc and cc not in cc_emails:
                cc_emails.append(cc)
        # Never CC the agent their own report is addressed to.
        cc_emails = [c for c in cc_emails if c != email]

        job = AgentEmailJob(
            superior=str(superior),
            agent=str(agent),
            email=email,
            cc_emails=cc_emails,
            row_count=len(result_clean),
            display_df=result_clean,
            subject=subject,
            status="Ready" if email else "Missing email",
        )
        if email:
            job.html_body = _build_html_body(result_clean)
        jobs.append(job)

    return jobs


POWER_AUTOMATE_EXPORT_COLUMNS = [
    "Agent", "Superior", "Email", "CC", "Subject", "HTMLBody", "Status", "Error",
]


def build_power_automate_export(jobs: list) -> pd.DataFrame:
    """
    Build the DataFrame to hand off to a Power Automate flow: one row per
    agent with the full HTML body, ready for a "List rows" + "Apply to
    each" + "Send an email (V2)" flow. Agents with no email are skipped
    (nothing for the flow to send). Each row's CC is that agent's own
    superior + the default CC list, resolved per-agent in build_tracker.
    """
    rows = []
    for j in jobs:
        if not j.email:
            continue
        rows.append({
            "Agent": j.agent,
            "Superior": j.superior,
            "Email": j.email,
            "CC": "; ".join(j.cc_emails),
            "Subject": j.subject or DEFAULT_SUBJECT,
            "HTMLBody": j.html_body,
            "Status": "Pending",
            "Error": "",
        })
    return pd.DataFrame(rows, columns=POWER_AUTOMATE_EXPORT_COLUMNS)


def export_power_automate_excel(export_df: pd.DataFrame) -> bytes:
    """Return the (possibly user-edited) export DataFrame as .xlsx bytes."""
    import io as _io
    buf = _io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="AgentEmailQueue")
    buf.seek(0)
    return buf.getvalue()


def tracker_to_dataframe(jobs: list) -> pd.DataFrame:
    """Render the list of AgentEmailJob into the DataFrame shown in the UI."""
    status_icons = {
        "Ready": "\u23f3 Ready",
        "Missing email": "\u26a0\ufe0f Missing email",
    }
    if not jobs:
        return pd.DataFrame(
            columns=["DSS / Superior", "Agent", "Email", "CC", "Rows Included", "Status"]
        )
    return pd.DataFrame(
        [
            {
                "DSS / Superior": j.superior,
                "Agent": j.agent,
                "Email": j.email or "\u26a0\ufe0f not found",
                "CC": "; ".join(j.cc_emails) if j.cc_emails else "",
                "Rows Included": j.row_count,
                "Status": status_icons.get(j.status, j.status),
            }
            for j in jobs
        ]
    )