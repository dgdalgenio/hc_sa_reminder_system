"""
app.py
------
Streamlit UI for "Visit Views per SA". All data logic lives in data_logic.py,
and the styled Excel/HTML report building lives in report_builder.py; this
file is only responsible for layout, inputs, and rendering.
"""

import io
from datetime import date, datetime

import streamlit as st
import streamlit.components.v1 as components

from data_logic import (
    CLEAN_DAYS_TO_INST_COL,
    CLEAN_SUPERIOR_COL,
    CLEAN_VISITED_COL,
    DataLoadError,
    DateWindow,
    attach_visit_status,
    build_reference_data,
    clean_output_columns,
    coerce_display_dtypes,
    filter_by_visit_status,
    filter_reference_data,
    get_agent_options,
    get_superior_options,
    load_addresses,
    load_contracts,
    load_form_data,
    sort_output,
    style_output_table,
    toggle_columns,
)
from report_builder import build_visit_report_workbook, table_to_email
from email_utils import (
    DEFAULT_SUBJECT,
    build_power_automate_export,
    build_tracker,
    export_power_automate_excel,
    load_person_email_map,
    tracker_to_dataframe,
)

st.set_page_config(page_title="Visit Views per SA", layout="wide")
st.title("📍 Visit Views per SA")
st.caption(
    "Upload addresses, contracts, and visit-report (form) data to see which "
    "accounts have been visited, filtered by DSS, agent, and instalment date."
)

# ---------------------------------------------------------------------------
# Sidebar: Section 1 - file uploads
# ---------------------------------------------------------------------------
with st.sidebar.expander("1. Upload data", expanded=True):
    addresses_file = st.file_uploader(
        "Addresses data", type=["csv", "xlsx", "xls"], key="addresses_file"
    )
    contracts_file = st.file_uploader(
        "Contracts data", type=["csv", "xlsx", "xls"], key="contracts_file"
    )
    form_file = st.file_uploader(
        "Form data (visit reports)", type=["csv", "xlsx", "xls"], key="form_file"
    )

if not (addresses_file and contracts_file and form_file):
    st.info("👈 Upload addresses, contracts, and form data in the sidebar to get started.")
    st.stop()

# ---------------------------------------------------------------------------
# Load & merge data (cached so re-filtering doesn't re-parse files)
# ---------------------------------------------------------------------------


@st.cache_data(show_spinner="Loading and merging data...")
def _load_and_merge(addresses_bytes, addresses_name, contracts_bytes, contracts_name,
                     form_bytes, form_name):
    class _NamedBuffer(io.BytesIO):
        pass

    addr_buf = _NamedBuffer(addresses_bytes)
    addr_buf.name = addresses_name
    con_buf = _NamedBuffer(contracts_bytes)
    con_buf.name = contracts_name
    frm_buf = _NamedBuffer(form_bytes)
    frm_buf.name = form_name

    addresses_data = load_addresses(addr_buf)
    contracts_data = load_contracts(con_buf)
    form_data = load_form_data(frm_buf)
    reference_data = build_reference_data(addresses_data, contracts_data)
    return reference_data, form_data


try:
    reference_data, form_data = _load_and_merge(
        addresses_file.getvalue(), addresses_file.name,
        contracts_file.getvalue(), contracts_file.name,
        form_file.getvalue(), form_file.name,
    )
except DataLoadError as exc:
    st.error(f"⚠️ {exc}")
    st.stop()
except Exception as exc:  # pragma: no cover - defensive catch-all
    st.error(f"⚠️ Unexpected error while loading data: {exc}")
    st.stop()

if reference_data.empty:
    st.error(
        "No matching records found between the addresses data and contracts data. "
        "Please check that the contract numbers line up between the two files."
    )
    st.stop()

st.sidebar.success("✅ Data loaded successfully")

# ---------------------------------------------------------------------------
# Sidebar: Section 2 - filters
# ---------------------------------------------------------------------------
with st.sidebar.expander("2. Filters", expanded=True):
    superior_options = ["All"] + get_superior_options(reference_data)
    dss_filter = st.selectbox("DSS / Superior", superior_options, index=0)

    agent_options = ["All"] + get_agent_options(
        reference_data, None if dss_filter == "All" else dss_filter
    ) + [dss_filter]
    agents_filter = st.selectbox("Agent", agent_options, index=0)

    visit_inclusion = st.radio(
        "Include which rows?",
        ["All", "Visited only", "Not visited only"],
        index=0,
    )

# ---------------------------------------------------------------------------
# Sidebar: Section 3 - date window
# ---------------------------------------------------------------------------
with st.sidebar.expander("3. Date window", expanded=True):
    selection_date = st.date_input("Reference date", value=date.today())
    history_days = st.number_input(
        "Days to look back", min_value=0, max_value=365, value=0, step=1
    )
    forward_view = st.number_input(
        "Days to look forward", min_value=0, max_value=365, value=10, step=1
    )

window = DateWindow(
    selection_date=selection_date,
    history_days=int(history_days),
    forward_days=int(forward_view),
)

# ---------------------------------------------------------------------------
# Sidebar: Section 4 - bulk email prep (Power Automate export settings)
# ---------------------------------------------------------------------------
with st.sidebar.expander("4. Bulk email prep (Power Automate)", expanded=True):
    report_signature = st.text_input(
        "Sign off name",
        placeholder="e.g., Andrea, Yanni",
        help="Changes the trailing sign-off name at the bottom of the generated email."
    )

    email_subject = st.text_input("Email subject", value=DEFAULT_SUBJECT)

    person_email_file = st.file_uploader(
        "Person → Email mapping file",
        type=["csv", "xlsx", "xls"],
        key="person_email_file",
        help=(
            "One flat mapping covering BOTH agents and superiors. Must "
            "contain a 'Person' column and an 'Email' column. Used to look "
            "up each agent's own address, and to CC that agent's own "
            "superior automatically."
        ),
    )

    default_cc_input = st.text_area(
        "Default CC (comma-separated emails)",
        placeholder="e.g., ops@company.com, msme.team@company.com",
        help="Added to every row's CC in addition to that agent's own superior.",
    )
    default_cc_list = [e.strip() for e in default_cc_input.split(",") if e.strip()]


# ---------------------------------------------------------------------------
# Build the base output table (filters + visit join + dtypes), independent
# of sort/column-visibility choices made inside the Table View tab.
# ---------------------------------------------------------------------------
person_email_map = {}
if person_email_file is not None:
    try:
        person_email_map = load_person_email_map(person_email_file)
    except Exception as exc:
        st.sidebar.error(f"⚠️ Could not read person email mapping: {exc}")

try:
    filtered_data = filter_reference_data(reference_data, dss_filter, agents_filter, window)
    result = attach_visit_status(filtered_data, form_data)
    result_clean = clean_output_columns(result)
    result_clean = filter_by_visit_status(result_clean, visit_inclusion)
    result_clean = coerce_display_dtypes(result_clean)
except DataLoadError as exc:
    st.error(f"⚠️ {exc}")
    st.stop()
except KeyError as exc:
    st.error(f"⚠️ Missing expected column while processing form data: {exc}")
    st.stop()

# ---------------------------------------------------------------------------
# Main panel: tabs (order: Table View, Email-Ready View, Statistics & Options)
# ---------------------------------------------------------------------------
tab_table, tab_email, tab_stats, tab_bulk = st.tabs(
    ["📋 Table View", "✉️ Email-Ready View", "📊 Statistics & Options", "📨 Bulk Email Tracker"]
)

# --- Table View tab: column/sort controls are defined here first, since --
# --- both the Table View and Email-Ready View need the resulting frame. --
with tab_table:
    st.subheader("Table view")

    if result_clean.empty:
        st.warning("No records match the selected filters and date window.")
        display_df = result_clean
        report_wb = report_bytes = report_html = None
    else:
        show_superior = st.checkbox("Show 'Superior' column", value=False)
        show_agent = st.checkbox("Show 'Agent' column", value=True)
        display_df = toggle_columns(result_clean, include_superior=show_superior, include_agent=show_agent)

        sort_col_options = list(display_df.columns)
        default_sort_idx = sort_col_options.index(CLEAN_DAYS_TO_INST_COL) \
            if CLEAN_DAYS_TO_INST_COL in sort_col_options else 0

        sc1, sc2 = st.columns([3, 1])
        sort_col = sc1.selectbox("Sort by", sort_col_options, index=default_sort_idx, key="sort_col")
        sort_dir = sc2.radio("Order", ["Ascending", "Descending"], index=0, key="sort_dir")
        ascending = sort_dir == "Ascending"

        display_df = sort_output(display_df, sort_col, ascending)

        st.dataframe(style_output_table(display_df), use_container_width=True, hide_index=True)

        report_wb, report_bytes, report_html = build_visit_report_workbook(
            display_df#, report_date=selection_date
        )
        st.download_button(
            "⬇️ Download table as Excel (.xlsx)",
            data=report_bytes,
            file_name=f"visit_views_DSS[{dss_filter}]_agent[{agents_filter}].xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# --- Email-Ready View tab: mirrors the sort/column choices from Table View -
with tab_email:
    st.subheader("Email-ready view")


    if result_clean.empty or display_df.empty:
        st.warning("No records to show.")
    else:
        # components.html(report_html, height=600, scrolling=True)
        email_html = table_to_email(report_html, agent=agents_filter, report_signature=report_signature)

        st.components.v1.html(email_html, height=600, scrolling=True)

        st.download_button(
            label=f"📥 Download: visit_views_DSS[{dss_filter}]_agent[{agents_filter}].html",
            data=report_html,
            file_name=f"visit_views_DSS[{dss_filter}]_agent[{agents_filter}].html",
            mime="text/html",
            use_container_width=True
        )

# --- Statistics & Options tab: date-window transparency + summary metrics -
with tab_stats:
    st.subheader("Selected date window")
    col1, col2, col3 = st.columns(3)
    col1.metric("Start date", window.start_date.strftime("%Y-%m-%d"))
    col2.metric("Reference date", selection_date.strftime("%Y-%m-%d"))
    col3.metric("End date", window.end_date.strftime("%Y-%m-%d"))
    st.caption(
        f"Showing instalment dates from **{window.start_date.strftime('%Y-%m-%d')}** "
        f"to **{window.end_date.strftime('%Y-%m-%d')}** "
        f"(reference date minus {history_days} day(s), plus {forward_view} day(s))."
    )

    st.divider()
    st.subheader(f"Results ({len(result_clean)} record(s))")

    if result_clean.empty:
        st.warning("No records match the selected filters and date window.")
    else:
        visited_count = (result_clean[CLEAN_VISITED_COL] == "Yes").sum()
        not_visited_count = (result_clean[CLEAN_VISITED_COL] == "No").sum()
        due_soon_red = result_clean[CLEAN_DAYS_TO_INST_COL].between(0, 2).sum()
        due_soon_yellow = result_clean[CLEAN_DAYS_TO_INST_COL].between(3, 5).sum()

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total accounts", len(result_clean))
        m2.metric("Visited", int(visited_count))
        m3.metric("Not visited", int(not_visited_count))
        m4.metric("🔴 Due in 0-2 days", int(due_soon_red))
        m5.metric("🟡 Due in 3-5 days", int(due_soon_yellow))

        st.caption(
            "🔴 rows are due in 0-2 days, 🟡 rows are due in 3-5 days "
            "(see the Email-Ready View and Table View tabs). "
            f"'{CLEAN_SUPERIOR_COL}' is hidden from the Table View by default — "
            "toggle it on there if needed."
        )

# ---------------------------------------------------------------------------
# Bulk Email Tracker tab
# ---------------------------------------------------------------------------
with tab_bulk:
    st.subheader("Bulk email tracker")

    col_a, col_b = st.columns(2)
    bulk_show_superior = col_a.checkbox(
        "Show 'Superior' column (in emailed table)", value=False, key="bulk_show_superior"
    )
    bulk_show_agent = col_b.checkbox(
        "Show 'Agent' column (in emailed table)", value=True, key="bulk_show_agent"
    )

    # Everything that affects the tracker's contents. Any change to these
    # widgets automatically re-runs Streamlit's script top-to-bottom, so we
    # simply recompute the tracker whenever this signature changes — no
    # separate 'Prepare tracker' click is required.
    bulk_settings_sig = (
        window.start_date, window.end_date, visit_inclusion,
        tuple(sorted(person_email_map.items())),
        email_subject, tuple(default_cc_list), report_signature,
        bulk_show_superior, bulk_show_agent,
        id(reference_data), id(form_data),
    )

    if st.session_state.get("bulk_settings_sig") != bulk_settings_sig:
        st.session_state.bulk_jobs = build_tracker(
            reference_data=reference_data,
            form_data=form_data,
            dss_filter="All",
            window=window,
            visit_inclusion=visit_inclusion,
            person_email_map=person_email_map,
            subject=email_subject,
            default_cc_list=default_cc_list,
            report_signature=report_signature,
            include_superior=bulk_show_superior,
            include_agent=bulk_show_agent,
        )
        st.session_state.bulk_settings_sig = bulk_settings_sig
        # Any previously edited export table is now stale.
        st.session_state.pop("bulk_export_df", None)

    jobs = st.session_state.get("bulk_jobs", [])

    st.dataframe(
        tracker_to_dataframe(jobs),
        use_container_width=True,
        hide_index=True,
    )

    if jobs:
        ready = sum(1 for j in jobs if j.status == "Ready")
        missing = sum(1 for j in jobs if j.status == "Missing email")
        s1, s2, s3 = st.columns(3)
        s1.metric("Total agents", len(jobs))
        s2.metric("Ready for export", ready)
        s3.metric("⚠️ Missing email", missing)

        st.divider()
        st.markdown("**Export queue for Power Automate**")
        st.caption(
            "Edit any cell below if needed (e.g. tweak a Subject or CC) — "
            "your edits are reflected in the downloaded file."
        )

        base_export_df = build_power_automate_export(jobs)

        # Preserve edits across reruns unless the underlying tracker changed.
        if "bulk_export_df" not in st.session_state:
            st.session_state.bulk_export_df = base_export_df

        edited_export_df = st.data_editor(
            st.session_state.bulk_export_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            key="bulk_export_editor",
            column_config={
                "HTMLBody": st.column_config.TextColumn(
                    "HTMLBody", help="Full HTML email body (long — scroll to see)."
                ),
            },
        )
        st.session_state.bulk_export_df = edited_export_df

        export_bytes = export_power_automate_excel(edited_export_df)
        timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M")
        queue_filename = f"{timestamp} AgentEmailQueue.xlsx"
        st.download_button(
            f"⬇️ Download {queue_filename}",
            data=export_bytes,
            file_name=queue_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(
            "Drop this file into the OneDrive/SharePoint folder your "
            "Power Automate flow watches. The flow sends each row via "
            "your own Outlook mailbox, using the To/CC/Subject/HTMLBody "
            "columns above."
        )
    else:
        st.info("No agents with rows found for the current date window / row-inclusion setting.")
        
st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        text-align: center;
        padding: 10px;
        font-size: 14px;
    }
    </style>
    <div class="footer">
        © July 2026, by Deanne Algenio
    </div>
    """,
    unsafe_allow_html=True
)