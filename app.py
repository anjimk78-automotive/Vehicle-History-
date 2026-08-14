import time
from datetime import date

import pandas as pd
import streamlit as st

import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

# =========================================================================
# CONFIG
# =========================================================================
st.set_page_config(page_title="Vehicle History Monitoring System - KMN", layout="wide", page_icon="🚗")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Falls back to this Sheet if [gsheet] sheet_id isn't set in secrets.
DEFAULT_SHEET_ID = "1uHRV6X1xYkid9XhUYULK3xrDBFVHW7oStMaLi7oY2es"
DEFAULT_WORKSHEET_NAME = "VehicleHistory"

EVENT_TYPES = ["Service", "Repair", "Accident", "Recall", "Ownership Change", "Inspection", "Other"]
STATUSES = ["Open", "Completed", "Pending", "Cancelled"]

COLUMN_ORDER = [
    "Timestamp", "VIN", "Make", "Model", "Year", "License Plate",
    "Owner Name", "Owner Contact", "Mileage",
    "Event Date", "Event Type", "Description", "Cost", "Status",
    "Deleted",
]

# =========================================================================
# STYLE (same treatment as the other KMN apps, so this looks/feels like
# the same product suite)
# =========================================================================
st.markdown("""
<style>
ul[role="listbox"], div[role="listbox"] {
    width: max-content !important;
    min-width: 220px !important;
    max-width: 92vw !important;
}
[role="option"] {
    width: auto !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    word-break: break-word !important;
}
[role="option"] * {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
}

button[kind="primary"], button[kind="primaryFormSubmit"] {
    background-color: #e63946 !important;
    border-color: #e63946 !important;
    color: #ffffff !important;
}
button[kind="primary"]:hover, button[kind="primaryFormSubmit"]:hover {
    background-color: #c1121f !important;
    border-color: #c1121f !important;
    color: #ffffff !important;
}

.status-saved {
    display: inline-block;
    width: 100%;
    text-align: center;
    background-color: #2a9d8f;
    color: #ffffff;
    font-weight: 600;
    padding: 0.45rem 0.6rem;
    border-radius: 0.5rem;
}

@media (max-width: 700px) {
    div[data-testid="stHorizontalBlock"] {
        flex-direction: column !important;
    }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        width: 100% !important;
        min-width: 100% !important;
    }
}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center;'>Vehicle History Monitoring System</h1>", unsafe_allow_html=True)
st.subheader("KMN Automotive")
st.markdown("---")


# =========================================================================
# GOOGLE SHEETS BACKEND
# =========================================================================
def _gsheet_configured():
    return "gcp_service_account" in st.secrets


def _sheet_id():
    return st.secrets.get("gsheet", {}).get("sheet_id", DEFAULT_SHEET_ID)


def _worksheet_name():
    return st.secrets.get("gsheet", {}).get("worksheet_name", DEFAULT_WORKSHEET_NAME)


def _normalize_private_key(creds_dict: dict) -> dict:
    """Defends against the #1 cause of 'Unable to load PEM file' errors:
    a private_key that was pasted into secrets.toml with literal two-
    character '\\n' text instead of real line breaks. If the key doesn't
    already contain real newlines, this converts any literal backslash-n
    sequences into actual newlines before handing it to google-auth."""
    creds_dict = dict(creds_dict)
    key = creds_dict.get("private_key", "")
    if isinstance(key, str) and "\n" not in key and "\\n" in key:
        key = key.replace("\\n", "\n")
    creds_dict["private_key"] = key
    return creds_dict


@st.cache_resource(show_spinner=False)
def get_worksheet():
    creds_dict = _normalize_private_key(dict(st.secrets["gcp_service_account"]))
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(_sheet_id())
    worksheet_name = _worksheet_name()
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows=2000, cols=len(COLUMN_ORDER) + 2)
        ws.append_row(COLUMN_ORDER, value_input_option="USER_ENTERED")
    header = ws.row_values(1)
    if header != COLUMN_ORDER:
        ws.update("A1", [COLUMN_ORDER])
    return ws


def bump_data_version():
    st.session_state["_data_version"] = st.session_state.get("_data_version", 0) + 1


def _load_data_cached(data_version, sheet_id):
    # No caching beyond this session-scoped version bump: always reflects
    # the latest Sheet contents, including edits made directly in Sheets.
    ws = get_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in COLUMN_ORDER:
        if c not in df.columns:
            df[c] = ""
    if len(df) > 0:
        df = df[COLUMN_ORDER]
    df = df.astype(str).replace("nan", "")
    return df


def load_data():
    """Returns all records with soft-deleted rows filtered out."""
    df = _load_data_cached(st.session_state.get("_data_version", 0), _sheet_id())
    if "Deleted" in df.columns:
        is_deleted = df["Deleted"].astype(str).str.strip().str.lower().isin(["yes", "true", "1"])
        df = df[~is_deleted].reset_index(drop=True)
    return df


def append_record(record: dict):
    ws = get_worksheet()
    row = [str(record.get(c, "")) for c in COLUMN_ORDER]
    ws.append_row(row, value_input_option="USER_ENTERED")
    bump_data_version()


def update_record_by_timestamp(timestamp, record: dict):
    ws = get_worksheet()
    cell = ws.find(str(timestamp), in_column=1)
    row = [str(record.get(c, "")) for c in COLUMN_ORDER]
    if cell:
        end_a1 = rowcol_to_a1(cell.row, len(COLUMN_ORDER))
        ws.update(f"A{cell.row}:{end_a1}", [row], value_input_option="USER_ENTERED")
    else:
        ws.append_row(row, value_input_option="USER_ENTERED")
    bump_data_version()


def mark_deleted_by_timestamp(timestamp):
    """Soft-delete: flag Deleted='Yes' rather than removing the row, same
    pattern as the other KMN apps — keeps a full audit trail in the Sheet."""
    ws = get_worksheet()
    cell = ws.find(str(timestamp), in_column=1)
    if cell:
        deleted_col = COLUMN_ORDER.index("Deleted") + 1
        ws.update_cell(cell.row, deleted_col, "Yes")
        bump_data_version()


def to_number(value, as_int=False):
    value = str(value).strip()
    if value == "" or value.lower() == "nan":
        return 0 if as_int else 0.0
    try:
        return int(float(value)) if as_int else float(value)
    except ValueError:
        return 0 if as_int else 0.0


# =========================================================================
# SETUP CHECK
# =========================================================================
if not _gsheet_configured():
    st.error("❌ Google Sheets is not configured yet.")
    with st.expander("⚙️ How to connect this app to a Google Sheet", expanded=True):
        st.markdown(
            "1. Create a Google Cloud project, enable the **Google Sheets API** and "
            "**Google Drive API**, and create a **Service Account**.\n"
            "2. Create a JSON key for that service account and copy its contents.\n"
            "3. Open the target Google Sheet and share it (Editor access) with the "
            "service account's `client_email` address.\n"
            "4. Add the following to your app's Streamlit secrets:\n"
        )
        st.code(
            '[gcp_service_account]\n'
            'type = "service_account"\n'
            'project_id = "..."\n'
            'private_key_id = "..."\n'
            'private_key = \'\'\'-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n\'\'\'\n'
            'client_email = "...@....iam.gserviceaccount.com"\n'
            'client_id = "..."\n'
            'auth_uri = "https://accounts.google.com/o/oauth2/auth"\n'
            'token_uri = "https://oauth2.googleapis.com/token"\n'
            'auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"\n'
            'client_x509_cert_url = "..."\n\n'
            '[gsheet]\n'
            'sheet_id = "the-id-from-the-sheet-url"\n'
            'worksheet_name = "VehicleHistory"\n',
            language="toml",
        )
    st.stop()

try:
    get_worksheet()
except Exception as e:
    st.error(f"❌ Could not connect to the Google Sheet. Check your secrets and sharing settings.\n\n{e}")
    st.stop()

# =========================================================================
# PHASE SELECTOR — Enter / Edit & Delete / View, all in one interface
# =========================================================================
phase = st.radio(
    "Phase",
    ["📋 Enter Vehicle Record", "✏️ Edit / Delete Records", "📊 View / Search History"],
    horizontal=True,
    label_visibility="collapsed",
)
st.markdown("---")

# =========================================================================
# PHASE 1 — ENTER
# =========================================================================
if phase == "📋 Enter Vehicle Record":
    st.markdown("#### 📋 Enter Vehicle Record")

    with st.form("add_entry_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            vin = st.text_input("VIN *")
            make = st.text_input("Make")
            model = st.text_input("Model")
        with c2:
            year = st.number_input("Year", min_value=1900, max_value=2100, value=2020, step=1)
            plate = st.text_input("License Plate")
            owner_name = st.text_input("Owner Name")
        with c3:
            owner_contact = st.text_input("Owner Contact (phone/email)")
            mileage = st.number_input("Mileage", min_value=0, step=1)
            cost = st.number_input("Cost", min_value=0.0, step=0.01, format="%.2f")

        c4, c5, c6 = st.columns(3)
        with c4:
            event_date = st.date_input("Event Date", value=date.today())
        with c5:
            event_type = st.selectbox("Event Type", EVENT_TYPES)
        with c6:
            status = st.selectbox("Status", STATUSES)

        description = st.text_area("Description / Notes")

        submitted = st.form_submit_button("💾 Save Entry", type="primary", use_container_width=True)
        if submitted:
            if not vin.strip():
                st.error("❌ VIN is required.")
            else:
                record = {
                    "Timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S.%f"),
                    "VIN": vin.strip().upper(),
                    "Make": make.strip(),
                    "Model": model.strip(),
                    "Year": int(year),
                    "License Plate": plate.strip().upper(),
                    "Owner Name": owner_name.strip(),
                    "Owner Contact": owner_contact.strip(),
                    "Mileage": int(mileage),
                    "Event Date": event_date.isoformat(),
                    "Event Type": event_type,
                    "Description": description.strip(),
                    "Cost": float(cost),
                    "Status": status,
                    "Deleted": "",
                }
                append_record(record)
                st.success(f"✅ Saved entry for VIN {record['VIN']}.")
                time.sleep(1)
                st.rerun()

# =========================================================================
# PHASE 2 — EDIT / DELETE
# =========================================================================
elif phase == "✏️ Edit / Delete Records":
    st.markdown("#### ✏️ Edit / Delete Records")

    df = load_data()

    if df.empty:
        st.info("No records yet. Add one from the Enter phase.")
    else:
        vin_options = sorted([v for v in df["VIN"].unique().tolist() if v])
        search_vin = st.selectbox("Select a VIN to edit or delete", vin_options)

        vin_rows = df[df["VIN"] == search_vin].reset_index(drop=True)
        st.caption(f"{len(vin_rows)} record(s) on file for VIN **{search_vin}**.")

        vin_rows_display = vin_rows.drop(columns=["Timestamp", "Deleted"]).copy()
        vin_rows_display.insert(0, "#", range(1, len(vin_rows_display) + 1))
        st.dataframe(vin_rows_display, use_container_width=True, hide_index=True)

        row_choice = st.selectbox(
            "Select the record # to edit/delete (see # column above — same-day events can repeat, so pick by number, not date)",
            vin_rows.index,
            format_func=lambda i: f"# {i + 1}",
        )
        selected = vin_rows.loc[row_choice]

        st.markdown("<span class='status-saved'>✅ Saved record — editing below</span>", unsafe_allow_html=True)

        with st.form("edit_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                vin_e = st.text_input("VIN", value=selected["VIN"])
                make_e = st.text_input("Make", value=selected["Make"])
                model_e = st.text_input("Model", value=selected["Model"])
            with c2:
                year_e = st.number_input("Year", min_value=1900, max_value=2100,
                                          value=int(to_number(selected["Year"], as_int=True)) or 2020)
                plate_e = st.text_input("License Plate", value=selected["License Plate"])
                owner_name_e = st.text_input("Owner Name", value=selected["Owner Name"])
            with c3:
                owner_contact_e = st.text_input("Owner Contact", value=selected["Owner Contact"])
                mileage_e = st.number_input("Mileage", min_value=0, value=int(to_number(selected["Mileage"], as_int=True)))
                cost_e = st.number_input("Cost", min_value=0.0, value=to_number(selected["Cost"]), format="%.2f")

            c4, c5 = st.columns(2)
            with c4:
                event_type_e = st.selectbox(
                    "Event Type", EVENT_TYPES,
                    index=EVENT_TYPES.index(selected["Event Type"]) if selected["Event Type"] in EVENT_TYPES else 0,
                )
            with c5:
                status_e = st.selectbox(
                    "Status", STATUSES,
                    index=STATUSES.index(selected["Status"]) if selected["Status"] in STATUSES else 0,
                )
            description_e = st.text_area("Description / Notes", value=selected["Description"])

            b1, b2 = st.columns(2)
            update_clicked = b1.form_submit_button("💾 Update Record", type="primary", use_container_width=True)
            delete_clicked = b2.form_submit_button("🗑️ Delete Record", use_container_width=True)

            if update_clicked:
                record = {
                    "Timestamp": selected["Timestamp"],
                    "VIN": vin_e.strip().upper(), "Make": make_e.strip(), "Model": model_e.strip(),
                    "Year": int(year_e), "License Plate": plate_e.strip().upper(),
                    "Owner Name": owner_name_e.strip(), "Owner Contact": owner_contact_e.strip(),
                    "Mileage": int(mileage_e), "Event Date": selected["Event Date"],
                    "Event Type": event_type_e, "Description": description_e.strip(),
                    "Cost": float(cost_e), "Status": status_e, "Deleted": "",
                }
                update_record_by_timestamp(selected["Timestamp"], record)
                st.success("✅ Record updated.")
                time.sleep(1)
                st.rerun()

            if delete_clicked:
                mark_deleted_by_timestamp(selected["Timestamp"])
                st.success("✅ Record deleted.")
                time.sleep(1)
                st.rerun()

# =========================================================================
# PHASE 3 — VIEW / SEARCH
# =========================================================================
elif phase == "📊 View / Search History":
    st.markdown("#### 📊 View / Search Vehicle History")

    df = load_data()

    if df.empty:
        st.info("No records yet. Add one from the Enter phase.")
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            vin_filter = st.text_input("Filter by VIN contains")
        with c2:
            type_filter = st.multiselect("Filter by Event Type", EVENT_TYPES)
        with c3:
            status_filter = st.multiselect("Filter by Status", STATUSES)

        filtered = df.copy()
        if vin_filter:
            filtered = filtered[filtered["VIN"].str.contains(vin_filter, case=False, na=False)]
        if type_filter:
            filtered = filtered[filtered["Event Type"].isin(type_filter)]
        if status_filter:
            filtered = filtered[filtered["Status"].isin(status_filter)]

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Records", len(filtered))
        m2.metric("Unique Vehicles", filtered["VIN"].nunique() if not filtered.empty else 0)
        total_cost = pd.to_numeric(filtered["Cost"], errors="coerce").sum() if not filtered.empty else 0
        m3.metric("Total Cost", f"{total_cost:,.2f}")

        display_cols = [c for c in COLUMN_ORDER if c not in ("Timestamp", "Deleted")]
        st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Download filtered results as CSV",
            filtered[display_cols].to_csv(index=False).encode("utf-8"),
            file_name="vehicle_history_export.csv",
            mime="text/csv",
        )

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>KMN Automotive - Vehicle History Monitoring System</p>",
    unsafe_allow_html=True,
)
