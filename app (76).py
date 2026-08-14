import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
import json
import os

# ---------- CONFIG ----------
SHEET_ID = "1uHRV6X1xYkid9XhUYULK3xrDBFVHW7oStMaLi7oY2es"
WORKSHEET_NAME = "Sheet1"  # change if your tab is named differently

COLUMNS = [
    "VIN", "Make", "Model", "Year", "License Plate",
    "Owner Name", "Owner Contact", "Mileage",
    "Event Date", "Event Type", "Description", "Cost", "Status",
]

EVENT_TYPES = ["Service", "Repair", "Accident", "Recall", "Ownership Change", "Inspection", "Other"]
STATUSES = ["Open", "Completed", "Pending", "Cancelled"]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

st.set_page_config(page_title="Vehicle History Monitoring", page_icon="🚗", layout="wide")


# ---------- AUTH / CONNECTION ----------
@st.cache_resource(show_spinner="Connecting to Google Sheets...")
def get_client():
    """
    Loads service account credentials from (in order of priority):
      1. Streamlit secrets: st.secrets["gcp_service_account"]
      2. A local file named service_account.json next to this script
    Never hardcode key contents directly in this file.
    """
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        local_path = os.path.join(os.path.dirname(__file__), "service_account.json")
        if not os.path.exists(local_path):
            st.error(
                "No credentials found. Add a `service_account.json` file next to app.py, "
                "or configure `gcp_service_account` in Streamlit secrets."
            )
            st.stop()
        creds = Credentials.from_service_account_file(local_path, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def get_worksheet():
    client = get_client()
    sh = client.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=len(COLUMNS))
    if ws.row_values(1) != COLUMNS:
        ws.update("A1", [COLUMNS])
    return ws


def load_data():
    ws = get_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records, columns=COLUMNS)
    return df


def append_record(record: dict):
    ws = get_worksheet()
    row = [record.get(col, "") for col in COLUMNS]
    ws.append_row(row, value_input_option="USER_ENTERED")


def update_row(row_index_in_sheet: int, record: dict):
    """row_index_in_sheet is 1-based including header (so data row 1 -> sheet row 2)."""
    ws = get_worksheet()
    row = [record.get(col, "") for col in COLUMNS]
    ws.update(f"A{row_index_in_sheet}:{chr(64 + len(COLUMNS))}{row_index_in_sheet}", [row])


def delete_row(row_index_in_sheet: int):
    ws = get_worksheet()
    ws.delete_rows(row_index_in_sheet)


# ---------- UI ----------
st.title("🚗 Vehicle History Monitoring System")

page = st.sidebar.radio("Navigate", ["➕ Add Entry", "📋 View / Search History", "✏️ Edit / Delete"])

if page == "➕ Add Entry":
    st.subheader("Add a new vehicle history entry")
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
            cost = st.number_input("Cost ($)", min_value=0.0, step=0.01, format="%.2f")

        c4, c5, c6 = st.columns(3)
        with c4:
            event_date = st.date_input("Event Date", value=date.today())
        with c5:
            event_type = st.selectbox("Event Type", EVENT_TYPES)
        with c6:
            status = st.selectbox("Status", STATUSES)

        description = st.text_area("Description / Notes")

        submitted = st.form_submit_button("Save Entry")
        if submitted:
            if not vin.strip():
                st.error("VIN is required.")
            else:
                record = {
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
                }
                append_record(record)
                st.success(f"Saved entry for VIN {record['VIN']}.")
                st.cache_data.clear()

elif page == "📋 View / Search History":
    st.subheader("Vehicle history records")
    df = load_data()

    if df.empty:
        st.info("No records yet. Add one from the sidebar.")
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
        m3.metric("Total Cost ($)", f"{pd.to_numeric(filtered['Cost'], errors='coerce').sum():,.2f}" if not filtered.empty else "0.00")

        st.dataframe(filtered, use_container_width=True, hide_index=True)

        st.download_button(
            "Download filtered results as CSV",
            filtered.to_csv(index=False).encode("utf-8"),
            file_name="vehicle_history_export.csv",
            mime="text/csv",
        )

elif page == "✏️ Edit / Delete":
    st.subheader("Edit or delete an existing record")
    df = load_data()

    if df.empty:
        st.info("No records yet.")
    else:
        df_display = df.copy()
        df_display.insert(0, "Row #", range(2, len(df_display) + 2))  # sheet row, header = row 1
        st.dataframe(df_display, use_container_width=True, hide_index=True)

        row_num = st.number_input(
            "Enter Row # to edit or delete", min_value=2, max_value=len(df) + 1, step=1
        )
        selected = df.iloc[row_num - 2]

        with st.form("edit_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                vin = st.text_input("VIN", value=selected["VIN"])
                make = st.text_input("Make", value=selected["Make"])
                model = st.text_input("Model", value=selected["Model"])
            with c2:
                year = st.number_input("Year", min_value=1900, max_value=2100, value=int(selected["Year"] or 2020))
                plate = st.text_input("License Plate", value=selected["License Plate"])
                owner_name = st.text_input("Owner Name", value=selected["Owner Name"])
            with c3:
                owner_contact = st.text_input("Owner Contact", value=selected["Owner Contact"])
                mileage = st.number_input("Mileage", min_value=0, value=int(selected["Mileage"] or 0))
                cost = st.number_input("Cost ($)", min_value=0.0, value=float(selected["Cost"] or 0.0), format="%.2f")

            event_type = st.selectbox("Event Type", EVENT_TYPES, index=EVENT_TYPES.index(selected["Event Type"]) if selected["Event Type"] in EVENT_TYPES else 0)
            status = st.selectbox("Status", STATUSES, index=STATUSES.index(selected["Status"]) if selected["Status"] in STATUSES else 0)
            description = st.text_area("Description / Notes", value=selected["Description"])

            b1, b2 = st.columns(2)
            update_clicked = b1.form_submit_button("💾 Update Record")
            delete_clicked = b2.form_submit_button("🗑️ Delete Record")

            if update_clicked:
                record = {
                    "VIN": vin.strip().upper(), "Make": make.strip(), "Model": model.strip(),
                    "Year": int(year), "License Plate": plate.strip().upper(),
                    "Owner Name": owner_name.strip(), "Owner Contact": owner_contact.strip(),
                    "Mileage": int(mileage), "Event Date": selected["Event Date"],
                    "Event Type": event_type, "Description": description.strip(),
                    "Cost": float(cost), "Status": status,
                }
                update_row(int(row_num), record)
                st.success("Record updated.")
                st.cache_data.clear()

            if delete_clicked:
                delete_row(int(row_num))
                st.success("Record deleted.")
                st.cache_data.clear()
