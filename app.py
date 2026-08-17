import time
from datetime import date

import pandas as pd
import streamlit as st

import gspread
from google.oauth2.service_account import Credentials

# =========================================================================
# CONFIG
# =========================================================================
st.set_page_config(page_title="Vehicle History Monitoring System - KMN", layout="wide", page_icon="🚗")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Falls back to this Sheet if [gsheet] sheet_id isn't set in secrets.
DEFAULT_SHEET_ID = "1uHRV6X1xYkid9XhUYULK3xrDBFVHW7oStMaLi7oY2es"
DEFAULT_WORKSHEET_NAME = "VehicleHistory"
DEFAULT_VEHICLE_SHEET_NAME = "Sheet2"

EVENT_TYPES = ["Service", "Repair", "Accident", "Recall", "Inspection", "Other"]

# NOTE: "User" added to record who entered the record (from the logged-in session).
COLUMN_ORDER = [
    "Timestamp", "Vehicle No", "Vehicle Type", "Date", "Mileage (KM)",
    "Event Type", "Place", "Description of Goods / Service", "Cost", "User",
]

# Login credentials (only two users use this app)
USERS = {
    "Narmada": "narmada123",
    "Dilantha": "dilantha",
}

# Background image used behind the blurred login card (auto-maintenance themed).
LOGIN_BG_IMAGE_URL = "https://images.unsplash.com/photo-1503736334956-4c8f8e92946d?auto=format&fit=crop&w=1950&q=80"

# =========================================================================
# SESSION STATE INIT
# =========================================================================
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""

# =========================================================================
# STYLE
# =========================================================================
def inject_base_style():
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


def inject_login_style():
    st.markdown(f"""
    <style>
    [data-testid="stAppViewContainer"] {{
        background-image: url('{LOGIN_BG_IMAGE_URL}');
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }}
    [data-testid="stAppViewContainer"]::before {{
        content: "";
        position: fixed;
        inset: 0;
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        background: rgba(10, 15, 20, 0.55);
        z-index: 0;
    }}
    [data-testid="stHeader"],
    [data-testid="stToolbar"],
    [data-testid="stDecoration"],
    [data-testid="stStatusWidget"],
    #MainMenu,
    header {{
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
    }}
    [data-testid="stAppViewContainer"] > .main {{
        padding-top: 0 !important;
    }}
    /* Neutralize any transform on Streamlit's own wrapper elements — a
       transform on an ancestor silently changes what position:fixed
       centers against, which is what was throwing off centering. */
    html, body, .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stMain"],
    .main, .block-container {{
        transform: none !important;
        filter: none !important;
    }}
    .block-container {{
        padding-top: 2rem !important;
    }}
    .login-title {{
        text-align: center !important;
        width: 100%;
        font-weight: 800;
        font-size: clamp(0.95rem, 3.4vw, 1.3rem);
        white-space: nowrap;
        margin-bottom: 0.7rem;
        color: #c1121f;
        letter-spacing: 0.2px;
    }}
    div[data-testid="stForm"] .stMarkdown {{
        width: 100%;
    }}
    div[data-testid="stForm"] {{
        position: fixed !important;
        top: 50vh !important;
        left: 50vw !important;
        transform: translate(-50%, -50%) !important;
        margin: 0 !important;
        z-index: 2;
        width: 90%;
        max-width: 640px;
        height: auto !important;
        min-height: 0 !important;
        max-height: 90vh;
        overflow-y: auto;
        border: 1px solid #d6d6d6;
        border-radius: 0.8rem;
        padding: 1.9rem 2.2rem 1.3rem 2.2rem;
        background: rgba(255, 255, 255, 0.94);
        box-shadow: 0 12px 40px rgba(0,0,0,0.35);
    }}
    div[data-testid="stForm"] div {{
        height: auto !important;
        min-height: 0 !important;
        flex-grow: 0 !important;
    }}
    div[data-testid="stForm"] > div {{
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
    }}
    div[data-testid="stForm"] .stTextInput input {{
        border: 2px solid #000000 !important;
        border-radius: 0.75rem !important;
    }}
    div[data-testid="stForm"] .stTextInput,
    div[data-testid="stForm"] .stFormSubmitButton {{
        width: 100%;
    }}
    </style>
    """, unsafe_allow_html=True)


# =========================================================================
# GOOGLE SHEETS BACKEND
# =========================================================================
def _gsheet_configured():
    return "gcp_service_account" in st.secrets


def _sheet_id():
    return st.secrets.get("gsheet", {}).get("sheet_id", DEFAULT_SHEET_ID)


def _worksheet_name():
    return st.secrets.get("gsheet", {}).get("worksheet_name", DEFAULT_WORKSHEET_NAME)


def _vehicle_worksheet_name():
    return st.secrets.get("gsheet", {}).get("vehicle_worksheet_name", DEFAULT_VEHICLE_SHEET_NAME)


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
def get_spreadsheet():
    creds_dict = _normalize_private_key(dict(st.secrets["gcp_service_account"]))
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(_sheet_id())


@st.cache_resource(show_spinner=False)
def get_worksheet():
    sh = get_spreadsheet()
    worksheet_name = _worksheet_name()
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows=2000, cols=len(COLUMN_ORDER) + 2)
        ws.append_row(COLUMN_ORDER, value_input_option="USER_ENTERED")
    header = ws.row_values(1)
    if header != COLUMN_ORDER:
        # gspread's Worksheet.update() signature was swapped between major
        # versions — old versions accept (range_name, values), newer 6.x
        # versions accept (values, range_name). Calling it positionally as
        # update("A1", [COLUMN_ORDER]) means that on a newer gspread
        # install, "A1" gets treated as the *values* (and iterated
        # character-by-character) while [COLUMN_ORDER] gets treated as the
        # *range* — silently wiping/garbling the header row instead of
        # writing it. Using explicit keyword args makes it version-proof.
        ws.update(range_name="A1", values=[COLUMN_ORDER])
    return ws


@st.cache_resource(show_spinner=False)
def get_vehicle_worksheet():
    sh = get_spreadsheet()
    return sh.worksheet(_vehicle_worksheet_name())


def clear_data_caches():
    """Invalidates all cached reads so the next load fetches fresh data
    from the sheet — called after a write, and from the manual refresh
    button on the View page."""
    _load_data_cached.clear()
    load_vehicle_details.clear()


@st.cache_data(show_spinner=False, ttl=20)
def _load_data_cached(sheet_id):
    ws = get_worksheet()
    # get_all_values() reads the raw grid and never raises on duplicate or
    # blank header cells the way get_all_records() does, so it's used here
    # instead — the sheet's header row is already enforced to COLUMN_ORDER
    # by get_worksheet(), so records are mapped back to it by position.
    values = ws.get_all_values()
    n = len(COLUMN_ORDER)
    if len(values) <= 1:
        return pd.DataFrame(columns=COLUMN_ORDER)
    data_rows = values[1:]
    padded_rows = [(row + [""] * n)[:n] for row in data_rows]
    df = pd.DataFrame(padded_rows, columns=COLUMN_ORDER)
    df = df.astype(str).replace("nan", "")
    return df


def load_data():
    try:
        return _load_data_cached(_sheet_id())
    except Exception as e:
        st.error(f"❌ Could not read data from the Google Sheet.\n\n{e}")
        return pd.DataFrame(columns=COLUMN_ORDER)


def append_record(record: dict):
    ws = get_worksheet()
    row = [str(record.get(c, "")) for c in COLUMN_ORDER]
    ws.append_row(row, value_input_option="USER_ENTERED")
    _load_data_cached.clear()


def _find_column_index(headers, *candidates):
    """Case/whitespace-insensitive match of a header name to its column index."""
    norm = {}
    for i, h in enumerate(headers):
        key = str(h).strip().lower()
        if key and key not in norm:  # keep the first occurrence of duplicates
            norm[key] = i
    for cand in candidates:
        key = cand.strip().lower()
        if key in norm:
            return norm[key]
    return None


@st.cache_data(show_spinner=False, ttl=60)
def load_vehicle_details():
    """Returns {vehicle_no: {"Vehicle Type": ..., "Rider": ..., "Brand": ...}}
    sourced from Sheet2. Uses raw grid values (not get_all_records()) so
    duplicate or blank header cells in that tab can't crash the app."""
    ws = get_vehicle_worksheet()
    values = ws.get_all_values()
    if not values:
        return {}
    header = values[0]
    vno_idx = _find_column_index(header, "Vehicle No", "Vehicle No.", "VehicleNo", "Vehicle Number")
    vtype_idx = _find_column_index(header, "Vehicle Type", "VehicleType", "Type")
    rider_idx = _find_column_index(header, "Rider", "Rider Name", "Driver", "Driver Name")
    brand_idx = _find_column_index(header, "Brand", "Vehicle Brand", "Make")
    if vno_idx is None:
        return {}

    def _cell(row, idx):
        return row[idx].strip() if (idx is not None and idx < len(row)) else ""

    details = {}
    for row in values[1:]:
        vno = _cell(row, vno_idx)
        if not vno:
            continue
        details[vno] = {
            "Vehicle Type": _cell(row, vtype_idx),
            "Rider": _cell(row, rider_idx),
            "Brand": _cell(row, brand_idx),
        }
    return details


def to_number(value, as_int=False):
    value = str(value).strip()
    if value == "" or value.lower() == "nan":
        return 0 if as_int else 0.0
    try:
        return int(float(value)) if as_int else float(value)
    except ValueError:
        return 0 if as_int else 0.0


# =========================================================================
# LOGIN GATE
# =========================================================================
def render_login():
    inject_login_style()

    with st.form("login_form"):
        st.markdown('<div class="login-title">Welcome 👋 Vehicle History Monitoring System 🔧🚗</div>', unsafe_allow_html=True)
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("🔐 Login", type="primary", use_container_width=True)
        if submitted:
            if username in USERS and USERS[username] == password:
                st.session_state["authenticated"] = True
                st.session_state["username"] = username
                st.rerun()
            else:
                st.error("❌ Invalid username or password.")


if not st.session_state["authenticated"]:
    render_login()
    st.stop()

# =========================================================================
# MAIN APP (authenticated)
# =========================================================================
inject_base_style()

top_l, top_r = st.columns([5, 1])
with top_l:
    st.markdown("<h1 style='text-align: center;'>Vehicle History Monitoring System</h1>", unsafe_allow_html=True)
    st.subheader("KMN Automotive")
with top_r:
    st.write("")
    st.write(f"👤 {st.session_state['username']}")
    if st.button("Log out", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state["username"] = ""
        st.rerun()
st.markdown("---")

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
            "4. Make sure the Sheet has a tab named **Sheet2** with columns "
            "**Vehicle No**, **Vehicle Type**, **Rider**, and **Brand** listing every vehicle.\n"
            "5. Add the following to your app's Streamlit secrets:\n"
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
            'worksheet_name = "VehicleHistory"\n'
            'vehicle_worksheet_name = "Sheet2"\n',
            language="toml",
        )
    st.stop()

try:
    get_worksheet()
except Exception as e:
    st.error(f"❌ Could not connect to the Google Sheet. Check your secrets and sharing settings.\n\n{e}")
    st.stop()

try:
    vehicle_details = load_vehicle_details()
except Exception as e:
    vehicle_details = {}
    st.warning(
        f"⚠️ Could not read the vehicle list from **{_vehicle_worksheet_name()}**. "
        f"Make sure that tab exists with a 'Vehicle No' column.\n\n{e}"
    )

vehicle_map = {vno: info.get("Vehicle Type", "") for vno, info in vehicle_details.items()}

# =========================================================================
# PHASE SELECTOR
# =========================================================================
phase = st.radio(
    "Phase",
    ["📋 Record Entering", "📊 View", "🚙 Vehicle Details"],
    horizontal=True,
    label_visibility="collapsed",
)
st.markdown("---")

# =========================================================================
# PHASE 1 — RECORD ENTERING
# =========================================================================
if phase == "📋 Record Entering":
    st.markdown("#### 📋 Record Entering")

    vehicle_options = sorted(vehicle_map.keys())

    if not vehicle_options:
        st.info(
            f"No vehicles found in the **{_vehicle_worksheet_name()}** tab yet. "
            "Add vehicles there (with a 'Vehicle No' column) before creating records."
        )
    else:
        # Vehicle No lives outside the form so the Vehicle Type preview
        # updates live as soon as a vehicle is selected.
        vehicle_no = st.selectbox("Vehicle No *", vehicle_options, key="entry_vehicle_no")
        vehicle_type = vehicle_map.get(vehicle_no, "")
        st.text_input("Vehicle Type", value=vehicle_type, disabled=True)

        # Only Description and Cost get reset after a save — Vehicle No,
        # Date, Event Type, Mileage, and Place are meant to carry over into
        # the next entry, so they're cleared here (before the widgets below
        # are created) rather than via clear_on_submit, which would wipe
        # everything.
        if st.session_state.pop("_reset_desc_cost", False):
            st.session_state["entry_description"] = ""
            st.session_state["entry_cost"] = 0.0

        with st.form("add_entry_form", clear_on_submit=False):
            c1, c2 = st.columns(2)
            with c1:
                event_date = st.date_input("Date *", value=date.today(), key="entry_date")
                mileage = st.number_input("Mileage (KM) *", min_value=0, step=1, key="entry_mileage")
            with c2:
                event_type = st.selectbox("Event Type", EVENT_TYPES, key="entry_event_type")
                cost = st.number_input("Cost *", min_value=0.0, step=0.01, format="%.2f", key="entry_cost")

            place = st.text_input("Place *", key="entry_place")
            description = st.text_area("Description of Goods / Service *", key="entry_description")

            submitted = st.form_submit_button("💾 Save Entry", type="primary", use_container_width=True)
            if submitted:
                errors = []
                if not vehicle_no:
                    errors.append("Vehicle No is required.")
                if not place.strip():
                    errors.append("Place is required.")
                if not description.strip():
                    errors.append("Description of Goods / Service is required.")
                if mileage <= 0:
                    errors.append("Mileage (KM) is required.")
                if cost <= 0:
                    errors.append("Cost is required.")

                if errors:
                    for err in errors:
                        st.error(f"❌ {err}")
                else:
                    record = {
                        "Timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S.%f"),
                        "Vehicle No": vehicle_no,
                        "Vehicle Type": vehicle_type,
                        "Date": event_date.isoformat(),
                        "Mileage (KM)": int(mileage),
                        "Event Type": event_type,
                        "Place": place.strip(),
                        "Description of Goods / Service": description.strip(),
                        "Cost": float(cost),
                        # Who entered this record — pulled from the logged-in session,
                        # not user-editable.
                        "User": st.session_state["username"],
                    }
                    append_record(record)
                    st.success(f"✅ Saved entry for Vehicle No {vehicle_no}.")
                    time.sleep(1)
                    st.session_state["_reset_desc_cost"] = True
                    st.rerun()

# =========================================================================
# PHASE 2 — VIEW
# =========================================================================
elif phase == "📊 View":
    h_l, h_r = st.columns([5, 1])
    with h_l:
        st.markdown("#### 📊 View Vehicle History")
    with h_r:
        if st.button("🔄 Refresh", use_container_width=True):
            clear_data_caches()
            st.rerun()

    df = load_data()

    if df.empty:
        st.info("No records yet. Add one from the Record Entering section.")
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            vn_options = sorted([v for v in df["Vehicle No"].unique().tolist() if v])
            vn_filter = st.multiselect("Filter by VN", vn_options)
        with c2:
            type_filter = st.multiselect("Filter by Event Type", EVENT_TYPES)
        with c3:
            place_options = sorted([p for p in df["Place"].unique().tolist() if p])
            place_filter = st.multiselect("Filter by Place", place_options)

        filtered = df.copy()
        if vn_filter:
            filtered = filtered[filtered["Vehicle No"].isin(vn_filter)]
        if type_filter:
            filtered = filtered[filtered["Event Type"].isin(type_filter)]
        if place_filter:
            filtered = filtered[filtered["Place"].isin(place_filter)]

        display_cols = [c for c in COLUMN_ORDER if c != "Timestamp"]
        st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)

        total_amount = sum(to_number(c) for c in filtered["Cost"])
        st.markdown(
            f"<p style='text-align: right; font-weight: 700; font-size: 1.05rem;'>"
            f"Total Amount: {total_amount:,.2f}</p>",
            unsafe_allow_html=True,
        )

        # Build the export with a trailing Total row, and encode with a
        # UTF-8 BOM (utf-8-sig) — plain "utf-8" is what was making Sinhala
        # (and any other non-Latin) text show up as garbled characters when
        # opened in Excel, since Excel doesn't reliably auto-detect UTF-8
        # without the BOM.
        export_df = filtered[display_cols].copy()
        total_row = {col: "" for col in display_cols}
        total_row[display_cols[0]] = "Total"
        total_row["Cost"] = f"{total_amount:,.2f}"
        export_df = pd.concat([export_df, pd.DataFrame([total_row])], ignore_index=True)

        st.download_button(
            "⬇️ Download filtered results as CSV",
            export_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="vehicle_history_export.csv",
            mime="text/csv",
        )

# =========================================================================
# PHASE 3 — VEHICLE DETAILS
# =========================================================================
elif phase == "🚙 Vehicle Details":
    st.markdown("#### 🚙 Vehicle Details")

    detail_options = sorted(vehicle_details.keys())

    if not detail_options:
        st.info(
            f"No vehicles found in the **{_vehicle_worksheet_name()}** tab yet. "
            "Add vehicles there (with a 'Vehicle No' column) first."
        )
    else:
        selected_vno = st.selectbox("Vehicle No", detail_options, key="details_vehicle_no")
        info = vehicle_details.get(selected_vno, {})

        st.text_input("Vehicle Type", value=info.get("Vehicle Type", ""), disabled=True)
        st.text_input("Rider", value=info.get("Rider", ""), disabled=True)
        st.text_input("Brand", value=info.get("Brand", ""), disabled=True)

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>KMN Automotive - Vehicle History Monitoring System</p>",
    unsafe_allow_html=True,
)
