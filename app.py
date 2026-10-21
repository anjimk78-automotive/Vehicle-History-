import base64
import html
import inspect
import io
import os
import re
import csv
import time
from datetime import date

import altair as alt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import gspread
from google.oauth2.service_account import Credentials

# PDF export for the Serviced Report. Imported defensively so the rest of the
# app keeps working even if reportlab isn't installed yet (add "reportlab" to
# requirements.txt).
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas
    _REPORTLAB_AVAILABLE = True
except ImportError:
    _REPORTLAB_AVAILABLE = False

# =========================================================================
# CONFIG
# =========================================================================
st.set_page_config(page_title="Vehicle History Monitoring System - KMN", layout="wide", page_icon="🚗", initial_sidebar_state="expanded")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Falls back to this Sheet if [gsheet] sheet_id isn't set in secrets.
DEFAULT_SHEET_ID = "1uHRV6X1xYkid9XhUYULK3xrDBFVHW7oStMaLi7oY2es"
DEFAULT_WORKSHEET_NAME = "VehicleHistory"
DEFAULT_VEHICLE_SHEET_NAME = "Sheet2"
# NOTE: a NEW tab name is used for the redesigned Weekly Mileage Inspection
# (one row per mileage entry). The old "WeeklyMileage" tab used a different
# column layout, so it is left untouched instead of being overwritten.
DEFAULT_MILEAGE_LOG_SHEET_NAME = "MileageLog"

# Storage columns for the Weekly Mileage Inspection sheet: one row per entry.
# "Oil Type" is appended last so existing rows/columns never shift position.
MILEAGE_LOG_COLUMNS = ["Timestamp", "Vehicle No", "Vehicle Type", "Mileage Type", "Date", "Mileage (KM)", "User", "Oil Type"]
MILEAGE_TYPES = ["Weekly Mileage", "Serviced Mileage"]

# Oil Type options and how many KM after a service the next one is due.
OIL_SERVICE_INTERVALS = {"Castrol": 5000, "Bajaj Genuine": 2500, "Yamaha Genuine":2000, "Caltex Havoline 20W 50":2500}
OIL_TYPES = list(OIL_SERVICE_INTERVALS.keys())

# Lorries that service every 5000 km. Every other Lorry services every 6000 km.
# Vehicle Nos are compared ignoring spaces and letter case.
LORRY_5000_KM_VEHICLES = {"LB8797", "227-4969"}


def _normalize_vno(vno):
    return re.sub(r"\s+", "", str(vno)).upper()


EVENT_TYPES = ["Service", "Repair", "Accident", "Recall", "Inspection", "Other"]

# Vehicle Part options, dependent on the selected Event Type.
REPAIR_PARTS = [
    "Body", "Tires", "Battery", "AC", "Engine", "Suspension Repairs", "Electric Repairs", "Others", "Belt Changing", "Chasis"
]
SERVICE_PARTS = ["Oil Changing", "AC/Air Filter", "Tune Up & Additive", "Other"]


def vehicle_part_options(event_type):
    if event_type == "Repair":
        return REPAIR_PARTS
    if event_type == "Service":
        return SERVICE_PARTS
    return ["N/A"]


# NOTE: "User" added to record who entered the record (from the logged-in
# session). "Vehicle Part" is appended last so existing rows/columns never
# shift position — the View page reorders columns for display separately
# (see VIEW_DISPLAY_COLUMNS below) without touching this storage order.
# "Edited Date" and "Edited By" are appended after that: they stay blank
# until a record is changed from the Edit Entry section, which then fills in
# when it was edited and which logged-in user edited it.
COLUMN_ORDER = [
    "Timestamp", "Vehicle No", "Vehicle Type", "Date", "Mileage (KM)",
    "Event Type", "Place", "Description of Goods / Service", "Cost", "User",
    "Vehicle Part", "Edited Date", "Edited By",
]

# Display-only column order for the View page and Vehicle Details' service
# history — Vehicle Part shows as the 6th column here (right after Event
# Type) even though it's stored as the last column in the sheet.
VIEW_DISPLAY_COLUMNS = [
    "Vehicle No", "Vehicle Type", "Date", "Mileage (KM)", "Event Type",
    "Vehicle Part", "Place", "Description of Goods / Service", "Cost", "User",
]

# Login credentials come from Streamlit secrets (the [users] section), not
# from the code. Add them in the app's Secrets like this:
#   [users]
#   Username = "password"
# If the section is missing, nobody can log in.
try:
    USERS = {str(k): str(v) for k, v in st.secrets["users"].items()}
except Exception:
    USERS = {}

# Background image used behind the blurred login card (auto-maintenance themed).
LOGIN_BG_IMAGE_URL = "https://images.unsplash.com/photo-1520340356584-f9917d1eea6f?auto=format&fit=crop&w=1950&q=80"

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


def _mileage_log_worksheet_name():
    return st.secrets.get("gsheet", {}).get("mileage_log_worksheet_name", DEFAULT_MILEAGE_LOG_SHEET_NAME)


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
    # If this worksheet was created back when COLUMN_ORDER had fewer
    # columns (e.g. before "Vehicle Part" existed), it may not physically
    # have enough columns for the current schema — grow it if needed
    # before writing/reading, or a newly added column can silently fail
    # to hold data.
    if ws.col_count < len(COLUMN_ORDER):
        ws.resize(cols=len(COLUMN_ORDER))
    header = ws.row_values(1)
    if header != COLUMN_ORDER:
        # gspread's Worksheet.update() signature was swapped between major
        # versions — explicit keyword args make this version-proof.
        ws.update(range_name="A1", values=[COLUMN_ORDER])
    return ws


@st.cache_resource(show_spinner=False)
def get_vehicle_worksheet():
    sh = get_spreadsheet()
    return sh.worksheet(_vehicle_worksheet_name())


@st.cache_resource(show_spinner=False)
def get_mileage_log_worksheet():
    """Weekly Mileage Inspection entries live in their own sheet/tab
    (separate from VehicleHistory and Sheet2), created automatically the
    first time it's needed. One row per mileage entry."""
    sh = get_spreadsheet()
    worksheet_name = _mileage_log_worksheet_name()
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows=2000, cols=len(MILEAGE_LOG_COLUMNS) + 2)
        ws.append_row(MILEAGE_LOG_COLUMNS, value_input_option="USER_ENTERED")
    if ws.col_count < len(MILEAGE_LOG_COLUMNS):
        ws.resize(cols=len(MILEAGE_LOG_COLUMNS))
    header = ws.row_values(1)
    if header != MILEAGE_LOG_COLUMNS:
        ws.update(range_name="A1", values=[MILEAGE_LOG_COLUMNS])
    return ws


def clear_data_caches():
    """Invalidates all cached reads so the next load fetches fresh data
    from the sheet — called after a write, and from the manual refresh
    button on the View page."""
    _load_data_cached.clear()
    load_vehicle_details.clear()
    load_vehicle_all_columns.clear()
    load_mileage_log.clear()


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


def load_data_with_rows():
    """Fresh (uncached) read of the history sheet that also returns each
    record's physical sheet row number in a "_row" column — used by the Edit
    Entry section so it updates exactly the row the user picked."""
    ws = get_worksheet()
    values = ws.get_all_values()
    n = len(COLUMN_ORDER)
    cols = COLUMN_ORDER + ["_row"]
    if len(values) <= 1:
        return pd.DataFrame(columns=cols)
    rows = []
    for i, row in enumerate(values[1:], start=2):
        rows.append((row + [""] * n)[:n] + [i])
    df = pd.DataFrame(rows, columns=cols)
    for c in COLUMN_ORDER:
        df[c] = df[c].astype(str).replace("nan", "")
    return df


def append_record(record: dict):
    ws = get_worksheet()
    row = [str(record.get(c, "")) for c in COLUMN_ORDER]
    ws.append_row(row, value_input_option="USER_ENTERED")
    _load_data_cached.clear()


def append_records(records: list):
    """Batch version of append_record — used to save several Description/Cost
    line items from the same entry (same Vehicle No, Date, Mileage, Place)
    as separate sheet rows in a single API call."""
    ws = get_worksheet()
    rows = [[str(r.get(c, "")) for c in COLUMN_ORDER] for r in records]
    ws.append_rows(rows, value_input_option="USER_ENTERED")
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


# Newer Streamlit versions let st.selectbox suggest existing options while
# still accepting freshly typed text (a combobox). Checked once at import
# time so the Place field can use it when available and fall back to a
# plain text box on older Streamlit installs, instead of ever crashing.
_SELECTBOX_SUPPORTS_NEW_OPTIONS = "accept_new_options" in inspect.signature(st.selectbox).parameters


def known_places():
    """Distinct, non-blank Place values seen in past records, for the
    Place field's autocomplete suggestions."""
    try:
        return sorted({p for p in load_data()["Place"].tolist() if p})
    except Exception:
        return []


# Column names (Sheet2) that hold a 3D model — either a bare embed URL or
# a full pasted embed snippet (e.g. Sketchfab's iframe block). Checked
# case-insensitively against Sheet2's headers.
_3D_MODEL_COLUMN_NAMES = {"3d model", "3d", "model 3d", "sketchfab", "3d view", "3d model url"}


def extract_embed_src(value):
    """Pulls the iframe src URL out of either a bare link or a full pasted
    embed snippet (like Sketchfab's <iframe ...> block), so the 3D Model
    column can hold whichever the user finds easier to paste into Sheet2."""
    value = str(value).strip()
    if not value:
        return None
    match = re.search(r'src=["\']([^"\']+)["\']', value)
    if match:
        return match.group(1)
    if value.startswith("http"):
        return value
    return None


# -------------------------------------------------------------------------
# Vehicle preview card (Vehicle Details)
# -------------------------------------------------------------------------
# Sheet2 column names (case-insensitive) used by the preview card. Each
# picture cell can hold an image link (https://...), a Google Drive share link
# (the file must be shared as "Anyone with the link"), or the path of an
# image file kept in the app's GitHub repo (e.g. images/BJF 0588.jpg).
_VEHICLE_IMAGE_COLUMN_NAMES = ("vehicle image", "vehicle photo", "vehicle picture", "image", "photo")
_RIDER_IMAGE_COLUMN_NAMES = ("rider image", "rider photo", "rider picture")
# Line shown at the bottom of the card; the first filled column is used.
_MODEL_COLUMN_NAMES = ("vehicle model", "model", "brand")
_IMAGE_FILE_TYPES = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                     ".gif": "image/gif", ".webp": "image/webp"}
_MAX_REPO_IMAGE_BYTES = 5 * 1024 * 1024


def _field(info, names):
    """First non-blank value among the given Sheet2 column names."""
    lowered = {str(k).strip().lower(): v for k, v in info.items()}
    for name in names:
        value = str(lowered.get(name, "")).strip()
        if value:
            return value
    return ""


def image_src_from_cell(value):
    """Turns a Sheet2 picture cell into something an <img> tag can show: a web
    link, a Google Drive share link, or an image file kept in the app's repo.
    Returns None when the cell is blank or can't be used."""
    value = str(value).strip()
    if not value:
        return None
    if value.lower().startswith(("http://", "https://")):
        drive = (re.search(r"drive\.google\.com/file/d/([\w-]+)", value)
                 or re.search(r"drive\.google\.com/\S*[?&]id=([\w-]+)", value))
        if drive:
            return f"https://drive.google.com/thumbnail?id={drive.group(1)}&sz=w1000"
        return value
    # A file inside the app's own folder (anything outside it is refused).
    try:
        base = os.path.realpath(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.realpath(os.path.join(base, value))
        mime = _IMAGE_FILE_TYPES.get(os.path.splitext(path)[1].lower())
        if (mime and os.path.commonpath([base, path]) == base and os.path.isfile(path)
                and os.path.getsize(path) <= _MAX_REPO_IMAGE_BYTES):
            with open(path, "rb") as fh:
                return f"data:{mime};base64," + base64.b64encode(fh.read()).decode("ascii")
    except (ValueError, OSError):
        pass
    return None


def build_vehicle_card_html(vehicle_no, vehicle_src, rider_src, model_text):
    """The preview card: Vehicle No on top, the vehicle picture on the left,
    the rider picture with a 'Rider' label on the right, and the model line
    underneath. A missing picture shows a grey 'No image' box."""
    def pic(src, alt):
        if src:
            return f'<img src="{html.escape(src, quote=True)}" alt="{html.escape(alt)}" referrerpolicy="no-referrer">'
        return '<div class="empty">No image</div>'

    return f"""
    <style>
      body {{ margin: 0; font-family: Arial, Helvetica, sans-serif; background: transparent; }}
      .card {{ max-width: 560px; border: 1px solid #000; background: #fff; color: #000; }}
      .title {{ text-align: center; font-size: 18px; padding: 8px 6px; border-bottom: 1px solid #000; }}
      .body {{ display: grid; grid-template-columns: 1.25fr 1fr; border-bottom: 1px solid #000; }}
      .vehicle {{
        height: 250px; box-sizing: border-box; padding: 6px; border-right: 1px solid #000;
        display: flex; align-items: center; justify-content: center;
      }}
      .vehicle img {{ max-width: 100%; max-height: 100%; object-fit: contain; }}
      .rider {{ height: 250px; display: flex; flex-direction: column; }}
      .photo {{
        flex: 1; min-height: 0; box-sizing: border-box; padding: 6px; border-bottom: 1px solid #000;
        display: flex; align-items: center; justify-content: center;
      }}
      .photo img {{ max-width: 100%; max-height: 100%; object-fit: contain; }}
      .label {{ flex: 0 0 40px; line-height: 40px; text-align: center; font-size: 18px; }}
      .model {{ padding: 8px 6px; font-size: 18px; text-transform: uppercase; }}
      .empty {{
        width: 100%; height: 100%; display: flex; align-items: center; justify-content: center;
        background: #eee; color: #888; font-size: 14px;
      }}
    </style>
    <div class="card">
      <div class="title">{html.escape(str(vehicle_no))}</div>
      <div class="body">
        <div class="vehicle">{pic(vehicle_src, "Vehicle")}</div>
        <div class="rider">
          <div class="photo">{pic(rider_src, "Rider")}</div>
          <div class="label">Rider</div>
        </div>
      </div>
      <div class="model">{html.escape(str(model_text)) or "&nbsp;"}</div>
    </div>
    """


def build_details_table_html(headers, info):
    """Compact two-column table (Field | Value), one line per Sheet2 column,
    kept the same width as the preview card. Colours follow the app theme."""
    rows = "".join(
        f"<tr><th>{html.escape(str(h))}</th><td>{html.escape(str(info.get(h, '')))}</td></tr>"
        for h in headers
    )
    return (
        "<style>"
        ".vd-table{width:100%;max-width:560px;border-collapse:collapse;font-size:13px;line-height:1.25;}"
        ".vd-table th,.vd-table td{border:1px solid rgba(128,128,128,0.45);padding:3px 8px;"
        "text-align:left;vertical-align:top;word-break:break-word;}"
        ".vd-table th{width:38%;font-weight:600;background:rgba(128,128,128,0.12);}"
        "</style>"
        f'<table class="vd-table"><tbody>{rows}</tbody></table>'
    )


@st.cache_data(show_spinner=False, ttl=60)
def load_vehicle_all_columns():
    """Returns ({vehicle_no: {header: value, ...}}, [headers]) covering
    every column present in Sheet2 (not just Vehicle Type/Rider/Brand) —
    used by the Vehicle Details page so it shows whatever columns that
    sheet actually has, including ones added later."""
    ws = get_vehicle_worksheet()
    values = ws.get_all_values()
    if not values:
        return {}, []
    header = values[0]
    vno_idx = _find_column_index(header, "Vehicle No", "Vehicle No.", "VehicleNo", "Vehicle Number")
    if vno_idx is None:
        return {}, []

    def _cell(row, idx):
        return row[idx].strip() if idx < len(row) else ""

    headers = [h for i, h in enumerate(header) if str(h).strip() and i != vno_idx]
    details = {}
    for row in values[1:]:
        vno = _cell(row, vno_idx)
        if not vno:
            continue
        details[vno] = {h: _cell(row, i) for i, h in enumerate(header) if str(h).strip()}
    return details, headers


# -------------------------------------------------------------------------
# Weekly Mileage Inspection helpers
# -------------------------------------------------------------------------
@st.cache_data(show_spinner=False, ttl=20)
def load_mileage_log(sheet_id):
    ws = get_mileage_log_worksheet()
    values = ws.get_all_values()
    n = len(MILEAGE_LOG_COLUMNS)
    if len(values) <= 1:
        return pd.DataFrame(columns=MILEAGE_LOG_COLUMNS)
    data_rows = values[1:]
    padded_rows = [(row + [""] * n)[:n] for row in data_rows]
    df = pd.DataFrame(padded_rows, columns=MILEAGE_LOG_COLUMNS)
    df = df.astype(str).replace("nan", "")
    return df


def append_mileage_record(record: dict):
    ws = get_mileage_log_worksheet()
    row = [str(record.get(c, "")) for c in MILEAGE_LOG_COLUMNS]
    ws.append_row(row, value_input_option="USER_ENTERED")
    load_mileage_log.clear()


def latest_mileage_entry(log_df, vehicle_no, mileage_type):
    """Returns (mileage, date_str) of the most recent entry (by Date, then
    by Timestamp) for this vehicle and mileage type, or None."""
    if log_df.empty:
        return None
    sub = log_df[(log_df["Vehicle No"] == vehicle_no) & (log_df["Mileage Type"] == mileage_type)]
    if sub.empty:
        return None
    sub = sub.sort_values(["Date", "Timestamp"], ascending=[False, False])
    top = sub.iloc[0]
    return top["Mileage (KM)"], top["Date"]


def latest_serviced_oil_type(log_df, vehicle_no):
    """Oil Type recorded on the most recent Serviced Mileage entry for this
    vehicle, or "" if there is none (or it was saved without one)."""
    if log_df.empty:
        return ""
    sub = log_df[(log_df["Vehicle No"] == vehicle_no) & (log_df["Mileage Type"] == "Serviced Mileage")]
    if sub.empty:
        return ""
    sub = sub.sort_values(["Date", "Timestamp"], ascending=[False, False])
    return str(sub.iloc[0]["Oil Type"]).strip()


def is_fixed_interval_type(vehicle_type):
    """True for Lorry and Car/Cab vehicle types. For these the next service
    interval is fixed (Lorry 5000/6000 km, Car/Cab 6000 km) and does not
    depend on the oil type, so Oil Type is a free-text, optional field for
    them."""
    vtype = str(vehicle_type).strip().lower()
    return "lorry" in vtype or "car" in vtype or "cab" in vtype


def next_service_interval(log_df, vehicle_no):
    """Returns (km_interval, label) for the next service, or (None, "").
    Lorry -> 5000 km for LB 8797 and 227 - 4969, 6000 km for all other
    lorries. Car/Cab -> 6000 km. Any other type (bikes) depends on the
    Oil Type of the last service (see OIL_SERVICE_INTERVALS)."""
    vtype = str(vehicle_map.get(vehicle_no, "")).strip().lower()
    if "lorry" in vtype:
        if _normalize_vno(vehicle_no) in {_normalize_vno(v) for v in LORRY_5000_KM_VEHICLES}:
            return 5000, "Lorry"
        return 6000, "Lorry"
    if "car" in vtype or "cab" in vtype:
        return 6000, "Car/Cab"
    oil_type = latest_serviced_oil_type(log_df, vehicle_no)
    interval = OIL_SERVICE_INTERVALS.get(oil_type)
    if interval:
        return interval, oil_type
    return None, ""


def _fmt_mileage(value):
    num = to_number(value)
    if str(value).strip() == "":
        return "—"
    return f"{int(num)}" if num == int(num) else f"{num}"


def build_mileage_lines_html(log_df, vehicles, red_below_km=None):
    """One horizontal 'line' per vehicle: a thick black track with an orange
    Last Serviced box and a white Weekly Mileage box (each with its date
    underneath), mirroring the reference layout. The Rider's name (from
    Sheet2) is shown in brackets on the line below the Vehicle No."""
    def _km_to_next_service(v):
        """Next Service mileage minus latest Weekly Mileage, or None when it
        can't be worked out (no serviced/weekly entry or no oil type)."""
        s = latest_mileage_entry(log_df, v, "Serviced Mileage")
        w = latest_mileage_entry(log_df, v, "Weekly Mileage")
        iv, _ = next_service_interval(log_df, v)
        if s and w and iv:
            return (to_number(s[0]) + iv) - to_number(w[0])
        return None

    # Lines are ordered by km left to the next service, smallest first
    # (overdue vehicles with negative values come first); vehicles whose
    # value can't be calculated go last, in Vehicle No order.
    vehicles = sorted(
        vehicles,
        key=lambda v: (
            _km_to_next_service(v) is None,
            _km_to_next_service(v) if _km_to_next_service(v) is not None else 0,
            v,
        ),
    )

    rows_html = []
    for vno in vehicles:
        serviced = latest_mileage_entry(log_df, vno, "Serviced Mileage")
        weekly = latest_mileage_entry(log_df, vno, "Weekly Mileage")

        # Rider name (from Sheet2) shown in brackets under the Vehicle No.
        rider_name = str(vehicle_details.get(vno, {}).get("Rider", "")).strip()
        rider_html = f'<div class="rider">({html.escape(rider_name)})</div>' if rider_name else ""

        s_val = _fmt_mileage(serviced[0]) if serviced else "—"
        s_date = html.escape(serviced[1]) if serviced else "&nbsp;"
        w_val = _fmt_mileage(weekly[0]) if weekly else "—"
        w_date = html.escape(weekly[1]) if weekly else "&nbsp;"

        # Next service = last serviced mileage + the interval for the oil
        # used at that service (Castrol 5000 km, Bajaj Genuine 2500 km).
        interval, interval_label = next_service_interval(log_df, vno)
        if serviced and interval:
            n_val = _fmt_mileage(to_number(serviced[0]) + interval)
            n_note = f"{html.escape(interval_label)} (+{interval} km)"
        else:
            n_val = "—"
            n_note = "Oil type not set" if serviced else "&nbsp;"

        # Difference between the latest Weekly Mileage and the Last Serviced
        # mileage (KM run since the service), shown in red on the line
        # between those two boxes. Plain line if either value is missing.
        if serviced and weekly:
            diff_val = to_number(weekly[0]) - to_number(serviced[0])
            diff_html = f'<div class="diff"><span>{html.escape(_fmt_mileage(diff_val))} km</span></div>'
        else:
            diff_html = '<div class="seg mid"></div>'

        # KM still to go until the next service: Next Service mileage minus
        # the latest Weekly Mileage, shown in red on the line between those
        # two boxes. Plain line if the next service can't be worked out yet.
        if serviced and interval and weekly:
            remaining_val = (to_number(serviced[0]) + interval) - to_number(weekly[0])
            remaining_html = f'<div class="diff"><span>{html.escape(_fmt_mileage(remaining_val))} km</span></div>'
        else:
            remaining_html = '<div class="seg mid"></div>'

        # Only when asked for (the Serviced Report): colour the Next Service
        # box red once the KM left to the next service is below the limit.
        next_box_class = "nextsvc"
        if red_below_km is not None and serviced and interval and weekly and remaining_val < red_below_km:
            next_box_class = "nextsvc low"

        rows_html.append(f"""
        <div class="row">
          <div class="vlabel">
            <div class="vno">{html.escape(vno)}</div>
            {rider_html}
          </div>
          <div class="track">
            <div class="seg short"></div>
            <div class="col">
              <div class="cap">Last Serviced</div>
              <div class="box serviced">{html.escape(s_val)}</div>
              <div class="dt">{s_date}</div>
            </div>
            {diff_html}
            <div class="col">
              <div class="cap">Weekly Mileage</div>
              <div class="box weekly">{html.escape(w_val)}</div>
              <div class="dt">{w_date}</div>
            </div>
            {remaining_html}
            <div class="col wide">
              <div class="cap">Next Service</div>
              <div class="box {next_box_class}">{html.escape(n_val)}</div>
              <div class="dt">{n_note}</div>
            </div>
            <div class="seg long"></div>
          </div>
        </div>
        """)

    return f"""
    <style>
      body {{ margin: 0; font-family: "Source Sans Pro", Arial, sans-serif; background: transparent; }}
      .row {{
        display: flex; align-items: center; height: 104px;
        background: #ffffff; border: 1px solid #cfcfcf; border-radius: 6px;
        margin-bottom: 10px; padding: 0 10px; box-sizing: border-box;
      }}
      .vlabel {{
        flex: 0 0 140px; display: flex; flex-direction: column;
        align-items: center; justify-content: center; text-align: center;
      }}
      .vno {{ font-weight: 700; font-size: 15px; color: #111; }}
      .rider {{
        font-weight: 600; font-size: 12px; color: #555; margin-top: 3px;
        max-width: 130px; word-break: break-word; line-height: 1.2;
      }}
      .track {{ flex: 1; display: flex; align-items: center; min-width: 0; }}
      .seg {{ height: 6px; background: #000; }}
      .seg.short {{ flex: 0 0 24px; }}
      .seg.mid {{ flex: 0 0 40px; }}
      .seg.long {{ flex: 1; }}
      .col {{ flex: 0 0 112px; text-align: center; }}
      .cap {{ font-size: 13px; color: #111; margin-bottom: 3px; white-space: nowrap; }}
      .box {{
        height: 38px; line-height: 38px; font-weight: 700; font-size: 18px;
        box-sizing: border-box; color: #111;
      }}
      .box.serviced {{ background: #ed7d31; }}
      .box.weekly {{ background: #ffffff; border: 2px solid #f2b705; }}
      .col.wide {{ flex: 0 0 150px; }}
      .box.nextsvc {{ background: #1e6fd9; color: #ffffff; }}
      .box.nextsvc.low {{ background: #e00000; }}
      .diff {{ flex: 0 0 80px; position: relative; height: 6px; background: #000; }}
      .diff span {{
        position: absolute; left: -10px; right: -10px; bottom: 9px;
        text-align: center; color: #e00000; font-weight: 700; font-size: 13px; white-space: nowrap;
      }}
      .dt {{ font-size: 12px; font-weight: 700; color: #111; margin-top: 3px; }}
    </style>
    {''.join(rows_html)}
    """


# -------------------------------------------------------------------------
# Serviced Report helpers
# -------------------------------------------------------------------------
# Vehicle Types left out of the Weekly Mileage Inspection and Serviced Report.
EXCLUDED_VEHICLE_TYPES = {"other", "others"}


def is_excluded_vehicle_type(vehicle_type):
    return str(vehicle_type).strip().lower() in EXCLUDED_VEHICLE_TYPES


def vehicle_type_group_rank(vehicle_type):
    """Serviced Report group order: Bike (0), Lorry (1), Car/Cab (2); a
    vehicle with a blank Vehicle Type goes last (3)."""
    vtype = str(vehicle_type).strip().lower()
    if not vtype:
        return 3
    if "lorry" in vtype:
        return 1
    if "car" in vtype or "cab" in vtype:
        return 2
    return 0


def km_to_next_service(log_df, vehicle_no):
    """KM still to go until the next service: (last Serviced Mileage +
    service interval) minus the latest Weekly Mileage — the same figure
    shown in red on the Mileage Lines. Negative means overdue. Returns None
    when it can't be worked out (no serviced/weekly entry, or no interval)."""
    serviced = latest_mileage_entry(log_df, vehicle_no, "Serviced Mileage")
    weekly = latest_mileage_entry(log_df, vehicle_no, "Weekly Mileage")
    interval, _ = next_service_interval(log_df, vehicle_no)
    if serviced and weekly and interval:
        return (to_number(serviced[0]) + interval) - to_number(weekly[0])
    return None


def build_serviced_report_rows(log_df, vehicle_types=None):
    """One row per vehicle (optionally limited to the given Vehicle Types),
    leaving out the Others vehicle type. Rows are grouped Bike, Lorry, then
    Car/Cab, and within each group ordered by KM left to the next service,
    smallest first (overdue vehicles first); vehicles whose value can't be
    calculated go last in their group."""
    rows = []
    for vno, info in vehicle_details.items():
        vtype = info.get("Vehicle Type", "")
        if is_excluded_vehicle_type(vtype):
            continue
        if vehicle_types and vtype not in vehicle_types:
            continue
        rows.append({
            "Vehicle No": vno,
            "Vehicle Type": vtype,
            "Rider": info.get("Rider", ""),
            "km": km_to_next_service(log_df, vno),
        })
    rows.sort(key=lambda r: (
        vehicle_type_group_rank(r["Vehicle Type"]),
        r["km"] is None,
        r["km"] if r["km"] is not None else 0,
        r["Vehicle No"],
    ))
    return rows


def serviced_line_data(log_df, vehicle_no):
    """The figures shown on one Serviced Report line — the same values the
    Mileage Lines show for a vehicle (Last Serviced, Weekly Mileage, the two
    red KM figures and Next Service)."""
    serviced = latest_mileage_entry(log_df, vehicle_no, "Serviced Mileage")
    weekly = latest_mileage_entry(log_df, vehicle_no, "Weekly Mileage")
    interval, interval_label = next_service_interval(log_df, vehicle_no)

    data = {
        "s_val": _fmt_mileage(serviced[0]) if serviced else "—",
        "s_date": serviced[1] if serviced else "",
        "w_val": _fmt_mileage(weekly[0]) if weekly else "—",
        "w_date": weekly[1] if weekly else "",
        "diff": None,
        "remaining": None,
        "remaining_km": None,
    }
    if serviced and interval:
        data["n_val"] = _fmt_mileage(to_number(serviced[0]) + interval)
        data["n_note"] = f"{interval_label} (+{interval} km)"
    else:
        data["n_val"] = "—"
        data["n_note"] = "Oil type not set" if serviced else ""
    if serviced and weekly:
        data["diff"] = f"{_fmt_mileage(to_number(weekly[0]) - to_number(serviced[0]))} km"
    if serviced and interval and weekly:
        remaining_km = (to_number(serviced[0]) + interval) - to_number(weekly[0])
        data["remaining_km"] = remaining_km
        data["remaining"] = f"{_fmt_mileage(remaining_km)} km"
    return data


def build_serviced_report_pdf(rows, log_df):
    """Renders the Serviced Report as a portrait PDF laid out like the
    Mileage Lines: one line per vehicle (Vehicle No with the Rider in
    brackets, orange Last Serviced box, white Weekly Mileage box, blue Next
    Service box, red KM figures on the line), grouped Bike, Lorry, Car/Cab.
    `rows` must already be ordered (see build_serviced_report_rows)."""
    group_names = {0: "Bike", 1: "Lorry", 2: "Car/Cab", 3: "No Vehicle Type"}
    buffer = io.BytesIO()
    page_w, page_h = A4
    c = canvas.Canvas(buffer, pagesize=(page_w, page_h))
    c.setTitle("Serviced Report")

    margin = 12 * mm
    row_h = 26 * mm
    gap = 3 * mm
    head_h = 9 * mm
    left, right = margin, page_w - margin
    bottom_limit = margin + 6 * mm
    black = colors.black
    grey = colors.HexColor("#555555")
    red = colors.HexColor("#e00000")

    def fit_text(text, font, size, max_w):
        text = str(text)
        if stringWidth(text, font, size) <= max_w:
            return text
        while text and stringWidth(text + "..", font, size) > max_w:
            text = text[:-1]
        return text + ".."

    def new_page_top():
        c.showPage()
        return page_h - margin

    def draw_heading(y, text):
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(left, y - 6 * mm, text)
        return y - head_h

    def draw_row(yb, r):
        d = serviced_line_data(log_df, r["Vehicle No"])
        cy = yb + row_h / 2

        # Row card
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#cfcfcf"))
        c.setLineWidth(0.6)
        c.roundRect(left, yb, right - left, row_h, 1.5 * mm, stroke=1, fill=1)

        # Vehicle No, with the Rider in brackets underneath
        label_cx = left + 17 * mm
        rider = str(r["Rider"]).strip()
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", 10)
        c.drawCentredString(
            label_cx, cy + (1.5 * mm if rider else -1.2 * mm),
            fit_text(r["Vehicle No"], "Helvetica-Bold", 10, 30 * mm),
        )
        if rider:
            c.setFont("Helvetica-Bold", 8)
            c.setFillColor(grey)
            c.drawCentredString(label_cx, cy - 3 * mm, fit_text(f"({rider})", "Helvetica-Bold", 8, 30 * mm))

        # The thick black track
        x_track = left + 34 * mm
        x_end = right - 4 * mm
        c.setFillColor(black)
        c.rect(x_track, cy - 0.7 * mm, x_end - x_track, 1.4 * mm, stroke=0, fill=1)

        box_h = 9 * mm

        def draw_col(x_left, col_w, caption, value, note, fill, text_color, stroke=None):
            cx = x_left + col_w / 2
            bw = col_w - 2 * mm
            c.setFillColor(black)
            c.setFont("Helvetica", 8)
            c.drawCentredString(cx, cy + box_h / 2 + 2 * mm, caption)
            c.setFillColor(fill)
            if stroke is not None:
                c.setStrokeColor(stroke)
                c.setLineWidth(1)
                c.rect(cx - bw / 2, cy - box_h / 2, bw, box_h, stroke=1, fill=1)
            else:
                c.rect(cx - bw / 2, cy - box_h / 2, bw, box_h, stroke=0, fill=1)
            c.setFillColor(text_color)
            c.setFont("Helvetica-Bold", 12)
            c.drawCentredString(cx, cy - 1.5 * mm, value)
            c.setFillColor(black)
            c.setFont("Helvetica-Bold", 7.5)
            c.drawCentredString(cx, cy - box_h / 2 - 3.5 * mm, fit_text(note, "Helvetica-Bold", 7.5, col_w + 8 * mm))

        def draw_seg_text(x_left, seg_w, text):
            if text:
                c.setFillColor(red)
                c.setFont("Helvetica-Bold", 8)
                c.drawCentredString(x_left + seg_w / 2, cy + 2 * mm, text)

        # The Next Service box turns red when 500 km or fewer are left
        # (below 501 km), including vehicles that are already overdue.
        low_km = d["remaining_km"] is not None and d["remaining_km"] < 501
        next_fill = red if low_km else colors.HexColor("#1e6fd9")

        x = x_track + 3 * mm
        draw_col(x, 26 * mm, "Last Serviced", d["s_val"], d["s_date"], colors.HexColor("#ed7d31"), black)
        x += 26 * mm
        draw_seg_text(x, 17 * mm, d["diff"])
        x += 17 * mm
        draw_col(x, 26 * mm, "Weekly Mileage", d["w_val"], d["w_date"], colors.white, black, stroke=colors.HexColor("#f2b705"))
        x += 26 * mm
        draw_seg_text(x, 17 * mm, d["remaining"])
        x += 17 * mm
        draw_col(x, 42 * mm, "Next Service", d["n_val"], d["n_note"], next_fill, colors.white)

    # Page title
    c.setFillColor(black)
    c.setFont("Helvetica-Bold", 15)
    c.drawCentredString(page_w / 2, page_h - margin - 5 * mm, "Vehicle History Monitoring System - KMN")
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left, page_h - margin - 14 * mm, "Serviced Report")
    c.setFont("Helvetica", 8)
    c.setFillColor(grey)
    c.drawRightString(right, page_h - margin - 14 * mm, f"Generated on {date.today().isoformat()}")
    y = page_h - margin - 19 * mm

    current_rank = None
    for r in rows:
        rank = vehicle_type_group_rank(r["Vehicle Type"])
        if rank != current_rank:
            if y - (head_h + row_h) < bottom_limit:
                y = new_page_top()
            y = draw_heading(y, group_names.get(rank, ""))
            current_rank = rank
        elif y - row_h < bottom_limit:
            y = new_page_top()
            y = draw_heading(y, f"{group_names.get(rank, '')} (continued)")
        draw_row(y - row_h, r)
        y -= row_h + gap

    if y - 13 * mm < margin:
        y = new_page_top()
    c.setFont("Helvetica", 7.5)
    c.setFillColor(grey)
    c.drawString(
        left, y - 3 * mm,
        "Red figures on the line: KM run since the last service (between Last Serviced and Weekly Mileage) and",
    )
    c.drawString(
        left, y - 6.5 * mm,
        "KM left until the next service (between Weekly Mileage and Next Service). Zero or negative = due or overdue.",
    )
    c.drawString(
        left, y - 10 * mm,
        "A red Next Service box means 500 km or fewer are left.",
    )
    c.save()
    return buffer.getvalue()


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

st.markdown("<h1 style='text-align: center;'>Vehicle History Monitoring System - KMN</h1>", unsafe_allow_html=True)

# Sidebar: the signed-in user and the Log out button. (The Refresh button and
# the section list are added to the sidebar further down, once the setup
# checks have passed.)
with st.sidebar:
    st.write(f"👤 {st.session_state['username']}")
    if st.button("Log out", use_container_width=True, key="sidebar_logout"):
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
            'vehicle_worksheet_name = "Sheet2"\n'
            'mileage_log_worksheet_name = "MileageLog"\n',
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
with st.sidebar:
    # One Refresh for the whole app: clears every cached read so the next
    # load fetches fresh data from the Google Sheet.
    if st.button("🔄 Refresh", use_container_width=True, key="sidebar_refresh"):
        clear_data_caches()
        st.rerun()
    st.markdown("---")
    phase = st.radio(
        "Sections",
        ["📋 Record Entering", "📊 View", "🚙 Vehicle Details", "📈 Analytics", "🛣️ Weekly Mileage Inspection"],
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
        # updates live as soon as a vehicle is selected. Event Type is also
        # kept outside the form for the same reason — the Vehicle Part
        # options below depend on it, and widgets inside a form don't
        # trigger a rerun on change.
        vehicle_no = st.selectbox("Vehicle No *", vehicle_options, key="entry_vehicle_no")
        vehicle_type = vehicle_map.get(vehicle_no, "")
        st.text_input("Vehicle Type", value=vehicle_type, disabled=True)
        event_type = st.selectbox("Event Type", EVENT_TYPES, key="entry_event_type")

        # The Description/Cost table resets after a save by rendering it
        # under a brand-new widget key each time.
        if "_items_key_version" not in st.session_state:
            st.session_state["_items_key_version"] = 0
        items_key = f"entry_items_{st.session_state['_items_key_version']}"

        with st.form("add_entry_form", clear_on_submit=False):
            c1, c2 = st.columns(2)
            with c1:
                event_date = st.date_input("Date *", value=date.today(), key="entry_date")
                mileage = st.text_input(
                    "Mileage (KM) *", key="entry_mileage",
                    placeholder="e.g. 15000 or NA",
                )
            with c2:
                part_options = vehicle_part_options(event_type)
                if _SELECTBOX_SUPPORTS_NEW_OPTIONS:
                    place = st.selectbox(
                        "Place *",
                        options=known_places(),
                        index=None,
                        placeholder="Start typing a place...",
                        accept_new_options=True,
                        key="entry_place",
                    )
                    place = place or ""
                else:
                    place = st.text_input("Place *", key="entry_place")

            st.markdown("**Description of Goods / Service \\*, Vehicle Part, and Cost \\***")
            items_df = st.data_editor(
                pd.DataFrame({
                    "Description of Goods / Service *": [""],
                    "Cost *": [0.0],
                    "Vehicle Part": [part_options[0]]
                }),
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key=items_key,
                column_config={
                    "Description of Goods / Service *": st.column_config.TextColumn(width="large"),
                    "Cost *": st.column_config.NumberColumn(step=0.01, format="%.2f"),
                    "Vehicle Part": st.column_config.SelectboxColumn(options=part_options)
                },
            )

            submitted = st.form_submit_button("💾 Save Entry", type="primary", use_container_width=True)
            if submitted:
                errors = []
                if not vehicle_no:
                    errors.append("Vehicle No is required.")
                if not place.strip():
                    errors.append("Place is required.")
                mileage_str = mileage.strip()
                mileage_value = ""
                if not mileage_str:
                    errors.append("Mileage (KM) is required.")
                elif mileage_str.upper() == "NA":
                    mileage_value = "NA"
                else:
                    try:
                        as_float = float(mileage_str)
                        if as_float < 0:
                            errors.append("Mileage (KM) must be a positive number or 'NA'.")
                        else:
                            mileage_value = int(as_float) if as_float == int(as_float) else as_float
                    except ValueError:
                        errors.append("Mileage (KM) must be a number or 'NA'.")

                valid_items = []
                incomplete_row = False
                for _, item in items_df.iterrows():
                    desc = str(item.get("Description of Goods / Service *", "") or "").strip()
                    row_part = str(item.get("Vehicle Part", "") or "").strip() or part_options[0]
                    try:
                        cost_val = float(item.get("Cost *", 0.0) or 0.0)
                    except (TypeError, ValueError):
                        cost_val = 0.0

                    if not desc and cost_val == 0:
                        continue
                    if not desc:
                        incomplete_row = True
                        continue
                    valid_items.append((desc, row_part, cost_val))

                if incomplete_row:
                    errors.append("Each row with a Cost needs a Description of Goods / Service too.")
                if not valid_items:
                    errors.append("Add at least one Description of Goods / Service and Cost entry.")

                if errors:
                    for err in errors:
                        st.error(f"❌ {err}")
                else:
                    records = [
                        {
                            "Timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S.%f"),
                            "Vehicle No": vehicle_no,
                            "Vehicle Type": vehicle_type,
                            "Date": event_date.isoformat(),
                            "Mileage (KM)": mileage_value,
                            "Event Type": event_type,
                            "Place": place.strip(),
                            "Description of Goods / Service": desc,
                            "Cost": cost_val,
                            "User": st.session_state["username"],
                            "Vehicle Part": row_part,
                        }
                        for desc, row_part, cost_val in valid_items
                    ]
                    append_records(records)
                    entry_word = "entry" if len(records) == 1 else "entries"
                    st.success(f"✅ Saved {len(records)} {entry_word} for Vehicle No {vehicle_no}.")
                    time.sleep(1)
                    st.session_state["_items_key_version"] += 1
                    st.rerun()

        # -----------------------------------------------------------------
        # EDIT ENTRY — fix an entry that was entered by mistake. The record
        # is updated in place in the Google Sheet; the original Timestamp,
        # Vehicle No, Vehicle Type and User are kept, and the "Edited Date"
        # and "Edited By" columns record when it was edited and by whom.
        # -----------------------------------------------------------------
        st.markdown("---")
        st.markdown("#### ✏️ Edit Entry")
        st.caption("Pick an entry that was entered by mistake and correct it. The sheet records the edited date and the user who edited it.")

        edit_vno = st.selectbox("Vehicle No", vehicle_options, key="edit_vehicle_no")

        try:
            edit_all_df = load_data_with_rows()
        except Exception as e:
            edit_all_df = pd.DataFrame(columns=COLUMN_ORDER + ["_row"])
            st.error(f"❌ Could not read data from the Google Sheet.\n\n{e}")

        edit_df = edit_all_df[edit_all_df["Vehicle No"] == edit_vno]

        if edit_df.empty:
            st.info("No entries recorded for this vehicle yet.")
        else:
            edit_df = edit_df.sort_values(["Date", "Timestamp"], ascending=[False, False])
            edit_labels = {
                int(r["_row"]): (
                    f"{r['Date']} | {r['Event Type']} | {r['Vehicle Part']} | "
                    f"{r['Description of Goods / Service']} | {r['Cost']}"
                )
                for _, r in edit_df.iterrows()
            }
            edit_row_no = st.selectbox(
                "Select entry to edit",
                list(edit_labels.keys()),
                format_func=lambda k: edit_labels[k],
                key=f"edit_row_select_{edit_vno}",
            )
            sel = edit_df[edit_df["_row"] == edit_row_no].iloc[0]

            # Event Type sits outside the form so the Vehicle Part options
            # below update live when it changes.
            edit_event_options = list(EVENT_TYPES)
            if sel["Event Type"] and sel["Event Type"] not in edit_event_options:
                edit_event_options.append(sel["Event Type"])
            edit_event_default = edit_event_options.index(sel["Event Type"]) if sel["Event Type"] in edit_event_options else 0
            edit_event_type = st.selectbox(
                "Event Type",
                edit_event_options,
                index=edit_event_default,
                key=f"edit_event_type_{edit_row_no}",
            )

            edit_part_options = list(vehicle_part_options(edit_event_type))
            if sel["Vehicle Part"] and sel["Vehicle Part"] not in edit_part_options and edit_event_type == sel["Event Type"]:
                edit_part_options.append(sel["Vehicle Part"])
            edit_part_default = edit_part_options.index(sel["Vehicle Part"]) if sel["Vehicle Part"] in edit_part_options else 0

            parsed_date = pd.to_datetime(sel["Date"], errors="coerce")
            edit_date_default = parsed_date.date() if pd.notna(parsed_date) else date.today()

            with st.form("edit_entry_form", clear_on_submit=False):
                e1, e2 = st.columns(2)
                with e1:
                    edit_date = st.date_input("Date *", value=edit_date_default, key=f"edit_date_{edit_row_no}")
                    edit_mileage = st.text_input(
                        "Mileage (KM) *", value=sel["Mileage (KM)"], key=f"edit_mileage_{edit_row_no}",
                        placeholder="e.g. 15000 or NA",
                    )
                    edit_part = st.selectbox(
                        "Vehicle Part", edit_part_options, index=edit_part_default,
                        key=f"edit_part_{edit_row_no}_{edit_event_type}",
                    )
                with e2:
                    edit_place = st.text_input("Place *", value=sel["Place"], key=f"edit_place_{edit_row_no}")
                    edit_desc = st.text_input(
                        "Description of Goods / Service *", value=sel["Description of Goods / Service"],
                        key=f"edit_desc_{edit_row_no}",
                    )
                    edit_cost = st.number_input(
                        "Cost *", value=float(to_number(sel["Cost"])), step=0.01, format="%.2f",
                        key=f"edit_cost_{edit_row_no}",
                    )

                if sel["Edited Date"] or sel["Edited By"]:
                    st.caption(f"Last edited: {sel['Edited Date']} by {sel['Edited By']}")

                edit_submitted = st.form_submit_button("💾 Save Changes", type="primary", use_container_width=True)
                if edit_submitted:
                    edit_errors = []
                    edit_mileage_value = ""
                    edit_mileage_str = edit_mileage.strip()
                    if not edit_place.strip():
                        edit_errors.append("Place is required.")
                    if not edit_desc.strip():
                        edit_errors.append("Description of Goods / Service is required.")
                    if not edit_mileage_str:
                        edit_errors.append("Mileage (KM) is required.")
                    elif edit_mileage_str.upper() == "NA":
                        edit_mileage_value = "NA"
                    else:
                        try:
                            as_float = float(edit_mileage_str)
                            if as_float < 0:
                                edit_errors.append("Mileage (KM) must be a positive number or 'NA'.")
                            else:
                                edit_mileage_value = int(as_float) if as_float == int(as_float) else as_float
                        except ValueError:
                            edit_errors.append("Mileage (KM) must be a number or 'NA'.")

                    if edit_errors:
                        for err in edit_errors:
                            st.error(f"❌ {err}")
                    else:
                        try:
                            ws = get_worksheet()
                            n_cols = len(COLUMN_ORDER)
                            current_row = (ws.row_values(int(edit_row_no)) + [""] * n_cols)[:n_cols]
                            if current_row[0] != sel["Timestamp"]:
                                st.error("❌ This entry changed in the sheet since it was loaded. Press Refresh and try again.")
                            else:
                                updated = dict(zip(COLUMN_ORDER, current_row))
                                updated.update({
                                    "Date": edit_date.isoformat(),
                                    "Mileage (KM)": edit_mileage_value,
                                    "Event Type": edit_event_type,
                                    "Place": edit_place.strip(),
                                    "Description of Goods / Service": edit_desc.strip(),
                                    "Cost": edit_cost,
                                    "Vehicle Part": edit_part,
                                    "Edited Date": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "Edited By": st.session_state["username"],
                                })
                                new_row = [str(updated.get(c, "")) for c in COLUMN_ORDER]
                                ws.update(
                                    range_name=f"A{int(edit_row_no)}",
                                    values=[new_row],
                                    value_input_option="USER_ENTERED",
                                )
                                _load_data_cached.clear()
                                st.success(f"✅ Entry updated for Vehicle No {edit_vno}.")
                                time.sleep(1)
                                st.rerun()
                        except Exception as e:
                            st.error(f"❌ Could not update the entry.\n\n{e}")

# =========================================================================
# PHASE 2 — VIEW
# =========================================================================
elif phase == "📊 View":
    st.markdown("#### 📊 View Vehicle History")

    df = load_data()

    if df.empty:
        st.info("No records yet. Add one from the Record Entering section.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            vn_options = sorted([v for v in df["Vehicle No"].unique().tolist() if v])
            vn_filter = st.multiselect("Filter by VN", vn_options)
        with c2:
            type_filter = st.multiselect("Filter by Event Type", EVENT_TYPES)
        with c3:
            place_options = sorted([p for p in df["Place"].unique().tolist() if p])
            place_filter = st.multiselect("Filter by Place", place_options)
        with c4:
            if type_filter:
                part_options_view = sorted({
                    p for et in type_filter for p in vehicle_part_options(et)
                })
            else:
                part_options_view = sorted([p for p in df["Vehicle Part"].unique().tolist() if p])
            part_filter = st.multiselect("Filter by Vehicle Part", part_options_view)

        search_text = st.text_input(
            "🔍 Search",
            placeholder="Search anything: vehicle no, place, description, part, user, date...",
            key="view_search",
        )

        filtered = df.copy()
        if vn_filter:
            filtered = filtered[filtered["Vehicle No"].isin(vn_filter)]
        if type_filter:
            filtered = filtered[filtered["Event Type"].isin(type_filter)]
        if place_filter:
            filtered = filtered[filtered["Place"].isin(place_filter)]
        if part_filter:
            filtered = filtered[filtered["Vehicle Part"].isin(part_filter)]

        if search_text.strip():
            haystack = (
                filtered[VIEW_DISPLAY_COLUMNS]
                .astype(str)
                .agg(" ".join, axis=1)
                .str.lower()
            )
            search_mask = pd.Series(True, index=filtered.index)
            for term in search_text.strip().lower().split():
                search_mask &= haystack.str.contains(term, regex=False)
            filtered = filtered[search_mask]

        display_cols = list(VIEW_DISPLAY_COLUMNS)

        group_cols = ["Vehicle No", "Date", "Place"]
        table_df = filtered[display_cols].copy()
        cost_numeric = filtered["Cost"].apply(to_number)
        group_totals = cost_numeric.groupby(
            [filtered[c] for c in group_cols]
        ).transform("sum")
        is_last_in_group = ~filtered.duplicated(subset=group_cols, keep="last")
        table_df["Total Amount"] = ""
        table_df.loc[is_last_in_group, "Total Amount"] = group_totals[is_last_in_group].map(
            lambda x: f"{x:,.2f}"
        )

        st.dataframe(table_df, use_container_width=True, hide_index=True)

        total_amount = sum(to_number(c) for c in filtered["Cost"])
        st.markdown(
            f"<p style='text-align: right; font-weight: 700; font-size: 1.05rem;'>"
            f"Total Amount: {total_amount:,.2f}</p>",
            unsafe_allow_html=True,
        )

        vn_type_pairs = sorted({
            f"{vn} ({vt})" if vt else vn
            for vn, vt in zip(filtered["Vehicle No"], filtered["Vehicle Type"])
            if vn
        })
        event_part_pairs = sorted({
            f"{et} - {vp}" if vp else et
            for et, vp in zip(filtered["Event Type"], filtered["Vehicle Part"])
            if et
        })

        export_table_cols = [
            "Date", "Mileage (KM)", "Place",
            "Description of Goods / Service", "Cost", "Total Amount",
        ]
        export_table_df = table_df[export_table_cols]

        export_buffer = io.StringIO()
        csv_writer = csv.writer(export_buffer)
        csv_writer.writerow(["Vehicle History Monitoring System - KMN"])
        csv_writer.writerow(["; ".join(vn_type_pairs)])
        csv_writer.writerow(["; ".join(event_part_pairs)])
        csv_writer.writerow([])
        csv_writer.writerow(export_table_cols)
        for row in export_table_df.itertuples(index=False, name=None):
            csv_writer.writerow(row)
        csv_writer.writerow(["Total", "", "", "", f"{total_amount:,.2f}", ""])

        st.download_button(
            "⬇️ Download filtered results as CSV",
            export_buffer.getvalue().encode("utf-8-sig"),
            file_name="vehicle_history_export.csv",
            mime="text/csv",
        )

# =========================================================================
# PHASE 3 — VEHICLE DETAILS
# =========================================================================
elif phase == "🚙 Vehicle Details":
    st.markdown("#### 🚙 Vehicle Details")

    try:
        all_details, detail_headers = load_vehicle_all_columns()
    except Exception as e:
        all_details, detail_headers = {}, []
        st.warning(
            f"⚠️ Could not read the full vehicle details from "
            f"**{_vehicle_worksheet_name()}**.\n\n{e}"
        )

    detail_options = sorted(all_details.keys())

    if not detail_options:
        st.info(
            f"No vehicles found in the **{_vehicle_worksheet_name()}** tab yet. "
            "Add vehicles there (with a 'Vehicle No' column) first."
        )
    else:
        selected_vno = st.selectbox("Vehicle No", detail_options, key="details_vehicle_no")
        info = all_details.get(selected_vno, {})

        # Preview card: Vehicle No, vehicle picture, rider picture and model,
        # taken from the Vehicle Image / Rider Image / Model columns of Sheet2.
        components.html(
            build_vehicle_card_html(
                selected_vno,
                image_src_from_cell(_field(info, _VEHICLE_IMAGE_COLUMN_NAMES)),
                image_src_from_cell(_field(info, _RIDER_IMAGE_COLUMN_NAMES)),
                _field(info, _MODEL_COLUMN_NAMES),
            ),
            height=350,
        )

        if not detail_headers:
            st.info("No additional columns found for this vehicle.")
        else:
            model_col = next(
                (h for h in detail_headers if h.strip().lower() in _3D_MODEL_COLUMN_NAMES),
                None,
            )
            # The picture columns are shown on the preview card, so their long
            # links are left out of the details table.
            picture_cols = set(_VEHICLE_IMAGE_COLUMN_NAMES) | set(_RIDER_IMAGE_COLUMN_NAMES)
            table_headers = [
                h for h in detail_headers
                if h != model_col and h.strip().lower() not in picture_cols
            ]

            if table_headers:
                # A compact static table (not the scrolling grid) so every field
                # is listed line by line, all visible at once.
                st.markdown(build_details_table_html(table_headers, info), unsafe_allow_html=True)

            if model_col:
                embed_src = extract_embed_src(info.get(model_col, ""))
                if embed_src:
                    st.markdown("**🧊 3D Model**")
                    components.iframe(embed_src, height=480, scrolling=True)
                elif info.get(model_col, "").strip():
                    st.caption(f"3D Model: {info.get(model_col, '')}")

        st.markdown("---")
        st.markdown(f"**📜 Service History — {selected_vno}**")

        history_df = load_data()
        vehicle_history = history_df[history_df["Vehicle No"] == selected_vno]

        if vehicle_history.empty:
            st.info("No service history recorded for this vehicle yet.")
        else:
            history_display_cols = [c for c in VIEW_DISPLAY_COLUMNS if c != "Vehicle No"]
            st.dataframe(vehicle_history[history_display_cols], use_container_width=True, hide_index=True)

            vehicle_total = sum(to_number(c) for c in vehicle_history["Cost"])
            st.markdown(
                f"<p style='text-align: right; font-weight: 700; font-size: 1.05rem;'>"
                f"Total Amount: {vehicle_total:,.2f}</p>",
                unsafe_allow_html=True,
            )

# =========================================================================
# PHASE 4 — ANALYTICS
# =========================================================================
elif phase == "📈 Analytics":
    st.markdown("#### 📈 Analytics")

    analytics_df = load_data()

    if analytics_df.empty:
        st.info("No records yet. Add one from the Record Entering section.")
    else:
        a1, a2, a3, a4 = st.columns(4)
        with a1:
            an_vn_options = sorted([v for v in analytics_df["Vehicle No"].unique().tolist() if v])
            an_vn_filter = st.multiselect("Vehicle No", an_vn_options, key="analytics_vehicle_no")
        with a2:
            an_timeframe = st.selectbox("Time Frame", ["Monthly", "Yearly"], key="analytics_timeframe")
        with a3:
            an_event_filter = st.multiselect("Event Type", EVENT_TYPES, key="analytics_event_type")
        with a4:
            an_part_options = sorted([p for p in analytics_df["Vehicle Part"].unique().tolist() if p])
            an_part_filter = st.multiselect("Vehicle Part", an_part_options, key="analytics_vehicle_part")

        an_filtered = analytics_df.copy()
        if an_vn_filter:
            an_filtered = an_filtered[an_filtered["Vehicle No"].isin(an_vn_filter)]
        if an_event_filter:
            an_filtered = an_filtered[an_filtered["Event Type"].isin(an_event_filter)]
        if an_part_filter:
            an_filtered = an_filtered[an_filtered["Vehicle Part"].isin(an_part_filter)]

        if an_filtered.empty:
            st.info("No records match the selected filters.")
        else:
            an_dates = pd.to_datetime(an_filtered["Date"], errors="coerce")
            an_costs = an_filtered["Cost"].apply(to_number)
            valid_mask = an_dates.notna()
            an_dates = an_dates[valid_mask]
            an_costs = an_costs[valid_mask]

            if an_dates.empty:
                st.info("No records with a valid Date match the selected filters.")
            else:
                if an_timeframe == "Monthly":
                    period_labels = an_dates.dt.strftime("%Y-%m")
                else:
                    period_labels = an_dates.dt.strftime("%Y")

                summary = (
                    pd.DataFrame({"Period": period_labels.values, "Cost": an_costs.values})
                    .groupby("Period", as_index=False)["Cost"].sum()
                    .sort_values("Period")
                )

                axis_title = "Month" if an_timeframe == "Monthly" else "Year"
                base = alt.Chart(summary).encode(
                    x=alt.X("Period:N", title=axis_title, sort=None),
                    y=alt.Y("Cost:Q", title="Total Cost"),
                )
                line = base.mark_line(color="#e63946", strokeWidth=2)
                points = base.mark_point(color="#e63946", size=70, filled=True)
                labels = base.mark_text(
                    align="center", baseline="line-bottom", dy=-8, color="#333333"
                ).encode(text=alt.Text("Cost:Q", format=",.2f"))
                chart = (line + points + labels).properties(height=420)
                st.altair_chart(chart, use_container_width=True)

                summary_display = summary.rename(
                    columns={"Period": an_timeframe, "Cost": "Total Cost"}
                )
                summary_display["Total Cost"] = summary_display["Total Cost"].map(lambda x: f"{x:,.2f}")
                with st.expander("📄 Show data table", expanded=False):
                    st.dataframe(summary_display, use_container_width=True, hide_index=True)

# =========================================================================
# PHASE 5 — WEEKLY MILEAGE INSPECTION (three sub-phases)
# =========================================================================
elif phase == "🛣️ Weekly Mileage Inspection":
    st.markdown("#### 🛣️ Weekly Mileage Inspection")

    st.caption(
        "Entries are kept in their own "
        f"**{_mileage_log_worksheet_name()}** sheet, separate from the vehicle history records."
    )

    mileage_phase = st.radio(
        "Mileage phase",
        ["✍️ Entering", "📍 Mileage Lines", "📑 Serviced Report"],
        horizontal=True,
        label_visibility="collapsed",
        key="mileage_subphase",
    )
    st.markdown("---")

    vehicle_types_available = sorted({t for t in vehicle_map.values() if t and not is_excluded_vehicle_type(t)})

    if not vehicle_map:
        st.info(
            f"No vehicles found in the **{_vehicle_worksheet_name()}** tab yet. "
            "Add vehicles there (with a 'Vehicle No' column) before recording mileage."
        )
    elif not vehicle_types_available:
        st.info(
            f"No Vehicle Type values found in the **{_vehicle_worksheet_name()}** tab. "
            "Fill the 'Vehicle Type' column for your vehicles first."
        )
    else:
        try:
            mileage_log_df = load_mileage_log(_sheet_id())
        except Exception as e:
            mileage_log_df = pd.DataFrame(columns=MILEAGE_LOG_COLUMNS)
            st.error(f"❌ Could not read the **{_mileage_log_worksheet_name()}** sheet.\n\n{e}")

        # ---------------------------------------------------------------
        # SUB-PHASE 1 — ENTERING
        # ---------------------------------------------------------------
        if mileage_phase == "✍️ Entering":
            # Selections live outside the form so Vehicle No's options
            # update live as soon as the Vehicle Type changes.
            ml_vehicle_type = st.selectbox("Vehicle Type *", vehicle_types_available, key="mil_vehicle_type")
            ml_vehicle_options = sorted(v for v, t in vehicle_map.items() if t == ml_vehicle_type)
            ml_vehicle_no = st.selectbox("Vehicle No *", ml_vehicle_options, key="mil_vehicle_no")
            ml_mileage_type = st.selectbox("Mileage Type *", MILEAGE_TYPES, key="mil_mileage_type")
            # Oil Type only applies to a service, so it appears when
            # Serviced Mileage is selected (outside the form so it shows live).
            # Lorry and Car/Cab: free-text and optional (their next service
            # interval is fixed and doesn't depend on the oil). Other types
            # (bikes): a required selection, as the oil decides the interval.
            ml_oil_type = ""
            if ml_mileage_type == "Serviced Mileage":
                if is_fixed_interval_type(ml_vehicle_type):
                    ml_oil_type = st.text_input(
                        "Oil Type (optional)",
                        key="mil_oil_type_text",
                        placeholder="Type the oil used (optional)",
                    ).strip()
                else:
                    ml_oil_type = st.selectbox("Oil Type *", OIL_TYPES, key="mil_oil_type")

            with st.form("mileage_entry_form", clear_on_submit=True):
                f1, f2 = st.columns(2)
                with f1:
                    ml_date = st.date_input("Date *", value=date.today(), key="mil_date")
                with f2:
                    ml_mileage = st.text_input("Mileage (KM) *", key="mil_mileage", placeholder="e.g. 12600")

                ml_submitted = st.form_submit_button("💾 Save", type="primary", use_container_width=True)
                if ml_submitted:
                    ml_errors = []
                    ml_value = ""
                    ml_str = ml_mileage.strip()
                    if not ml_vehicle_no:
                        ml_errors.append("Vehicle No is required.")
                    if not ml_str:
                        ml_errors.append("Mileage (KM) is required.")
                    else:
                        try:
                            as_float = float(ml_str)
                            if as_float < 0:
                                ml_errors.append("Mileage (KM) must be a positive number.")
                            else:
                                ml_value = int(as_float) if as_float == int(as_float) else as_float
                        except ValueError:
                            ml_errors.append("Mileage (KM) must be a number.")

                    if ml_errors:
                        for err in ml_errors:
                            st.error(f"❌ {err}")
                    else:
                        append_mileage_record({
                            "Timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S.%f"),
                            "Vehicle No": ml_vehicle_no,
                            "Vehicle Type": ml_vehicle_type,
                            "Mileage Type": ml_mileage_type,
                            "Date": ml_date.isoformat(),
                            "Mileage (KM)": ml_value,
                            "User": st.session_state["username"],
                            "Oil Type": ml_oil_type,
                        })
                        st.success(f"✅ Saved {ml_mileage_type} for {ml_vehicle_no}.")
                        time.sleep(1)
                        st.rerun()

            # All entries for the selected vehicle, shown below the form.
            st.markdown(f"**📜 Mileage Entries — {ml_vehicle_no}**")
            vehicle_entries = mileage_log_df[mileage_log_df["Vehicle No"] == ml_vehicle_no] \
                if not mileage_log_df.empty else mileage_log_df
            if vehicle_entries.empty:
                st.info("No mileage entries recorded for this vehicle yet.")
            else:
                entries_view = (
                    vehicle_entries.sort_values(["Date", "Timestamp"], ascending=[False, False])
                    [["Vehicle No", "Mileage Type", "Date", "Mileage (KM)"]]
                    .rename(columns={"Mileage (KM)": "Mileage"})
                )
                st.dataframe(entries_view, use_container_width=True, hide_index=True)

        # ---------------------------------------------------------------
        # SUB-PHASE 2 — MILEAGE LINES
        # ---------------------------------------------------------------
        elif mileage_phase == "📍 Mileage Lines":
            lines_vehicle_type = st.selectbox("Vehicle Type", vehicle_types_available, key="mil_lines_vehicle_type")
            lines_vehicles = sorted(v for v, t in vehicle_map.items() if t == lines_vehicle_type)

            if not lines_vehicles:
                st.info("No vehicles found for this Vehicle Type.")
            else:
                lines_html = build_mileage_lines_html(mileage_log_df, lines_vehicles, red_below_km=501)
                components.html(
                    lines_html,
                    height=min(len(lines_vehicles) * 114 + 20, 900),
                    scrolling=True,
                )
                st.caption(
                    "Shows each vehicle's latest Serviced Mileage (orange) and latest "
                    "Weekly Mileage (white) with their dates. Next Service is the last serviced "
                    "mileage + 6000 km for Car/Cab, + 5000 km for Lorry LB 8797 and 227 - 4969 "
                    "(+ 6000 km for all other Lorries), and for bikes "
                    "+ 5000 km (Castrol) or + 2500 km (Bajaj Genuine). "
                    "A red Next Service box means 500 km or fewer are left."
                )

        # ---------------------------------------------------------------
        # SUB-PHASE 3 — SERVICED REPORT
        # ---------------------------------------------------------------
        else:
            report_type_options = sorted(
                {t for t in vehicle_map.values() if t and not is_excluded_vehicle_type(t)},
                key=lambda t: (vehicle_type_group_rank(t), t),
            )
            report_type_filter = st.multiselect(
                "Filter by Vehicle Type (leave empty for all)",
                report_type_options,
                key="serviced_report_types",
            )

            report_rows = build_serviced_report_rows(mileage_log_df, report_type_filter)

            if not report_rows:
                st.info("No vehicles match the selected Vehicle Type.")
            else:
                # Same line layout as the Weekly Mileage Inspection > Mileage Lines,
                # one block per vehicle type group (Bike, Lorry, Car/Cab). Each block
                # is ordered by KM left to the next service, smallest first.
                report_group_names = {0: "Bike", 1: "Lorry", 2: "Car/Cab", 3: "No Vehicle Type"}
                for report_rank in sorted({vehicle_type_group_rank(r["Vehicle Type"]) for r in report_rows}):
                    group_vehicles = [
                        r["Vehicle No"] for r in report_rows
                        if vehicle_type_group_rank(r["Vehicle Type"]) == report_rank
                    ]
                    st.markdown(f"##### {report_group_names[report_rank]}")
                    components.html(
                        build_mileage_lines_html(mileage_log_df, group_vehicles, red_below_km=501),
                        height=min(len(group_vehicles) * 114 + 20, 900),
                        scrolling=True,
                    )
                st.caption(
                    "Shows each vehicle's latest Serviced Mileage (orange) and latest "
                    "Weekly Mileage (white) with their dates, and the Next Service mileage (blue). "
                    "The red figures are the KM run since the last service and the KM left until the next "
                    "service (zero or negative means due or overdue). Grouped Bike, Lorry, then Car/Cab, "
                    "with the fewest KM left first within each group. A red Next Service box means "
                    "500 km or fewer are left."
                )

                if _REPORTLAB_AVAILABLE:
                    st.download_button(
                        "⬇️ Download Serviced Report as PDF",
                        build_serviced_report_pdf(report_rows, mileage_log_df),
                        file_name=f"serviced_report_{date.today().isoformat()}.pdf",
                        mime="application/pdf",
                    )
                else:
                    st.warning(
                        "PDF export needs the **reportlab** package. Add `reportlab` to your "
                        "requirements.txt and redeploy, then the download button will appear here."
                    )

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>Vehicle History Monitoring System</p>",
    unsafe_allow_html=True,
)
