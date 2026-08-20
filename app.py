import uuid
import time
import pandas as pd
import streamlit as st
from datetime import datetime, date
from io import BytesIO

import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

# =========================================================================
# CONFIG
# =========================================================================
import streamlit as st

st.set_page_config(page_title="Water Quality & Harvest Report - Data Collection", layout="wide", page_icon="🌐")

CUSTOMER_FILE = "Customer List.xlsx"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# =========================================================================
# STYLE
# Only tables (st.dataframe) get horizontal scrolling — that's native
# Streamlit behavior and needs no extra CSS. Everything else (the
# customer/farm pickers, etc.) should stack into a single column, one
# field after another, on narrow / mobile screens.
# =========================================================================
st.markdown("""
<style>
/* Keep dropdown menus readable/wide enough on small screens */
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
div[data-baseweb="tag"] { white-space: normal !important; max-width: 100% !important; }
span[data-baseweb="tag"] { white-space: normal !important; }

/* Red "Save" buttons (Streamlit primary-type buttons) */
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

/* "Saved" status pill (kept for other parts of the app that may use it) */
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

.pond-card {
    border: 1px solid rgba(128,128,128,0.25);
    border-radius: 10px;
    padding: 0.9rem 1rem 0.3rem 1rem;
    margin-bottom: 0.9rem;
    background-color: rgba(128,128,128,0.03);
}

/* Mobile: stack every row of widgets into a single column, one field
   after another. Tables (st.dataframe / st.data_editor) are NOT built
   from these horizontal blocks, so they are untouched and keep their
   own native horizontal scrollbar. */
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

st.markdown("<h1 style='text-align: center;'>Water Quality & Harvest Report - Data Collection</h1>",
            unsafe_allow_html=True)
st.subheader("KMN Aqua Services")
st.markdown("---")

# =========================================================================
# STATIC SELECTION OPTIONS
# =========================================================================
SPECIES_CULTURE = ["Vannamei", "Monodon", "Other"]
CYCLE_TYPE = ["Soon to be", "Running"]
WATER_COLOR_OPTIONS = ["Milky Color", "Light Green", "Dark Green", "Light Yellow",
                       "Light Brown", "Dark Brown", "Other"]
GRADE_OPTIONS = ["A", "B", "C"]
TECHNICIAN_OPTIONS = ["Mr. Vishmika", "Mr. Ashen", "Mr. Janaka", "Mr. Shashika", "Mr. Janushan"]

DISEASES_OPTIONS = ["WSS", "EHP", "WHITE FECES", "BLACK GILLS", "SOFT SHELL",
                     "MORTALITY ISSUE", "OXYGEN DROP", "GROWTH ISSUE", "ZOOTHAMNIUM", "Other"]
FEED_ISSUE_OPTIONS = ["Over feeding", "Under feeding", "Feed Drop", "Other"]
WATER_QUALITY_OPTIONS = ["PH issue", "Salinity issue", "Alkalinity issue", "Ammonia issue",
                          "Calcium Hardness Issue", "Magnesium Hardness Issue", "Other"]
ENVIRONMENT_ISSUE_OPTIONS = ["Heavy Rain", "High Temperature", "Other"]
MANAGEMENT_ISSUE_OPTIONS = ["Aeration System Failure", "Water Exchange Problem",
                             "Sludge & Bottom Soil Issue", "Chemical/Probiotic Overdose",
                             "Predator Attack", "Other"]

# All issue categories combined, used as the selectable options for the
# "Issues" column directly inside the pond spreadsheet.
ISSUES_OPTIONS = (
    [f"Disease: {x}" for x in DISEASES_OPTIONS] +
    [f"Feed: {x}" for x in FEED_ISSUE_OPTIONS] +
    [f"Water Quality: {x}" for x in WATER_QUALITY_OPTIONS] +
    [f"Environment: {x}" for x in ENVIRONMENT_ISSUE_OPTIONS] +
    [f"Management: {x}" for x in MANAGEMENT_ISSUE_OPTIONS]
)
ISSUES_SEP = ", "

# Options for the new Harvest Details section.
HARVEST_TYPE_OPTIONS = ["Full H", "Partial H"]

# =========================================================================
# FCR REFERENCE TABLE (from FCR.xlsx) — used to auto-calculate
# "Expect Harvest (KG)" and "Survival QTY" from ABW + Feed Per Day +
# Species Culture, the same way the spreadsheet does it with VLOOKUP.
#
# FCR.xlsx layout:
#   Table1 columns:  ABW (g) | F.P. (%) | Feed per day for PL 1lax (KG)
#   I5 (Expect Harvest KG) = H5/VLOOKUP(G5,Table1[],3,FALSE)*100000*G5/1000
#   J5 (Survival QTY)      = H5/VLOOKUP(G5,Table1[],3,FALSE)*100000
#   J11:J18 (Monodon)      = 0.9 * I11:I18 (Vannamei) -> Monodon uses a 0.9
#   factor versus the Vannamei baseline; the sheet has no separate table
#   for "Other", so it also uses the Vannamei baseline.
# =========================================================================
FEED_LOOKUP_TABLE = {
    3: 15.78, 4: 20, 5: 23.75, 6: 27, 7: 29.75, 8: 32, 9: 35.1, 10: 38,
    11: 40.7, 12: 43.2, 13: 45.5, 14: 47.6, 15: 49.5, 16: 51.2, 17: 52.7,
    18: 54, 19: 55.1, 20: 56, 21: 56.7, 22: 57.2, 23: 57.5, 24: 57.6,
    25: 57.5, 26: 57.2, 27: 56.7, 28: 56, 29: 55.1, 30: 54,
}
SPECIES_HARVEST_FACTOR = {"Vannamei": 1.0, "Monodon": 0.9, "Other": 1.0}

def _feed_rate_for_abw(abw):
    """Mirrors the Excel VLOOKUP(ABW, Table1[], 3, FALSE) lookup: exact match
    against the ABW(g) -> Feed/day(KG) reference table. ABW is rounded to
    the nearest whole gram and clamped to the table's 3-30g range, since
    the sheet's table only has whole-gram rows."""
    try:
        abw_val = float(abw)
    except (TypeError, ValueError):
        return None
    if abw_val <= 0:
        return None
    abw_int = max(3, min(30, round(abw_val)))
    return FEED_LOOKUP_TABLE.get(abw_int)

def calc_expected_harvest_and_survival(abw, feed_per_day, species):
    """Returns (Expect Harvest KG, Survival QTY), or (None, None) if ABW /
    Feed Per Day aren't usable numbers yet. Formulas match FCR.xlsx (see
    comment above the FEED_LOOKUP_TABLE definition)."""
    feed_rate = _feed_rate_for_abw(abw)
    if not feed_rate:
        return None, None
    try:
        feed_val = float(feed_per_day)
        abw_val = float(abw)
    except (TypeError, ValueError):
        return None, None
    if feed_val <= 0 or abw_val <= 0:
        return None, None
    factor = SPECIES_HARVEST_FACTOR.get(str(species).strip(), 1.0)
    survival_qty = (feed_val / feed_rate) * 100000 * factor
    expected_harvest_kg = survival_qty * abw_val / 1000
    return round(expected_harvest_kg, 2), round(survival_qty, 0)

COLUMN_ORDER = [
    "Timestamp", "Customer", "Farm Name with Code", "Zone", "Area",
    "Pond Number", "Date", "Species Culture", "Cycle Type",
    "DOC", "Density", "Feed Per Day", "ABW",
    "Expect Harvest (KG)", "Survival QTY",
    "Issues", "Water Color", "Grade", "Remark", "Technician",
    "Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW",
    "Harvest Date 2", "Harvest Type 2", "Harvest KG 2", "Harvest ABW 2",
    "Deleted",
]

# Columns shown/edited in the pond spreadsheet (the rest — Customer, Farm,
# Zone, Area, Pond, Technician, Timestamp — come from the selectors above
# the table and are attached automatically when a row is saved).
POND_COLS = ["Date", "DOC", "Species Culture", "Density", "Feed Per Day", "ABW",
             "Expect Harvest (KG)", "Survival QTY",
             "Cycle Type", "Issues", "Water Color", "Grade", "Remark"]

# =========================================================================
# GOOGLE SHEETS BACKEND
# =========================================================================
def _gsheet_configured():
    return "gcp_service_account" in st.secrets and "gsheet" in st.secrets and "sheet_id" in st.secrets["gsheet"]

@st.cache_resource(show_spinner=False)
def get_worksheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet_id = st.secrets["gsheet"]["sheet_id"]
    worksheet_name = st.secrets["gsheet"].get("worksheet_name", "WaterQualityData")
    sh = client.open_by_key(sheet_id)
    try:
        ws = sh.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_name, rows=2000, cols=len(COLUMN_ORDER) + 2)
        ws.append_row(COLUMN_ORDER, value_input_option="USER_ENTERED")
    # Make sure the header row matches what we expect (self-heals a blank sheet,
    # and adds any new columns — e.g. Harvest Date / Harvest Type / Harvest KG
    # / Harvest ABW / Harvest Date 2 / Harvest Type 2 / Harvest KG 2 /
    # Harvest ABW 2 / Expect Harvest (KG) / Survival QTY — to a
    # sheet that was created before they existed).
    header = ws.row_values(1)
    if header != COLUMN_ORDER:
        ws.update("A1", [COLUMN_ORDER])
    return ws

def _load_data_cached(data_version, sheet_id):
    # Always read fresh from the Google Sheet — no caching. This is what
    # makes edits made directly in the Google Sheet (e.g. correcting a
    # Density value, or editing any other saved cell) show up back in the
    # app's pond history the next time the page loads/reruns, instead of
    # being stuck showing the first value that was ever saved.
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

def bump_data_version():
    st.session_state["_data_version"] = st.session_state.get("_data_version", 0) + 1

def load_data():
    """Returns the sheet's data with any soft-deleted rows filtered out. The
    rows themselves are never removed from the Google Sheet — the Deleted
    column just flags them, so this is the single place that hides them
    from the rest of the app (pond dropdown, history table, full export)."""
    sheet_id = st.secrets["gsheet"]["sheet_id"]
    df = _load_data_cached(st.session_state.get("_data_version", 0), sheet_id)
    if "Deleted" in df.columns:
        is_deleted = df["Deleted"].astype(str).str.strip().str.lower().isin(["yes", "true", "1"])
        df = df[~is_deleted].reset_index(drop=True)
    return df

def mark_deleted_by_timestamp(timestamp):
    """Soft-delete: flag the row as Deleted='Yes' in the Google Sheet. The
    row and all its data stay in the sheet permanently for audit purposes —
    it just gets hidden from the app (see load_data above) so it disappears
    from the screen and never reappears, even after a refresh."""
    ws = get_worksheet()
    cell = ws.find(str(timestamp), in_column=1)
    if cell:
        deleted_col_index = COLUMN_ORDER.index("Deleted") + 1
        ws.update_cell(cell.row, deleted_col_index, "Yes")
        bump_data_version()

def update_harvest_by_timestamp(timestamp, harvest_date, harvest_type, harvest_kg="", harvest_abw="", slot=1):
    """Writes Harvest Date/Type/KG/ABW onto the existing saved row that
    matches this Timestamp (i.e. an existing Pond Details record), instead
    of creating a brand-new row. Harvest KG and Harvest ABW are optional —
    an empty string just clears that cell. `slot` picks which set of
    Harvest columns to write: 1 -> "Harvest Date/Type/KG/ABW" (the first
    harvest), 2 -> "Harvest Date 2/Type 2/KG 2/ABW 2" (a second harvest for
    the same pond row, e.g. a partial harvest followed by a later one)."""
    ws = get_worksheet()
    cell = ws.find(str(timestamp), in_column=1)
    if not cell:
        return False
    suffix = "" if slot == 1 else " 2"
    harvest_date_col = COLUMN_ORDER.index(f"Harvest Date{suffix}") + 1
    harvest_type_col = COLUMN_ORDER.index(f"Harvest Type{suffix}") + 1
    harvest_kg_col = COLUMN_ORDER.index(f"Harvest KG{suffix}") + 1
    harvest_abw_col = COLUMN_ORDER.index(f"Harvest ABW{suffix}") + 1
    ws.update_cell(cell.row, harvest_date_col, harvest_date)
    ws.update_cell(cell.row, harvest_type_col, harvest_type)
    ws.update_cell(cell.row, harvest_kg_col, harvest_kg)
    ws.update_cell(cell.row, harvest_abw_col, harvest_abw)
    bump_data_version()
    return True

def append_record(record):
    ws = get_worksheet()
    row = [str(record.get(c, "")) for c in COLUMN_ORDER]
    ws.append_row(row, value_input_option="USER_ENTERED")
    bump_data_version()

def append_records(records):
    """Append several records to the sheet in a single batched call."""
    if not records:
        return
    ws = get_worksheet()
    rows = [[str(r.get(c, "")) for c in COLUMN_ORDER] for r in records]
    ws.append_rows(rows, value_input_option="USER_ENTERED")
    bump_data_version()

def update_record_by_timestamp(timestamp, record):
    ws = get_worksheet()
    cell = ws.find(str(timestamp), in_column=1)
    row = [str(record.get(c, "")) for c in COLUMN_ORDER]
    if cell:
        end_a1 = rowcol_to_a1(cell.row, len(COLUMN_ORDER))
        ws.update(f"A{cell.row}:{end_a1}", [row], value_input_option="USER_ENTERED")
    else:
        ws.append_row(row, value_input_option="USER_ENTERED")
    bump_data_version()

def delete_record_by_timestamp(timestamp):
    ws = get_worksheet()
    cell = ws.find(str(timestamp), in_column=1)
    if cell:
        ws.delete_rows(cell.row)
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
# GOOGLE SHEETS SETUP CHECK
# =========================================================================
if not _gsheet_configured():
    st.error("❌ Google Sheets is not configured yet.")
    with st.expander("⚙️ How to connect this app to a Google Sheet", expanded=True):
        st.markdown(
            "1. Create a Google Cloud project, enable the **Google Sheets API** and "
            "**Google Drive API**, and create a **Service Account**.\n"
            "2. Create a JSON key for that service account and copy its contents.\n"
            "3. Create a Google Sheet, and share it (Editor access) with the service "
            "account's `client_email` address.\n"
            "4. Add the following to your app's `.streamlit/secrets.toml`:\n"
        )
        st.code(
            '[gcp_service_account]\n'
            'type = "service_account"\n'
            'project_id = "..."\n'
            'private_key_id = "..."\n'
            'private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"\n'
            'client_email = "...@....iam.gserviceaccount.com"\n'
            'client_id = "..."\n'
            'auth_uri = "https://accounts.google.com/o/oauth2/auth"\n'
            'token_uri = "https://oauth2.googleapis.com/token"\n'
            'auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"\n'
            'client_x509_cert_url = "..."\n\n'
            '[gsheet]\n'
            'sheet_id = "the-id-from-the-sheet-url"\n'
            'worksheet_name = "WaterQualityData"\n',
            language="toml",
        )
    st.stop()

try:
    get_worksheet()
except Exception as e:
    st.error(f"❌ Could not connect to the Google Sheet. Check your secrets and sharing settings.\n\n{e}")
    st.stop()

# =========================================================================
# LOAD CUSTOMER LIST
# =========================================================================
@st.cache_data
def load_customer_data():
    return pd.read_excel(CUSTOMER_FILE)

try:
    customer_df = load_customer_data()
except Exception as e:
    st.error(f"❌ Could not load '{CUSTOMER_FILE}'. Make sure it's in the app folder. ({e})")
    st.stop()

REQUIRED_COLS = ["Customer Name", "Farm Name with Code", "Zone", "Area"]
missing_cols = [c for c in REQUIRED_COLS if c not in customer_df.columns]
if missing_cols:
    st.error(f"❌ 'Customer List.xlsx' is missing required column(s): {', '.join(missing_cols)}")
    st.stop()

for _col in REQUIRED_COLS:
    customer_df[_col] = customer_df[_col].apply(
        lambda v: "" if pd.isna(v) else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))
    )

all_customers = sorted(customer_df["Customer Name"].replace("", pd.NA).dropna().unique().tolist())
all_zones = sorted(customer_df["Zone"].replace("", pd.NA).dropna().unique().tolist())
all_areas = sorted(customer_df["Area"].replace("", pd.NA).dropna().unique().tolist())

# =========================================================================
# EXISTING PONDS LOOKUP
# =========================================================================
def get_existing_ponds(customer, farm):
    df = load_data()
    required = {"Customer", "Farm Name with Code", "Pond Number"}
    if len(df) > 0 and required.issubset(df.columns):
        farm_hist = df[(df["Customer"] == customer) & (df["Farm Name with Code"] == farm)]
        if len(farm_hist) > 0:
            return sorted(
                [p for p in farm_hist["Pond Number"].dropna().unique().tolist() if str(p).strip() != ""]
            )
    return []

# =========================================================================
# STEP 1: CUSTOMER
# =========================================================================
st.subheader("📋 Enter Water Quality Data")

col1, col2 = st.columns(2)
with col1:
    customer = st.selectbox("Customer Name *", all_customers, key="customer_select")

farm_options = sorted(
    customer_df.loc[customer_df["Customer Name"] == customer, "Farm Name with Code"]
    .dropna().unique().tolist()
)
if not farm_options:
    farm_options = ["-- No farms found for this customer --"]

with col2:
    farm = st.selectbox("Farm Name with Code *", farm_options, key=f"farm_select_{customer}")

farm_row_match = customer_df[
    (customer_df["Customer Name"] == customer) & (customer_df["Farm Name with Code"] == farm)
]
if len(farm_row_match) > 0 and "Marketing Manager" in customer_df.columns:
    mm = farm_row_match.iloc[0].get("Marketing Manager", "")
    if str(mm).strip():
        st.caption(f"Marketing Manager: {mm}")

# =========================================================================
# STEP 2: ZONE / AREA
# =========================================================================
default_zone = farm_row_match.iloc[0]["Zone"] if len(farm_row_match) > 0 else (all_zones[0] if all_zones else "")
default_area = farm_row_match.iloc[0]["Area"] if len(farm_row_match) > 0 else (all_areas[0] if all_areas else "")

col3, col4 = st.columns(2)
with col3:
    zone_index = all_zones.index(default_zone) if default_zone in all_zones else 0
    zone = st.selectbox("Zone *", all_zones, index=zone_index, key=f"zone_select_{farm}")
with col4:
    area_index = all_areas.index(default_area) if default_area in all_areas else 0
    area = st.selectbox("Area *", all_areas, index=area_index, key=f"area_select_{farm}")

# =========================================================================
# STEP 3: TECHNICIAN
# =========================================================================
technician = st.selectbox("Technician *", TECHNICIAN_OPTIONS, key="technician_select")

# =========================================================================
# STEP 4: POND DETAILS — pond selection bar, saved (read-only) history,
# then a small spreadsheet for adding new records
# =========================================================================
st.markdown("---")
st.markdown("#### 🐟 Pond Details")

ADD_NEW_LABEL = "➕ Add New Pond"
existing_ponds = get_existing_ponds(customer, farm)

if existing_ponds:
    st.caption(f"{len(existing_ponds)} pond(s) on record for this farm.")
    pond_bar_options = existing_ponds + [ADD_NEW_LABEL]
else:
    st.info("No ponds found for this farm yet. Choose **Add New Pond** below to add the first one.")
    pond_bar_options = [ADD_NEW_LABEL]

selected_pond_choice = st.selectbox("Select Pond *", pond_bar_options, key=f"pond_bar_{farm}")

if selected_pond_choice == ADD_NEW_LABEL:
    pond_number = st.text_input("New Pond Number *", key=f"new_pond_number_{farm}").strip()
else:
    pond_number = selected_pond_choice
    st.caption(f"Adding / editing records for Pond **{pond_number}**")

widget_scope = f"{farm}_{selected_pond_choice}"

# =========================================================================
# LOAD THIS POND'S HISTORY FROM THE GOOGLE SHEET
# (this always reflects only the selected customer + farm + pond; switching
# ponds never touches any other pond's saved data)
# =========================================================================
df_pond_hist_full = load_data()
required_cols = {"Customer", "Farm Name with Code", "Pond Number"}
if pond_number and len(df_pond_hist_full) > 0 and required_cols.issubset(df_pond_hist_full.columns):
    df_pond_hist_full = df_pond_hist_full[
        (df_pond_hist_full["Customer"] == customer)
        & (df_pond_hist_full["Farm Name with Code"] == farm)
        & (df_pond_hist_full["Pond Number"] == pond_number)
    ].copy()
else:
    df_pond_hist_full = pd.DataFrame(columns=COLUMN_ORDER)

prev_date = None
prev_doc = None
prev_species = None
prev_cycle = None
prev_density = None
if len(df_pond_hist_full) > 0 and "Date" in df_pond_hist_full.columns:
    df_pond_hist_full["_ParsedDate"] = pd.to_datetime(df_pond_hist_full["Date"], errors="coerce")
    df_pond_hist_full = df_pond_hist_full.sort_values(by="_ParsedDate").reset_index(drop=True)
    if df_pond_hist_full["_ParsedDate"].notna().any():
        latest_idx = df_pond_hist_full["_ParsedDate"].last_valid_index()
        prev_date = df_pond_hist_full.loc[latest_idx, "_ParsedDate"].date()
        prev_doc = to_number(df_pond_hist_full.loc[latest_idx, "DOC"], as_int=True)
        prev_species = str(df_pond_hist_full.loc[latest_idx].get("Species Culture") or "").strip() or None
        prev_cycle = str(df_pond_hist_full.loc[latest_idx].get("Cycle Type") or "").strip() or None
        prev_density = to_number(df_pond_hist_full.loc[latest_idx, "Density"], as_int=True)
        # A "Soon to be" row does not seed the DOC auto-calc chain for the
        # next row that gets added after it — the cycle hasn't actually
        # started yet, so there's nothing to count days-of-culture from.
        if prev_cycle == "Soon to be":
            prev_date, prev_doc = None, None

default_species = prev_species if prev_species in SPECIES_CULTURE else SPECIES_CULTURE[0]
default_cycle = prev_cycle if prev_cycle in CYCLE_TYPE else "Running"
default_density = prev_density if prev_density else 0

st.markdown(f"##### 📜 History — Pond {pond_number}" if pond_number else "##### 📜 History")

@st.fragment
def _pond_editor_fragment():
    # Wrapped in @st.fragment so that editing a cell only reruns this piece
    # of the page instead of the entire script (customer/farm lookups,
    # Google Sheet reads, etc.) — this cuts down on how much of the page
    # gets touched per keystroke, which helps with the grid resetting its
    # scroll position (a known Streamlit limitation with dependent columns
    # like our DOC auto-calc / Status: https://github.com/streamlit/streamlit/issues/10181).
    # "Timestamp" is kept as a hidden internal column so we know which rows
    # are already saved (non-blank Timestamp -> locked, Status = Saved) vs.
    # brand-new rows added in this session (blank Timestamp -> editable,
    # Status = New (unsaved)).
    # The Sheet stores one "Issues" cell per row (issues joined with
    # ISSUES_SEP). The spreadsheet lets you pick more than one issue via a
    # single Multiselect column — so here the saved "Issues" string is
    # split back out into a Python list for display/editing.
    EDITOR_POND_COLS = ["Date", "DOC", "Density", "Feed Per Day", "ABW",
                         "Expect Harvest (KG)", "Survival QTY", "Species Culture",
                         "Cycle Type", "Issues", "Water Color", "Grade", "Remark"]
    display_cols = ["Timestamp"] + EDITOR_POND_COLS
    if len(df_pond_hist_full) > 0:
        _src = df_pond_hist_full.copy()
        if "Issues" in _src.columns:
            _src["Issues"] = _src["Issues"].fillna("").astype(str).apply(
                lambda s: [p.strip() for p in s.split(ISSUES_SEP) if p.strip()]
            )
        else:
            _src["Issues"] = [[] for _ in range(len(_src))]
        existing_pond_cols = [c for c in display_cols if c in _src.columns]
        df_pond_hist_display = _src[existing_pond_cols].copy()
    else:
        df_pond_hist_display = pd.DataFrame(columns=display_cols)
    original_row_count = len(df_pond_hist_display)

    if original_row_count == 0:
        st.info(f"No history yet for Pond {pond_number}. Add its first record in the spreadsheet below.")

    _TEXT_COLS = ["Timestamp", "ABW", "Species Culture", "Cycle Type",
                  "Water Color", "Grade", "Remark"]
    _NUM_COLS = ["DOC", "Density", "Feed Per Day", "Expect Harvest (KG)", "Survival QTY"]

    def _normalize_issues_cell(v):
        if isinstance(v, list):
            return [str(x).strip() for x in v if str(x).strip()]
        if v is None:
            return []
        try:
            if pd.isna(v):
                return []
        except (TypeError, ValueError):
            pass
        v = str(v).strip()
        return [p.strip() for p in v.split(ISSUES_SEP) if p.strip()] if v else []

    def _normalize_pond_dtypes(df):
        df = df.copy()
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        for numcol in _NUM_COLS:
            if numcol in df.columns:
                df[numcol] = pd.to_numeric(df[numcol], errors="coerce")
        for c in _TEXT_COLS:
            if c in df.columns:
                df[c] = df[c].fillna("").astype(str)
        if "Issues" in df.columns:
            df["Issues"] = df["Issues"].apply(_normalize_issues_cell)
        return df

    if len(df_pond_hist_display) > 0:
        editor_df = _normalize_pond_dtypes(df_pond_hist_display)
    else:
        _empty_dtypes = {
            "Timestamp": "object", "Date": "object",
            "DOC": "float64", "Density": "float64", "Feed Per Day": "float64",
            "ABW": "object",
            "Expect Harvest (KG)": "float64", "Survival QTY": "float64",
            "Species Culture": "object", "Cycle Type": "object",
            "Issues": "object",
            "Water Color": "object", "Grade": "object",
            "Remark": "object",
        }
        editor_df = pd.DataFrame({c: pd.Series(dtype=_empty_dtypes.get(c, "object")) for c in display_cols})
        editor_df = _normalize_pond_dtypes(editor_df)

    editor_df["Status"] = editor_df["Timestamp"].apply(
        lambda t: "✅ Saved" if str(t).strip() else "🆕 New (unsaved)"
    )

    st.caption("Rows marked **✅ Saved** are already in the Google Sheet — their fields can't be "
               "edited here, but you can delete one (🗑️ row menu) and it will disappear from this "
               "screen for good, even after a refresh. The underlying data stays in the Google Sheet "
               "(it's flagged, not erased). Add new rows at the bottom (**🆕 New (unsaved)**) and "
               "click Save to write them.")

    column_config = {
        "Date": st.column_config.DateColumn("Date *", required=True),
        "DOC": st.column_config.NumberColumn("DOC (auto)", help="Filled in automatically once you pick a Date — edit it to override", step=1),
        "Density": st.column_config.NumberColumn("Density", step=1, default=default_density),
        "Feed Per Day": st.column_config.NumberColumn("Feed/Day"),
        "ABW": st.column_config.TextColumn("ABW"),
        "Expect Harvest (KG)": st.column_config.NumberColumn(
            "Expect Harvest (KG) (auto)", format="%.2f",
            help="Auto-calculated from ABW, Feed/Day & Species Culture (FCR.xlsx reference) — edit to override"),
        "Survival QTY": st.column_config.NumberColumn(
            "Survival QTY (auto)", format="%.0f",
            help="Auto-calculated from ABW, Feed/Day & Species Culture (FCR.xlsx reference) — edit to override"),
        "Species Culture": st.column_config.SelectboxColumn("Species Culture *", options=SPECIES_CULTURE,
                                                              required=True, default=default_species),
        "Cycle Type": st.column_config.SelectboxColumn("Cycle Type *", options=CYCLE_TYPE,
                                                         required=True, default=default_cycle),
        "Issues": st.column_config.MultiselectColumn("Issues", options=ISSUES_OPTIONS, default=[]),
        "Water Color": st.column_config.SelectboxColumn("Water Color", options=WATER_COLOR_OPTIONS, required=False),
        "Grade": st.column_config.SelectboxColumn("Grade", options=GRADE_OPTIONS, required=False),
        "Remark": st.column_config.TextColumn("Remark"),
        "Status": st.column_config.TextColumn("Status", disabled=True),
    }
    # "Timestamp" is deliberately left out of column_order so it stays in
    # the underlying data (for locking rows / matching to sheet rows)
    # without being shown or editable — "Status" is shown last instead.
    # This is purely the VISUAL order the grid displays columns in — it's
    # independent of EDITOR_POND_COLS / the working dataframe's own column
    # order and of the Google Sheet's COLUMN_ORDER (data storage is
    # unaffected). Species Culture is placed 3rd, right after Date/DOC.
    column_order = ["Date", "DOC", "Species Culture", "Density", "Feed Per Day",
                     "ABW", "Expect Harvest (KG)", "Survival QTY",
                     "Cycle Type", "Issues", "Water Color", "Grade", "Remark", "Status"]

    editor_key = f"editor_{widget_scope}"
    working_key = f"__pond_working_{widget_scope}"
    working_sig_key = f"__pond_working_sig_{widget_scope}"
    # Tracks how many trailing rows in `working_key` came from the widget's
    # own `added_rows` list the last time we folded it in. `added_rows`
    # is NOT a "what's new since last time" delta — it's Streamlit's
    # running list of every row added since the grid was last remounted,
    # and each entry keeps gaining fields as more of its cells are typed
    # into. Blindly appending everything in `added_rows` on every
    # on_change therefore appended the same new row again and again,
    # producing duplicate saved records. This counter lets us drop the
    # rows we appended last time and rebuild the tail fresh each time
    # instead of stacking on top of it.
    added_count_key = f"__pond_added_count_{widget_scope}"

    def _parse_cell_date(val):
        if val is None:
            return None
        try:
            if pd.isna(val):
                return None
        except (TypeError, ValueError):
            pass
        if isinstance(val, pd.Timestamp):
            return val.date()
        if isinstance(val, date):
            return val
        parsed = pd.to_datetime(val, errors="coerce")
        return parsed.date() if pd.notna(parsed) else None

    def _doc_is_blank(val):
        return val is None or val == "" or (isinstance(val, float) and pd.isna(val))

    def _recompute_docs(df, seed_date, seed_doc):
        """Return a copy of df with every blank DOC cell filled in, using
        (previous row's DOC + days since previous row's Date). Rows are
        chained in chronological (Date) order — NOT in whatever order they
        happen to sit in the grid — so DOC still auto-calculates correctly
        for "the next record" even if a new row is typed in with an
        earlier date than another still-unsaved row that was added before
        it (e.g. catching up on a couple of missed days out of order). A
        row whose Cycle Type is "Soon to be" does not seed the following
        row's auto-calculation — the running chain is reset right after
        such a row, so the following row needs its own DOC (or starts a
        fresh chain from its own Date/DOC) instead of inheriting days
        counted from a cycle that hasn't actually started yet."""
        df = df.reset_index(drop=True).copy()
        # Chronological processing order (stable: ties/blank dates keep
        # their original relative position instead of being reshuffled).
        dated_positions = [(pos, _parse_cell_date(df.at[pos, "Date"])) for pos in range(len(df))]
        chain_order = [pos for pos, _ in sorted(
            dated_positions, key=lambda item: (item[1] is None, item[1])
        )]
        first_pos = chain_order[0] if chain_order else None
        run_date, run_doc = seed_date, seed_doc
        for i in chain_order:
            row_date = _parse_cell_date(df.at[i, "Date"])
            if row_date is None:
                continue
            row_is_soon_to_be = str(df.at[i, "Cycle Type"]).strip() == "Soon to be"
            if run_date is None or run_doc is None:
                if not _doc_is_blank(df.at[i, "DOC"]):
                    try:
                        run_doc = int(df.at[i, "DOC"])
                        run_date = row_date
                    except Exception:
                        pass
                if row_is_soon_to_be:
                    run_date, run_doc = None, None
                continue
            if i == first_pos and not _doc_is_blank(df.at[i, "DOC"]):
                run_doc = int(df.at[i, "DOC"])
                run_date = row_date
                if row_is_soon_to_be:
                    run_date, run_doc = None, None
                continue
            computed = int(run_doc) + (row_date - run_date).days
            if _doc_is_blank(df.at[i, "DOC"]):
                df.at[i, "DOC"] = computed
                run_doc = computed
            else:
                try:
                    run_doc = int(df.at[i, "DOC"])
                except Exception:
                    df.at[i, "DOC"] = computed
                    run_doc = computed
            run_date = row_date
            if row_is_soon_to_be:
                run_date, run_doc = None, None
        return df

    def _recompute_harvest_survival_row(df, i):
        """Fill Expect Harvest (KG) / Survival QTY for row i, in place, if
        either is still blank — using that same row's ABW, Feed Per Day and
        Species Culture (see calc_expected_harvest_and_survival, based on
        the FCR.xlsx reference sheet). Never overwrites a value the user
        already typed/edited."""
        harvest_blank = _doc_is_blank(df.at[i, "Expect Harvest (KG)"])
        survival_blank = _doc_is_blank(df.at[i, "Survival QTY"])
        if not harvest_blank and not survival_blank:
            return
        hv, surv = calc_expected_harvest_and_survival(
            df.at[i, "ABW"], df.at[i, "Feed Per Day"], df.at[i, "Species Culture"]
        )
        if harvest_blank and hv is not None:
            df.at[i, "Expect Harvest (KG)"] = hv
        if survival_blank and surv is not None:
            df.at[i, "Survival QTY"] = surv

    def _recompute_harvest_survival(df):
        df = df.reset_index(drop=True).copy()
        for i in range(len(df)):
            _recompute_harvest_survival_row(df, i)
        return df

    def _recompute_status(df):
        df = df.copy()
        df["Status"] = df["Timestamp"].apply(
            lambda t: "✅ Saved" if str(t).strip() else "🆕 New (unsaved)"
        )
        return df

    # Keep our own persistent copy of the table (separate from the widget's
    # internal state). `editor_df` above was just rebuilt from a fresh
    # Google Sheet read, so it always holds the latest values for every
    # already-saved row. Switching pond/farm/customer starts a clean copy;
    # staying on the same pond keeps any new/unsaved rows the user is
    # entering, but always swaps in this fresh `editor_df` for the saved
    # rows — so an edit made directly in the Google Sheet (e.g. a
    # corrected Density) shows up right away instead of being stuck on
    # whatever value the app first loaded.
    switch_signature = (customer, farm, pond_number)
    if st.session_state.get(working_sig_key) != switch_signature:
        st.session_state[working_key] = editor_df.copy()
        st.session_state[working_sig_key] = switch_signature
        st.session_state[added_count_key] = 0
        if editor_key in st.session_state:
            del st.session_state[editor_key]
    else:
        existing_working = st.session_state.get(working_key)
        if existing_working is not None:
            unsaved_rows = existing_working[
                existing_working["Timestamp"].astype(str).str.strip() == ""
            ].copy()
            st.session_state[working_key] = pd.concat(
                [editor_df.copy(), unsaved_rows], ignore_index=True
            )
        else:
            st.session_state[working_key] = editor_df.copy()
        # NOTE: deliberately NOT deleting st.session_state[editor_key] here.
        # This branch runs on every rerun while staying on the same pond
        # (including the rerun right after a cell edit). apply_editor_changes
        # already folds edits into working_key and clears the widget's own
        # edited_rows/added_rows/deleted_rows in place, so the widget's
        # tracked state is already consistent with working_key. Deleting the
        # key here forced st.data_editor to remount from scratch on every
        # keystroke, which is what reset the grid's scroll position/focus.

    def apply_editor_changes():
        """on_change callback: fold whatever was just typed into our working
        copy, but IGNORE any attempted edit/delete on a row that's already
        saved (non-blank Timestamp) — those are locked.

        DOC auto-calc and the Status column only ever depend on the Date/DOC
        chain and on rows being added/removed; Expect Harvest (KG) / Survival
        QTY only depend on that same row's ABW / Feed Per Day / Species
        Culture. So we only run the full rebuild (_normalize_pond_dtypes /
        _recompute_docs / _recompute_harvest_survival / _recompute_status)
        when one of those actually happened. A plain edit to any other column
        (Density, Cycle, Water Color, Grade, Remark, Issues) just patches
        that one cell in place on the SAME dataframe object instead of
        rebuilding a fresh one — passing a materially-identical dataframe
        back into the grid on every keystroke was what made it reset its
        scroll position back to the first columns each time."""
        state = st.session_state.get(editor_key)
        if not state:
            return

        edited_rows = state.get("edited_rows", {})
        added_rows = state.get("added_rows", [])
        deleted_rows = state.get("deleted_rows", [])

        RECOMPUTE_TRIGGER_COLS = ("Date", "DOC", "ABW", "Feed Per Day", "Species Culture", "Cycle Type")
        needs_recompute = bool(added_rows) or bool(deleted_rows) or any(
            any(col in changes for col in RECOMPUTE_TRIGGER_COLS) for changes in edited_rows.values()
        )

        df = st.session_state[working_key]
        df.reset_index(drop=True, inplace=True)

        for idx, changes in edited_rows.items():
            idx = int(idx)
            if idx < len(df):
                if str(df.at[idx, "Timestamp"]).strip():
                    continue  # locked: already saved, ignore the edit
                for col, val in changes.items():
                    if col in df.columns and col not in ("Timestamp", "Status"):
                        df.at[idx, col] = val

        # `added_rows` is Streamlit's running list of new rows, not a
        # delta since the last on_change — the same not-yet-committed row
        # reappears here (with more fields filled in) every time another
        # of its cells is edited. Remove whatever we appended from it on
        # the previous run before appending the current list, so a
        # not-yet-saved row is only ever represented once in `df` instead
        # of accumulating a duplicate on every keystroke.
        prev_added_len = st.session_state.get(added_count_key, 0)
        if prev_added_len:
            df = df.iloc[: max(0, len(df) - prev_added_len)].reset_index(drop=True)

        for new_row in added_rows:
            row = {c: new_row.get(c) for c in df.columns}
            row["Timestamp"] = ""  # brand-new row -> unsaved until Save is clicked
            if not isinstance(row.get("Issues"), list):
                row["Issues"] = []
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

        st.session_state[added_count_key] = len(added_rows)

        for idx in sorted(deleted_rows, reverse=True):
            if idx < len(df):
                ts = str(df.at[idx, "Timestamp"]).strip()
                if ts:
                    # Already-saved row: soft-delete in the Google Sheet
                    # (flags it, keeps the underlying data), then remove it
                    # from the screen — it will stay gone after a refresh.
                    mark_deleted_by_timestamp(ts)
                df = df.drop(index=idx).reset_index(drop=True)
                # A deleted row is gone for good — if it happened to be
                # one of the still-pending "added" rows, don't let a
                # later on_change try to re-append it from added_rows.
                if idx >= len(df) - st.session_state.get(added_count_key, 0):
                    st.session_state[added_count_key] = max(
                        0, st.session_state.get(added_count_key, 0) - 1
                    )

        if needs_recompute:
            df = _normalize_pond_dtypes(df)
            df = _recompute_docs(df, prev_date, prev_doc)
            df = _recompute_harvest_survival(df)
            df = _recompute_status(df)

        st.session_state[working_key] = df
        # Clear only the widget's own edit-tracking (not the whole widget
        # state) so the grid doesn't get fully remounted on every keystroke
        # — that remount was what reset the scroll position/focus after
        # each cell edit.
        state["edited_rows"] = {}
        state["added_rows"] = []
        state["deleted_rows"] = []

    edited_df = st.data_editor(
        st.session_state[working_key],
        column_config=column_config,
        column_order=column_order,
        num_rows="dynamic",
        use_container_width=True,
        height=320,
        key=editor_key,
        on_change=apply_editor_changes,
    )

    saving_flag_key = f"__saving_{widget_scope}"
    save_in_progress = st.session_state.get(saving_flag_key, False)
    save_clicked = st.button(
        "💾 Save New Records", use_container_width=True,
        type="primary", key=f"save_{widget_scope}",
        disabled=save_in_progress,
    )

    if save_clicked and save_in_progress:
        # A save for this pond is already being processed (e.g. the person
        # double-clicked, or clicked again during the network round-trip
        # to Google Sheets) — ignore this second click instead of writing
        # the same rows to the sheet twice.
        save_clicked = False

    if save_clicked:
        # Lock the button immediately so a second click while this save is
        # still in flight (network calls to Google Sheets can take a few
        # seconds) can't append the same new rows a second time.
        st.session_state[saving_flag_key] = True

        errors = []
        if not customer or not farm or not zone or not area or not technician:
            errors.append("Please fill in all required top-level fields (marked with *)")
        if not pond_number:
            errors.append("Pond Number is required")

        new_records = []
        running_prev_date = prev_date
        running_prev_doc = prev_doc
        now_base = datetime.now()

        rows_all = st.session_state[working_key].reset_index(drop=True)
        # De-duplicate defensively: if, despite the added_rows fix above,
        # more than one row in the current unsaved batch is byte-for-byte
        # identical (every editable field the same), only keep the first
        # occurrence — a real duplicate save should never fall through to
        # the sheet write.
        _saved_mask0 = rows_all["Timestamp"].astype(str).str.strip() != ""
        _dedupe_cols = [c for c in rows_all.columns if c != "Timestamp"]

        def _hashable(v):
            return tuple(v) if isinstance(v, list) else v

        _keep_mask = pd.Series(True, index=rows_all.index)
        _seen = set()
        for idx, _r in rows_all[~_saved_mask0].iterrows():
            key = tuple(_hashable(_r[c]) for c in _dedupe_cols)
            if key in _seen:
                _keep_mask.loc[idx] = False
            else:
                _seen.add(key)
        rows_all = rows_all[_keep_mask].reset_index(drop=True)

        # Process rows in chronological (Date) order rather than grid/
        # insertion order: saved rows already come pre-sorted by Date, but
        # brand-new unsaved rows are simply appended to the bottom of the
        # grid in whatever order they were typed. If a person enters two
        # new rows out of date order (e.g. back-filling an earlier day
        # after already adding a later one), DOC needs to chain off the
        # true previous day, not off whichever row happens to sit above it
        # on screen — otherwise DOC for that "next" row comes out wrong or
        # never gets filled in.
        _saved_mask = rows_all["Timestamp"].astype(str).str.strip() != ""
        _saved_part = rows_all[_saved_mask]
        _unsaved_part = rows_all[~_saved_mask].copy()
        if len(_unsaved_part) > 0:
            _unsaved_part["_sort_date"] = _unsaved_part["Date"].apply(_parse_cell_date)
            _unsaved_part = _unsaved_part.sort_values(
                by="_sort_date", kind="stable", na_position="last"
            ).drop(columns=["_sort_date"])
        rows_all = pd.concat([_saved_part, _unsaved_part], ignore_index=True)

        any_new_rows = False

        for i, row in rows_all.iterrows():
            row_label = f"Row {i + 1}"
            is_saved = str(row.get("Timestamp") or "").strip() != ""

            row_date = _parse_cell_date(row.get("Date"))

            if is_saved:
                # Already in the sheet — skip re-saving it, but keep the
                # running DOC/Date chain going so any new rows after it
                # auto-calculate correctly. A "Soon to be" row does not
                # seed the chain for whatever row comes after it, since the
                # cycle hasn't actually started yet.
                doc_val = row.get("DOC")
                if row_date is not None and not _doc_is_blank(doc_val):
                    running_prev_date, running_prev_doc = row_date, int(doc_val)
                if str(row.get("Cycle Type") or "").strip() == "Soon to be":
                    running_prev_date, running_prev_doc = None, None
                continue

            any_new_rows = True

            # --- Date ---
            if row_date is None:
                errors.append(f"{row_label}: Date is required")
                continue

            # --- DOC (auto-calculate if blank) ---
            doc_raw = row.get("DOC")
            doc_blank = _doc_is_blank(doc_raw)
            if doc_blank:
                if running_prev_date is not None and running_prev_doc is not None:
                    doc_final = int(running_prev_doc) + (row_date - running_prev_date).days
                else:
                    errors.append(f"{row_label}: DOC is required (no earlier record to auto-calculate from)")
                    continue
            else:
                doc_final = to_number(doc_raw, as_int=True)

            # --- Required dropdown fields ---
            species_val_row = str(row.get("Species Culture") or "").strip()
            cycle_val_row = str(row.get("Cycle Type") or "").strip()
            water_color_val_row = str(row.get("Water Color") or "").strip()
            grade_val_row = str(row.get("Grade") or "").strip()
            if not species_val_row:
                errors.append(f"{row_label}: Species Culture is required"); continue
            if not cycle_val_row:
                errors.append(f"{row_label}: Cycle Type is required"); continue

            # --- Expect Harvest (KG) / Survival QTY (auto-calculate if
            # blank, from ABW + Feed Per Day + Species Culture — same
            # formula as the FCR.xlsx reference sheet) ---
            expect_harvest_val = row.get("Expect Harvest (KG)")
            survival_qty_val = row.get("Survival QTY")
            if _doc_is_blank(expect_harvest_val) or _doc_is_blank(survival_qty_val):
                calc_hv, calc_surv = calc_expected_harvest_and_survival(
                    row.get("ABW"), row.get("Feed Per Day"), species_val_row
                )
                if _doc_is_blank(expect_harvest_val):
                    expect_harvest_val = calc_hv
                if _doc_is_blank(survival_qty_val):
                    survival_qty_val = calc_surv
            expect_harvest_final = "" if _doc_is_blank(expect_harvest_val) else expect_harvest_val
            survival_qty_final = "" if _doc_is_blank(survival_qty_val) else survival_qty_val

            # --- Combine the selected Issues (multiselect) into one string ---
            issues_val = row.get("Issues")
            if isinstance(issues_val, list):
                issues_picked = [str(x).strip() for x in issues_val if str(x).strip()]
            else:
                issues_picked = [str(issues_val).strip()] if str(issues_val or "").strip() else []
            issues_final = ISSUES_SEP.join(dict.fromkeys(issues_picked))

            row_timestamp = (now_base + pd.Timedelta(milliseconds=i)).strftime("%Y-%m-%d %H:%M:%S.%f")

            new_records.append({
                "Timestamp": row_timestamp,
                "Customer": customer,
                "Farm Name with Code": farm,
                "Zone": zone,
                "Area": area,
                "Pond Number": pond_number,
                "Date": row_date.isoformat(),
                "DOC": doc_final,
                "Density": to_number(row.get("Density"), as_int=True),
                "Feed Per Day": to_number(row.get("Feed Per Day")),
                "ABW": str(row.get("ABW") or "").strip(),
                "Expect Harvest (KG)": expect_harvest_final,
                "Survival QTY": survival_qty_final,
                "Species Culture": species_val_row,
                "Cycle Type": cycle_val_row,
                "Issues": issues_final,
                "Water Color": water_color_val_row,
                "Grade": grade_val_row,
                "Remark": str(row.get("Remark") or "").strip(),
                "Technician": technician,
            })

            # A "Soon to be" row does not seed the DOC auto-calc chain for
            # the next row in this same save batch — the cycle hasn't
            # actually started yet, so there's nothing to count days from.
            if cycle_val_row == "Soon to be":
                running_prev_date, running_prev_doc = None, None
            else:
                running_prev_date, running_prev_doc = row_date, doc_final

        if not any_new_rows and not errors:
            errors.append("Add at least one new row before saving")

        if errors:
            st.session_state[saving_flag_key] = False
            st.error("❌ " + "  \n❌ ".join(errors))
        else:
            append_records(new_records)

            st.success(f"✅ Saved {len(new_records)} new record(s) for Pond {pond_number}!")
            del st.session_state[working_key]
            del st.session_state[working_sig_key]
            st.session_state[added_count_key] = 0
            if editor_key in st.session_state:
                del st.session_state[editor_key]
            st.session_state[saving_flag_key] = False
            time.sleep(1)
            st.rerun()

if pond_number:
    _pond_editor_fragment()

# =========================================================================
# STEP 5: HARVEST DETAILS — Date + Harvest Type (required), plus optional
# Harvest KG and Harvest ABW. This pond's row supports up to TWO harvest
# records: "Harvest Date/Type/KG/ABW" (the first harvest) and
# "Harvest Date 2/Type 2/KG 2/ABW 2" (a second harvest, e.g. a later
# partial/full harvest for the same pond). Whichever slot is still empty is
# the one this form saves into next — no date-matching, and no
# overwrite-warning caption; the app just quietly fills the next open slot,
# or edits the last slot again if both are already filled.
# =========================================================================
st.markdown("---")
st.markdown("#### 🌾 Harvest Details")

if not pond_number:
    st.info("Select a pond above to record harvest details.")
elif len(df_pond_hist_full) == 0:
    st.info(f"No saved Pond Details records yet for Pond {pond_number}. Add one above first, "
            "then come back here to record its harvest.")
else:
    # df_pond_hist_full is already scoped to this Customer + Farm + Pond,
    # and already sorted ascending by date — so the last row is this
    # pond's row to save harvest info onto.
    pond_row = df_pond_hist_full.iloc[-1]
    harvest_row_timestamp = str(pond_row.get("Timestamp") or "").strip()

    existing_harvest_date = str(pond_row.get("Harvest Date") or "").strip()
    existing_harvest_type = str(pond_row.get("Harvest Type") or "").strip()
    existing_harvest_kg = str(pond_row.get("Harvest KG") or "").strip()
    existing_harvest_abw = str(pond_row.get("Harvest ABW") or "").strip()

    existing_harvest_date2 = str(pond_row.get("Harvest Date 2") or "").strip()
    existing_harvest_type2 = str(pond_row.get("Harvest Type 2") or "").strip()
    existing_harvest_kg2 = str(pond_row.get("Harvest KG 2") or "").strip()
    existing_harvest_abw2 = str(pond_row.get("Harvest ABW 2") or "").strip()

    slot1_filled = bool(existing_harvest_date or existing_harvest_type)
    slot2_filled = bool(existing_harvest_date2 or existing_harvest_type2)

    if not slot1_filled:
        target_slot = 1
        prefill_date, prefill_type = existing_harvest_date, existing_harvest_type
        prefill_kg, prefill_abw = existing_harvest_kg, existing_harvest_abw
    elif not slot2_filled:
        # First harvest already saved — this submission records the
        # SECOND harvest, so start from a blank form.
        target_slot = 2
        prefill_date, prefill_type, prefill_kg, prefill_abw = "", "", "", ""
    else:
        # Both slots already used — keep editing the second one.
        target_slot = 2
        prefill_date, prefill_type = existing_harvest_date2, existing_harvest_type2
        prefill_kg, prefill_abw = existing_harvest_kg2, existing_harvest_abw2

    def _parse_existing_harvest_date(val):
        parsed = pd.to_datetime(val, errors="coerce")
        return parsed.date() if pd.notna(parsed) else date.today()

    st.caption(f"📌 This will be saved as Harvest {target_slot} for Pond {pond_number}.")

    hv_col1, hv_col2 = st.columns(2)
    with hv_col1:
        harvest_date_input = st.date_input(
            "Harvest Date *",
            value=_parse_existing_harvest_date(prefill_date) if prefill_date else date.today(),
            key=f"harvest_date_{widget_scope}_{target_slot}",
        )
    with hv_col2:
        default_harvest_index = (
            HARVEST_TYPE_OPTIONS.index(prefill_type)
            if prefill_type in HARVEST_TYPE_OPTIONS else 0
        )
        harvest_type = st.selectbox(
            "Harvest Type *", HARVEST_TYPE_OPTIONS, index=default_harvest_index,
            key=f"harvest_type_{widget_scope}_{target_slot}",
        )

    # Harvest KG and Harvest ABW are optional (not required) — plain text
    # inputs, blank by default unless something was already saved.
    hv_col3, hv_col4 = st.columns(2)
    with hv_col3:
        harvest_kg_input = st.text_input(
            "Harvest KG", value=prefill_kg, key=f"harvest_kg_{widget_scope}_{target_slot}",
        )
    with hv_col4:
        harvest_abw_input = st.text_input(
            "Harvest ABW", value=prefill_abw, key=f"harvest_abw_{widget_scope}_{target_slot}",
        )

    if st.button("✅ Submit Harvest", key=f"harvest_submit_{widget_scope}"):
        if not harvest_row_timestamp:
            st.error("❌ Could not find that record's timestamp — please refresh and try again.")
        elif update_harvest_by_timestamp(
            harvest_row_timestamp, harvest_date_input.isoformat(), harvest_type,
            harvest_kg_input.strip(), harvest_abw_input.strip(), slot=target_slot,
        ):
            st.success(
                f"✅ Harvest {target_slot} info saved for Pond {pond_number} — "
                f"{harvest_date_input.isoformat()} ({harvest_type})."
            )
            time.sleep(1)
            st.rerun()
        else:
            st.error("❌ Could not find that record in the Google Sheet.")

# =========================================================================
# FARM SUMMARY — every saved record for this Customer + Farm Name with
# Code, across ALL ponds (not just the pond currently selected above).
# Always reads fresh from the Google Sheet via load_data(), so it already
# reflects new saves, direct Sheet edits, and soft-deletes.
# =========================================================================
st.markdown("---")
st.markdown(f"#### 📊 All Saved Records — {farm}")

df_farm_summary = load_data()
_farm_required = {"Customer", "Farm Name with Code"}
if len(df_farm_summary) > 0 and _farm_required.issubset(df_farm_summary.columns):
    df_farm_summary = df_farm_summary[
        (df_farm_summary["Customer"] == customer) & (df_farm_summary["Farm Name with Code"] == farm)
    ].copy()
else:
    df_farm_summary = pd.DataFrame(columns=COLUMN_ORDER)

if len(df_farm_summary) > 0:
    total_expect_harvest_kg = None
    if "Date" in df_farm_summary.columns:
        df_farm_summary["_ParsedDate"] = pd.to_datetime(df_farm_summary["Date"], errors="coerce")

        # Total Expect Harvest (KG) for the farm = each pond's MOST RECENT
        # saved record's "Expect Harvest (KG)" value, summed across every
        # pond on this farm (not every historical row, which would double
        # count a pond's earlier daily estimates).
        if {"Pond Number", "Expect Harvest (KG)"}.issubset(df_farm_summary.columns):
            _latest_per_pond = (
                df_farm_summary.dropna(subset=["_ParsedDate"])
                .sort_values("_ParsedDate")
                .groupby("Pond Number", as_index=False)
                .last()
            )
            _harvest_vals = pd.to_numeric(_latest_per_pond["Expect Harvest (KG)"], errors="coerce").dropna()
            if len(_harvest_vals) > 0:
                total_expect_harvest_kg = float(_harvest_vals.sum())

        sort_cols = [c for c in ["Pond Number"] if c in df_farm_summary.columns] + ["_ParsedDate"]
        df_farm_summary = df_farm_summary.sort_values(by=sort_cols).drop(columns=["_ParsedDate"])

    # "DOC Today" = this row's saved DOC + however many days have passed
    # between its Date and today (i.e. what the DOC would be right now).
    # It's a live, always-changing number rather than something actually
    # saved in the Sheet, so it's shown in red/bold to stand out. Rows
    # whose Cycle Type is "Soon to be" haven't actually started yet, so
    # DOC Today just stays 0 for them instead of counting elapsed days.
    def _compute_doc_today(row):
        if str(row.get("Cycle Type") or "").strip() == "Soon to be":
            return "0"
        parsed = pd.to_datetime(row.get("Date"), errors="coerce")
        if pd.isna(parsed):
            return ""
        try:
            doc_num = int(float(row.get("DOC")))
        except (TypeError, ValueError):
            return ""
        days_passed = (pd.Timestamp(date.today()) - parsed).days
        return str(doc_num + days_passed)

    df_farm_summary["DOC Today"] = df_farm_summary.apply(_compute_doc_today, axis=1)

    _farm_display_cols = ["Pond Number", "Date", "Species Culture", "Cycle Type", "DOC", "DOC Today", "Density",
                           "Feed Per Day", "ABW", "Expect Harvest (KG)", "Survival QTY",
                           "Issues", "Water Color", "Grade", "Remark", "Technician",
                           "Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW",
                           "Harvest Date 2", "Harvest Type 2", "Harvest KG 2", "Harvest ABW 2"]
    _farm_display_cols = [c for c in _farm_display_cols if c in df_farm_summary.columns]

    # st.dataframe has no way to color/bold an individual column's text, so
    # this one table is rendered as a plain HTML table instead (only this
    # table — everything else keeps using st.dataframe / st.data_editor as
    # before).
    def _escape_html(v):
        return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def _render_highlighted_table(df, cols, highlight_col, highlight_style="color:red;font-weight:bold;"):
        header_html = "".join(
            f"<th style='padding:6px 10px;border-bottom:2px solid #ccc;text-align:left;white-space:nowrap;'>{_escape_html(c)}</th>"
            for c in cols
        )
        rows_html = ""
        for _, r in df.iterrows():
            cells_html = ""
            for c in cols:
                cell_style = "padding:6px 10px;border-bottom:1px solid #eee;white-space:nowrap;"
                if c == highlight_col:
                    cell_style += highlight_style
                cells_html += f"<td style='{cell_style}'>{_escape_html(r.get(c, ''))}</td>"
            rows_html += f"<tr>{cells_html}</tr>"
        return (
            "<div style='overflow-x:auto; width:100%;'>"
            "<table style='width:100%; border-collapse:collapse; font-size:0.9rem;'>"
            f"<thead><tr>{header_html}</tr></thead>"
            f"<tbody>{rows_html}</tbody>"
            "</table></div>"
        )

    st.markdown(
        _render_highlighted_table(df_farm_summary, _farm_display_cols, "DOC Today"),
        unsafe_allow_html=True,
    )
    st.caption(f"{len(df_farm_summary)} saved record(s) across all ponds for {farm}.")

    if total_expect_harvest_kg is not None:
        st.markdown(
            f"**🌾 Total Expect Harvest (KG) — {farm}: {total_expect_harvest_kg:,.2f} kg** "
            "(sum of each pond's latest Expect Harvest (KG) estimate)"
        )
else:
    st.info(f"No saved records yet for {farm}.")

st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>KMN Aqua Services - Water Quality Monitoring System</p>",
            unsafe_allow_html=True)
