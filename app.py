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

st.set_page_config(page_title="Water Quality Report - Data Collection", layout="wide", page_icon="💧")

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

st.markdown("<h1 style='text-align: center;'>Water Quality Report - Data Collection -</h1>",
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

COLUMN_ORDER = [
    "Timestamp", "Customer", "Farm Name with Code", "Zone", "Area",
    "Pond Number", "Date", "Species Culture", "Cycle Type",
    "DOC", "Density", "Feed Per Day", "ABW",
    "Issues", "Water Color", "Grade", "Remark", "Technician",
    "Deleted",
]

# Columns shown/edited in the pond spreadsheet (the rest — Customer, Farm,
# Zone, Area, Pond, Technician, Timestamp — come from the selectors above
# the table and are attached automatically when a row is saved).
POND_COLS = ["Date", "DOC", "Density", "Feed Per Day", "ABW", "Species Culture",
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
    # Make sure the header row matches what we expect (self-heals a blank sheet)
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
    EDITOR_POND_COLS = ["Date", "DOC", "Density", "Feed Per Day", "ABW", "Species Culture",
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
    _NUM_COLS = ["DOC", "Density", "Feed Per Day"]

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
            "ABW": "object", "Species Culture": "object", "Cycle Type": "object",
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
    column_order = EDITOR_POND_COLS + ["Status"]

    editor_key = f"editor_{widget_scope}"
    working_key = f"__pond_working_{widget_scope}"
    working_sig_key = f"__pond_working_sig_{widget_scope}"

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
        """Return a copy of df with every blank DOC cell filled in, in row
        order, using (previous row's DOC + days since previous row's Date)."""
        df = df.reset_index(drop=True).copy()
        run_date, run_doc = seed_date, seed_doc
        for i in range(len(df)):
            row_date = _parse_cell_date(df.at[i, "Date"])
            if row_date is None:
                continue
            if run_date is None or run_doc is None:
                if not _doc_is_blank(df.at[i, "DOC"]):
                    try:
                        run_doc = int(df.at[i, "DOC"])
                        run_date = row_date
                    except Exception:
                        pass
                continue
            if i == 0 and not _doc_is_blank(df.at[i, "DOC"]):
                run_doc = int(df.at[i, "DOC"])
                run_date = row_date
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
        chain and on rows being added/removed. So we only run the full
        rebuild (_normalize_pond_dtypes / _recompute_docs / _recompute_status)
        when one of those actually happened. A plain edit to any other column
        (Density, Species, Cycle, Water Color, Grade, Remark, Issues, ABW)
        just patches that one cell in place on the SAME dataframe object
        instead of rebuilding a fresh one — passing a materially-identical
        dataframe back into the grid on every keystroke was what made it
        reset its scroll position back to the first columns each time."""
        state = st.session_state.get(editor_key)
        if not state:
            return

        edited_rows = state.get("edited_rows", {})
        added_rows = state.get("added_rows", [])
        deleted_rows = state.get("deleted_rows", [])

        needs_recompute = bool(added_rows) or bool(deleted_rows) or any(
            ("Date" in changes or "DOC" in changes) for changes in edited_rows.values()
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

        for new_row in added_rows:
            row = {c: new_row.get(c) for c in df.columns}
            row["Timestamp"] = ""  # brand-new row -> unsaved until Save is clicked
            if not isinstance(row.get("Issues"), list):
                row["Issues"] = []
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

        for idx in sorted(deleted_rows, reverse=True):
            if idx < len(df):
                ts = str(df.at[idx, "Timestamp"]).strip()
                if ts:
                    # Already-saved row: soft-delete in the Google Sheet
                    # (flags it, keeps the underlying data), then remove it
                    # from the screen — it will stay gone after a refresh.
                    mark_deleted_by_timestamp(ts)
                df = df.drop(index=idx).reset_index(drop=True)

        if needs_recompute:
            df = _normalize_pond_dtypes(df)
            df = _recompute_docs(df, prev_date, prev_doc)
            df = _recompute_status(df)

        st.session_state[working_key] = df
        # Clear only the widget's own edit-tracking (not the whole widget
        # state) so the grid doesn't get fully remounted on every keystroke
        # — that remount was what reset the scroll position/focus back to
        # the top of the spreadsheet after each cell edit.
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

    save_clicked = st.button("💾 Save New Records", use_container_width=True,
                              type="primary", key=f"save_{widget_scope}")

    if save_clicked:
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
        any_new_rows = False

        for i, row in rows_all.iterrows():
            row_label = f"Row {i + 1}"
            is_saved = str(row.get("Timestamp") or "").strip() != ""

            row_date = _parse_cell_date(row.get("Date"))

            if is_saved:
                # Already in the sheet — skip re-saving it, but keep the
                # running DOC/Date chain going so any new rows after it
                # auto-calculate correctly.
                doc_val = row.get("DOC")
                if row_date is not None and not _doc_is_blank(doc_val):
                    running_prev_date, running_prev_doc = row_date, int(doc_val)
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
                "Species Culture": species_val_row,
                "Cycle Type": cycle_val_row,
                "Issues": issues_final,
                "Water Color": water_color_val_row,
                "Grade": grade_val_row,
                "Remark": str(row.get("Remark") or "").strip(),
                "Technician": technician,
            })

            running_prev_date, running_prev_doc = row_date, doc_final

        if not any_new_rows and not errors:
            errors.append("Add at least one new row before saving")

        if errors:
            st.error("❌ " + "  \n❌ ".join(errors))
        else:
            append_records(new_records)

            st.success(f"✅ Saved {len(new_records)} new record(s) for Pond {pond_number}!")
            del st.session_state[working_key]
            del st.session_state[working_sig_key]
            if editor_key in st.session_state:
                del st.session_state[editor_key]
            time.sleep(1)
            st.rerun()

if pond_number:
    _pond_editor_fragment()

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
    if "Date" in df_farm_summary.columns:
        df_farm_summary["_ParsedDate"] = pd.to_datetime(df_farm_summary["Date"], errors="coerce")
        sort_cols = [c for c in ["Pond Number"] if c in df_farm_summary.columns] + ["_ParsedDate"]
        df_farm_summary = df_farm_summary.sort_values(by=sort_cols).drop(columns=["_ParsedDate"])
    _farm_display_cols = ["Pond Number", "Date", "Species Culture", "Cycle Type", "DOC", "Density",
                           "Feed Per Day", "ABW", "Issues", "Water Color", "Grade", "Remark", "Technician"]
    _farm_display_cols = [c for c in _farm_display_cols if c in df_farm_summary.columns]
    st.dataframe(df_farm_summary[_farm_display_cols], use_container_width=True, hide_index=True)
    st.caption(f"{len(df_farm_summary)} saved record(s) across all ponds for {farm}.")
else:
    st.info(f"No saved records yet for {farm}.")

st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>KMN Aqua Services - Water Quality Monitoring System</p>",
            unsafe_allow_html=True)
