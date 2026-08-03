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

st.title("💧 Water Quality Report - Data Collection")
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

# All issue categories combined. Inside the spreadsheet editor each cell can
# only hold a single selection (a data_editor limitation), so this is used
# as a single-select dropdown in the Issues column, same as the reference
# spreadsheet layout.
ISSUES_OPTIONS = (
    [f"Disease: {x}" for x in DISEASES_OPTIONS] +
    [f"Feed: {x}" for x in FEED_ISSUE_OPTIONS] +
    [f"Water Quality: {x}" for x in WATER_QUALITY_OPTIONS] +
    [f"Environment: {x}" for x in ENVIRONMENT_ISSUE_OPTIONS] +
    [f"Management: {x}" for x in MANAGEMENT_ISSUE_OPTIONS]
)
ISSUES_SEP = "; "

COLUMN_ORDER = [
    "Timestamp", "Customer", "Farm Name with Code", "Zone", "Area",
    "Pond Number", "Date", "Species Culture", "Cycle Type",
    "DOC", "Density", "Feed Per Day", "ABW",
    "Issues", "Water Color", "Grade", "Remark", "Technician",
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

@st.cache_data(show_spinner=False)
def _load_data_cached(_version, _sheet_id):
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
    sheet_id = st.secrets["gsheet"]["sheet_id"]
    return _load_data_cached(st.session_state.get("_data_version", 0), sheet_id)

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

def delete_all_pond_records(customer, farm, pond_number):
    """Delete every saved row for this customer+farm+pond (used right before
    re-writing that pond's whole spreadsheet in one shot on Save)."""
    ws = get_worksheet()
    all_values = ws.get_all_values()
    if not all_values:
        return
    header = all_values[0]
    try:
        idx_customer = header.index("Customer")
        idx_farm = header.index("Farm Name with Code")
        idx_pond = header.index("Pond Number")
    except ValueError:
        return
    rows_to_delete = []
    for i, row in enumerate(all_values[1:], start=2):
        row = row + [""] * (len(header) - len(row))
        if (row[idx_customer] == customer
                and row[idx_farm] == farm
                and row[idx_pond] == pond_number):
            rows_to_delete.append(i)
    for r in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(r)
    if rows_to_delete:
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
# STEP 4: POND DETAILS — pond selection bar, then a spreadsheet
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
if len(df_pond_hist_full) > 0 and "Date" in df_pond_hist_full.columns:
    df_pond_hist_full["_ParsedDate"] = pd.to_datetime(df_pond_hist_full["Date"], errors="coerce")
    df_pond_hist_full = df_pond_hist_full.sort_values(by="_ParsedDate").reset_index(drop=True)
    if df_pond_hist_full["_ParsedDate"].notna().any():
        latest_idx = df_pond_hist_full["_ParsedDate"].last_valid_index()
        prev_date = df_pond_hist_full.loc[latest_idx, "_ParsedDate"].date()
        prev_doc = to_number(df_pond_hist_full.loc[latest_idx, "DOC"], as_int=True)
        prev_species = str(df_pond_hist_full.loc[latest_idx].get("Species Culture") or "").strip() or None
        prev_cycle = str(df_pond_hist_full.loc[latest_idx].get("Cycle Type") or "").strip() or None

default_species = prev_species if prev_species in SPECIES_CULTURE else SPECIES_CULTURE[0]
default_cycle = prev_cycle if prev_cycle in CYCLE_TYPE else CYCLE_TYPE[0]

st.markdown(f"##### 📜 History — Pond {pond_number}" if pond_number else "##### 📜 History")

if pond_number:
    # "Timestamp" is kept as a hidden internal column so we know which rows
    # already exist in the sheet (Saved) vs. which were just added in this
    # session (New / unsaved) — that's what drives the Status column below.
    display_cols = ["Timestamp"] + POND_COLS
    existing_pond_cols = [c for c in display_cols if c in df_pond_hist_full.columns]
    df_pond_hist_display = df_pond_hist_full[existing_pond_cols].copy() if len(df_pond_hist_full) > 0 \
        else pd.DataFrame(columns=display_cols)
    original_row_count = len(df_pond_hist_display)

    if original_row_count == 0:
        st.info(f"No history yet for Pond {pond_number}. Add its first record in the spreadsheet below.")

    editor_df = df_pond_hist_display.copy()
    if len(editor_df) > 0:
        editor_df["Date"] = pd.to_datetime(editor_df["Date"], errors="coerce").dt.date
        for numcol in ["DOC", "Density", "Feed Per Day"]:
            editor_df[numcol] = pd.to_numeric(editor_df[numcol], errors="coerce")
    else:
        editor_df = pd.DataFrame({c: pd.Series(dtype="object") for c in display_cols})
        editor_df["Date"] = editor_df["Date"].astype("object")

    if "Status" not in editor_df.columns:
        editor_df["Status"] = editor_df["Timestamp"].apply(
            lambda t: "✅ Saved" if str(t).strip() else "🆕 New (unsaved)"
        )

    column_config = {
        "Date": st.column_config.DateColumn("Date *", required=True),
        "DOC": st.column_config.NumberColumn("DOC (auto)", help="Filled in automatically once you pick a Date — edit it to override", step=1),
        "Density": st.column_config.NumberColumn("Density", step=1),
        "Feed Per Day": st.column_config.NumberColumn("Feed/Day"),
        "ABW": st.column_config.TextColumn("ABW"),
        "Species Culture": st.column_config.SelectboxColumn("Species Culture *", options=SPECIES_CULTURE,
                                                              required=True, default=default_species),
        "Cycle Type": st.column_config.SelectboxColumn("Cycle Type *", options=CYCLE_TYPE,
                                                         required=True, default=default_cycle),
        "Issues": st.column_config.SelectboxColumn("Issues", options=["(none)"] + ISSUES_OPTIONS, required=False),
        "Water Color": st.column_config.SelectboxColumn("Water Color", options=WATER_COLOR_OPTIONS, required=False),
        "Grade": st.column_config.SelectboxColumn("Grade", options=GRADE_OPTIONS, required=False),
        "Remark": st.column_config.TextColumn("Remark"),
        "Status": st.column_config.TextColumn("Status", disabled=True),
    }
    # "Timestamp" is deliberately left out of column_order so it stays in
    # the underlying data (for matching rows back to sheet rows) without
    # being shown or editable — "Status" is shown last instead.
    column_order = POND_COLS + ["Status"]

    editor_key = f"editor_{widget_scope}"
    working_key = f"__pond_working_{widget_scope}"
    working_sig_key = f"__pond_working_sig_{widget_scope}"

    def _parse_cell_date(val):
        if val is None or val == "":
            return None
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
    # internal state). Re-seed it whenever we switch pond/farm/customer or
    # right after a save changes the saved history.
    base_signature = (customer, farm, pond_number, original_row_count)
    if st.session_state.get(working_sig_key) != base_signature:
        st.session_state[working_key] = editor_df.copy()
        st.session_state[working_sig_key] = base_signature
        if editor_key in st.session_state:
            del st.session_state[editor_key]

    def apply_editor_changes():
        """on_change callback: fold whatever was just typed into our working
        copy, recompute DOC + Status, then delete the widget's own delta
        state so Streamlit redraws the table fresh from our updated copy."""
        state = st.session_state.get(editor_key)
        if not state:
            return
        df = st.session_state[working_key].reset_index(drop=True).copy()

        for idx, changes in state.get("edited_rows", {}).items():
            idx = int(idx)
            if idx < len(df):
                for col, val in changes.items():
                    if col in df.columns:
                        df.at[idx, col] = val

        for new_row in state.get("added_rows", []):
            row = {c: new_row.get(c) for c in df.columns}
            row["Timestamp"] = ""  # brand-new row -> unsaved until Save is clicked
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

        for idx in sorted(state.get("deleted_rows", []), reverse=True):
            if idx < len(df):
                df = df.drop(index=idx).reset_index(drop=True)

        df = _recompute_docs(df, prev_date, prev_doc)
        df = _recompute_status(df)

        st.session_state[working_key] = df
        del st.session_state[editor_key]

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

    save_clicked = st.button("💾 Save Changes to Pond History", use_container_width=True,
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

        edited_rows = st.session_state[working_key].reset_index(drop=True)
        for i, row in edited_rows.iterrows():
            row_label = f"Row {i + 1}"

            # --- Date ---
            row_date = row.get("Date")
            if isinstance(row_date, str) and row_date.strip():
                parsed = pd.to_datetime(row_date, errors="coerce")
                row_date = parsed.date() if pd.notna(parsed) else None
            if row_date is None or (isinstance(row_date, float) and pd.isna(row_date)):
                errors.append(f"{row_label}: Date is required")
                continue

            # --- DOC (auto-calculate if blank) ---
            doc_raw = row.get("DOC")
            doc_blank = doc_raw is None or (isinstance(doc_raw, float) and pd.isna(doc_raw)) or str(doc_raw).strip() == ""
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

            # Existing rows keep their original Timestamp; new/blank rows get one now
            row_timestamp = str(row.get("Timestamp") or "").strip()
            if not row_timestamp:
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
                "Issues": "" if str(row.get("Issues") or "").strip() in ("", "(none)") else str(row.get("Issues")).strip(),
                "Water Color": water_color_val_row,
                "Grade": grade_val_row,
                "Remark": str(row.get("Remark") or "").strip(),
                "Technician": technician,
            })

            running_prev_date, running_prev_doc = row_date, doc_final

        if errors:
            st.error("❌ " + "  \n❌ ".join(errors))
        else:
            delete_all_pond_records(customer, farm, pond_number)
            append_records(new_records)

            st.success(f"✅ Saved {len(new_records)} record(s) for Pond {pond_number}!")
            del st.session_state[working_key]
            del st.session_state[working_sig_key]
            if editor_key in st.session_state:
                del st.session_state[editor_key]
            time.sleep(1)
            st.rerun()

    # Downloads reflect the last-saved state of this pond's history
    if original_row_count > 0:
        pdl1, pdl2 = st.columns(2)
        with pdl1:
            pond_csv = df_pond_hist_display[POND_COLS].to_csv(index=False)
            st.download_button(
                "📥 Download this pond's history (CSV)", data=pond_csv,
                file_name=f"pond_{pond_number}_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv", use_container_width=True, key=f"dl_csv_{farm}_{pond_number}",
            )
        with pdl2:
            pond_buf = BytesIO()
            df_pond_hist_display[POND_COLS].to_excel(pond_buf, index=False, sheet_name="Pond History")
            pond_buf.seek(0)
            st.download_button(
                "📥 Download this pond's history (Excel)", data=pond_buf,
                file_name=f"pond_{pond_number}_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key=f"dl_xlsx_{farm}_{pond_number}",
            )

# =========================================================================
# HISTORY FOR THIS FARM (table — allowed to scroll horizontally)
# =========================================================================
st.markdown("---")
st.subheader(f"📊 Saved Pond History — {farm}")

df_all = load_data()

if len(df_all) > 0 and {"Customer", "Farm Name with Code"}.issubset(df_all.columns):
    df_farm = df_all[(df_all["Customer"] == customer) & (df_all["Farm Name with Code"] == farm)]
    existing_cols = [c for c in COLUMN_ORDER if c in df_farm.columns]
    extra_cols = [c for c in df_farm.columns if c not in COLUMN_ORDER]
    df_farm_display = df_farm[existing_cols + extra_cols]

    if len(df_farm_display) > 0:
        st.write(f"Records for this farm: **{len(df_farm_display)}**")
        st.dataframe(df_farm_display, use_container_width=True, height=350)

        dl1, dl2 = st.columns(2)
        with dl1:
            csv = df_farm_display.to_csv(index=False)
            st.download_button(
                "📥 Download this farm's history (CSV)", data=csv,
                file_name=f"{farm}_water_quality_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv", use_container_width=True,
            )
        with dl2:
            buf = BytesIO()
            df_farm_display.to_excel(buf, index=False, sheet_name="Pond History")
            buf.seek(0)
            st.download_button(
                "📥 Download this farm's history (Excel)", data=buf,
                file_name=f"{farm}_water_quality_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        st.info("ℹ️ No records saved yet for this farm.")
else:
    st.info("ℹ️ No data saved yet. Fill out the form above to get started!")

with st.expander("📁 View / download full dataset (all customers & farms)"):
    if len(df_all) > 0:
        existing_cols_all = [c for c in COLUMN_ORDER if c in df_all.columns]
        extra_cols_all = [c for c in df_all.columns if c not in COLUMN_ORDER]
        df_all_display = df_all[existing_cols_all + extra_cols_all]
        st.dataframe(df_all_display, use_container_width=True, height=400)
        csv_all = df_all_display.to_csv(index=False)
        st.download_button(
            "📥 Download full dataset (CSV)", data=csv_all,
            file_name=f"water_quality_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv", use_container_width=True,
        )
    else:
        st.write("No data yet.")

st.markdown("<p style='text-align: center; color: gray;'>KMN Aqua Services - Water Quality Monitoring System</p>",
            unsafe_allow_html=True)
