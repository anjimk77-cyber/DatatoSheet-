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
# customer/farm pickers, the pond record cards) should stack into a single
# column, one field after another, on narrow / mobile screens.
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

/* "Saved" status pill */
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
   after another. Tables (st.dataframe) are NOT built from these
   horizontal blocks, so they are untouched and keep their own native
   horizontal scrollbar. */
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

# Issues is now a genuine multi-select field (a pond can have several issues
# on the same visit). Values are stored in the sheet joined by "; ".
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

# Columns shown/edited on each pond record card (the rest — Customer, Farm,
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
# STEP 4: POND SELECTION
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

# =========================================================================
# LOAD THIS POND'S HISTORY
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

if len(df_pond_hist_full) > 0 and "Date" in df_pond_hist_full.columns:
    df_pond_hist_full["_ParsedDate"] = pd.to_datetime(df_pond_hist_full["Date"], errors="coerce")
    df_pond_hist_full = df_pond_hist_full.sort_values(by="_ParsedDate").reset_index(drop=True)

st.markdown(f"##### 📜 History — Pond {pond_number}" if pond_number else "##### 📜 History")

# =========================================================================
# PER-ROW RECORD CARDS (replaces the old spreadsheet grid + single
# "Save Changes" button). Each card has its own red "Save" button; once
# saved, that card shows a green "Saved" status instead of the button.
# =========================================================================
def _row_from_series(r):
    issues_raw = str(r.get("Issues") or "").strip()
    issues_list = [x.strip() for x in issues_raw.split(ISSUES_SEP) if x.strip()] if issues_raw else []
    return {
        "row_id": str(uuid.uuid4()),
        "timestamp": str(r.get("Timestamp") or ""),
        "saved": True,
        "data": {
            "Date": pd.to_datetime(r.get("Date"), errors="coerce").date() if str(r.get("Date") or "").strip() else None,
            "DOC": to_number(r.get("DOC"), as_int=True),
            "Density": to_number(r.get("Density"), as_int=True),
            "Feed Per Day": to_number(r.get("Feed Per Day")),
            "ABW": str(r.get("ABW") or ""),
            "Species Culture": str(r.get("Species Culture") or "") or SPECIES_CULTURE[0],
            "Cycle Type": str(r.get("Cycle Type") or "") or CYCLE_TYPE[0],
            "Issues": issues_list,
            "Water Color": str(r.get("Water Color") or ""),
            "Grade": str(r.get("Grade") or ""),
            "Remark": str(r.get("Remark") or ""),
        },
    }

def _blank_row(prev_row_data=None):
    prev_row_data = prev_row_data or {}
    return {
        "row_id": str(uuid.uuid4()),
        "timestamp": "",
        "saved": False,
        "data": {
            "Date": None,
            "DOC": None,
            # Density carries forward from the previous record, exactly like
            # Species Culture and Cycle Type already did.
            "Density": prev_row_data.get("Density", 0),
            "Feed Per Day": None,
            "ABW": "",
            "Species Culture": prev_row_data.get("Species Culture", SPECIES_CULTURE[0]),
            "Cycle Type": prev_row_data.get("Cycle Type", CYCLE_TYPE[0]),
            "Issues": [],
            "Water Color": prev_row_data.get("Water Color", ""),
            "Grade": "",
            "Remark": "",
        },
    }

rows_state_key = f"__pond_rows__{customer}__{farm}__{pond_number}"
rows_sig_key = "__pond_rows_sig"
scope_sig = (customer, farm, pond_number)

if st.session_state.get(rows_sig_key) != scope_sig or rows_state_key not in st.session_state:
    st.session_state[rows_state_key] = [_row_from_series(r) for _, r in df_pond_hist_full.iterrows()]
    st.session_state[rows_sig_key] = scope_sig

rows_list = st.session_state[rows_state_key]

if pond_number:
    if len(rows_list) == 0:
        st.info(f"No history yet for Pond {pond_number}. Add its first record below.")

    running_date, running_doc = None, None
    running_prev_data = {}

    for idx, row in enumerate(rows_list):
        rid = row["row_id"]
        d = row["data"]

        with st.container():
            st.markdown('<div class="pond-card">', unsafe_allow_html=True)
            c_date, c_doc, c_dens, c_feed, c_abw = st.columns(5)
            with c_date:
                date_val = st.date_input("Date *", value=d["Date"], key=f"date_{rid}")
            with c_doc:
                doc_default = d["DOC"]
                if doc_default in (None, 0) and running_date is not None and running_doc is not None and date_val:
                    doc_default = int(running_doc) + (date_val - running_date).days
                doc_val = st.number_input("DOC", value=int(doc_default) if doc_default not in (None, "") else 0,
                                           step=1, key=f"doc_{rid}")
            with c_dens:
                dens_val = st.number_input("Density", value=int(d["Density"] or 0), step=1, key=f"density_{rid}")
            with c_feed:
                feed_val = st.number_input("Feed/Day", value=float(d["Feed Per Day"] or 0.0), key=f"feed_{rid}")
            with c_abw:
                abw_val = st.text_input("ABW", value=d["ABW"], key=f"abw_{rid}")

            c_sp, c_cyc, c_wc, c_gr = st.columns(4)
            with c_sp:
                sp_default = d["Species Culture"] if d["Species Culture"] in SPECIES_CULTURE else SPECIES_CULTURE[0]
                sp_val = st.selectbox("Species Culture *", SPECIES_CULTURE,
                                       index=SPECIES_CULTURE.index(sp_default), key=f"species_{rid}")
            with c_cyc:
                cyc_default = d["Cycle Type"] if d["Cycle Type"] in CYCLE_TYPE else CYCLE_TYPE[0]
                cyc_val = st.selectbox("Cycle Type *", CYCLE_TYPE,
                                        index=CYCLE_TYPE.index(cyc_default), key=f"cycle_{rid}")
            with c_wc:
                wc_options = [""] + WATER_COLOR_OPTIONS
                wc_default = d["Water Color"] if d["Water Color"] in wc_options else ""
                wc_val = st.selectbox("Water Color", wc_options, index=wc_options.index(wc_default), key=f"watercolor_{rid}")
            with c_gr:
                gr_options = [""] + GRADE_OPTIONS
                gr_default = d["Grade"] if d["Grade"] in gr_options else ""
                gr_val = st.selectbox("Grade", gr_options, index=gr_options.index(gr_default), key=f"grade_{rid}")

            c_issues, c_remark = st.columns([2, 2])
            with c_issues:
                issues_default = [x for x in d["Issues"] if x in ISSUES_OPTIONS]
                issues_val = st.multiselect("Issues (select all that apply)", ISSUES_OPTIONS,
                                             default=issues_default, key=f"issues_{rid}")
            with c_remark:
                remark_val = st.text_input("Remark", value=d["Remark"], key=f"remark_{rid}")

            current_values = {
                "Date": date_val, "DOC": doc_val, "Density": dens_val, "Feed Per Day": feed_val,
                "ABW": abw_val, "Species Culture": sp_val, "Cycle Type": cyc_val,
                "Issues": issues_val, "Water Color": wc_val, "Grade": gr_val, "Remark": remark_val,
            }
            is_currently_saved = row["saved"] and row.get("saved_snapshot") == current_values

            c_status, c_remove = st.columns([3, 1])
            with c_status:
                if is_currently_saved:
                    st.markdown('<div class="status-saved">✅ Saved</div>', unsafe_allow_html=True)
                else:
                    save_clicked = st.button("💾 Save", key=f"save_{rid}", type="primary", use_container_width=True)
                    if save_clicked:
                        if date_val is None:
                            st.error("Date is required for this row.")
                        elif not customer or not farm or not zone or not area or not technician or not pond_number:
                            st.error("Please fill in all required top-level fields (marked with *) above.")
                        else:
                            ts = row["timestamp"] or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            record = {
                                "Timestamp": ts,
                                "Customer": customer,
                                "Farm Name with Code": farm,
                                "Zone": zone,
                                "Area": area,
                                "Pond Number": pond_number,
                                "Date": date_val.isoformat(),
                                "DOC": int(doc_val),
                                "Density": int(dens_val),
                                "Feed Per Day": feed_val,
                                "ABW": abw_val.strip(),
                                "Species Culture": sp_val,
                                "Cycle Type": cyc_val,
                                "Issues": ISSUES_SEP.join(issues_val),
                                "Water Color": wc_val,
                                "Grade": gr_val,
                                "Remark": remark_val.strip(),
                                "Technician": technician,
                            }
                            update_record_by_timestamp(ts, record)
                            row["timestamp"] = ts
                            row["saved"] = True
                            row["saved_snapshot"] = current_values
                            st.rerun()
            with c_remove:
                if st.button("🗑️", key=f"remove_{rid}", use_container_width=True, help="Remove this row"):
                    if row["saved"] and row["timestamp"]:
                        delete_record_by_timestamp(row["timestamp"])
                    rows_list.pop(idx)
                    st.rerun()

            st.markdown('</div>', unsafe_allow_html=True)

        # Chain DOC / carry-forward defaults for the next row, whether this
        # one is saved yet or not.
        if date_val is not None:
            running_date, running_doc = date_val, doc_val
        running_prev_data = current_values

    if st.button("➕ Add New Row", use_container_width=True, key=f"add_row_{customer}_{farm}_{pond_number}"):
        rows_list.append(_blank_row(running_prev_data))
        st.rerun()

    # Downloads reflect the current saved state of this pond's history
    df_pond_hist_display = load_data()
    if len(df_pond_hist_display) > 0 and required_cols.issubset(df_pond_hist_display.columns):
        df_pond_hist_display = df_pond_hist_display[
            (df_pond_hist_display["Customer"] == customer)
            & (df_pond_hist_display["Farm Name with Code"] == farm)
            & (df_pond_hist_display["Pond Number"] == pond_number)
        ][[c for c in POND_COLS if c in df_pond_hist_display.columns]]

    if len(df_pond_hist_display) > 0:
        pdl1, pdl2 = st.columns(2)
        with pdl1:
            pond_csv = df_pond_hist_display.to_csv(index=False)
            st.download_button(
                "📥 Download this pond's history (CSV)", data=pond_csv,
                file_name=f"pond_{pond_number}_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv", use_container_width=True, key=f"dl_csv_{farm}_{pond_number}",
            )
        with pdl2:
            pond_buf = BytesIO()
            df_pond_hist_display.to_excel(pond_buf, index=False, sheet_name="Pond History")
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
