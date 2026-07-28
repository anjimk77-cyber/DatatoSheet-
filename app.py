import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# Configure page
st.set_page_config(page_title="Water Quality Report - Data Collection", layout="wide")

# Fix: on narrow/mobile widths, Streamlit's selectbox/multiselect dropdown menus
# inherit the width of the (very narrow) closed field, so long option text gets
# clipped with an ellipsis. This CSS forces the open option list to size itself
# to its content and lets option text wrap, so the full option name is always
# visible when a field is tapped/clicked, regardless of how narrow the field is.
st.markdown("""
<style>
/* The popup list that appears when you tap a selectbox/multiselect */
ul[role="listbox"],
div[role="listbox"] {
    width: max-content !important;
    min-width: 220px !important;
    max-width: 92vw !important;
}
/* Each option row inside that list */
[role="option"] {
    width: auto !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
    word-break: break-word !important;
}
/* Any inner span/div Streamlit uses to render the option label */
[role="option"] * {
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: unset !important;
}

/* Selected tags shown inside a closed multiselect cell: let them wrap onto
   multiple lines and grow the box taller, instead of clipping in one row */
div[data-baseweb="select"] > div {
    flex-wrap: wrap !important;
    height: auto !important;
    min-height: 38px !important;
}
div[data-baseweb="tag"] {
    white-space: normal !important;
    max-width: 100% !important;
}
span[data-baseweb="tag"] {
    white-space: normal !important;
}

/* Let each column size itself to fit whatever is inside it, instead of a
   fixed width. Short entries stay narrow, long entries/selections make the
   column grow. The overall row still scrolls horizontally if it doesn't
   fit the screen. */
div[data-testid="stHorizontalBlock"] {
    flex-wrap: nowrap !important;
    overflow-x: auto !important;
    -webkit-overflow-scrolling: touch !important;
    align-items: flex-start !important;
}
div[data-testid="column"] {
    flex: 0 0 auto !important;
    width: auto !important;
    min-width: 110px !important;
    box-sizing: border-box !important;
}
/* Let the text inputs and select boxes shrink/grow with their content
   instead of stretching to fill a fixed-width parent */
div[data-testid="stTextInput"] input {
    width: auto !important;
    min-width: 90px !important;
}
div[data-baseweb="select"] {
    width: auto !important;
    min-width: 110px !important;
}
div[data-baseweb="select"] > div {
    width: auto !important;
}
/* Parent wrappers must not clip the horizontal scroll area */
div[data-testid="stVerticalBlockBorderWrapper"],
div[data-testid="stVerticalBlock"] {
    overflow-x: visible !important;
}
</style>
""", unsafe_allow_html=True)

# Title
st.title("💧 Water Quality Report - Data Collection")
st.subheader("KMN Aqua Services")
st.markdown("---")

# ------------------------------------------------------------------
# Google Sheets connection
# ------------------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

WORKSHEET_NAME = "Data"  # tab name inside the Google Sheet

SHEET_COLUMNS = [
    "Timestamp", "Customer", "Farm Name", "Zone", "Area", "Species Culture",
    "Cycle Type", "Pond Number", "Density", "DOC", "Feed Per Day", "ABW",
    "Diseases Issue", "Feed Issue", "Water Quality Issue", "Environment Issue",
    "Water Color", "Management Issue", "Remark", "Technician",
]


@st.cache_resource
def get_gsheet_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(creds)


@st.cache_resource
def get_worksheet():
    client = get_gsheet_client()
    sheet = client.open_by_url(st.secrets["sheet"]["url"])
    try:
        ws = sheet.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=len(SHEET_COLUMNS))
        ws.append_row(SHEET_COLUMNS)
    return ws

# Load customer data from Excel
@st.cache_data
def load_customer_data():
    df = pd.read_excel("Customer List.xlsx")
    return df

# Get unique customers, farms, zones, areas
customer_df = load_customer_data()
customers = customer_df["Customer ID"].astype(str) + " - " + customer_df["Customer Name"]
unique_customers = customers.unique().tolist()
farms = customer_df["Farm Name"].dropna().unique().tolist()
zones = customer_df["Zone"].dropna().unique().tolist()
areas = customer_df["Area"].dropna().unique().tolist()

# Define options
SPECIES_CULTURE = ["Vannamei", "Monodon", "Other"]
CYCLE_TYPE = ["Soon to be", "Running"]
DISEASES_OPTIONS = ["WSS", "EHP", "WHITE FECES", "BLACK GILLS", "SOFT SHELL", "MORTALITY ISSUE", "OXYGEN DROP", "GROWTH ISSUE", "ZOOTHAMNIUM", "Other"]
FEED_ISSUE_OPTIONS = ["Over feeding", "Under feeding", "Feed Drop", "Other"]
WATER_QUALITY_OPTIONS = ["PH issue", "Salinity issue", "Alkalinity issue", "Ammonia issue", "Calcium Hardness Issue", "Magnesium Hardness Issue", "Other"]
ENVIRONMENT_ISSUE_OPTIONS = ["Heavy Rain", "High Temperature", "Other"]
WATER_COLOR_OPTIONS = ["Milky Color", "Light Green", "Dark Green", "Light Yellow", "Light Brown", "Dark Brown", "Other"]
MANAGEMENT_ISSUE_OPTIONS = ["Aeration System Failure", "Water Exchange Problem", "Sludge & Bottom Soil Issue", "Chemical/Probiotic Overdose", "Predator Attack", "Other"]
TECHNICIAN_OPTIONS = ["Mr. Vishmika", "Mr. Ashen", "Mr. Janaka", "Mr. Shashika", "Mr. Janushan"]

# Load existing data
def load_data() -> pd.DataFrame:
    ws = get_worksheet()
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=SHEET_COLUMNS)
    return pd.DataFrame(records)

# Save data — appends one or more new rows to the sheet in a single batch call
def save_rows(new_rows: list):
    ws = get_worksheet()
    if not ws.get_all_values():
        ws.append_row(SHEET_COLUMNS)
    values = [[str(row.get(col, "")) for col in SHEET_COLUMNS] for row in new_rows]
    ws.append_rows(values)

# Initialize session state
if 'form_submitted' not in st.session_state:
    st.session_state.form_submitted = False
if 'submission_count' not in st.session_state:
    st.session_state.submission_count = 0
if 'selected_customer' not in st.session_state:
    st.session_state.selected_customer = unique_customers[0] if unique_customers else ""
if 'selected_farm' not in st.session_state:
    st.session_state.selected_farm = farms[0] if farms else ""
if 'selected_zone' not in st.session_state:
    st.session_state.selected_zone = zones[0] if zones else ""
if 'selected_area' not in st.session_state:
    st.session_state.selected_area = areas[0] if areas else ""
if 'selected_technician' not in st.session_state:
    st.session_state.selected_technician = TECHNICIAN_OPTIONS[0] if TECHNICIAN_OPTIONS else ""

# Get indices for pre-selected values
customer_index = unique_customers.index(st.session_state.selected_customer) if st.session_state.selected_customer in unique_customers else 0
farm_index = farms.index(st.session_state.selected_farm) if st.session_state.selected_farm in farms else 0
zone_index = zones.index(st.session_state.selected_zone) if st.session_state.selected_zone in zones else 0
area_index = areas.index(st.session_state.selected_area) if st.session_state.selected_area in areas else 0

# ---------------------------------------------------------------------------
# Spreadsheet-style pond entry grid (st.data_editor)
# ---------------------------------------------------------------------------
STARTING_ROWS = 5

POND_COLUMNS = [
    "Pond Number", "Density", "DOC", "Feed/Day", "ABW",
    "Species Culture", "Cycle Type",
    "Diseases Issue", "Feed Issue", "Water Quality Issue",
    "Environment Issue", "Water Color", "Management Issue",
]

def blank_pond_df(n_rows=STARTING_ROWS):
    return pd.DataFrame(
        [{col: "" for col in POND_COLUMNS} for _ in range(n_rows)]
    )

# editor_version is bumped after every successful submit to force
# st.data_editor to reinitialize with a fresh blank grid (data_editor has
# no built-in "clear" method, so changing its key is how you reset it).
if 'editor_version' not in st.session_state:
    st.session_state.editor_version = 0
if 'pond_grid_data' not in st.session_state:
    st.session_state.pond_grid_data = blank_pond_df()

st.subheader("📋 Enter Water Quality Data")

# Row 1: Customer and Farm
col1, col2 = st.columns(2)
with col1:
    customer = st.selectbox("Customer *", unique_customers, index=customer_index)
with col2:
    farm = st.selectbox("Farm Name", farms, index=farm_index if farms else None)

# Row 2: Zone and Area
col3, col4 = st.columns(2)
with col3:
    zone = st.selectbox("Zone *", zones if zones else [""], index=zone_index)
with col4:
    area = st.selectbox("Area *", areas if areas else [""], index=area_index)

st.markdown("#### 🐟 Pond Details")
st.caption(
    "This is a live spreadsheet — click a cell to edit it. Species Culture, Cycle Type, "
    "and Water Color are single-choice dropdowns — click the cell and pick one."
)
st.caption(
    "🧾 **Issue columns (Diseases, Feed, Water Quality, Environment, Management) now "
    "accept multiple selections** — click the cell and type the issues that apply, "
    "separated by commas (e.g. `WSS, EHP`). Hover a column header to see its full list "
    "of valid options. Anything you type that isn't on the list will be flagged when you submit."
)
st.caption(
    "🗑️ **To remove a row:** hover the row → check the box on its left edge → "
    "click the trash icon that appears above the grid. "
    "➕ **To add a row:** use the blank row at the bottom of the grid."
)

def issue_help(options):
    return "Type one or more, separated by commas. Valid options: " + ", ".join(options)

edited_pond_df = st.data_editor(
    st.session_state.pond_grid_data,
    key=f"pond_editor_{st.session_state.editor_version}",
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Pond Number": st.column_config.TextColumn("Pond Number", width="small"),
        "Density": st.column_config.NumberColumn("Density", width="small", step=1, format="%d"),
        "DOC": st.column_config.NumberColumn("DOC", width="small", step=1, format="%d"),
        "Feed/Day": st.column_config.NumberColumn("Feed/Day", width="small", step=0.1, format="%.2f"),
        "ABW": st.column_config.TextColumn("ABW", width="small"),
        "Species Culture": st.column_config.SelectboxColumn(
            "Species Culture *", options=SPECIES_CULTURE, width="medium", required=False
        ),
        "Cycle Type": st.column_config.SelectboxColumn(
            "Cycle Type *", options=CYCLE_TYPE, width="medium", required=False
        ),
        # NOTE: st.data_editor has no native multi-select cell type — SelectboxColumn
        # only ever allows a single value per cell. These five columns are TextColumn
        # instead, so a user can type several comma-separated values into one cell
        # (e.g. "WSS, EHP"). The values are validated against the allowed lists below
        # (find_invalid_selections) when the form is submitted.
        "Diseases Issue": st.column_config.TextColumn(
            "Diseases Issue", width="medium", help=issue_help(DISEASES_OPTIONS)
        ),
        "Feed Issue": st.column_config.TextColumn(
            "Feed Issue", width="medium", help=issue_help(FEED_ISSUE_OPTIONS)
        ),
        "Water Quality Issue": st.column_config.TextColumn(
            "Water Quality Issue", width="medium", help=issue_help(WATER_QUALITY_OPTIONS)
        ),
        "Environment Issue": st.column_config.TextColumn(
            "Environment Issue", width="medium", help=issue_help(ENVIRONMENT_ISSUE_OPTIONS)
        ),
        "Water Color": st.column_config.SelectboxColumn(
            "Water Color *", options=WATER_COLOR_OPTIONS, width="medium", required=False
        ),
        "Management Issue": st.column_config.TextColumn(
            "Management Issue", width="medium", help=issue_help(MANAGEMENT_ISSUE_OPTIONS)
        ),
    },
)

# Remark
st.markdown("#### 📝 Remark")
remark = st.text_area("Additional remarks or notes", placeholder="Enter any additional information", height=80)

# Technician
st.markdown("#### 👤 Technician *")
technician_index = TECHNICIAN_OPTIONS.index(st.session_state.selected_technician) if st.session_state.selected_technician in TECHNICIAN_OPTIONS else 0
technician = st.selectbox("Select technician", TECHNICIAN_OPTIONS, index=technician_index, key="tech")

# Submit button
submitted = st.button("✅ Submit Data", use_container_width=True)

def to_number(value, as_int=False):
    """Safely convert a cell value to a number; blank/invalid becomes 0."""
    value = str(value).strip()
    if value == "" or value.lower() == "nan":
        return 0 if as_int else 0.0
    try:
        return int(float(value)) if as_int else float(value)
    except ValueError:
        return 0 if as_int else 0.0

def clean_text(value):
    value = "" if pd.isna(value) else str(value).strip()
    return "" if value.lower() == "nan" else value

def parse_multi(value):
    """Split a comma-separated cell into a clean list of individual selections."""
    value = clean_text(value)
    if not value:
        return []
    return [v.strip() for v in value.split(",") if v.strip()]

MULTI_ISSUE_COLUMNS = {
    "Diseases Issue": DISEASES_OPTIONS,
    "Feed Issue": FEED_ISSUE_OPTIONS,
    "Water Quality Issue": WATER_QUALITY_OPTIONS,
    "Environment Issue": ENVIRONMENT_ISSUE_OPTIONS,
    "Management Issue": MANAGEMENT_ISSUE_OPTIONS,
}

def find_invalid_selections(rows):
    """Check every multi-issue cell against its allowed option list (case-insensitive).
    Returns a list of human-readable error strings, one per bad value found."""
    errors = []
    for row_num, r in enumerate(rows, start=1):
        pond_label = clean_text(r.get("Pond Number", "")) or f"row {row_num}"
        for column, allowed_options in MULTI_ISSUE_COLUMNS.items():
            allowed_lower = {opt.lower(): opt for opt in allowed_options}
            for selection in parse_multi(r.get(column, "")):
                if selection.lower() not in allowed_lower:
                    errors.append(
                        f"Pond '{pond_label}' — **{column}**: \"{selection}\" is not a valid option"
                    )
    return errors

def normalize_multi(value, allowed_options):
    """Return the comma-joined, canonically-cased list of selections for a cell."""
    allowed_lower = {opt.lower(): opt for opt in allowed_options}
    normalized = [allowed_lower.get(v.lower(), v) for v in parse_multi(value)]
    return ", ".join(normalized)

if submitted:
    # Keep only rows where at least pond number, density, or DOC was entered
    rows_to_save = []
    for _, r in edited_pond_df.iterrows():
        pond_number = clean_text(r.get("Pond Number", ""))
        density = clean_text(r.get("Density", ""))
        doc = clean_text(r.get("DOC", ""))
        if pond_number or density or doc:
            rows_to_save.append(r)

    invalid_selections = find_invalid_selections(rows_to_save)

    if not customer or not zone or not area or not technician:
        st.error("❌ Please fill in all required top-level fields (marked with *)")
    elif len(rows_to_save) == 0:
        st.error("❌ Please enter at least one pond row before submitting")
    elif any(not clean_text(r.get("Water Color", "")) for r in rows_to_save):
        st.error("❌ Water Color is required for every row")
    elif any(not clean_text(r.get("Species Culture", "")) for r in rows_to_save):
        st.error("❌ Species Culture is required for every row")
    elif any(not clean_text(r.get("Cycle Type", "")) for r in rows_to_save):
        st.error("❌ Cycle Type is required for every row")
    elif invalid_selections:
        st.error("❌ Some issue columns contain values that aren't in the allowed list:")
        for err in invalid_selections:
            st.markdown(f"- {err}")
        st.caption("Fix the values above (check spelling/commas) and submit again.")
    else:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_rows = []
        for r in rows_to_save:
            new_rows.append({
                "Timestamp": timestamp,
                "Customer": customer,
                "Farm Name": farm,
                "Zone": zone,
                "Area": area,
                "Species Culture": clean_text(r.get("Species Culture", "")),
                "Cycle Type": clean_text(r.get("Cycle Type", "")),
                "Pond Number": clean_text(r.get("Pond Number", "")),
                "Density": to_number(r.get("Density", ""), as_int=True),
                "DOC": to_number(r.get("DOC", ""), as_int=True),
                "Feed Per Day": to_number(r.get("Feed/Day", "")),
                "ABW": clean_text(r.get("ABW", "")),
                "Diseases Issue": normalize_multi(r.get("Diseases Issue", ""), DISEASES_OPTIONS),
                "Feed Issue": normalize_multi(r.get("Feed Issue", ""), FEED_ISSUE_OPTIONS),
                "Water Quality Issue": normalize_multi(r.get("Water Quality Issue", ""), WATER_QUALITY_OPTIONS),
                "Environment Issue": normalize_multi(r.get("Environment Issue", ""), ENVIRONMENT_ISSUE_OPTIONS),
                "Water Color": clean_text(r.get("Water Color", "")),
                "Management Issue": normalize_multi(r.get("Management Issue", ""), MANAGEMENT_ISSUE_OPTIONS),
                "Remark": remark,
                "Technician": technician
            })

        # Save all new rows to Google Sheets in one batch
        try:
            save_rows(new_rows)
        except Exception as e:
            st.error(f"❌ Could not save to Google Sheet: {e}")
            st.stop()

        # Save the selected values for next form
        st.session_state.selected_customer = customer
        st.session_state.selected_farm = farm
        st.session_state.selected_zone = zone
        st.session_state.selected_area = area
        st.session_state.selected_technician = technician

        # Reset the pond grid back to fresh blank rows, and bump the editor
        # version so st.data_editor reinitializes instead of keeping old edits
        st.session_state.pond_grid_data = blank_pond_df()
        st.session_state.editor_version += 1

        st.session_state.form_submitted = True
        st.success(f"✅ {len(new_rows)} row(s) saved successfully!")

        # Auto-refresh after brief delay to see success message
        import time
        time.sleep(1)
        st.rerun()

# Display saved data
st.markdown("---")
st.subheader("📊 Saved Data")

df = None
try:
    df = load_data()
except Exception as e:
    st.error(f"❌ Could not load data from Google Sheet: {e}")
    df = pd.DataFrame(columns=SHEET_COLUMNS)

if len(df) > 0:
    st.write(f"Total records: **{len(df)}**")
    
    # Reorder columns for display
    column_order = [
        "Timestamp", "Customer", "Farm Name", "Zone", "Area",
        "Species Culture", "Cycle Type",
        "Pond Number", "Density", "DOC", "Feed Per Day", "ABW",
        "Diseases Issue", "Feed Issue", "Water Quality Issue",
        "Environment Issue", "Water Color", "Management Issue",
        "Remark", "Technician"
    ]
    # Only keep columns that actually exist in the dataframe, preserving order
    existing_columns = [col for col in column_order if col in df.columns]
    # Add any extra columns not in our defined order (safety net)
    extra_columns = [col for col in df.columns if col not in column_order]
    df_display = df[existing_columns + extra_columns]
    
    # Display data in a table
    st.dataframe(df_display, use_container_width=True, height=400)
    
    # Download options
    col1, col2 = st.columns(2)
    
    with col1:
        # Download as CSV
        csv = df_display.to_csv(index=False)
        st.download_button(
            label="📥 Download as CSV",
            data=csv,
            file_name=f"water_quality_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with col2:
        # Download as Excel
        from io import BytesIO
        excel_buffer = BytesIO()
        df_display.to_excel(excel_buffer, index=False, sheet_name="Water Quality Data")
        excel_buffer.seek(0)
        st.download_button(
            label="📥 Download as Excel",
            data=excel_buffer,
            file_name=f"water_quality_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
else:
    st.info("ℹ️ No data saved yet. Fill out the form above to get started!")

st.markdown("<p style='text-align: center; color: gray;'>KMN Aqua Services - Water Quality Monitoring System</p>", unsafe_allow_html=True)
