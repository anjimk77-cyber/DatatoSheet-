import streamlit as st
import pandas as pd
import os
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

/* Keep every column at a readable minimum width instead of letting it shrink
   to fit the screen. If the row is wider than the screen, it scrolls
   horizontally rather than squeezing each cell unreadably small. */
div[data-testid="stHorizontalBlock"] {
    flex-wrap: nowrap !important;
    overflow-x: auto !important;
    -webkit-overflow-scrolling: touch !important;
}
div[data-testid="column"] {
    min-width: 170px !important;
    flex: 0 0 auto !important;
    width: 170px !important;
}
</style>
""", unsafe_allow_html=True)

# Title
st.title("💧 Water Quality Report - Data Collection")
st.subheader("KMN Aqua Services")
st.markdown("---")

# Define the data file path
DATA_FILE = "water_quality_data.csv"

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
def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_csv(DATA_FILE)
        # Migrate old "AB" column to "ABW" if present
        if "AB" in df.columns and "ABW" not in df.columns:
            df = df.rename(columns={"AB": "ABW"})
            df.to_csv(DATA_FILE, index=False)
        return df
    return pd.DataFrame()

# Save data
def save_data(df):
    df.to_csv(DATA_FILE, index=False)

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
# Form fields that should reset after submission
if 'pond_number_val' not in st.session_state:
    st.session_state.pond_number_val = ""
if 'density_val' not in st.session_state:
    st.session_state.density_val = 0
if 'doc_val' not in st.session_state:
    st.session_state.doc_val = 0
if 'feed_per_day_val' not in st.session_state:
    st.session_state.feed_per_day_val = 0.0
if 'abw_val' not in st.session_state:
    st.session_state.abw_val = ""
if 'diseases_val' not in st.session_state:
    st.session_state.diseases_val = []
if 'feed_issue_val' not in st.session_state:
    st.session_state.feed_issue_val = []
if 'water_quality_val' not in st.session_state:
    st.session_state.water_quality_val = []
if 'env_issue_val' not in st.session_state:
    st.session_state.env_issue_val = []
if 'management_issue_val' not in st.session_state:
    st.session_state.management_issue_val = []
if 'remark_val' not in st.session_state:
    st.session_state.remark_val = ""
if 'selected_technician' not in st.session_state:
    st.session_state.selected_technician = TECHNICIAN_OPTIONS[0] if TECHNICIAN_OPTIONS else ""
if 'water_color_val' not in st.session_state:
    st.session_state.water_color_val = ""
if 'species_val' not in st.session_state:
    st.session_state.species_val = 0
if 'cycle_val' not in st.session_state:
    st.session_state.cycle_val = 0

# Get indices for pre-selected values
customer_index = unique_customers.index(st.session_state.selected_customer) if st.session_state.selected_customer in unique_customers else 0
farm_index = farms.index(st.session_state.selected_farm) if st.session_state.selected_farm in farms else 0
zone_index = zones.index(st.session_state.selected_zone) if st.session_state.selected_zone in zones else 0
area_index = areas.index(st.session_state.selected_area) if st.session_state.selected_area in areas else 0

# Row-based pond entry table (supports true multi-select per row, plus add/remove rows)
STARTING_ROWS = 5

def blank_row():
    return {
        "pond_number": "",
        "density": "",
        "doc": "",
        "feed_per_day": "",
        "abw": "",
        "diseases": [],
        "feed_issue": [],
        "water_quality": [],
        "env_issue": [],
        "water_color": "",
        "management_issue": [],
    }

if 'next_row_id' not in st.session_state:
    st.session_state.next_row_id = STARTING_ROWS
if 'row_ids' not in st.session_state:
    st.session_state.row_ids = list(range(STARTING_ROWS))
if 'rows_data' not in st.session_state:
    st.session_state.rows_data = {rid: blank_row() for rid in st.session_state.row_ids}

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

# Row 3: Species Culture and Cycle Type
col5, col6 = st.columns(2)
with col5:
    species = st.selectbox("Species Culture *", SPECIES_CULTURE)
with col6:
    cycle = st.selectbox("Cycle Type *", CYCLE_TYPE)

st.markdown("#### 🐟 Pond Details")

# Header row (labels only)
COLUMN_WIDTHS = [0.9, 0.6, 0.6, 0.8, 0.9, 1.8, 1.6, 1.8, 1.6, 1.1, 1.8, 0.5]
h1, h2, h3, h4, h5, h6, h7, h8, h9, h10, h11, h12 = st.columns(COLUMN_WIDTHS)
for h, label in zip(
    [h1, h2, h3, h4, h5, h6, h7, h8, h9, h10, h11],
    ["Pond Number", "Density", "DOC", "Feed/Day", "ABW", "Diseases Issue", "FEED Issue",
     "Water Quality Issue", "Environment Issue", "Water Color *", "Management & Equipment Issue"]
):
    h.markdown(f"**{label}**")

water_color_choices = ["-- Select --"] + WATER_COLOR_OPTIONS

with st.container(border=True):
    for rid in st.session_state.row_ids:
        row = st.session_state.rows_data[rid]
        c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12 = st.columns(COLUMN_WIDTHS)
        with c1:
            row["pond_number"] = st.text_input("Pond Number", value=row["pond_number"], key=f"pond_number_{rid}", label_visibility="collapsed")
        with c2:
            row["density"] = st.text_input("Density", value=row["density"], key=f"density_{rid}", label_visibility="collapsed")
        with c3:
            row["doc"] = st.text_input("DOC", value=row["doc"], key=f"doc_{rid}", label_visibility="collapsed")
        with c4:
            row["feed_per_day"] = st.text_input("Feed/Day", value=row["feed_per_day"], key=f"feed_per_day_{rid}", label_visibility="collapsed")
        with c5:
            row["abw"] = st.text_input("ABW", value=row["abw"], key=f"abw_{rid}", label_visibility="collapsed")
        with c6:
            row["diseases"] = st.multiselect("Diseases", DISEASES_OPTIONS, default=row["diseases"], key=f"diseases_{rid}", label_visibility="collapsed")
        with c7:
            row["feed_issue"] = st.multiselect("Feed Issue", FEED_ISSUE_OPTIONS, default=row["feed_issue"], key=f"feed_issue_{rid}", label_visibility="collapsed")
        with c8:
            row["water_quality"] = st.multiselect("Water Quality", WATER_QUALITY_OPTIONS, default=row["water_quality"], key=f"water_quality_{rid}", label_visibility="collapsed")
        with c9:
            row["env_issue"] = st.multiselect("Environment", ENVIRONMENT_ISSUE_OPTIONS, default=row["env_issue"], key=f"env_issue_{rid}", label_visibility="collapsed")
        with c10:
            wc_default = row["water_color"] if row["water_color"] in WATER_COLOR_OPTIONS else "-- Select --"
            wc_selected = st.selectbox("Water Color", water_color_choices, index=water_color_choices.index(wc_default), key=f"water_color_{rid}", label_visibility="collapsed")
            row["water_color"] = wc_selected if wc_selected != "-- Select --" else ""
        with c11:
            row["management_issue"] = st.multiselect("Management", MANAGEMENT_ISSUE_OPTIONS, default=row["management_issue"], key=f"management_issue_{rid}", label_visibility="collapsed")
        with c12:
            if st.button("🗑️", key=f"remove_{rid}", help="Remove this row"):
                st.session_state.row_ids.remove(rid)
                del st.session_state.rows_data[rid]
                st.rerun()

if st.button("➕ Add Row"):
    new_id = st.session_state.next_row_id
    st.session_state.row_ids.append(new_id)
    st.session_state.rows_data[new_id] = blank_row()
    st.session_state.next_row_id += 1
    st.rerun()

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
    """Safely convert a text field to a number; blank/invalid becomes 0."""
    value = str(value).strip()
    if value == "":
        return 0 if as_int else 0.0
    try:
        return int(value) if as_int else float(value)
    except ValueError:
        return 0 if as_int else 0.0

if submitted:
    # Collect rows that have at least a pond number, density, or DOC entered
    rows_to_save = [
        row for row in st.session_state.rows_data.values()
        if str(row["pond_number"]).strip() != "" or str(row["density"]).strip() != "" or str(row["doc"]).strip() != ""
    ]

    if not customer or not zone or not area or not species or not cycle or not technician:
        st.error("❌ Please fill in all required top-level fields (marked with *)")
    elif len(rows_to_save) == 0:
        st.error("❌ Please enter at least one pond row before submitting")
    elif any(not r["water_color"] for r in rows_to_save):
        st.error("❌ Water Color is required for every row")
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
                "Species Culture": species,
                "Cycle Type": cycle,
                "Pond Number": r["pond_number"],
                "Density": to_number(r["density"], as_int=True),
                "DOC": to_number(r["doc"], as_int=True),
                "Feed Per Day": to_number(r["feed_per_day"]),
                "ABW": r["abw"],
                "Diseases Issue": ", ".join(r["diseases"]) if r["diseases"] else "",
                "Feed Issue": ", ".join(r["feed_issue"]) if r["feed_issue"] else "",
                "Water Quality Issue": ", ".join(r["water_quality"]) if r["water_quality"] else "",
                "Environment Issue": ", ".join(r["env_issue"]) if r["env_issue"] else "",
                "Water Color": r["water_color"],
                "Management Issue": ", ".join(r["management_issue"]) if r["management_issue"] else "",
                "Remark": remark,
                "Technician": technician
            })

        # Load existing data and append all new rows
        df = load_data()
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

        # Save to CSV
        save_data(df)

        # Save the selected values for next form
        st.session_state.selected_customer = customer
        st.session_state.selected_farm = farm
        st.session_state.selected_zone = zone
        st.session_state.selected_area = area
        st.session_state.selected_technician = technician

        # Reset the pond rows back to 5 fresh blank rows
        st.session_state.row_ids = list(range(st.session_state.next_row_id, st.session_state.next_row_id + STARTING_ROWS))
        st.session_state.rows_data = {rid: blank_row() for rid in st.session_state.row_ids}
        st.session_state.next_row_id += STARTING_ROWS

        st.session_state.form_submitted = True
        st.success(f"✅ {len(new_rows)} row(s) saved successfully!")

        # Auto-refresh after brief delay to see success message
        import time
        time.sleep(1)
        st.rerun()

# Display saved data
st.markdown("---")
st.subheader("📊 Saved Data")

df = load_data()

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

# Clear button always visible
st.markdown("---")
st.markdown("### 🗑️ Delete All Records")

col_delete = st.columns([1, 1, 1])
with col_delete[1]:
    if st.button("Delete All Records", use_container_width=True):
        if os.path.exists(DATA_FILE):
            os.remove(DATA_FILE)
            st.success("✅ All records deleted successfully!")
            st.rerun()

st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>KMN Aqua Services - Water Quality Monitoring System</p>", unsafe_allow_html=True)
