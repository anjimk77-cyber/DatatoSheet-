import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from io import BytesIO

# Configure page
st.set_page_config(page_title="Water Quality Report - Data Collection", layout="wide")

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

COLUMNS = [
    "Timestamp", "Customer", "Farm Name", "Zone", "Area", "Species Culture",
    "Cycle Type", "Pond Number", "Density", "DOC", "Feed Per Day", "AB",
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
        ws = sheet.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=len(COLUMNS))
        ws.append_row(COLUMNS)
    return ws


def load_data() -> pd.DataFrame:
    ws = get_worksheet()
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=COLUMNS)
    return pd.DataFrame(records)


def save_row(new_row: dict):
    ws = get_worksheet()
    if not ws.get_all_values():
        ws.append_row(COLUMNS)
    ws.append_row([str(new_row.get(col, "")) for col in COLUMNS])


def clear_all_data():
    ws = get_worksheet()
    ws.clear()
    ws.append_row(COLUMNS)


# ------------------------------------------------------------------
# Customer list (still loaded from Excel bundled in the repo)
# ------------------------------------------------------------------
@st.cache_data
def load_customer_data():
    df = pd.read_excel("Customer List.xlsx")
    return df


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

# Initialize session state
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
if 'water_color_val' not in st.session_state:
    st.session_state.water_color_val = ""

# Get indices for pre-selected values
customer_index = unique_customers.index(st.session_state.selected_customer) if st.session_state.selected_customer in unique_customers else 0
farm_index = farms.index(st.session_state.selected_farm) if st.session_state.selected_farm in farms else 0
zone_index = zones.index(st.session_state.selected_zone) if st.session_state.selected_zone in zones else 0
area_index = areas.index(st.session_state.selected_area) if st.session_state.selected_area in areas else 0

# ------------------------------------------------------------------
# Form
# ------------------------------------------------------------------
with st.form(f"water_quality_form_{st.session_state.submission_count}"):
    st.subheader("📋 Enter Water Quality Data")

    col1, col2 = st.columns(2)
    with col1:
        customer = st.selectbox("Customer *", unique_customers, index=customer_index)
    with col2:
        farm = st.selectbox("Farm Name", farms, index=farm_index if farms else None)

    col3, col4 = st.columns(2)
    with col3:
        zone = st.selectbox("Zone *", zones if zones else [""], index=zone_index)
    with col4:
        area = st.selectbox("Area *", areas if areas else [""], index=area_index)

    col5, col6 = st.columns(2)
    with col5:
        species = st.selectbox("Species Culture *", SPECIES_CULTURE)
    with col6:
        cycle = st.selectbox("Cycle Type *", CYCLE_TYPE)

    col7, col8, col9 = st.columns(3)
    with col7:
        pond_number = st.text_input("Pond number (Ex: 1,2,3 or A,B,C...)")
    with col8:
        density = st.number_input("Density (PL stocking)", min_value=0, step=1)
    with col9:
        doc = st.number_input("DOC (Days of Culture)", min_value=0, step=1)

    col10, col11 = st.columns(2)
    with col10:
        feed_per_day = st.number_input("Feed Per Day (kg)", min_value=0.0, step=0.1)
    with col11:
        ab = st.text_input("AB (Additional Details)")

    st.markdown("#### 🦐 Diseases Issue")
    diseases = st.multiselect("Select applicable issues", DISEASES_OPTIONS)

    st.markdown("#### 🍚 FEED Issue")
    feed_issue = st.multiselect("Select applicable feed issues", FEED_ISSUE_OPTIONS)

    st.markdown("#### 💧 Water Quality Issue")
    water_quality = st.multiselect("Select applicable water quality issues", WATER_QUALITY_OPTIONS)

    st.markdown("#### 🌍 Environment Issue")
    env_issue = st.multiselect("Select applicable environment issues", ENVIRONMENT_ISSUE_OPTIONS)

    st.markdown("#### 🎨 Water Color *")
    water_color_options = ["-- Select Water Color --"] + WATER_COLOR_OPTIONS
    water_color_display = st.session_state.water_color_val if st.session_state.water_color_val in WATER_COLOR_OPTIONS else "-- Select Water Color --"
    water_color_index = water_color_options.index(water_color_display)
    selected_option = st.selectbox("Select water color", water_color_options, index=water_color_index, key="wc")
    water_color = selected_option if selected_option != "-- Select Water Color --" else ""

    st.markdown("#### ⚙️ Management & Equipment Issue")
    management_issue = st.multiselect("Select applicable management/equipment issues", MANAGEMENT_ISSUE_OPTIONS)

    st.markdown("#### 📝 Remark")
    remark = st.text_area("Additional remarks or notes", placeholder="Enter any additional information", height=80)

    st.markdown("#### 👤 Technician *")
    technician_index = TECHNICIAN_OPTIONS.index(st.session_state.selected_technician) if st.session_state.selected_technician in TECHNICIAN_OPTIONS else 0
    technician = st.selectbox("Select technician", TECHNICIAN_OPTIONS, index=technician_index, key="tech")

    submitted = st.form_submit_button("✅ Submit Data", width='stretch')

    if submitted:
        if not customer or not zone or not area or not species or not cycle or not water_color or not technician:
            st.error("❌ Please fill in all required fields (marked with *)")
        else:
            new_row = {
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Customer": customer,
                "Farm Name": farm,
                "Zone": zone,
                "Area": area,
                "Species Culture": species,
                "Cycle Type": cycle,
                "Pond Number": pond_number,
                "Density": density,
                "DOC": doc,
                "Feed Per Day": feed_per_day,
                "AB": ab,
                "Diseases Issue": ", ".join(diseases) if diseases else "",
                "Feed Issue": ", ".join(feed_issue) if feed_issue else "",
                "Water Quality Issue": ", ".join(water_quality) if water_quality else "",
                "Environment Issue": ", ".join(env_issue) if env_issue else "",
                "Water Color": water_color,
                "Management Issue": ", ".join(management_issue) if management_issue else "",
                "Remark": remark,
                "Technician": technician,
            }

            try:
                save_row(new_row)
            except Exception as e:
                st.error(f"❌ Could not save to Google Sheet: {e}")
            else:
                st.session_state.selected_customer = customer
                st.session_state.selected_farm = farm
                st.session_state.selected_zone = zone
                st.session_state.selected_area = area
                st.session_state.selected_technician = technician
                st.session_state.water_color_val = ""
                st.session_state.submission_count += 1

                st.success("✅ Data saved successfully!")
                import time
                time.sleep(1)
                st.rerun()

# ------------------------------------------------------------------
# Display saved data
# ------------------------------------------------------------------
st.markdown("---")
st.subheader("📊 Saved Data")

try:
    df = load_data()
except Exception as e:
    st.error(f"❌ Could not load data from Google Sheet: {e}")
    df = pd.DataFrame(columns=COLUMNS)

if len(df) > 0:
    st.write(f"Total records: **{len(df)}**")

    st.dataframe(df, width='stretch', height=400)

    col1, col2 = st.columns(2)

    with col1:
        csv = df.to_csv(index=False)
        st.download_button(
            label="📥 Download as CSV",
            data=csv,
            file_name=f"water_quality_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            width='stretch',
        )

    with col2:
        excel_buffer = BytesIO()
        df.to_excel(excel_buffer, index=False, sheet_name="Water Quality Data")
        excel_buffer.seek(0)
        st.download_button(
            label="📥 Download as Excel",
            data=excel_buffer,
            file_name=f"water_quality_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
        )
else:
    st.info("ℹ️ No data saved yet. Fill out the form above to get started!")

# Clear button
st.markdown("---")
st.markdown("### 🗑️ Delete All Records")

col_delete = st.columns([1, 1, 1])
with col_delete[1]:
    if st.button("Delete All Records", width='stretch'):
        try:
            clear_all_data()
            st.success("✅ All records deleted successfully!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Could not clear Google Sheet: {e}")

st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>KMN Aqua Services - Water Quality Monitoring System</p>", unsafe_allow_html=True)
