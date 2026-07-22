import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configure page
st.set_page_config(page_title="Water Quality Report - Data Collection", layout="wide")

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
        return pd.read_csv(DATA_FILE)
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
if 'ab_val' not in st.session_state:
    st.session_state.ab_val = ""
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

# Create form for input
with st.form(f"water_quality_form_{st.session_state.submission_count}"):
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
    
    # Row 4: Pond Number, Density, DOC
    col7, col8, col9 = st.columns(3)
    with col7:
        pond_number = st.text_input("Pond number (Ex: 1,2,3 or A,B,C...)")
    with col8:
        density = st.number_input("Density (PL stocking)", min_value=0, step=1)
    with col9:
        doc = st.number_input("DOC (Days of Culture)", min_value=0, step=1)
    
    # Row 5: Feed Per Day and AB
    col10, col11 = st.columns(2)
    with col10:
        feed_per_day = st.number_input("Feed Per Day (kg)", min_value=0.0, step=0.1)
    with col11:
        ab = st.text_input("AB (Additional Details)")
    
    # Diseases Issue
    st.markdown("#### 🦐 Diseases Issue")
    diseases = st.multiselect("Select applicable issues", DISEASES_OPTIONS)
    
    # Feed Issue
    st.markdown("#### 🍚 FEED Issue")
    feed_issue = st.multiselect("Select applicable feed issues", FEED_ISSUE_OPTIONS)
    
    # Water Quality Issue
    st.markdown("#### 💧 Water Quality Issue")
    water_quality = st.multiselect("Select applicable water quality issues", WATER_QUALITY_OPTIONS)
    
    # Environment Issue
    st.markdown("#### 🌍 Environment Issue")
    env_issue = st.multiselect("Select applicable environment issues", ENVIRONMENT_ISSUE_OPTIONS)
    
    # Water Color
    st.markdown("#### 🎨 Water Color *")
    water_color_options = ["-- Select Water Color --"] + WATER_COLOR_OPTIONS
    water_color_display = st.session_state.water_color_val if st.session_state.water_color_val in WATER_COLOR_OPTIONS else "-- Select Water Color --"
    water_color_index = water_color_options.index(water_color_display)
    selected_option = st.selectbox("Select water color", water_color_options, index=water_color_index, key="wc")
    water_color = selected_option if selected_option != "-- Select Water Color --" else ""
    
    # Management & Equipment Issue
    st.markdown("#### ⚙️ Management & Equipment Issue")
    management_issue = st.multiselect("Select applicable management/equipment issues", MANAGEMENT_ISSUE_OPTIONS)
    
    # Remark
    st.markdown("#### 📝 Remark")
    remark = st.text_area("Additional remarks or notes", placeholder="Enter any additional information", height=80)
    
    # Technician
    st.markdown("#### 👤 Technician *")
    technician_index = TECHNICIAN_OPTIONS.index(st.session_state.selected_technician) if st.session_state.selected_technician in TECHNICIAN_OPTIONS else 0
    technician = st.selectbox("Select technician", TECHNICIAN_OPTIONS, index=technician_index, key="tech")
    
    # Submit button
    submitted = st.form_submit_button("✅ Submit Data", use_container_width=True)
    
    if submitted:
        # Validate required fields
        if not customer or not zone or not area or not species or not cycle or not water_color or not technician:
            st.error("❌ Please fill in all required fields (marked with *)")
        else:
            # Create new row
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
                "Technician": technician
            }
            
            # Load existing data and append new row
            df = load_data()
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            
            # Save to CSV
            save_data(df)
            
            # Save the selected values for next form
            st.session_state.selected_customer = customer
            st.session_state.selected_farm = farm
            st.session_state.selected_zone = zone
            st.session_state.selected_area = area
            st.session_state.selected_technician = technician
            
            # Clear all other form fields for next entry
            st.session_state.pond_number_val = ""
            st.session_state.density_val = 0
            st.session_state.doc_val = 0
            st.session_state.feed_per_day_val = 0.0
            st.session_state.ab_val = ""
            st.session_state.diseases_val = []
            st.session_state.feed_issue_val = []
            st.session_state.water_quality_val = []
            st.session_state.env_issue_val = []
            st.session_state.management_issue_val = []
            st.session_state.remark_val = ""
            st.session_state.water_color_val = ""
            st.session_state.species_val = 0
            st.session_state.cycle_val = 0
            
            # Increment submission count to reset form
            st.session_state.submission_count += 1
            
            st.session_state.form_submitted = True
            st.success("✅ Data saved successfully!")
            
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
    
    # Display data in a table
    st.dataframe(df, use_container_width=True, height=400)
    
    # Download options
    col1, col2 = st.columns(2)
    
    with col1:
        # Download as CSV
        csv = df.to_csv(index=False)
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
        df.to_excel(excel_buffer, index=False, sheet_name="Water Quality Data")
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
