import streamlit as st
import pandas as pd
from io import BytesIO

# =========================
# BUSINESS UNIT MAPPING
# =========================
business_unit_map = {
    "Austria": {"Sender Name": "Austria", "Sender Location Id": "5000692765"},
    "Denmark": {"Sender Name": "Denmark", "Sender Location Id": "5000538928"},
    "Driffield": {"Sender Name": "Driffield", "Sender Location Id": "5000503209"},
    "France": {"Sender Name": "France", "Sender Location Id": "0101076563"},
    "Ireland": {"Sender Name": "Ireland", "Sender Location Id": "5000515873"},
    "Italy": {"Sender Name": "Italy", "Sender Location Id": "0101230808"},
    "Netherlands": {"Sender Name": "Netherlands", "Sender Location Id": "0100646888"},
    "Spain": {"Sender Name": "Spain", "Sender Location Id": "5000449357"},
    "HQ": {"Sender Name": "HQ", "Sender Location Id": "1000358868"},
    "Coca-Cola HBC Northern Ireland Ltd": {"Sender Name": "Coca-Cola HBC Northern Ireland Ltd", "Sender Location Id": "5000513592"}
}

# =========================
# USER INPUTS
# =========================
business_unit = input("Enter Business Unit: ")

if business_unit not in business_unit_map:
    raise ValueError(" Invalid Business Unit")

batch_number = input("Enter Batch Number: ")

# =========================
# LOAD MAIN FILE
# =========================
df = pd.read_excel("/content/060526 CCH CHEP.xlsx")

# Clean columns
df.columns = df.columns.str.strip().str.replace('\xa0', '', regex=True)

# Rename required fields
df.rename(columns={
    'Customer': 'Customer',
    'Delivery': 'Reference',
    'Billing doc. date': 'Date'
}, inplace=True)

# Ensure clean key
df['Customer'] = df['Customer'].astype(str).str.strip()

# =========================
# LOAD LOOKUP FILE (XLOOKUP SOURCE)
# =========================
lookup_df = pd.read_excel("/content/CCH IPP and CHEP.xlsx")

lookup_df.columns = lookup_df.columns.str.strip().str.replace('\xa0', '', regex=True)

lookup_df.rename(columns={
    'Customer': 'Customer',
    'Location ID': 'Location Id'
}, inplace=True)

lookup_df['Customer'] = lookup_df['Customer'].astype(str).str.strip()

# =========================
# XLOOKUP (CORRECT LEFT JOIN)
# =========================
df = df.merge(
    lookup_df[['Customer', 'Location Id']],
    on='Customer',
    how='left'
)

# =========================
# VALIDATION CHECK (IMPORTANT)
# =========================
if 'Location Id' not in df.columns:
    raise ValueError(" Location Id not created from lookup")

# =========================
# FIX PALLET TYPE
# =========================
df['Pallet Type'] = 'CHEP 01 - UK'


'''# =========================
# FIX EXCEL SERIAL DATE
# =========================

df['Date'] = pd.to_datetime(
    df['Date'],
    unit='D',
    origin='1899-12-30',
    errors='coerce'
)

df['Date'] = df['Date'].dt.strftime('%d/%m/%Y')'''

# =========================
# BUILD TRACKING FILE
# =========================
tracking_df = pd.DataFrame()

tracking_df['Movement Date'] = df['Date']
tracking_df['Business Unit'] = business_unit
tracking_df['Pooler'] = 'CHEP'
tracking_df['Movement Direction'] = 'Out'
tracking_df['Pallet Type'] = df['Pallet Type']
tracking_df['Reference 1'] = df['Reference']
tracking_df['Reference 2'] = ''
tracking_df['Reference 3'] = ''
tracking_df['Batch Number'] = batch_number

# Sender
tracking_df['Sender Location Id'] = business_unit_map[business_unit]['Sender Location Id']
tracking_df['Sender Name'] = business_unit_map[business_unit]['Sender Name']
tracking_df['Sender Town'] = ''
tracking_df['Sender Postcode'] = ''

# Receiver (FROM LOOKUP)
tracking_df['Receiver Location Id'] = df['Location Id']
tracking_df['Receiver Name'] = df['Customer']
tracking_df['Receiver Town'] = ''
tracking_df['Receiver Postcode'] = ''

# Movement
tracking_df['Movement Type'] = 'Out - CHEP Drop Point'
tracking_df['Quantity'] = df['Quantity']
tracking_df['Savings'] = ''
tracking_df['Declared Status'] = 'Declared'

# =========================
# CLEAN & EXPORT
# =========================
tracking_df = tracking_df.dropna(subset=['Movement Date'])

output_file = f"{batch_number}.xlsx"
tracking_df.to_excel(output_file, index=False)

print(f"✅ Tracking sheet created: {output_file}")



st.title("Tracking Sheet Generator")

business_unit = st.selectbox("Select Business Unit", list(business_unit_map.keys()))
pooler = "CHEP"
batch_number = st.text_input("Enter Batch Number")

main_file = st.file_uploader("Upload Main Excel File", type=["xlsx"])

if main_file:
    df = pd.read_excel(main_file)
    df.columns = df.columns.astype(str).str.strip()

    st.subheader("Map Input Columns")

    columns = df.columns.tolist()

    movement_date_col = st.selectbox("Select column for Movement Date", columns)
    quantity_col = st.selectbox("Select column for Quantity", columns)
    pallet_type_col = st.selectbox("Select column for Pallet Type", columns)
    reference_col = st.selectbox("Select column for Reference", columns)
    customer_col = st.selectbox("Select column for Customer", columns)

    gid_available = st.radio(
        "Is GID / Location ID already available in the main file?",
        ["Yes", "No"]
    )

    gid_col = None
    lookup_file = None

    if gid_available == "Yes":
        gid_col = st.selectbox("Select column for GID / Location ID", columns)
    else:
        lookup_file = st.file_uploader(
            "Upload Customer to GID Matching File",
            type=["xlsx"]
        )

         
def map_pallet_type(value):
    value = str(value).strip().upper()

    if (
        "PALLET 1000X1200 MM" in value
        or "1-B1210A" in value
        or value == "01"
    ):
        return "CHEP 01 - UK"

    elif (
        "3-B1208A" in value
        or value == "03"
    ):
        return "CHEP 03 - Euro"

    elif (
        "8-B0806A" in value
        or value == "08"
    ):
        return "CHEP 08 - Half"

    return value    

    if st.button("Generate Tracking File"):

        work_df = df.copy()

        work_df["Movement Date"] = work_df[movement_date_col]
        work_df["Quantity"] = pd.to_numeric(work_df[quantity_col], errors="coerce").fillna(0)
        work_df["Pallet Type"] = work_df[pallet_type_col].apply(map_pallet_type)
        work_df["Reference"] = work_df[reference_col]
        work_df["Customer"] = work_df[customer_col]

        if gid_available == "Yes":
            work_df["GID"] = work_df[gid_col]
        else:
            if lookup_file is None:
                st.error("Please upload the Customer to GID matching file.")
                st.stop()

            lookup_df = pd.read_excel(lookup_file)
            lookup_df.columns = lookup_df.columns.astype(str).str.strip()

            st.subheader("Map Lookup File Columns")

            lookup_customer_col = st.selectbox(
                "Select Customer column in lookup file",
                lookup_df.columns.tolist()
            )

            lookup_gid_col = st.selectbox(
                "Select GID / Location ID column in lookup file",
                lookup_df.columns.tolist()
            )

            lookup_df = lookup_df[[lookup_customer_col, lookup_gid_col]].drop_duplicates()

            lookup_df.rename(columns={
                lookup_customer_col: "Customer",
                lookup_gid_col: "GID"
            }, inplace=True)

            work_df = work_df.merge(
                lookup_df,
                on="Customer",
                how="left"
            )

        grouped = work_df.groupby(
            ["Reference", "Pallet Type"],
            as_index=False
        ).agg({
            "Quantity": "sum",
            "Movement Date": "first",
            "Customer": "first",
            "GID": "first"
        })

        tracking_df = pd.DataFrame({
            "Movement Date": grouped["Movement Date"],
            "Business Unit": business_unit_map[business_unit]["Sender Name"],
            "Pooler": pooler,
            "Movement Direction": "Out",
            "Pallet Type": grouped["Pallet Type"],
            "Reference 1": grouped["Reference"],
            "Reference 2": "",
            "Reference 3": "",
            "Batch Number": batch_number,
            "Sender Location Id": business_unit_map[business_unit]["Sender Location Id"],
            "Sender Name": business_unit_map[business_unit]["Sender Name"],
            "Sender Town": "",
            "Sender Postcode": "",
            "Receiver Location Id": grouped["GID"],
            "Receiver Name": grouped["Customer"],
            "Receiver Town": "",
            "Receiver Postcode": "",
            "Movement Type": f"Out - {pooler} Drop Point",
            "Quantity": grouped["Quantity"],
            "Savings": "",
            "Declared Status": "Declared"
        })

        buffer = BytesIO()
        tracking_df.to_excel(buffer, index=False)
        buffer.seek(0)

        st.success("Tracking file generated successfully!")

        st.download_button(
            label="Download Tracking File",
            data=buffer,
            file_name=f"{batch_number}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
