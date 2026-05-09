import streamlit as st
import pandas as pd
from io import BytesIO

business_unit_map = {
    "AUSTRIA": {"Sender Name": "Austria", "Sender Location Id": "5000692765"},
    "DENMARK": {"Sender Name": "Denmark", "Sender Location Id": "5000538928"},
    "DRIFFIELD": {"Sender Name": "Driffield", "Sender Location Id": "5000503209"},
    "FRANCE": {"Sender Name": "France", "Sender Location Id": "0101076563"},
    "IRELAND": {"Sender Name": "Ireland", "Sender Location Id": "5000515873"},
    "ITALY": {"Sender Name": "Italy", "Sender Location Id": "0101230808"},
    "NETHERLANDS": {"Sender Name": "Netherlands", "Sender Location Id": "0100646888"},
    "SPAIN": {"Sender Name": "Spain", "Sender Location Id": "5000449357"},
    "HQ": {"Sender Name": "HQ", "Sender Location Id": "1000358868"}
}

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
        value = str(value).strip()

        pallet_map = {
            "03": "CHEP 03 - Euro",
            "01": "CHEP 01 - UK",
            "3-B1208A": "CHEP 03 - Euro"
        }

        for key, mapped_value in pallet_map.items():
            if key in value:
                return mapped_value

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