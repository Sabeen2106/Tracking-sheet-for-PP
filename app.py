import streamlit as st
import pandas as pd
from io import BytesIO
import re

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
    "Coca-Cola HBC Northern Ireland Ltd": {
        "Sender Name": "Coca-Cola HBC Northern Ireland Ltd",
        "Sender Location Id": "5000513592"
    }
}


def clean_columns(df):
    df.columns = df.columns.astype(str).str.strip().str.replace("\xa0", "", regex=False)
    return df


def map_pallet_type(value):
    value = str(value).strip().upper()

    if "PALLET 1000X1200 MM" in value or "1-B1210A" in value or value in ["01", "UK"]:
        return "CHEP 01 - UK"

    if "3-B1208A" in value or value in ["03", "EUR", "EURO", "CHEP 80"]:
        return "CHEP 03 - Euro"

    if "8-B0806A" in value or value == "08":
        return "CHEP 08 - Half"

    return value


def convert_date_to_ddmmyyyy(value):
    if pd.isna(value):
        return pd.NaT

    try:
        if isinstance(value, (int, float)) and value > 1000:
            return pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")

        value_str = str(value).strip()

        if value_str.endswith(".0"):
            value_str = value_str[:-2]

        if value_str.isdigit() and len(value_str) == 8:
            year = int(value_str[:4])
            if year > 1900:
                return pd.Timestamp(
                    year=year,
                    month=int(value_str[4:6]),
                    day=int(value_str[6:8])
                )

        if value_str.isdigit() and len(value_str) == 6:
            return pd.to_datetime(value_str, format="%d%m%y", errors="coerce")

        if "-" in value_str:
            parts = value_str[:10].split("-")
            if len(parts) == 3 and len(parts[0]) == 4:
                return pd.Timestamp(year=int(parts[0]), month=int(parts[1]), day=int(parts[2]))

        if "/" in value_str:
            parts = value_str[:10].split("/")
            if len(parts) == 3 and len(parts[0]) == 4:
                return pd.Timestamp(year=int(parts[0]), month=int(parts[1]), day=int(parts[2]))

        return pd.to_datetime(value_str, errors="coerce", dayfirst=True)

    except Exception:
        return pd.NaT


def clean_reference_number(value):
    value = str(value).strip()

    if value.endswith(".0"):
        value = value[:-2]

    if len(value) > 13:
        value = value[-13:]

    return value


def split_spain_customer_gid(value):
    value = str(value).strip()

    if value.endswith(".0"):
        value = value[:-2]

    value = value.lstrip("-").strip()

    match = re.match(r"^(\d+)[\s\-]*(.*)$", value)

    if match:
        return pd.Series({
            "Customer": match.group(2).strip().lstrip("-").strip(),
            "GID": match.group(1).strip()
        })

    return pd.Series({
        "Customer": value,
        "GID": ""
    })


def is_missing_gid(series):
    return (
        series.isna()
        | (series.astype(str).str.strip() == "")
        | (series.astype(str).str.lower().str.strip() == "nan")
    )


def excel_buffer(df):
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer


def build_tracking_file(final_df, business_unit, pooler, batch_number):
    columns = [
        "Movement Date", "Business Unit", "Pooler", "Movement Direction",
        "Pallet Type", "Reference 1", "Reference 2", "Reference 3",
        "Batch Number", "Sender Location Id", "Sender Name", "Sender Town",
        "Sender Postcode", "Receiver Location Id", "Receiver Name",
        "Receiver Town", "Receiver Postcode", "Movement Type", "Quantity",
        "Savings", "Declared Status"
    ]

    if final_df.empty:
        return pd.DataFrame(columns=columns)

    grouped = final_df.groupby(
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

    return tracking_df[columns]


st.title("PP Tracking Sheet")

business_unit = st.selectbox("Select Business Unit", list(business_unit_map.keys()))
pooler = "CHEP"
batch_number = st.text_input("Enter Batch Number")

main_file = st.file_uploader("Upload Main Excel File", type=["xlsx"])

if main_file:
    df = pd.read_excel(main_file)
    df = clean_columns(df)

    st.subheader("Map Main File Columns")

    columns = df.columns.tolist()

    movement_date_col = st.selectbox("Select column for Movement Date", columns)
    quantity_col = st.selectbox("Select column for Quantity", columns)
    pallet_type_col = st.selectbox("Select column for Pallet Type", columns)
    reference_col = st.selectbox("Select column for Reference", columns)
    customer_col = st.selectbox("Select column for Customer", columns)

    if business_unit == "Spain":
        st.info("Spain selected: GID will be extracted from the Customer column automatically.")
        gid_available = "Spain"
        gid_col = None
        lookup_file = None
    else:
        gid_available = st.radio(
            "Is GID / Location ID already available in the main file?",
            ["Yes", "No"]
        )

        if gid_available == "Yes":
            gid_col = st.selectbox("Select column for GID / Location ID", columns)
            lookup_file = None
        else:
            gid_col = None

            if business_unit == "HQ":
                lookup_file = st.file_uploader(
                    "Upload HQ Location Mapping Table",
                    type=["xlsx"]
                )
            else:
                lookup_file = st.file_uploader(
                    "Upload Customer to GID Matching File",
                    type=["xlsx"]
                )

    if st.button("Prepare Data"):
        if not batch_number:
            st.error("Please enter Batch Number.")
            st.stop()

        work_df = df.copy()

        work_df["Movement Date Parsed"] = work_df[movement_date_col].apply(convert_date_to_ddmmyyyy)

        work_df["Movement Date Parsed"] = pd.to_datetime(
            work_df["Movement Date Parsed"],
            errors="coerce"
        )

        work_df["Movement Date"] = work_df["Movement Date Parsed"].apply(
            lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else ""
        )

        work_df["Quantity"] = pd.to_numeric(
            work_df[quantity_col],
            errors="coerce"
        ).fillna(0)

        work_df = work_df[work_df["Quantity"] != 0].copy()

        work_df["Pallet Type"] = work_df[pallet_type_col].apply(map_pallet_type)

        work_df["Reference"] = work_df[reference_col].apply(clean_reference_number)

        if business_unit == "Spain":
            spain_split = work_df[customer_col].apply(split_spain_customer_gid)

            work_df["Customer"] = spain_split["Customer"]
            work_df["GID"] = spain_split["GID"]

            st.session_state["work_df"] = work_df
            st.session_state["need_lookup_mapping"] = False

        else:
            work_df["Customer"] = work_df[customer_col].astype(str).str.strip()

            if gid_available == "Yes":
                work_df["GID"] = work_df[gid_col]
                st.session_state["work_df"] = work_df
                st.session_state["need_lookup_mapping"] = False

            else:
                if lookup_file is None:
                    st.error("Please upload the required GID matching file.")
                    st.stop()

                lookup_df = pd.read_excel(lookup_file)
                lookup_df = clean_columns(lookup_df)

                st.session_state["lookup_df"] = lookup_df
                st.session_state["work_df_without_gid"] = work_df
                st.session_state["need_lookup_mapping"] = True
                st.session_state["business_unit_for_lookup"] = business_unit


if st.session_state.get("need_lookup_mapping", False):
    lookup_df = st.session_state["lookup_df"]
    work_df = st.session_state["work_df_without_gid"]
    lookup_business_unit = st.session_state["business_unit_for_lookup"]

    if lookup_business_unit == "HQ":
        st.subheader("HQ Location Mapping")

        main_address_col = st.selectbox(
            "Select Address Line 1 column from Main File",
            work_df.columns.tolist(),
            key="hq_main_address_col"
        )

        lookup_address_col = st.selectbox(
            "Select Address1 column from HQ Mapping Table",
            lookup_df.columns.tolist(),
            key="hq_lookup_address_col"
        )

        lookup_gid_col = st.selectbox(
            "Select CHEP GID column from HQ Mapping Table",
            lookup_df.columns.tolist(),
            key="hq_lookup_gid_col"
        )

        if st.button("Apply HQ Mapping"):
            work_df["Address Match Key"] = (
                work_df[main_address_col]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            lookup_df["Address Match Key"] = (
                lookup_df[lookup_address_col]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            lookup_df = lookup_df[["Address Match Key", lookup_gid_col]].copy()

            lookup_df.rename(columns={
                lookup_gid_col: "GID"
            }, inplace=True)

            lookup_df = lookup_df.dropna(subset=["Address Match Key"])

            lookup_df = lookup_df.drop_duplicates(
                subset=["Address Match Key"],
                keep="first"
            )

            work_df = work_df.merge(
                lookup_df,
                on="Address Match Key",
                how="left"
            )

            st.session_state["work_df"] = work_df
            st.session_state["need_lookup_mapping"] = False

    else:
        st.subheader("Customer to GID Mapping")

        lookup_customer_col = st.selectbox(
            "Select Customer column in lookup file",
            lookup_df.columns.tolist(),
            key="lookup_customer_col"
        )

        lookup_gid_col = st.selectbox(
            "Select GID / Location ID column",
            lookup_df.columns.tolist(),
            key="lookup_gid_col"
        )

        if st.button("Apply Customer Lookup"):
            lookup_df = lookup_df[[lookup_customer_col, lookup_gid_col]].copy()

            lookup_df.rename(columns={
                lookup_customer_col: "Customer",
                lookup_gid_col: "GID"
            }, inplace=True)

            lookup_df["Customer"] = lookup_df["Customer"].astype(str).str.strip()

            lookup_df = lookup_df.dropna(subset=["Customer"])

            lookup_df = lookup_df.drop_duplicates(
                subset=["Customer"],
                keep="first"
            )

            work_df = work_df.merge(
                lookup_df,
                on="Customer",
                how="left"
            )

            st.session_state["work_df"] = work_df
            st.session_state["need_lookup_mapping"] = False


if "work_df" in st.session_state:
    work_df = st.session_state["work_df"].copy()

    today = pd.Timestamp.today().normalize()

    work_df["Days Old"] = (
        today - work_df["Movement Date Parsed"]
    ).dt.days

    old_date_df = work_df[work_df["Days Old"] > 89].copy()

    if not old_date_df.empty:
        st.warning("Some movement dates are more than 89 days old.")
        st.subheader("Review Dates More Than 89 Days Old")

        for idx, row in old_date_df.iterrows():
            reference = row["Reference"]
            old_date = row["Movement Date"]

            st.write(f"Reference Number {reference} has date {old_date}")

            amend_date = st.checkbox(
                f"Amend date for reference {reference}",
                key=f"amend_date_{idx}"
            )

            if amend_date:
                new_date = st.date_input(
                    f"Enter new date for reference {reference}",
                    key=f"new_date_{idx}"
                )

                work_df.loc[idx, "Movement Date Parsed"] = pd.to_datetime(new_date)
                work_df.loc[idx, "Movement Date"] = pd.to_datetime(new_date).strftime("%d/%m/%Y")

    missing_gid_df = work_df[is_missing_gid(work_df["GID"])].copy()

    gid_inputs = {}
    remove_rows = []

    if not missing_gid_df.empty:
        st.warning("Some rows are missing GID / Location ID.")
        st.subheader("Missing GID Input by Customer and Reference")

        for idx, row in missing_gid_df.iterrows():
            customer = row["Customer"]
            reference = row["Reference"]
            quantity = row["Quantity"]
            movement_date = row["Movement Date"]

            with st.container():
                st.markdown("---")
                st.write(f"Customer: {customer}")
                st.write(f"Reference Number: {reference}")
                st.write(f"Movement Date: {movement_date}")
                st.write(f"Quantity: {quantity}")

                gid_inputs[idx] = st.text_input(
                    "Please enter GID for this row",
                    key=f"gid_input_{idx}"
                )

                remove = st.checkbox(
                    "GID not found - remove this row",
                    key=f"remove_row_{idx}"
                )

                if remove:
                    remove_rows.append(idx)

    else:
        st.success("No missing GIDs found.")

    if st.button("Generate Final Files"):
        for idx, gid_value in gid_inputs.items():
            if gid_value.strip():
                work_df.loc[idx, "GID"] = gid_value.strip()

        removed_rows_df = work_df.loc[remove_rows].copy()

        gid_required_df = work_df[
            is_missing_gid(work_df["GID"])
            & ~work_df.index.isin(remove_rows)
        ].copy()

        final_df = work_df[
            ~is_missing_gid(work_df["GID"])
            & ~work_df.index.isin(remove_rows)
        ].copy()

        tracking_df = build_tracking_file(final_df, business_unit, pooler, batch_number)
        removed_tracking_df = build_tracking_file(removed_rows_df, business_unit, pooler, batch_number)
        gid_required_tracking_df = build_tracking_file(gid_required_df, business_unit, pooler, batch_number)

        st.success("Files generated successfully.")

        st.download_button(
            label="Download Tracking Sheet",
            data=excel_buffer(tracking_df),
            file_name=f"{batch_number}_tracking_sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download GID Required File",
            data=excel_buffer(gid_required_tracking_df),
            file_name=f"{batch_number}_gid_required.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download Removed Rows",
            data=excel_buffer(removed_tracking_df),
            file_name="Removed rows.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
