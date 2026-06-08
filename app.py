# streamlit run app.py

import streamlit as st
import pandas as pd

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Smart Inventory System",
    layout="wide"
)

# ---------------- TITLE ----------------
st.title("📦 Smart Inventory Management System")
st.write("Upload Excel file to calculate stock shortages.")

# ---------------- FILE UPLOAD ----------------
uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

# ---------------- PROCESS FILE ----------------
if uploaded_file is not None:

    # Read Excel
    df = pd.read_excel(uploaded_file)

    # Clean column names
    df.columns = df.columns.astype(str).str.strip()

    # ---------------- DEBUG ----------------
   # st.subheader("🔍 Detected Columns")
    #st.write(df.columns.tolist())

    # ---------------- ORIGINAL DATA ----------------
    st.subheader("📄 Original Data")
    st.dataframe(df)

    # ---------------- STOCK CALCULATION ----------------

    # Convert stock columns to numeric
    df["Status : Current"] = pd.to_numeric(
        df["Status : Current"],
        errors="coerce"
    ).fillna(0)

    df["IPS Consign"] = pd.to_numeric(
        df["IPS Consign"],
        errors="coerce"
    ).fillna(0)

    # Total Stock
    df["Total Stock"] = (
        df["Status : Current"] +
        df["IPS Consign"]
    )

    # ---------------- FIND DATE COLUMNS ----------------

    forecast_columns = []

    for col in df.columns:

        try:
            parsed = pd.to_datetime(col)

            # Only treat columns that look like dates
            if pd.notnull(parsed):
                forecast_columns.append(col)

        except:
            pass

    #st.subheader("📅 Forecast Columns Detected")
    #st.write(forecast_columns)

    # Convert forecast columns to numeric
    for col in forecast_columns:
        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        ).fillna(0)

    # Total Required
    df["Total Required"] = df[forecast_columns].sum(axis=1)

    # To Order
    df["To Order"] = (
        df["Total Required"] -
        df["Total Stock"]
    ).clip(lower=0)

     # ---------------- SHORTAGE DATE & SHORTAGE QTY ----------------

    shortage_qty_list = []
    shortage_date_list = []

    for _, row in df.iterrows():

        remaining_stock = row["Total Stock"]

        shortage_qty = 0
        shortage_date = "No Shortage"

        for date_col in forecast_columns:

            demand = row[date_col]

            remaining_stock -= demand

            if remaining_stock < 0:

                shortage_qty = abs(int(remaining_stock))
                shortage_date = date_col

                break

        shortage_qty_list.append(shortage_qty)
        shortage_date_list.append(shortage_date)

    df["Shortage Qty"] = shortage_qty_list
    df["Shortage Date"] = shortage_date_list
    
    # ---------------- STATUS COLUMN ----------------

    df["Status"] = df["To Order"].apply(
    lambda x: "🟡 Need Order" if x > 0 else "🟢 Sufficient Stock"
)

# ---------------- OUTPUT TABLE ----------------

    output_df = df[
        [
            "Item",
            "Status : Current",
            "IPS Consign",
            "Total Stock",
            "Total Required",
            "To Order",
            "Shortage Qty",
            "Shortage Date",
            "Status"
        ]
    ]

    # ---------------- SUMMARY ----------------

    st.subheader("📊 Inventory Summary")

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Total Products",
        len(output_df)
    )

    col2.metric(
        "Total Required Qty",
        int(output_df["Total Required"].sum())
    )

    col3.metric(
        "Total Order Qty",
        int(output_df["To Order"].sum())
    )

    # ---------------- DISPLAY ----------------

    st.subheader("✅ Processed Inventory Report")

    st.dataframe(
        output_df,
        use_container_width=True
    )

    # ---------------- DOWNLOAD ----------------

    output_file = "processed_inventory.xlsx"

    output_df.to_excel(
        output_file,
        index=False
    )

    with open(output_file, "rb") as file:

        st.download_button(
            label="📥 Download Processed Report",
            data=file,
            file_name="processed_inventory.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        # make a shortage col , which will have the amount of short (number) and the date on which it is falling short on.
        #current stck+IPS -(req1+req2) , so if you get negative , stop and that will go in the shortage col with date
        