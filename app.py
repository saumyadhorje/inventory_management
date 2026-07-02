# streamlit run app.py

#Stk+IPS+OpenPO+Transit (Total Avail Column name) -->done
#New Column (Sea1 +Sea2+Sea3+Fedex1+Fedex2) name it transit in phase two ->done
#Change Column Name (Qty (Stk-weekly f/cast)) to (Shortage Qty for Shipment)-->done

import streamlit as st
import pandas as pd
from datetime import datetime

def format_date(date_str):
    dt = pd.to_datetime(date_str)

    day = dt.day

    if 10 <= day % 100 <= 20:
        suffix = "th"
    else:
        suffix = {
            1: "st",
            2: "nd",
            3: "rd"
        }.get(day % 10, "th")

    return f"{day}{suffix} {dt.strftime('%B, %Y')}"

st.set_page_config(
    page_title="Smart Inventory System",
    layout="wide"
)

st.title("📦 Smart Inventory Management System")

phase = st.radio(
    "Select Phase",
    [
        "Phase 1 - Create Processed Report",
        "Phase 2 - Generate Final Report"
    ]
)

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

# =====================================================
# PHASE 1
# =====================================================

if phase == "Phase 1 - Create Processed Report":

    if uploaded_file is not None:

        df = pd.read_excel(uploaded_file)
        #st.write(df.columns.tolist())
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.astype(str).str.strip()
        st.subheader("📄 Original Data")
        st.dataframe(df)

        df["Status : Current"] = pd.to_numeric(
            df["Status : Current"],
            errors="coerce"
        ).fillna(0)

        df["IPS Consign"] = pd.to_numeric(
            df["IPS Consign"],
            errors="coerce"
        ).fillna(0)

        df["Total Stock"] = (
            df["Status : Current"] +
            df["IPS Consign"]
        )

        forecast_columns = []

        for col in df.columns:
            try:
                pd.to_datetime(col)
                forecast_columns.append(col)
            except:
                pass

        for col in forecast_columns:
            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            ).fillna(0)

        df["Total Required"] = df[
            forecast_columns
        ].sum(axis=1)

        df["To Order"] = (
            df["Total Required"] -
            df["Total Stock"]
        ).clip(lower=0)

        shortage_qty_list = []
        shortage_date_list = []

        for _, row in df.iterrows():

            remaining_stock = row["Total Stock"]

            shortage_qty = 0
            shortage_date = "No Shortage"

            for date_col in forecast_columns:

                remaining_stock -= row[date_col]

                if remaining_stock < 0:

                    shortage_qty = (
                        int(remaining_stock)
                    )

                   # shortage_date = date_col
                    shortage_date = format_date(date_col)

                    break

            shortage_qty_list.append(
                shortage_qty
            )

            shortage_date_list.append(
                shortage_date
            )

        df["Shortage Qty"] = shortage_qty_list
        df["Shortage Date"] = shortage_date_list

        df["Status"] = df["To Order"].apply(
            lambda x:
            "🟡 Need Order"
            if x > 0
            else "🟢 Sufficient Stock"
        )

        phase1_df = pd.DataFrame()

        phase1_df["Item"] = df["Item"]

        phase1_df["Status : Current"] = df[
            "Status : Current"
        ]

        phase1_df["IPS Consign"] = df[
            "IPS Consign"
        ]

        phase1_df["Stock +IPS"] = df[
            "Total Stock"
        ]
#sum of all weekly forecast quantities.
        phase1_df["Total Forecast Qty"] = df[
            "Total Required"
        ]

#F/cast qty less total stock =Total Forecast Qty - Stock +IPS
        phase1_df[
            "F/cast qty less total stock"
        ] = df["To Order"]
        
    

        phase1_df[
            "Shortage Qty for Shipment"
        ] = df["Shortage Qty"]

        phase1_df[
            "Shortage Date"
        ] = df["Shortage Date"]

        phase1_df["Status"] = df["Status"]

        phase1_df["Open PO Qty"] = 0
        phase1_df["Supplier stk"] = 0
        phase1_df["Sea transit 1"] = 0
        phase1_df["Sea transit 2"] = 0
        phase1_df["Sea transit 3"] = 0
        phase1_df["Transit fedex 1"] = 0
        phase1_df["Transit fedex 2"] = 0

        # Hidden forecast columns
        for col in forecast_columns:
            phase1_df[col] = df[col]

        st.subheader("📊 Inventory Summary")

        c1, c2, c3 = st.columns(3)

        c1.metric(
            "Total Products",
            len(phase1_df)
        )

        c2.metric(
            "Total Required Qty",
            int(
                phase1_df[
                    "Total Forecast Qty"
                ].sum()
            )
        )

        c3.metric(
            "Total Order Qty",
            int(
                phase1_df[
                    "F/cast qty less total stock"
                ].sum()
            )
        )

        visible_cols = [
            "Item",
            "Status : Current",
            "IPS Consign",
            "Stock +IPS",
            "Total Forecast Qty",
            "F/cast qty less total stock",
            "Shortage Qty for Shipment",
            "Shortage Date",
            "Status",
            "Open PO Qty",
            "Supplier stk",
            "Sea transit 1",
            "Sea transit 2",
            "Sea transit 3",
            "Transit fedex 1",
            "Transit fedex 2"
        ]

        st.subheader(
            "✅ Processed Inventory Report"
        )

        st.dataframe(
            phase1_df[visible_cols],
            use_container_width=True
        )

        output_file = (
            "processed_inventory.xlsx"
        )

        phase1_df.to_excel(
            output_file,
            index=False
        )

        with open(output_file, "rb") as file:

            st.download_button(
                "📥 Download Processed Report",
                data=file,
                file_name=output_file
            )

# =====================================================
# PHASE 2
# =====================================================

elif phase == "Phase 2 - Generate Final Report":

    if uploaded_file is not None:

        df = pd.read_excel(uploaded_file)

        df.columns = (
            df.columns.astype(str).str.strip()
        )

        # Debug - remove later if you want
        #st.write("Columns Found:")
        #st.write(df.columns.tolist())

        # Find hidden forecast date columns
        forecast_columns = []

        for col in df.columns:
            try:
                pd.to_datetime(col)
                forecast_columns.append(col)
            except:
                pass

        editable_cols = [
    "Open PO Qty",
    "Supplier stk",
    "Sea transit 1",
    "Sea transit 2",
    "Sea transit 3",
    "Transit fedex 1",
    "Transit fedex 2"
]
        

        # Convert editable columns to numeric
        for col in editable_cols:

            if col not in df.columns:
                st.error(f"Missing column: {col}")
                st.stop()

            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            ).fillna(0)

        # Transit Total

        df["Transit"] = (
            df["Sea transit 1"] +
            df["Sea transit 2"] +
            df["Sea transit 3"] +
            df["Transit fedex 1"] +
            df["Transit fedex 2"]
        )
        #total avail
        df["Stk+IPS+OpenPO+Transit"] = (
    df["Stock +IPS"] +
    df["Open PO Qty"] +
    df["Supplier stk"] +
    df["Transit"]
)

        # Total shortage
        df["Total shortage"] = (
            df["Total Forecast Qty"] -
            df["Stk+IPS+OpenPO+Transit"]
        ).clip(lower=0)

        # Shortage Qty & Date
        shortage_qty_list = []
        shortage_date_list = []

        for _, row in df.iterrows():

            available_stock = row["Stk+IPS+OpenPO+Transit"]

            shortage_qty = 0
            shortage_date = "No Shortage"

            for date_col in forecast_columns:

                available_stock -= row[date_col]

                if available_stock < 0:

                    shortage_qty = (
                        int(available_stock)
                    )

                    #shortage_date = date_col
                    shortage_date = format_date(date_col)

                    break

            shortage_qty_list.append(
                shortage_qty
            )

            shortage_date_list.append(
                shortage_date
            )

        df["Shortage Qty"] = shortage_qty_list
        df["Shortage Date"] = shortage_date_list

        # Status
        df["Status"] = df[
            "Total shortage"
        ].apply(
            lambda x:
            "🟡 Need Order"
            if x > 0
            else "🟢 Sufficient Stock"
        )

        # Final Output
       # Final Output
        final_df = df[
    [
        "Item",
        "Stock +IPS",
        "Open PO Qty",
        "Supplier stk",
        "Sea transit 1",
        "Sea transit 2",
        "Sea transit 3",
        "Transit fedex 1",
        "Transit fedex 2",
        "Transit",
        "Total Forecast Qty",
        "Stk+IPS+OpenPO+Transit",
        "Total shortage",
        "Shortage Qty",
        "Shortage Date",
        "Status"
    ]
]

        st.subheader(
            "✅ Final Inventory Report"
        )

        st.dataframe(
            final_df,
            use_container_width=True
        )

        final_df.to_excel(
            "final_inventory_report.xlsx",
            index=False
        )

        with open(
            "final_inventory_report.xlsx",
            "rb"
        ) as file:

            st.download_button(
                label="📥 Download Final Report",
                data=file,
                file_name="final_inventory_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )