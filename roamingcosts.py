# streamlit_app.py

import streamlit as st
import pandas as pd
from io import BytesIO

def clean_roaming_data(file):
    # Load the raw Excel file
    xls = pd.ExcelFile(file)
    df = xls.parse("Call Gate June", skiprows=5)

    # Rename columns (make them consistent and readable)
    df.columns = [
        "MSISDN", "Transporter", "VehicleReg", 
        "CallsRoaming", "CallsData", "TotalExclVAT", "Total"
    ]

    # Convert numeric columns
    for col in ["CallsRoaming", "CallsData", "TotalExclVAT", "Total"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Processed')
    output.seek(0)
    return output

# Streamlit UI
st.title("Roaming Cost Processor")

uploaded_file = st.file_uploader("Upload raw Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df_cleaned = clean_roaming_data(uploaded_file)
        st.success("File processed successfully!")
        st.dataframe(df_cleaned.head())

        download_file = to_excel(df_cleaned)
        st.download_button(
            label="Download cleaned Excel file",
            data=download_file,
            file_name="processed_roaming_cost.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"An error occurred: {e}")

# Run the app with: streamlit run code.py
