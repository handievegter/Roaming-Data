import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import random

def clean_roaming_data(file, cut_off=20):
    xls = pd.ExcelFile(file)
    df = xls.parse("Call Gate June", skiprows=5)

    # Rename columns
    df.columns = [
        "MSISDN", "Transporter", "VehicleReg",
        "CallsRoaming", "CallsData", "TotalExclVAT", "Old Total"
    ]

    for col in ["CallsRoaming", "CallsData", "TotalExclVAT", "Old Total"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["New Total"] = df["Old Total"]
    df["TransporterGroup"] = df["Transporter"].str.replace(r"\s*BUP$", "", regex=True).str.strip()
    df["Status"] = ""

    result_rows = []

    for transporter, group in df.groupby("TransporterGroup"):
        group = group.copy().sort_values(by="VehicleReg").reset_index(drop=True)

        is_small = (group["Old Total"] < cut_off) & (group["Old Total"] > 0)
        is_large = (group["Old Total"] >= cut_off)

        if is_large.any():
            for idx in group[is_small].index:
                small_value = group.at[idx, "Old Total"]
                candidates = group[is_large & (group.index != idx)]

                if not candidates.empty:
                    target_idx = random.choice(candidates.index.tolist())
                    group.at[target_idx, "New Total"] += small_value
                    group.at[target_idx, "Status"] = "added"
                    group.at[idx, "New Total"] = 0
                    group.at[idx, "Status"] = "zeroed"
        else:
            # No large values: consolidate all into one
            total_sum = group["Old Total"].sum()
            collector_idx = random.choice(group.index.tolist())

            for idx in group.index:
                if idx == collector_idx:
                    group.at[idx, "New Total"] = total_sum
                    group.at[idx, "Status"] = "collected"
                else:
                    group.at[idx, "New Total"] = 0
                    group.at[idx, "Status"] = "zeroed"

        group = group.drop(columns="TransporterGroup")

        total_row = {
            "MSISDN": "",
            "Transporter": f"{transporter} - Grand Total",
            "VehicleReg": "",
            "CallsRoaming": "",
            "CallsData": "",
            "TotalExclVAT": "",
            "Old Total": group["Old Total"].sum(),
            "New Total": group["New Total"].sum(),
            "Status": "total"
        }

        result_rows.append(group)
        result_rows.append(pd.DataFrame([total_row]))

        # Two empty spacer rows
        empty_row = pd.DataFrame([[""] * len(group.columns)] * 2, columns=group.columns)
        result_rows.append(empty_row)

    final_df = pd.concat(result_rows, ignore_index=True)
    return final_df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Processed')

    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    # Define styles
    bold_font = Font(bold=True)
    grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    blue_fill = PatternFill(start_color="B3E5FC", end_color="B3E5FC", fill_type="solid")

    # Apply styling based on hidden "Status" column
    status_col_idx = list(df.columns).index("Status") + 1
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        status = row[status_col_idx - 1].value
        if status == "total":
            for cell in row:
                cell.font = bold_font
                cell.fill = grey_fill
        elif status == "zeroed":
            for cell in row:
                cell.fill = red_fill
        elif status == "added":
            for cell in row:
                cell.fill = green_fill
        elif status == "collected":
            for cell in row:
                cell.fill = blue_fill

    # Delete the Status column
    ws.delete_cols(status_col_idx)

    # Format MSISDN column (A) as number with no decimals
    for cell in ws["A"][1:]:  # Skip header
        if isinstance(cell.value, (int, float)):
            cell.number_format = "0"

    # Autosize columns
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            val = str(cell.value) if cell.value is not None else ""
            max_length = max(max_length, len(val))
        ws.column_dimensions[col_letter].width = max_length + 2

    styled_output = BytesIO()
    wb.save(styled_output)
    styled_output.seek(0)
    return styled_output

# Streamlit UI
st.title("Roaming Cost Processor")

st.set_page_config(
   page_title="Roaming Cost Processor",
   page_icon="💸"
)


cut_off = st.number_input(
    "Cut-off for merging small totals",
    min_value=0,
    value=10,
    step=1,
    help="All values below this will be merged into larger ones or grouped if none are large."
)

uploaded_file = st.file_uploader("Upload raw Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df_cleaned = clean_roaming_data(uploaded_file, cut_off)
        st.success("File processed successfully. Download it below:")

        download_file = to_excel(df_cleaned)
        st.download_button(
            label="Download cleaned Excel file",
            data=download_file,
            file_name="processed_roaming_cost.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"An error occurred: {e}")
