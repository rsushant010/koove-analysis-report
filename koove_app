import streamlit as st
import pandas as pd
import io
import datetime
import numpy as np
import re
import zipfile

# --- Helper Functions for Data Processing ---
# These functions contain the core logic for processing each part of the report.

def find_sheet_by_keyword(xls, keyword):
    """Finds the first sheet in an Excel file containing a specific keyword (case-insensitive)."""
    for sheet_name in xls.sheet_names:
        if keyword.lower() in sheet_name.lower():
            return sheet_name
    return None

def process_oee_data(oee_df, target_date, analysis_df):
    """Processes the OEE data for the given date and updates the analysis DataFrame."""
    daily_oee_data = oee_df[oee_df['Date'] == target_date].copy()
    if daily_oee_data.empty:
        st.warning(f"No OEE data found for {target_date.strftime('%Y-%m-%d')}.")
        return analysis_df

    for i in range(1, 4):
        line_name = f'Line {i}'
        line_data = daily_oee_data[daily_oee_data['Line'] == line_name]
        hrs_sno, cap_sno, qual_sno, oee_sno = i * 4 - 3, i * 4 - 2, i * 4 - 1, i * 4

        if not line_data.empty and line_data.iloc[0]['OEE'] > 0:
            run_time, total_pcs, quality, oee, downtime = line_data.iloc[0][['Run-Time', 'Total-Pcs', 'Quality', 'OEE', 'Downtime (Hours)']]
            analysis_df.loc[hrs_sno, 'Actual'] = f"{run_time:.0f} Hrs"
            remark_hrs = f"Operation time {run_time:.0f} hrs."
            if run_time < 22 and downtime > 0: remark_hrs += f" {downtime:.1f} hrs downtime."
            analysis_df.loc[hrs_sno, 'Remark'] = remark_hrs
            prod_capacity = (total_pcs / run_time) if run_time > 0 else 0
            analysis_df.loc[cap_sno, 'Actual'] = f"{prod_capacity:,.0f} Pcs/hrs"
            analysis_df.loc[cap_sno, 'Remark'] = f"Actual Production rate {prod_capacity:,.0f} pcs/hr."
            analysis_df.loc[qual_sno, 'Actual'] = f"{quality * 100:.0f} %"
            analysis_df.loc[qual_sno, 'Remark'] = f"Quality is {quality * 100:.0f}%."
            analysis_df.loc[oee_sno, 'Actual'] = f"{oee * 100:.0f} %"
            analysis_df.loc[oee_sno, 'Remark'] = f"OEE is {oee * 100:.0f}%."
        else:
            for sno in [hrs_sno, cap_sno, qual_sno, oee_sno]:
                analysis_df.loc[sno, 'Actual'] = 'shutdown'
            analysis_df.loc[hrs_sno, 'Remark'] = 'Line was shutdown for the day.'
    return analysis_df

def process_production_target(xls, analysis_df):
    prod_sheet_name = find_sheet_by_keyword(xls, "Gloves Production")
    if not prod_sheet_name: 
        st.warning("'Gloves Production' sheet not found for production target processing.")
        return analysis_df
    prod_df = pd.read_excel(xls, sheet_name=prod_sheet_name, header=None)
    x, y = -1, -1
    for r in range(min(10, len(prod_df))):
        for c in range(len(prod_df.columns)):
            if "target mtd" in str(prod_df.iloc[r, c]).lower():
                x, y = r, c
                break
        if x != -1: break
    if x != -1:
        try:
            target_qty, actual_qty, actual_percent = prod_df.iloc[x, y + 1], prod_df.iloc[x + 1, y + 1], prod_df.iloc[x + 2, y + 1]
            analysis_df.loc[14, 'Standard'], analysis_df.loc[14, 'Actual'], analysis_df.loc[14, 'Remark'] = f"{target_qty:,.0f} Pcs", f"{actual_qty:,.0f} Pcs", f"Actual production is {actual_qty:,.0f} pcs."
            analysis_df.loc[15, 'Standard'], analysis_df.loc[15, 'Actual'], analysis_df.loc[15, 'Remark'] = "100 %", f"{actual_percent:.1f} %", f"Achievement is {actual_percent:.1f}%."
        except IndexError: st.warning("Could not find production target values in the expected cells.")
    else: st.warning("'Target MTD' keyword not found in 'Gloves Production' sheet.")
    return analysis_df

def process_consumption(xls, target_date, analysis_df):
    consump_sheet_name = find_sheet_by_keyword(xls, "coal & elec")
    if not consump_sheet_name: 
        st.warning("'coal & elec' sheet not found for consumption processing.")
        return analysis_df
    consump_df = pd.read_excel(xls, sheet_name=consump_sheet_name, header=None)
    date_col_index = None
    for r in range(consump_df.shape[0]):
        for c in range(consump_df.shape[1]):
            try:
                if pd.to_datetime(consump_df.iloc[r, c]).date() == target_date.date():
                    date_col_index = c
                    break
            except (ValueError, TypeError): continue
        if date_col_index is not None: break
    if date_col_index is None: 
        st.warning(f"No consumption data found for date {target_date.strftime('%Y-%m-%d')}.")
        return analysis_df
    coal_row_index, elec_row_index = None, None
    for i in range(consump_df.shape[0]):
        particular = str(consump_df.iloc[i, 0]).lower()
        if "coal" in particular: coal_row_index = i
        if "electri" in particular: elec_row_index = i
    if coal_row_index is not None:
        coal_val = consump_df.iloc[coal_row_index, date_col_index]
        analysis_df.loc[17, 'Actual'], analysis_df.loc[17, 'Remark'] = f"{coal_val:,.0f} Kg", f"Coal consumption is {coal_val:,.0f} Kg."
    if elec_row_index is not None:
        elec_val = consump_df.iloc[elec_row_index, date_col_index]
        analysis_df.loc[18, 'Actual'], analysis_df.loc[18, 'Remark'] = f"{elec_val:,.0f} Unit", f"Electricity consumption is {elec_val:,.0f} unit."
    return analysis_df

def process_abnormalities(xls, analysis_df):
    analysis_df.loc[13, 'Actual'], analysis_df.loc[13, 'Remark'] = "13 Nos", "13 nos. of Abnormalities are found."
    return analysis_df

def process_inventory(xls, analysis_df):
    prod_sheet_name = find_sheet_by_keyword(xls, "Gloves Production")
    if not prod_sheet_name: 
        st.warning("'Gloves Production' sheet not found for inventory check.")
        return analysis_df
    prod_df = pd.read_excel(xls, sheet_name=prod_sheet_name, header=None)
    start_row, start_col = -1, -1
    for r in range(len(prod_df)):
        for c in range(len(prod_df.columns)):
            if "xnbr latex" in str(prod_df.iloc[r, c]).lower():
                start_row, start_col = r, c
                break
        if start_row != -1: break
    if start_row == -1: 
        st.warning("'XNBR LATEX' keyword not found for inventory check.")
        return analysis_df
    inventory_rows = []
    current_sno = 19
    for r in range(start_row, len(prod_df)):
        particular = prod_df.iloc[r, start_col]
        if pd.isna(particular) or particular == "": break
        actual_val, days_val = prod_df.iloc[r, start_col + 1], prod_df.iloc[r, start_col + 2]
        try: remark = f"Stock available for {int(float(days_val))} days."
        except (ValueError, TypeError): remark = str(days_val)
        if "xnbr latex" in str(particular).lower():
            try:
                in_transit = prod_df.iloc[r, start_col + 3]
                if pd.notna(in_transit): remark += f" {in_transit} In transit."
            except IndexError: pass
        inventory_rows.append({'Sl. No.': current_sno, 'Particulars': str(particular).upper(), 'Unit': 'Kg', 'Standard': '', 'Actual': f"{actual_val:,.0f} Kg", 'Remark': remark})
        current_sno += 1
    if inventory_rows:
        analysis_df = pd.concat([analysis_df, pd.DataFrame(inventory_rows).set_index('Sl. No.')])
    return analysis_df

def process_order_details(xls, analysis_df):
    order_sheet_name = find_sheet_by_keyword(xls, "Clear order details")
    if not order_sheet_name: 
        st.warning("'Clear order details' sheet not found.")
        return analysis_df
    order_df = pd.read_excel(xls, sheet_name=order_sheet_name)
    last_row = order_df.iloc[-1]
    clear_order_col, dispatch_col, pending_col = None, None, None
    for col in order_df.columns:
        col_lower = str(col).lower()
        if "total payment receive" in col_lower: clear_order_col = col
        if "total dispatch price" in col_lower: dispatch_col = col
        if "advance payment" in col_lower: pending_col = col
    financial_rows = []
    if clear_order_col:
        val = last_row[clear_order_col]
        financial_rows.append({'Sl. No.': 24, 'Particulars': 'CLEAR ORDER VALUE', 'Unit': 'Rs.', 'Standard': '', 'Actual': f"Rs. {val:,.2f}", 'Remark': f"Total Clear Order value is Rs. {val:,.2f}"})
    if dispatch_col:
        val = last_row[dispatch_col]
        financial_rows.append({'Sl. No.': 25, 'Particulars': 'TOTAL DISPATCH VALUE', 'Unit': 'Rs.', 'Standard': '', 'Actual': f"Rs. {val:,.2f}", 'Remark': f"Total dispatch value is Rs. {val:,.2f}"})
    if pending_col:
        val = last_row[pending_col]
        financial_rows.append({'Sl. No.': 26, 'Particulars': 'PENDING', 'Unit': 'Rs.', 'Standard': '', 'Actual': f"Rs. {val:,.2f}", 'Remark': f"Pending quantity value is Rs. {val:,.2f}"})
    if financial_rows:
        analysis_df = pd.concat([analysis_df, pd.DataFrame(financial_rows).set_index('Sl. No.')])
    return analysis_df

# --- Streamlit App UI ---

st.set_page_config(layout="wide")
st.title("📊 Multi-File Analysis Report Generator")

st.sidebar.header("⚙️ Controls")
uploaded_files = st.sidebar.file_uploader(
    "1. Upload Excel Workbooks",
    type="xlsx",
    accept_multiple_files=True
)

file_dates = {}
if uploaded_files:
    st.sidebar.subheader("2. Set Analysis Dates")
    for uploaded_file in uploaded_files:
        file_dates[uploaded_file.name] = st.sidebar.date_input(
            f"Date for {uploaded_file.name}",
            datetime.date.today(),
            key=uploaded_file.name
        )

    if st.sidebar.button("🚀 Generate & Prepare Download", type="primary"):
        st.header("Processing Files...")
        progress_bar = st.progress(0)
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for i, uploaded_file in enumerate(uploaded_files):
                file_name = uploaded_file.name
                target_date = pd.to_datetime(file_dates[file_name])
                
                st.subheader(f"Processing '{file_name}' for {target_date.strftime('%Y-%m-%d')}")
                
                try:
                    xls = pd.ExcelFile(uploaded_file)
                    oee_sheet_name = find_sheet_by_keyword(xls, "oee")
                    if not oee_sheet_name:
                        st.error(f"Could not find 'OEE' sheet in '{file_name}'. Skipping.")
                        continue
                        
                    oee_df = pd.read_excel(xls, sheet_name=oee_sheet_name)
                    oee_df['Date'] = pd.to_datetime(oee_df['Date'], errors='coerce')
                    oee_df.dropna(subset=['Date'], inplace=True)
                    oee_df.fillna(0, inplace=True)

                    analysis_data = {'Sl. No.': [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,17,18], 'Particulars': ['LINE 1 Working Hrs.','LINE 1 Production Capacity','LINE 1 Quality','LINE 1 OEE','LINE 2 Working Hrs.','LINE 2 Production Capacity','LINE 2 Quality','LINE 2 OEE','LINE 3 Working Hrs.','LINE 3 Production Capacity','LINE 3 Quality','LINE 3 OEE','ABNORMALITIES (1+2+3)','PRODUCTION TARGET (Qty/day)','PRODUCTION TARGET (in %)','COAL CONSUMPTION PER DAY','ELECTRICITY CONSUMPTION PER DAY'], 'Unit': ['Hrs','Pcs/hrs','%','%','Hrs','Pcs/hrs','%','%','Hrs','Pcs/hrs','%','%','Nos','Pcs','%','Kg','Unit'], 'Standard': ['22 Hrs','18181 Pcs/hrs','100 %','80 %','22 Hrs','22727 Pcs/hrs','100 %','80 %','22 Hrs','22727 Pcs/hrs','100 %','80 %','0 Nos','','','',''], 'Actual': [''] * 17, 'Remark': [''] * 17}
                    analysis_df = pd.DataFrame(analysis_data).set_index('Sl. No.')
                    
                    # Run all processing functions
                    analysis_df = process_oee_data(oee_df, target_date, analysis_df)
                    analysis_df = process_abnormalities(xls, analysis_df)
                    analysis_df = process_production_target(xls, analysis_df)
                    analysis_df = process_consumption(xls, target_date, analysis_df)
                    analysis_df = process_inventory(xls, analysis_df)
                    analysis_df = process_order_details(xls, analysis_df)

                    final_report_df = analysis_df.sort_index().reset_index().drop(columns=['Sl. No.'])
                    
                    # Preview the generated report in the app
                    st.dataframe(final_report_df.fillna(''))

                    # Create the new workbook in memory
                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        final_report_df.to_excel(writer, sheet_name='Analysis Points', index=False)
                        for sheet_name in xls.sheet_names:
                            original_sheet_df = pd.read_excel(xls, sheet_name=sheet_name)
                            original_sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Add the new workbook to the zip file
                    zf.writestr(f"Report_{file_name}", output_buffer.getvalue())

                except Exception as e:
                    st.error(f"An error occurred while processing '{file_name}': {e}")
                
                progress_bar.progress((i + 1) / len(uploaded_files))

        st.header("✅ Processing Complete!")
        st.download_button(
            label="📥 Download All Reports (.zip)",
            data=zip_buffer.getvalue(),
            file_name="Analysis_Reports.zip",
            mime="application/zip",
        )
else:
    st.info("Please upload one or more Excel files to begin.")
