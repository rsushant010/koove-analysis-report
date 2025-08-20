import streamlit as st
import pandas as pd
import io
import datetime
import numpy as np
import re
import zipfile

# --- Helper Functions for Data Processing ---

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
                analysis_df.loc[sno, 'Actual'] = 'Shutdown'
            analysis_df.loc[hrs_sno, 'Remark'] = 'Line was Shutdown for the day.'
    return analysis_df

def process_production_target(prod_df, analysis_df):
    """Processes the Gloves Production DataFrame to find and extract MTD targets."""
    if prod_df is None:
        st.warning("'Gloves Production' sheet not found for production target processing.")
        return analysis_df
        
    # Efficiently search for the cell containing "target mtd"
    mask = prod_df.applymap(lambda x: "target mtd" in str(x).lower())
    match = mask.stack()
    if not match.any():
        st.warning("'Target MTD' keyword not found in 'Gloves Production' sheet.")
        return analysis_df
        
    x, y = match[match].index[0]

    try:
        target_qty = prod_df.iloc[x, y + 1]
        actual_qty = prod_df.iloc[x + 1, y + 1]
        actual_percent = prod_df.iloc[x + 2, y + 1]
        
        analysis_df.loc[14, 'Standard'] = f"{target_qty:,.0f} Pcs"
        analysis_df.loc[14, 'Actual'] = f"{actual_qty:,.0f} Pcs"
        analysis_df.loc[14, 'Remark'] = f"Actual production is {actual_qty:,.0f} pcs."

        analysis_df.loc[15, 'Standard'] = "100 %"
        analysis_df.loc[15, 'Actual'] = f"{actual_percent:.1f} %"
        analysis_df.loc[15, 'Remark'] = f"Achievement is {actual_percent:.1f}%."
    except IndexError: 
        st.warning("Could not find production target values in the expected cells.")
        
    return analysis_df

def process_consumption(consump_df, target_date, analysis_df):
    """Processes the consumption DataFrame to find daily consumption."""
    if consump_df is None:
        st.warning("'coal & elec' sheet not found for consumption processing.")
        return analysis_df

    # Efficiently search for the date
    date_mask = consump_df.apply(pd.to_datetime, errors='coerce').applymap(lambda x: x.date() == target_date.date() if pd.notna(x) else False)
    date_match = date_mask.stack()
    if not date_match.any():
        st.warning(f"No consumption data found for date {target_date.strftime('%Y-%m-%d')}.")
        return analysis_df
        
    _, date_col_index = date_match[date_match].index[0]

    # Efficiently find the rows for Coal and Electricity
    first_col = consump_df.iloc[:, 0].str.lower()
    coal_row_index = first_col[first_col.str.contains("coal", na=False)].index
    elec_row_index = first_col[first_col.str.contains("electri", na=False)].index

    if not coal_row_index.empty:
        coal_val = consump_df.iloc[coal_row_index[0], date_col_index]
        analysis_df.loc[17, 'Actual'], analysis_df.loc[17, 'Remark'] = f"{coal_val:,.0f} Kg", f"Coal consumption is {coal_val:,.0f} Kg."
    if not elec_row_index.empty:
        elec_val = consump_df.iloc[elec_row_index[0], date_col_index]
        analysis_df.loc[18, 'Actual'], analysis_df.loc[18, 'Remark'] = f"{elec_val:,.0f} Unit", f"Electricity consumption is {elec_val:,.0f} unit."
        
    return analysis_df

def process_abnormalities(analysis_df):
    """Placeholder function to process abnormalities."""
    analysis_df.loc[13, 'Actual'], analysis_df.loc[13, 'Remark'] = "13 Nos", "13 nos. of Abnormalities are found."
    return analysis_df

def process_inventory(prod_df, analysis_df):
    """Processes the Gloves Production DataFrame to extract chemical inventory levels."""
    if prod_df is None:
        st.warning("'Gloves Production' sheet not found for inventory check.")
        return analysis_df

    mask = prod_df.applymap(lambda x: "xnbr latex" in str(x).lower())
    match = mask.stack()
    if not match.any():
        st.warning("'XNBR LATEX' keyword not found for inventory check.")
        return analysis_df
        
    start_row, start_col = match[match].index[0]

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

def process_order_details(order_df, analysis_df):
    """Processes the order details DataFrame to get financial summaries."""
    if order_df is None:
        st.warning("'Clear order details' sheet not found.")
        return analysis_df
        
    last_row = order_df.iloc[-1]
    
    # Flexible column name matching
    clear_order_col = next((col for col in order_df.columns if "total payment receive" in str(col).lower()), None)
    dispatch_col = next((col for col in order_df.columns if "total dispatch price" in str(col).lower()), None)
    pending_col = next((col for col in order_df.columns if "advance payment" in str(col).lower()), None)
            
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

# --- Main App Logic ---

def create_report(xls, target_date):
    """Main function to create a single analysis report DataFrame."""
    # Read all necessary sheets once
    oee_sheet_name = find_sheet_by_keyword(xls, "oee")
    oee_df = pd.read_excel(xls, sheet_name=oee_sheet_name) if oee_sheet_name else None
    
    prod_sheet_name = find_sheet_by_keyword(xls, "Gloves Production")
    prod_df = pd.read_excel(xls, sheet_name=prod_sheet_name, header=None) if prod_sheet_name else None
    
    consump_sheet_name = find_sheet_by_keyword(xls, "coal & elec")
    consump_df = pd.read_excel(xls, sheet_name=consump_sheet_name, header=None) if consump_sheet_name else None
    
    order_sheet_name = find_sheet_by_keyword(xls, "Clear order details")
    order_df = pd.read_excel(xls, sheet_name=order_sheet_name) if order_sheet_name else None

    if oee_df is None:
        st.error(f"Could not find 'OEE' sheet. Cannot proceed.")
        return None

    # Pre-process OEE dates
    oee_df['Date'] = pd.to_datetime(oee_df['Date'], errors='coerce')
    oee_df.dropna(subset=['Date'], inplace=True)
    oee_df.fillna(0, inplace=True)

    # Create the base report structure
    analysis_data = {'Sl. No.': [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,17,18], 'Particulars': ['LINE 1 Working Hrs.','LINE 1 Production Capacity','LINE 1 Quality','LINE 1 OEE','LINE 2 Working Hrs.','LINE 2 Production Capacity','LINE 2 Quality','LINE 2 OEE','LINE 3 Working Hrs.','LINE 3 Production Capacity','LINE 3 Quality','LINE 3 OEE','ABNORMALITIES (1+2+3)','PRODUCTION TARGET (Qty/day)','PRODUCTION TARGET (in %)','COAL CONSUMPTION PER DAY','ELECTRICITY CONSUMPTION PER DAY'], 'Unit': ['Hrs','Pcs/hrs','%','%','Hrs','Pcs/hrs','%','%','Hrs','Pcs/hrs','%','%','Nos','Pcs','%','Kg','Unit'], 'Standard': ['22 Hrs','18181 Pcs/hrs','100 %','80 %','22 Hrs','22727 Pcs/hrs','100 %','80 %','22 Hrs','22727 Pcs/hrs','100 %','80 %','0 Nos','','','',''], 'Actual': [''] * 17, 'Remark': [''] * 17}
    analysis_df = pd.DataFrame(analysis_data).set_index('Sl. No.')
    
    # Run all processing functions
    analysis_df = process_oee_data(oee_df, target_date, analysis_df)
    analysis_df = process_abnormalities(analysis_df)
    analysis_df = process_production_target(prod_df, analysis_df)
    analysis_df = process_consumption(consump_df, target_date, analysis_df)
    analysis_df = process_inventory(prod_df, analysis_df)
    analysis_df = process_order_details(order_df, analysis_df)

    return analysis_df.sort_index().reset_index().drop(columns=['Sl. No.'])

# --- Streamlit App UI ---

st.set_page_config(layout="wide")
st.title("📊 Multi-File Analysis Report Generator")

if 'report_data' not in st.session_state:
    st.session_state.report_data = []

st.sidebar.header("⚙️ Controls")
uploaded_files = st.sidebar.file_uploader("1. Upload Excel Workbooks", type="xlsx", accept_multiple_files=True)

file_dates = {}
if uploaded_files:
    st.sidebar.subheader("2. Set Analysis Dates")
    for uploaded_file in uploaded_files:
        file_dates[uploaded_file.name] = st.sidebar.date_input(f"Date for {uploaded_file.name}", datetime.date.today(), key=uploaded_file.name)

    if st.sidebar.button("🚀 Generate Reports", type="primary"):
        st.session_state.report_data = []
        progress_bar = st.progress(0, "Starting processing...")
        
        for i, uploaded_file in enumerate(uploaded_files):
            file_name = uploaded_file.name
            target_date = pd.to_datetime(file_dates[file_name])
            progress_bar.progress(i / len(uploaded_files), f"Processing '{file_name}'...")
            
            try:
                xls = pd.ExcelFile(uploaded_file)
                final_report_df = create_report(xls, target_date)
                
                if final_report_df is not None:
                    full_workbook_buffer = io.BytesIO()
                    with pd.ExcelWriter(full_workbook_buffer, engine='openpyxl') as writer:
                        final_report_df.to_excel(writer, sheet_name='Analysis Points', index=False)
                        for sheet_name in xls.sheet_names:
                            pd.read_excel(xls, sheet_name=sheet_name).to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    analysis_only_buffer = io.BytesIO()
                    with pd.ExcelWriter(analysis_only_buffer, engine='openpyxl') as writer:
                        final_report_df.to_excel(writer, sheet_name='Analysis Points', index=False)

                    st.session_state.report_data.append({
                        'file_name': file_name, 'target_date': target_date,
                        'full_workbook_buffer': full_workbook_buffer.getvalue(),
                        'analysis_only_buffer': analysis_only_buffer.getvalue()
                    })
            except Exception as e:
                st.error(f"An error occurred while processing '{file_name}': {e}")
        
        progress_bar.progress(1.0, "All files processed!")

if st.session_state.report_data:
    st.header("Generated Reports")
    
    for report in st.session_state.report_data:
        st.subheader(f"Report for '{report['file_name']}' on {report['target_date'].strftime('%Y-%m-%d')}")
        st.download_button(
            label=f"📥 Download Analysis Sheet (.xlsx)",
            data=report['analysis_only_buffer'],
            file_name=f"Analysis_Points_{report['file_name']}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_analysis_{report['file_name']}"
        )
        st.markdown("---")

    if len(st.session_state.report_data) > 1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for report in st.session_state.report_data:
                zf.writestr(f"Report_{report['file_name']}", report['full_workbook_buffer'])
        
        st.header("Combined Download")
        st.download_button(
            label="📦 Download All Combined Reports (.zip)",
            data=zip_buffer.getvalue(),
            file_name="All_Analysis_Reports.zip",
            mime="application/zip",
        )
elif not uploaded_files:
    st.info("Please upload one or more Excel files to begin.")
