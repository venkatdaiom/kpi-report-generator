import pandas as pd
import numpy as np
import os
import streamlit as st
from io import BytesIO

# --- Page Configuration ---
st.set_page_config(
    page_title="KPI Report Generator",
    page_icon="📊",
    layout="wide"
)

# --- Core KPI Calculation Function (Unchanged) ---
# This powerful function is the engine of the report and does not need to be changed.
def calculate_kpis(sub_df):
    """Calculates a comprehensive list of KPIs for a given subset of leads."""
    if sub_df.empty:
        return {}
    kpis = {}
    num_leads_created = len(sub_df)
    qualified_leads_df = sub_df[sub_df['isQualified'] == 1]
    num_qualified_leads = len(qualified_leads_df)
    kpis['# of Leads Created (Salesforce)'] = num_leads_created
    kpis['# of Qualified leads'] = num_qualified_leads
    kpis['% Qualified Leads of Leads Created'] = (num_qualified_leads / num_leads_created * 100) if num_leads_created > 0 else 0
    attempted_qualified_leads_df = qualified_leads_df[qualified_leads_df['is_Lead_Called'] == 1]
    num_attempted_qualified = len(attempted_qualified_leads_df)
    kpis['# of Leads Attempted'] = num_attempted_qualified
    kpis['Attempt% of Qualified leads'] = (num_attempted_qualified / num_qualified_leads * 100) if num_qualified_leads > 0 else 0
    connected_qualified_leads_df = qualified_leads_df[qualified_leads_df['is_Lead_Connected'] == 1]
    num_connected_qualified = len(connected_qualified_leads_df)
    kpis['# of Leads Connected'] = num_connected_qualified
    kpis['Connection % of Qualified Leads'] = (num_connected_qualified / num_qualified_leads * 100) if num_qualified_leads > 0 else 0
    kpis['Time to First Attempt (P50) in hours'] = sub_df['TimeDiffLeadAttempt_hours'].dropna().quantile(0.50)
    kpis['Time to First Attempt (P90) in hours'] = sub_df['TimeDiffLeadAttempt_hours'].dropna().quantile(0.90)
    kpis['Time to First Connect (P50) in hours'] = sub_df['TimeDiffLeadConnect_hours'].dropna().quantile(0.50)
    kpis['Time to First Connect (P90) in hours'] = sub_df['TimeDiffLeadConnect_hours'].dropna().quantile(0.90)
    total_attempted_df = sub_df[sub_df['is_Lead_Called'] == 1]
    total_connected_df = sub_df[sub_df['is_Lead_Connected'] == 1]
    attempted_after_24h = total_attempted_df[total_attempted_df['TimeDiffLeadAttempt_hours'] > 24].shape[0]
    kpis['% Contri. of Leads Attempted after 24 hours'] = (attempted_after_24h / len(total_attempted_df) * 100) if not total_attempted_df.empty else 0
    connected_after_24h = total_connected_df[total_connected_df['TimeDiffLeadConnect_hours'] > 24].shape[0]
    kpis['% Contri. of Leads Contacted after 24 hours'] = (connected_after_24h / len(total_connected_df) * 100) if not total_connected_df.empty else 0
    return kpis

# --- Function to convert DataFrame to Excel in memory ---
def to_excel(df):
    """Converts a DataFrame to an Excel file in memory (bytes)."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='KPI_Report')
    processed_data = output.getvalue()
    return processed_data

# --- Main Application UI and Logic ---
st.title("📊 KPI Report Generator")
st.markdown("Upload your lead master Excel file to generate a downloadable, consolidated KPI report.")

# --- File Uploader ---
uploaded_file = st.file_uploader(
    "Choose your Excel file",
    type=['xlsx'],
    help="Please upload the 'CX Callling KPIs Lead Master' file."
)

if uploaded_file is not None:
    with st.spinner("Processing your file..."):
        try:
            # 1. Data Preparation
            df = pd.read_excel(uploaded_file)
            df = df.rename(columns={
                'is Lead Called?': 'is_Lead_Called',
                'is Lead Connected?': 'is_Lead_Connected',
                'TimeDiffLeadAttempt': 'TimeDiffLeadAttempt_hours',
                'TimeDiffLeadConnect': 'TimeDiffLeadConnect_hours'
            })
            df['LeadCreateDateTime_dt'] = pd.to_datetime(df['LeadCreateDateTime'], errors='coerce')
            df['LeadCreateMonth'] = df['LeadCreateMonth'].astype(str)
            df.dropna(subset=['LeadCreateDateTime_dt', 'LeadCreateMonth', 'Opportunity Source'], inplace=True)
            
            # 2. Main Loop and Report Consolidation
            kpi_group_mapping = {
                '# of Leads Created (Salesforce)': 'Lead Vol', '# of Qualified leads': 'Lead Vol', '% Qualified Leads of Leads Created': 'Lead Vol',
                '# of Leads Attempted': 'Lead Attempt', 'Attempt% of Qualified leads': 'Lead Attempt',
                'Time to First Attempt (P50) in hours': 'Lead Attempt', 'Time to First Attempt (P90) in hours': 'Lead Attempt',
                '% Contri. of Leads Attempted after 24 hours': 'Lead Attempt',
                '# of Leads Connected': 'Lead Connect', 'Connection % of Qualified Leads': 'Lead Connect',
                'Time to First Connect (P50) in hours': 'Lead Connect', 'Time to First Connect (P90) in hours': 'Lead Connect',
                '% Contri. of Leads Contacted after 24 hours': 'Lead Connect'
            }
            periods_to_analyze = {'April': '2025-04', 'May': '2025-05', 'June MTD (till 19th)': '2025-06'}
            sources_to_analyze = ['Overall'] + sorted([s for s in df['Opportunity Source'].unique() if pd.notna(s)])
            all_reports_list = []

            for source in sources_to_analyze:
                monthly_kpis_for_source = {}
                df_source = df[df['Opportunity Source'] == source].copy() if source != 'Overall' else df.copy()
                for period_name, month_string in periods_to_analyze.items():
                    if 'MTD' in period_name:
                        day_limit = int(period_name.split('till ')[1].replace('th','').replace('st','').replace('nd','').replace('rd','').replace(')',''))
                        df_period = df_source[(df_source['LeadCreateDateTime_dt'].dt.strftime('%Y-%m') == month_string) & (df_source['LeadCreateDateTime_dt'].dt.day <= day_limit)].copy()
                    else:
                        df_period = df_source[df_source['LeadCreateDateTime_dt'].dt.strftime('%Y-%m') == month_string].copy()
                    monthly_kpis_for_source[period_name] = calculate_kpis(df_period)
                filtered_results = {k: v for k, v in monthly_kpis_for_source.items() if v}
                if not filtered_results: continue
                
                kpi_df = pd.DataFrame(filtered_results).reset_index().rename(columns={'index': 'KPIs'})
                kpi_df.insert(0, 'Source', source)
                kpi_df.insert(1, 'Overall Leads', kpi_df['KPIs'].map(kpi_group_mapping))
                all_reports_list.append(kpi_df)
                
                # Add 2 blank rows after each segment's data
                blank_df = pd.DataFrame([[''] * len(kpi_df.columns)], columns=kpi_df.columns)
                all_reports_list.append(blank_df)
                all_reports_list.append(blank_df.copy())
            
            # 3. Prepare the final DataFrame and provide a download button
            if all_reports_list:
                # Remove the last two blank rows
                final_report_df = pd.concat(all_reports_list[:-2], ignore_index=True).round(1)
                
                st.success("🎉 Your report has been generated!")
                st.dataframe(final_report_df) # Show a preview of the report
                
                # Convert the final DataFrame to an Excel file in memory
                excel_data = to_excel(final_report_df)
                
                st.download_button(
                    label="📥 Download Report as Excel",
                    data=excel_data,
                    file_name="Consolidated_KPI_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No data found for any source in the specified periods. No report was generated.")

        except Exception as e:
            st.error(f"An error occurred during processing: {e}")
