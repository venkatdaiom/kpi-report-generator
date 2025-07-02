import pandas as pd
import numpy as np
import os
import streamlit as st
import gspread
import json
from google.oauth2.service_account import Credentials

# --- Page Configuration ---
st.set_page_config(
    page_title="KPI Report Generator",
    page_icon="📊",
    layout="wide"
)

# --- Google Sheets API Authentication ---
# This function connects to Google Sheets using Render's Environment Variables.
def authenticate_gspread():
    """Authenticates with Google Sheets API using secrets from Render's Environment Variables."""
    try:
        # Fetch the entire JSON credentials string from the environment variable
        creds_json_str = os.environ.get("GCP_SERVICE_ACCOUNT_JSON")
        
        if not creds_json_str:
            st.error("GCP_SERVICE_ACCOUNT_JSON environment variable not found on the server.")
            st.info("Please ensure the secret is configured correctly in the Render dashboard.")
            return None
        
        # Convert the JSON string back into a dictionary
        creds_dict = json.loads(creds_json_str)
        
        scopes = ['https://www.googleapis.com/auth/spreadsheets']
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Failed to authenticate with Google Sheets: {e}")
        return None

# --- Core KPI Calculation Function (Unchanged) ---
# This is the same powerful function from our previous discussions.
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

# --- Main Application UI and Logic ---
st.title("📊 KPI Report Generator")
st.markdown("Upload your lead master Excel file, and this tool will generate a consolidated KPI report and upload it to the designated Google Sheet.")

uploaded_file = st.file_uploader("Choose your Excel file", type=['xlsx'], help="Please upload the 'CX Callling KPIs Lead Master' file.")
TARGET_URL = "https://docs.google.com/spreadsheets/d/1BHqpUbEMlXasWrk6gRZOgEyGZiEnh7WK7z7YHgiR80k/edit"
st.text_input("Target Google Sheet (Output)", TARGET_URL, disabled=True)

if st.button("Generate and Upload Report", type="primary"):
    if uploaded_file is not None:
        with st.spinner("Processing your file... This may take a moment."):
            try:
                # 1. Data Preparation
                st.write("Step 1: Reading and preparing data...")
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
                st.success("Data prepared successfully.")
                
                # 2. Main Loop and Report Consolidation
                st.write("Step 2: Generating KPI reports for each segment...")
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
                    blank_df = pd.DataFrame([[''] * len(kpi_df.columns)], columns=kpi_df.columns)
                    all_reports_list.append(blank_df)
                    all_reports_list.append(blank_df.copy())
                if all_reports_list:
                    all_reports_list = all_reports_list[:-2]
                st.success("All segments processed.")

                # 3. Upload to Google Sheets
                st.write("Step 3: Uploading consolidated report to Google Sheets...")
                if all_reports_list:
                    final_report_df = pd.concat(all_reports_list, ignore_index=True).round(1)
                    gc = authenticate_gspread()
                    if gc:
                        spreadsheet = gc.open_by_url(TARGET_URL)
                        worksheet = spreadsheet.sheet1
                        worksheet.clear()
                        data_to_upload = [final_report_df.columns.values.tolist()] + final_report_df.astype(str).values.tolist()
                        worksheet.update('A1', data_to_upload, value_input_option='USER_ENTERED')
                        st.balloons()
                        st.success("🎉 Report successfully generated and uploaded to Google Sheets!")
                        st.markdown(f"**[Click here to view the report]({TARGET_URL})**")
                        st.dataframe(final_report_df)
                else:
                    st.warning("No data found for any source in the specified periods. No report was generated.")
            except Exception as e:
                st.error(f"An error occurred during processing: {e}")
    else:
        st.warning("Please upload an Excel file to proceed.")
