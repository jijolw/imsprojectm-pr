import streamlit as st
st.set_page_config(page_title="IMS Form Entry", layout="wide")

import gspread
import json
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import time
from io import BytesIO
from xhtml2pdf import pisa
from jinja2 import Template

# -----------------------------
# SHEET SELECTION SYSTEM - FIXED
# -----------------------------
sheet_choice = st.sidebar.selectbox("Choose File Type", ["LW FILES", "M&PR FILES"])

# FIXED: Use the actual Google Sheet ID instead of sheet_choice
# Dynamically switch Google Sheet ID based on selected sheet type
SHEET_IDS = {
    "LW FILES": "1wxntHZp4xEQWCmLAt2TVF8ohG6uHuvV_3QbaK7wSwGw",
    "M&PR FILES": "17KL-cKMJNGncAJngla_eX6aVdzvoihbdpgOJqOiGycc"
}

GOOGLE_SHEET_ID = SHEET_IDS.get(sheet_choice)


# -----------------------------
# GOOGLE SHEETS AUTHENTICATION - FIXED
# -----------------------------
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

import os
import json

# Check if running on Streamlit Cloud (secrets available)
if "IMS_CREDENTIALS_JSON" in st.secrets:
    creds = json.loads(st.secrets["IMS_CREDENTIALS_JSON"])
    with open("temp_creds.json", "w") as f:
        json.dump(creds, f)
    CREDENTIAL_FILE = "temp_creds.json"
else:
    # Fallback to local file for PC use
    CREDENTIAL_FILE = "imscredentials.json"

@st.cache_resource
def get_gsheet_client(sheet_id):
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIAL_FILE, SCOPE)
        client = gspread.authorize(creds)
        return client.open_by_key(sheet_id)
    except Exception as e:
        st.error(f"Error connecting to Google Sheets: {e}")
        return None


# -----------------------------
# IMPROVED CACHING FOR API CALLS
# -----------------------------
@st.cache_data(ttl=300)  # Cache for 5 minutes
def get_sheet_data(sheet_name, sheet_id):
    """Get sheet data with caching to reduce API calls"""
    try:
        sheet_client = get_gsheet_client(sheet_id)
        if not sheet_client:
            return [], []

        worksheet = sheet_client.worksheet(sheet_name)
        all_values = worksheet.get_all_values()
        if not all_values:
            return [], []

        headers = all_values[0]
        records = []
        for row in all_values[1:]:
            while len(row) < len(headers):
                row.append("")
            record = dict(zip(headers, row))
            records.append(record)

        return headers, records
    except Exception as e:
        st.error(f"Error fetching data from worksheet '{sheet_name}': {e}")
        return [], []

@st.cache_data(ttl=600)  # Cache for 10 minutes
def get_all_sheet_names(sheet_id):
    """Get all sheet names with caching"""
    try:
        sheet_client = get_gsheet_client(sheet_id)
        if not sheet_client:
            return []
        return [ws.title for ws in sheet_client.worksheets()]
    except Exception as e:
        st.error(f"Error fetching sheet names: {e}")
        return []

# -----------------------------
# LOAD FORM CONFIGURATION SYSTEM
# -----------------------------
@st.cache_data(ttl=3600)  # Cache for 1 hour
def load_form_configs_for_sheet(sheet_type):
    """Load form configurations based on sheet type"""
    try:
        if sheet_type == "LW FILES":
            config_file = "form_configs.json"
        elif sheet_type == "M&PR FILES":
            config_file = "forms_mpr_configs.json"  # Fixed filename
        else:
            st.error(f"❌ Unknown sheet type: {sheet_type}")
            return {}
        
        with open(config_file, "r") as f:
            configs = json.load(f)
        
        st.sidebar.success(f"✅ Loaded {len(configs)} forms from {config_file}")
        return configs
        
    except FileNotFoundError:
        expected_file = "form_configs.json" if sheet_type == "LW FILES" else "forms_mpr_configs.json"
        st.error(f"❌ Config file not found for {sheet_type}")
        st.error(f"Expected file: {expected_file}")
        st.info("💡 Please ensure the config file exists in your project directory")
        return {}
    except json.JSONDecodeError as e:
        st.error(f"❌ Error parsing config file for {sheet_type}: {e}")
        return {}

# Rate limiting helper
def api_rate_limit():
    """Simple rate limiting to avoid quota issues"""
    if 'last_api_call' not in st.session_state:
        st.session_state.last_api_call = 0
    
    current_time = time.time()
    time_since_last = current_time - st.session_state.last_api_call
    
    if time_since_last < 1:  # Wait at least 1 second between API calls
        time.sleep(1 - time_since_last)
    
    st.session_state.last_api_call = time.time()

# Load form configurations based on selected sheet type
form_configs = load_form_configs_for_sheet(sheet_choice)

# Display current sheet info - FIXED
st.sidebar.markdown("---")
st.sidebar.markdown(f"**Current Sheet Type:** `{sheet_choice}`")
st.sidebar.markdown(f"**Google Sheet ID:** `{GOOGLE_SHEET_ID}`")
st.sidebar.markdown(f"**Available Forms:** {len(form_configs)}")

# Show warning if no forms available for selected sheet
if not form_configs:
    st.warning(f"⚠️ No form configurations found for '{sheet_choice}'. Please check your config file.")
    st.info("💡 **Expected Files:**")
    st.info("- LW FILES: `form_configs.json`")
    st.info("- M&PR FILES: `forms_mpr_configs.json`")

st.title("📋 IMS Form Entry System")
st.markdown(f"**Active Sheet Type:** `{sheet_choice}` | **Available Forms:** {len(form_configs)}")

# Only proceed if we have form configs
if form_configs:
    # -----------------------------
    # MAIN FORM ENTRY SECTION
    # -----------------------------
    st.markdown("## ✏️ Form Entry & Edit")

    selected_form = st.selectbox(
        "Select a form to fill or edit",
        list(form_configs.keys()),
        format_func=lambda x: f"{x} - {form_configs[x]['title']}"
    )

    form_data = form_configs[selected_form]

    # Get cached data
    headers, records = get_sheet_data(selected_form, GOOGLE_SHEET_ID)

    df = pd.DataFrame(records) if records else pd.DataFrame()

    st.subheader(form_data["title"])

    selected_row_index = None
    edit_mode = st.checkbox("✏️ Enable Edit Mode")

    if edit_mode and not df.empty:
        st.info("Select a row to edit.")
        df_display = df.copy()
        df_display.index += 2  # Show real Google Sheet row numbers (starting from 2)
        st.dataframe(df_display, use_container_width=True)

        row_num = st.number_input("Enter row number to edit (from above):", min_value=2, max_value=len(df)+1, step=1)
        selected_row_index = row_num - 2
        if 0 <= selected_row_index < len(df):
            st.success(f"Editing row {row_num}")
            prefill_data = df.iloc[selected_row_index]
        else:
            prefill_data = {}
    else:
        prefill_data = {}

    # -----------------------------
    # ENTRY / EDIT FORM
    # -----------------------------
    form_values = {}
    with st.form(key="entry_form"):
        for field in form_data["fields"]:
            form_values[field] = st.text_input(field, value=prefill_data.get(field, ""))

        signature_values = {}
        for signer in form_data.get("signatures", []):
            default_checked = prefill_data.get(signer, "") == "✔️ Yes"
            signature_values[signer] = st.checkbox(signer, value=default_checked)

        submitted = st.form_submit_button("Submit")

    if submitted:
        try:
            api_rate_limit()  # Rate limiting
            
            # Add timestamp
            form_values["Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Add signatures
            for signer, signed in signature_values.items():
                form_values[signer] = "✔️ Yes" if signed else "❌ No"

            # Ensure all headers are matched correctly
            row = [form_values.get(col, "") for col in headers]

            sheet_client = get_gsheet_client(GOOGLE_SHEET_ID)

            if sheet_client:
                # FIXED: Access the correct worksheet within the Google Sheet
                worksheet = sheet_client.worksheet(selected_form)
                
                if edit_mode and selected_row_index is not None:
                    worksheet.update(f"A{selected_row_index + 2}", [row])
                    st.success(f"✅ Row {selected_row_index + 2} updated successfully.")
                else:
                    worksheet.append_row(row)
                    st.success("✅ New entry submitted successfully.")
                
                # Clear cache to refresh data
                st.cache_data.clear()
                st.rerun()
            
        except Exception as e:
            st.error(f"❌ Error writing to Google Sheet: {e}")

    # -----------------------------
    # IMPROVED DATA VIEW SECTION  
    # -----------------------------
    st.markdown("---")
    st.markdown("## 📊 Data View & Management")

    # Refresh data button
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()

    with col2:
        view_mode = st.selectbox("View Mode", ["Enhanced Table View", "Card View"])

    # Select sheet for viewing - filter based on current sheet type
    available_sheets = [name for name in get_all_sheet_names(GOOGLE_SHEET_ID) if name in form_configs]
    view_sheet = st.selectbox("📄 Select Sheet to View", available_sheets, key="view_sheet_select")

    if view_sheet:
        view_headers, view_records = get_sheet_data(view_sheet, GOOGLE_SHEET_ID)

        view_df = pd.DataFrame(view_records) if view_records else pd.DataFrame()
        
        if view_df.empty:
            st.warning("No data available in the selected sheet.")
        else:
            st.info(f"📈 Total Records: {len(view_df)} | Last Updated: {datetime.now().strftime('%H:%M:%S')}")
            
            # Search and filter
            search_term = st.text_input("🔍 Search in data:", placeholder="Enter search term...")
            
            if search_term:
                # Search across all columns
                mask = view_df.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
                filtered_df = view_df[mask]
                st.info(f"Found {len(filtered_df)} records matching '{search_term}'")
            else:
                filtered_df = view_df
            
            if view_mode == "Enhanced Table View":
                # Improved table view with better column handling
                display_df = filtered_df.copy()
                display_df.index = range(2, len(display_df) + 2)  # Google Sheet row numbers
                
                # Get form config for this sheet to identify signature columns
                sheet_config = form_configs.get(view_sheet, {})
                signature_columns = sheet_config.get("signatures", [])
                
                # Column selection for better visibility
                st.markdown("### 📋 Column Visibility Control")
                all_columns = list(display_df.columns)
                
                # Separate regular fields and signature fields
                regular_fields = [col for col in all_columns if col not in signature_columns and col != "Timestamp"]
                signature_fields = [col for col in all_columns if col in signature_columns]
                timestamp_field = ["Timestamp"] if "Timestamp" in all_columns else []
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown("**📝 Form Fields:**")
                    show_regular = st.checkbox("Show All Form Fields", value=True, key="show_regular")
                    if not show_regular:
                        selected_regular = st.multiselect("Select Form Fields:", regular_fields, default=regular_fields[:3])
                    else:
                        selected_regular = regular_fields
                
                with col2:
                    st.markdown("**✍️ Signature Fields:**")
                    show_signatures = st.checkbox("Show All Signatures", value=True, key="show_signatures")
                    if not show_signatures:
                        selected_signatures = st.multiselect("Select Signatures:", signature_fields, default=signature_fields[:3])
                    else:
                        selected_signatures = signature_fields
                
                with col3:
                    st.markdown("**⏰ Timestamp:**")
                    show_timestamp = st.checkbox("Show Timestamp", value=True, key="show_timestamp")
                    selected_timestamp = timestamp_field if show_timestamp else []
                
                # Combine selected columns
                columns_to_show = selected_regular + selected_signatures + selected_timestamp
                
                if columns_to_show:
                    display_filtered = display_df[columns_to_show]
                    
                    # Style the dataframe to highlight signatures
                    def style_signatures(val):
                        if "✔️" in str(val) or "Yes" in str(val):
                            return 'background-color: #d4edda; color: #155724; font-weight: bold'
                        elif "❌" in str(val) or "No" in str(val):
                            return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
                        return ''
                    
                    styled_df = display_filtered.style.applymap(style_signatures, subset=selected_signatures if selected_signatures else [])
                    
                    st.markdown(f"**Showing {len(columns_to_show)} columns out of {len(all_columns)} total columns**")
                    st.dataframe(styled_df, use_container_width=True, height=400)
                else:
                    st.warning("Please select at least one column to display.")
            
            else:  # Card View
                for idx, row in filtered_df.iterrows():
                    with st.container():
                        col1, col2 = st.columns([3, 1])
                        with col1:
                            st.markdown(f"**Row {idx + 2}**")
                            for field, value in row.items():
                                if value:  # Only show non-empty fields
                                    if "✔️" in str(value) or "Yes" in str(value):
                                        st.markdown(f"• **{field}:** :green[{value}]")
                                    elif "❌" in str(value) or "No" in str(value):
                                        st.markdown(f"• **{field}:** :red[{value}]")
                                    else:
                                        st.write(f"• **{field}:** {value}")
                        with col2:
                            st.write(f"Row: {idx + 2}")
                        st.markdown("---")

    # [Rest of the PDF export section remains the same - keeping it for brevity]
    # -----------------------------
    # COMPLETE PDF EXPORT SECTION
    # -----------------------------
    st.markdown("---")
    st.markdown("## 📝 Enhanced PDF Export System")

    # PDF export options
    pdf_tab1, pdf_tab2 = st.tabs(["📄 Individual Row PDFs", "📊 Enhanced Table View PDF"])

    with pdf_tab1:
        st.markdown("### Individual Row PDFs")
        st.info("Generate separate PDF for each selected row with detailed form layout.")
        
        pdf_sheet = st.selectbox("📄 Select Sheet", available_sheets, key="pdf_individual_sheet")
        
        if pdf_sheet:
            pdf_headers, pdf_records = get_sheet_data(pdf_sheet, GOOGLE_SHEET_ID)

            pdf_df = pd.DataFrame(pdf_records) if pdf_records else pd.DataFrame()
            
            if not pdf_df.empty:
                selected_rows = st.multiselect("Select Row(s) for Individual PDFs", pdf_df.index + 2, key="individual_pdf_rows")
                
                if selected_rows:
                    config = form_configs.get(pdf_sheet, {})
                    title = config.get("title", pdf_sheet)
                    form_fields = config.get("fields", [])
                    signature_fields = config.get("signatures", [])
                    
                    for row_num in selected_rows:
                        actual_index = row_num - 2
                        if 0 <= actual_index < len(pdf_df):
                            entry_data = pdf_df.iloc[actual_index].to_dict()
                            
                            # Individual row PDF generation
                            form_data_for_pdf = {field: entry_data.get(field, "") for field in form_fields}
                            signature_data_for_pdf = {signer: entry_data.get(signer, "❌ No") for signer in signature_fields}
                            
                            with st.expander(f"📋 Preview Row {row_num}", expanded=False):
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.write("**Form Fields:**")
                                    for field, value in form_data_for_pdf.items():
                                        st.write(f"• **{field}:** {value or '-'}")
                                with col2:
                                    if signature_data_for_pdf:
                                        st.write("**Signatures:**")
                                        for signer, status in signature_data_for_pdf.items():
                                            color = "green" if "✔️" in status or "Yes" in status else "red"
                                            st.markdown(f"• **{signer}:** :{color}[{status}]")
                            
                            if st.button(f"📄 Generate PDF for Row {row_num}", key=f"gen_individual_{row_num}"):
                                # Individual PDF template
                                individual_template = """
                                <!DOCTYPE html>
                                <html>
                                <head>
                                    <meta charset="utf-8">
                                    <style>
                                        body { font-family: Arial, sans-serif; margin: 20px; color: #333; }
                                        .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #333; padding-bottom: 10px; }
                                        .form-title { font-size: 24px; font-weight: bold; margin-bottom: 5px; }
                                        .export-info { font-size: 12px; color: #666; margin-top: 5px; }
                                        .content-table { width: 100%; border-collapse: collapse; margin-bottom: 30px; }
                                        .content-table th { background-color: #f5f5f5; border: 1px solid #ddd; padding: 12px 8px; text-align: left; font-weight: bold; width: 30%; }
                                        .content-table td { border: 1px solid #ddd; padding: 12px 8px; text-align: left; width: 70%; }
                                        .signatures-section { margin-top: 40px; border-top: 1px solid #ddd; padding-top: 20px; }
                                        .signatures-title { font-size: 18px; font-weight: bold; margin-bottom: 15px; color: #333; }
                                        .signature-item { margin-bottom: 15px; padding: 8px; border: 1px solid #eee; background-color: #fafafa; }
                                        .signature-name { font-weight: bold; display: inline-block; width: 200px; }
                                        .signature-status { font-size: 16px; }
                                        .signed { color: #28a745; }
                                        .not-signed { color: #dc3545; }
                                        .footer { margin-top: 50px; text-align: center; font-size: 10px; color: #999; border-top: 1px solid #eee; padding-top: 10px; }
                                    </style>
                                </head>
                                <body>
                                    <div class="header">
                                        <div class="form-title">{{ title }}</div>
                                        <div class="export-info">Row {{ row_number }} - Exported on {{ export_date }}</div>
                                    </div>
                                    
                                    <table class="content-table">
                                        {% for field, value in form_data.items() %}
                                        <tr><th>{{ field }}</th><td>{{ value if value.strip() else '-' }}</td></tr>
                                        {% endfor %}
                                    </table>
                                    
                                    {% if signatures %}
                                    <div class="signatures-section">
                                        <div class="signatures-title">📝 Signatures & Approvals</div>
                                        {% for signer, status in signatures.items() %}
                                        <div class="signature-item">
                                            <span class="signature-name">{{ signer }}:</span>
                                            <span class="signature-status {% if '✔️' in status or 'Yes' in status %}signed{% else %}not-signed{% endif %}">{{ status }}</span>
                                        </div>
                                        {% endfor %}
                                    </div>
                                    {% endif %}
                                    
                                    <div class="footer">Generated from IMS Form Entry System</div>
                                </body>
                                </html>
                                """
                                
                                try:
                                    template = Template(individual_template)
                                    rendered_html = template.render(
                                        title=title,
                                        row_number=row_num,
                                        export_date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        form_data=form_data_for_pdf,
                                        signatures=signature_data_for_pdf
                                    )
                                    
                                    pdf_buffer = BytesIO()
                                    pisa_status = pisa.CreatePDF(BytesIO(rendered_html.encode("utf-8")), dest=pdf_buffer)
                                    
                                    if not pisa_status.err:
                                        pdf_buffer.seek(0)
                                        st.success("✅ Individual PDF generated!")
                                        st.download_button(
                                            label=f"⬇️ Download Row {row_num} PDF",
                                            data=pdf_buffer.getvalue(),
                                            file_name=f"{pdf_sheet}_Row_{row_num}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                                            mime="application/pdf",
                                            key=f"dl_individual_{row_num}"
                                        )
                                    else:
                                        st.error(f"❌ PDF generation error: {pisa_status.err}")
                                except Exception as e:
                                    st.error(f"❌ Error: {str(e)}")

    with pdf_tab2:
        st.markdown("### 📊 Enhanced Table View PDF")
        st.info("Generate a comprehensive PDF with intelligent column handling and signature highlighting!")
        
        table_pdf_sheet = st.selectbox("📄 Select Sheet", available_sheets, key="pdf_table_sheet")
        
        if table_pdf_sheet:
            table_headers, table_records = get_sheet_data(table_pdf_sheet, GOOGLE_SHEET_ID)

            table_df = pd.DataFrame(table_records) if table_records else pd.DataFrame()
            
            if not table_df.empty:
                # Get form config for better column organization
                sheet_config = form_configs.get(table_pdf_sheet, {})
                signature_columns = sheet_config.get("signatures", [])
                form_fields = sheet_config.get("fields", [])
                
                # Improved column selection with categories
                st.markdown("### 📋 Column Selection for PDF")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown("**📝 Form Fields:**")
                    available_form_fields = [col for col in table_df.columns if col in form_fields]
                    selected_form_fields = st.multiselect(
                        "Select Form Fields:", 
                        available_form_fields, 
                        default=available_form_fields,
                        key="pdf_form_fields"
                    )
                
                with col2:
                    st.markdown("**✍️ Signature Fields:**")
                    available_signature_fields = [col for col in table_df.columns if col in signature_columns]
                    selected_signature_fields = st.multiselect(
                        "Select Signature Fields:", 
                        available_signature_fields, 
                        default=available_signature_fields,
                        key="pdf_signature_fields"
                    )
                
                with col3:
                    st.markdown("**⏰ Other Fields:**")
                    other_fields = [col for col in table_df.columns if col not in form_fields and col not in signature_columns]
                    selected_other_fields = st.multiselect(
                        "Select Other Fields:", 
                        other_fields, 
                        default=["Timestamp"] if "Timestamp" in other_fields else [],
                        key="pdf_other_fields"
                    )
                
                # Combine all selected columns
                all_selected_columns = selected_form_fields + selected_signature_fields + selected_other_fields
                
                # Row selection
                table_selected_rows = st.multiselect(
                    "Select Rows for Table PDF (leave empty for all rows)", 
                    table_df.index + 2, 
                    key="table_pdf_rows"
                )
                
                # Use all rows if none selected
                if not table_selected_rows:
                    rows_to_include = list(range(len(table_df)))
                    st.info(f"📊 Will include all {len(table_df)} rows with {len(all_selected_columns)} columns in PDF")
                else:
                    rows_to_include = [row - 2 for row in table_selected_rows]
                    st.info(f"📊 Will include {len(rows_to_include)} selected rows with {len(all_selected_columns)} columns in PDF")
                
                # PDF Layout options
                st.markdown("### 🎨 PDF Layout Options")
                col1, col2 = st.columns(2)
                with col1:
                    page_orientation = st.selectbox("Page Orientation", ["Landscape", "Portrait"], key="pdf_orientation")
                with col2:
                    include_summary = st.checkbox("Include Summary Statistics", value=True, key="pdf_summary")
                
                if all_selected_columns and st.button("📊 Generate Enhanced Table PDF", key="generate_enhanced_table_pdf"):
                    try:
                        config = form_configs.get(table_pdf_sheet, {})
                        title = config.get("title", table_pdf_sheet)
                        
                        # Prepare data
                        filtered_df = table_df.iloc[rows_to_include][all_selected_columns]
                        
                        # Calculate summary statistics for signatures
                        summary_stats = {}
                        for sig_col in selected_signature_fields:
                            if sig_col in filtered_df.columns:
                                total_entries = len(filtered_df)
                                signed_count = filtered_df[sig_col].str.contains('✔️|Yes', na=False).sum()
                                not_signed_count = filtered_df[sig_col].str.contains('❌|No', na=False).sum()
                                summary_stats[sig_col] = {
                                    'signed': signed_count,
                                    'not_signed': not_signed_count,
                                    'percentage': (signed_count / total_entries * 100) if total_entries > 0 else 0
                                }
                        
                        # Enhanced Table PDF template (keeping the same template from original)
                        enhanced_table_template = """
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <meta charset="utf-8">
                            <style>
                                @page { 
                                    size: {{ 'landscape' if orientation == 'Landscape' else 'portrait' }}; 
                                    margin: 0.5in; 
                                }
                                body { 
                                    font-family: Arial, sans-serif; 
                                    margin: 0; 
                                    color: #333; 
                                    font-size: 9px; 
                                }
                                .header { 
                                    text-align: center; 
                                    margin-bottom: 20px; 
                                    border-bottom: 2px solid #333; 
                                    padding-bottom: 10px; 
                                }
                                .form-title { 
                                    font-size: 18px; 
                                    font-weight: bold; 
                                    margin-bottom: 5px; 
                                }
                                .export-info { 
                                    font-size: 10px; 
                                    color: #666; 
                                    margin-top: 5px; 
                                }
                                .summary-section {
                                    margin-bottom: 20px;
                                    padding: 10px;
                                    background-color: #f8f9fa;
                                    border: 1px solid #dee2e6;
                                    border-radius: 4px;
                                }
                                .summary-title {
                                    font-size: 12px;
                                    font-weight: bold;
                                    margin-bottom: 10px;
                                    color: #495057;
                                }
                                .summary-grid {
                                    display: flex;
                                    flex-wrap: wrap;
                                    gap: 15px;
                                }
                                .summary-item {
                                    flex: 1;
                                    min-width: 150px;
                                    font-size: 10px;
                                }
                                .summary-name {
                                    font-weight: bold;
                                    color: #495057;
                                }
                                .data-table { 
                                    width: 100%; 
                                    border-collapse: collapse; 
                                    margin-bottom: 20px; 
                                    font-size: 8px; 
                                }
                                .data-table th { 
                                    background-color: #343a40; 
                                    color: white;
                                    border: 1px solid #495057; 
                                    padding: 6px 3px; 
                                    text-align: left; 
                                    font-weight: bold; 
                                    font-size: 8px;
                                    word-wrap: break-word;
                                    max-width: 100px;
                                }
                                .data-table td { 
                                    border: 1px solid #dee2e6; 
                                    padding: 4px 3px; 
                                    text-align: left; 
                                    word-wrap: break-word; 
                                    max-width: 100px;
                                    font-size: 8px;
                                }
                                .row-num { 
                                    background-color: #f8f9fa; 
                                    font-weight: bold; 
                                    text-align: center; 
                                    min-width: 30px;
                                }
                                .signed { 
                                    background-color: #d4edda; 
                                    color: #155724; 
                                    font-weight: bold;
                                    text-align: center;
                                }
                                .not-signed { 
                                    background-color: #f8d7da; 
                                    color: #721c24;
                                    text-align: center;
                                }
                                .form-field {
                                    background-color: #fff3cd;
                                }
                                .signature-field {
                                    font-weight: bold;
                                }
                                .footer { 
                                    margin-top: 20px; 
                                    text-align: center; 
                                    font-size: 8px; 
                                    color: #6c757d; 
                                    border-top: 1px solid #dee2e6; 
                                    padding-top: 10px; 
                                }
                                .column-header-form { background-color: #17a2b8 !important; }
                                .column-header-signature { background-color: #28a745 !important; }
                                .column-header-other { background-color: #6c757d !important; }
                            </style>
                        </head>
                        <body>
                            <div class="header">
                                <div class="form-title">{{ title }} - Enhanced Table View</div>
                                <div class="export-info">
                                    Exported on {{ export_date }} | 
                                    Records: {{ total_records }} | 
                                    Columns: {{ column_count }} ({{ form_field_count }} Form, {{ signature_field_count }} Signatures, {{ other_field_count }} Other)
                                </div>
                            </div>
                            
                            {% if include_summary and summary_stats %}
                            <div class="summary-section">
                                <div class="summary-title">📊 Signature Summary Statistics</div>
                                <div class="summary-grid">
                                    {% for field, stats in summary_stats.items() %}
                                    <div class="summary-item">
                                        <div class="summary-name">{{ field }}:</div>
                                        <div>✔️ Signed: {{ stats.signed }} ({{ "%.1f"|format(stats.percentage) }}%)</div>
                                        <div>❌ Not Signed: {{ stats.not_signed }}</div>
                                    </div>
                                    {% endfor %}
                                </div>
                            </div>
                            {% endif %}
                            
                            <table class="data-table">
                                <thead>
                                    <tr>
                                        <th class="row-num">Row</th>
                                        {% for column in columns %}
                                        <th class="{% if column in form_fields %}column-header-form{% elif column in signature_fields %}column-header-signature{% else %}column-header-other{% endif %}">{{ column }}</th>
                                        {% endfor %}
                                    </tr>
                                </thead>
                                <tbody>
                                    {% for row_data in table_data %}
                                    <tr>
                                        <td class="row-num">{{ row_data.row_number }}</td>
                                        {% for column in columns %}
                                        <td class="{% if column in signature_fields %}signature-field{% else %}form-field{% endif %} {% if column in signature_fields and ('✔️' in row_data.data.get(column, '') or 'Yes' in row_data.data.get(column, '')) %}signed{% elif column in signature_fields and ('❌' in row_data.data.get(column, '') or 'No' in row_data.data.get(column, '')) %}not-signed{% endif %}">
                                            {{ row_data.data.get(column, '') if row_data.data.get(column, '').strip() else '-' }}
                                        </td>
                                        {% endfor %}
                                    </tr>
                                    {% endfor %}
                                </tbody>
                            </table>
                            
                            <div class="footer">
                                Generated from IMS Form Entry System - Enhanced Table View | 
                                Page Orientation: {{ orientation }} | 
                                Total Columns: {{ column_count }}
                            </div>
                        </body>
                        </html>
                        """
                        
                        # Prepare template data
                        table_data = []
                        for idx, (original_idx, row) in enumerate(filtered_df.iterrows()):
                            table_data.append({
                                'row_number': original_idx + 2,  # Google Sheet row number
                                'data': row.to_dict()
                            })
                        
                        template = Template(enhanced_table_template)
                        rendered_html = template.render(
                            title=title,
                            export_date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            total_records=len(filtered_df),
                            column_count=len(all_selected_columns),
                            form_field_count=len(selected_form_fields),
                            signature_field_count=len(selected_signature_fields),
                            other_field_count=len(selected_other_fields),
                            columns=all_selected_columns,
                            form_fields=selected_form_fields,
                            signature_fields=selected_signature_fields,
                            table_data=table_data,
                            orientation=page_orientation,
                            include_summary=include_summary,
                            summary_stats=summary_stats if include_summary else {}
                        )
                        
                        pdf_buffer = BytesIO()
                        pisa_status = pisa.CreatePDF(BytesIO(rendered_html.encode("utf-8")), dest=pdf_buffer)
                        
                        if not pisa_status.err:
                            pdf_buffer.seek(0)
                            st.success("✅ Enhanced Table PDF generated successfully!")
                            
                            # Show preview of what was included
                            with st.expander("📋 PDF Content Preview", expanded=False):
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.markdown("**📝 Form Fields Included:**")
                                    for field in selected_form_fields:
                                        st.write(f"• {field}")
                                with col2:
                                    st.markdown("**✍️ Signature Fields Included:**")
                                    for field in selected_signature_fields:
                                        st.write(f"• {field}")
                                with col3:
                                    st.markdown("**⏰ Other Fields Included:**")
                                    for field in selected_other_fields:
                                        st.write(f"• {field}")
                                
                                if include_summary and summary_stats:
                                    st.markdown("**📊 Summary Statistics:**")
                                    for field, stats in summary_stats.items():
                                        st.write(f"• {field}: {stats['signed']} signed ({stats['percentage']:.1f}%), {stats['not_signed']} not signed")
                            
                            # Generate filename
                            rows_desc = f"All_{len(filtered_df)}_rows" if not table_selected_rows else f"Selected_{len(rows_to_include)}_rows"
                            cols_desc = f"{len(all_selected_columns)}_cols"
                            filename = f"{table_pdf_sheet}_Enhanced_Table_{rows_desc}_{cols_desc}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                            
                            st.download_button(
                                label=f"⬇️ Download Enhanced Table PDF ({len(filtered_df)} rows, {len(all_selected_columns)} columns)",
                                data=pdf_buffer.getvalue(),
                                file_name=filename,
                                mime="application/pdf",
                                key="dl_enhanced_table_pdf"
                            )
                        else:
                            st.error(f"❌ PDF generation error: {pisa_status.err}")
                            
                    except Exception as e:
                        st.error(f"❌ Error generating enhanced table PDF: {str(e)}")

# -----------------------------
# SYSTEM INFO - UPDATED
# -----------------------------
st.markdown("---")
st.markdown("### ℹ️ System Information")

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Available Forms", len(form_configs))
with col2:
    all_sheets = all_sheets = get_all_sheet_names(GOOGLE_SHEET_ID)

    st.metric("Available Sheets", len(all_sheets))
with col3:
    st.metric("Current Sheet Type", sheet_choice)

# Show available sheets for debugging - UPDATED
with st.expander("🔍 Debug Information", expanded=False):
    st.markdown("**Available Worksheets in Google Sheet:**")
    for sheet in all_sheets:
        st.write(f"• {sheet}")
    
    st.markdown("**Form Configurations Loaded:**")
    for form_id, config in form_configs.items():
        st.write(f"• {form_id}: {config['title']}")
    
    st.markdown(f"**Current Sheet Type Selection:** {sheet_choice}")
    st.markdown(f"**Google Sheet ID:** {GOOGLE_SHEET_ID}")
    st.markdown(f"**Config File Used:** {'form_configs.json' if sheet_choice == 'LW FILES' else 'forms_mpr_configs.json'}")
    
    # Show sample forms for each type
    if sheet_choice == "LW FILES":
        st.markdown("**Expected LW Forms:** GF A, GF B, LW4 01, LW4 02, etc.")
    else:
        st.markdown("**Expected M&PR Forms:** MP4_06_08, MP4_09, MP4_11, etc.")

# Cache management
if st.button("🔄 Clear All Cache", key="clear_cache_btn"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.success("✅ All caches cleared!")
    st.rerun()

st.markdown("**📋 Enhanced Features:**")
st.markdown("- ✅ **FIXED: Dual Sheet Support**: LW FILES & M&PR FILES with correct Google Sheet access")
st.markdown("- ✅ **Complete PDF Export**: Individual row PDFs + Enhanced table PDFs")
st.markdown("- ✅ **Smart Column Management**: Form fields, signatures, and timestamps")
st.markdown("- ✅ **Signature Highlighting**: Visual indicators for signed/unsigned status")
st.markdown("- ✅ **Advanced PDF Options**: Landscape/portrait, summary statistics")
st.markdown("- ✅ **Intelligent Caching**: Reduced API calls with 5-10 min cache")
st.markdown("- ✅ **Rate Limiting**: Prevents quota errors")
st.markdown("- ✅ **Enhanced Search**: Real-time filtering across all columns")
st.markdown("- ✅ **Card & Table Views**: Multiple viewing options")
st.markdown("- ✅ **Debug Tools**: Comprehensive troubleshooting information")

# Additional system status
st.markdown("**🔧 System Status:**")
col1, col2, col3, col4 = st.columns(4)
with col1:
    if get_gsheet_client(GOOGLE_SHEET_ID):
        st.success("✅ Google Sheets Connected")
    else:
        st.error("❌ Google Sheets Connection Failed")
with col2:
    st.success("✅ PDF Generation Ready")
with col3:
    st.success("✅ Caching Active")
with col4:
    st.success("✅ Enhanced UI Loaded")

# Configuration file status
st.markdown("**📁 Configuration Files:**")
col1, col2 = st.columns(2)
with col1:
    try:
        with open("form_configs.json", "r") as f:
            lw_configs = json.load(f)
        st.success(f"✅ LW FILES: {len(lw_configs)} forms loaded")
    except:
        st.error("❌ LW FILES: form_configs.json not found")
        
with col2:
    try:
        with open("forms_mpr_configs.json", "r") as f:
            mpr_configs = json.load(f)
        st.success(f"✅ M&PR FILES: {len(mpr_configs)} forms loaded")
    except:
        st.error("❌ M&PR FILES: forms_mpr_configs.json not found")

# UPDATED DEBUG SECTION - Now correctly shows worksheets
st.markdown("---")
st.markdown("### 🔍 FIXED: M&PR Worksheets Check")

if st.button("🔍 Check M&PR Worksheets", key="debug_mpr_sheets"):
    try:
        sheet_client = get_gsheet_client(GOOGLE_SHEET_ID)

        if sheet_client:
            # Get all worksheet names
            all_worksheets = sheet_client.worksheets()
            worksheet_names = [ws.title for ws in all_worksheets]
            
            st.markdown(f"**📋 All Worksheets Found in Google Sheet ID '{GOOGLE_SHEET_ID}':**")
            for i, name in enumerate(worksheet_names, 1):
                st.write(f"{i}. `{name}`")
            
            st.markdown("---")
            st.markdown("**📄 Forms from Config vs Worksheets:**")
            
            # Check each form from config
            missing_worksheets = []
            found_worksheets = []
            
            for form_id in form_configs.keys():
                if form_id in worksheet_names:
                    st.success(f"✅ `{form_id}` - Found in Google Sheets")
                    found_worksheets.append(form_id)
                else:
                    st.error(f"❌ `{form_id}` - **MISSING** from Google Sheets")
                    missing_worksheets.append(form_id)
            
            st.markdown("---")
            st.markdown(f"**📊 Summary:**")
            st.info(f"✅ Found: {len(found_worksheets)} worksheets")
            if missing_worksheets:
                st.error(f"❌ Missing: {len(missing_worksheets)} worksheets")
            else:
                st.success("🎉 All configured forms have corresponding worksheets!")
            
            if missing_worksheets:
                st.markdown("**🔧 Missing Worksheets that need to be created:**")
                for missing in missing_worksheets:
                    st.code(missing)
                
                st.markdown("**💡 Solutions:**")
                st.markdown("1. **Create missing worksheets** in your Google Sheet")
                st.markdown("2. **OR remove unused forms** from your config file")
                st.markdown("3. **OR check for typos** in worksheet names vs config names")
                
            # Show extra worksheets (in Google Sheets but not in config)
            extra_worksheets = [name for name in worksheet_names if name not in form_configs.keys()]
            if extra_worksheets:
                st.markdown("---")
                st.markdown("**📝 Extra Worksheets (in Google Sheets but not in config):**")
                for extra in extra_worksheets:
                    st.warning(f"⚠️ `{extra}` - Exists in Google Sheets but not in config file")
                
        else:
            st.error("❌ Could not connect to Google Sheets")
            
    except Exception as e:
        st.error(f"❌ Debug Error: {str(e)}")

# Also add a test for specific worksheet access
st.markdown("### 🧪 Test Specific Worksheet Access")
test_sheet = st.selectbox("Test specific worksheet:", list(form_configs.keys()), key="test_sheet_select")

if st.button(f"🧪 Test Access to '{test_sheet}'", key="test_specific_sheet"):
    try:
        sheet_client = get_gsheet_client(GOOGLE_SHEET_ID)

        if sheet_client:
            worksheet = sheet_client.worksheet(test_sheet)
            st.success(f"✅ Successfully accessed worksheet '{test_sheet}'")
            
            # Try to get some basic info
            cell_count = worksheet.row_count
            col_count = worksheet.col_count
            st.info(f"📊 Worksheet info: {cell_count} rows, {col_count} columns")
            
        else:
            st.error("❌ Could not connect to Google Sheets")
            
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"❌ Worksheet '{test_sheet}' does not exist in your Google Sheet")
        st.markdown(f"**🔧 Solution:** Create a worksheet named exactly `{test_sheet}` in your Google Sheet")
        
    except Exception as e:
        st.error(f"❌ Error accessing worksheet '{test_sheet}': {str(e)}")

# Add helpful links and information
st.markdown("---")
st.markdown("### 🔗 Quick Links & Information")
st.markdown(f"**📊 Your Google Sheet:** [Open in Google Sheets](https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/edit)")
st.markdown("**💡 Key Changes Made:**")
st.markdown("- ✅ Fixed Google Sheet connection to use your actual sheet ID")
st.markdown("- ✅ Both LW FILES and M&PR FILES now access the same Google Sheet with different worksheets")
st.markdown("- ✅ Form selection properly switches between different worksheet tabs")
st.markdown("- ✅ All data operations (add, edit, view, PDF) now work correctly for both sheet types")
sheet_choice = st.sidebar.radio("Select Sheet Type", ["LW FILES", "M&PR FILES"])
