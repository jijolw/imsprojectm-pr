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
# -----------------------------
# GOOGLE SHEETS AUTHENTICATION - FIXED
# -----------------------------

import streamlit as st
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials

# ✅ Google API Scope
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# ✅ Use credentials from Streamlit secrets (no fallback)
creds = json.loads(st.secrets["IMS_CREDENTIALS_JSON"])

# Write to a temporary file for gspread
with open("temp_creds.json", "w") as f:
    json.dump(creds, f)

CREDENTIAL_FILE = "temp_creds.json"

@st.cache_resource
def get_gsheet_client(sheet_id):
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIAL_FILE, SCOPE)
        client = gspread.authorize(creds)
        return client.open_by_key(sheet_id)
    except Exception as e:
        st.error(f"❌ Error connecting to Google Sheets: {e}")
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
# CACHING FUNCTIONS
# -----------------------------
@st.cache_data(ttl=300)
def get_sheet_data(sheet_name, sheet_id):
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

@st.cache_data(ttl=600)
def get_all_sheet_names(sheet_id):
    try:
        sheet_client = get_gsheet_client(sheet_id)
        if not sheet_client:
            return []
        return [ws.title for ws in sheet_client.worksheets()]
    except Exception as e:
        st.error(f"Error fetching sheet names: {e}")
        return []

# -----------------------------
# FORM CONFIGURATION
# -----------------------------
@st.cache_data(ttl=3600)
def load_form_configs_for_sheet(sheet_type):
    try:
        config_file = "form_configs.json" if sheet_type == "LW FILES" else "forms_mpr_configs.json"
        with open(config_file, "r") as f:
            configs = json.load(f)
        return configs
    except FileNotFoundError:
        st.error(f"Config file not found for {sheet_type}")
        return {}
    except json.JSONDecodeError as e:
        st.error(f"Error parsing config file for {sheet_type}: {e}")
        return {}

def api_rate_limit():
    if 'last_api_call' not in st.session_state:
        st.session_state.last_api_call = 0
    
    current_time = time.time()
    time_since_last = current_time - st.session_state.last_api_call
    
    if time_since_last < 1:
        time.sleep(1 - time_since_last)
    
    st.session_state.last_api_call = time.time()

# -----------------------------
# SHEET MANAGEMENT FUNCTIONS
# -----------------------------
def create_new_worksheet(sheet_name, headers):
    try:
        api_rate_limit()
        sheet_client = get_gsheet_client(GOOGLE_SHEET_ID)
        if not sheet_client:
            return False
        
        worksheet = sheet_client.add_worksheet(title=sheet_name, rows=100, cols=len(headers))
        worksheet.update('A1', [headers])
        
        st.success(f"✅ Created worksheet '{sheet_name}' with {len(headers)} columns")
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error creating worksheet: {e}")
        return False

def delete_worksheet(sheet_name):
    try:
        api_rate_limit()
        sheet_client = get_gsheet_client(GOOGLE_SHEET_ID)
        if not sheet_client:
            return False
        
        worksheet = sheet_client.worksheet(sheet_name)
        sheet_client.del_worksheet(worksheet)
        
        st.success(f"✅ Deleted worksheet '{sheet_name}'")
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error deleting worksheet: {e}")
        return False

# Load configurations
form_configs = load_form_configs_for_sheet(sheet_choice)

# -----------------------------
# MAIN APPLICATION
# -----------------------------
st.title("📋 IMS Form Entry System")
st.markdown(f"**Active Sheet Type:** `{sheet_choice}` | **Available Forms:** {len(form_configs)}")

if not form_configs:
    st.warning(f"⚠️ No form configurations found for '{sheet_choice}'.")
    st.stop()

# Create tabs for better organization
tab1, tab2, tab3, tab4 = st.tabs(["📝 Form Entry", "📊 Data View", "📄 PDF Export", "⚙️ Sheet Management"])

# -----------------------------
# TAB 1: FORM ENTRY
# -----------------------------
with tab1:
    st.markdown("## ✏️ Form Entry & Edit")

    selected_form = st.selectbox(
        "Select a form to fill or edit",
        list(form_configs.keys()),
        format_func=lambda x: f"{x} - {form_configs[x]['title']}"
    )

    form_data = form_configs[selected_form]
    headers, records = get_sheet_data(selected_form, GOOGLE_SHEET_ID)
    df = pd.DataFrame(records) if records else pd.DataFrame()

    st.subheader(form_data["title"])

    edit_mode = st.checkbox("✏️ Enable Edit Mode")
    selected_row_index = None
    prefill_data = {}

    if edit_mode and not df.empty:
        st.info("Select a row to edit.")
        df_display = df.copy()
        df_display.index += 2
        st.dataframe(df_display, use_container_width=True)

        row_num = st.number_input("Enter row number to edit:", min_value=2, max_value=len(df)+1, step=1)
        selected_row_index = row_num - 2
        if 0 <= selected_row_index < len(df):
            st.success(f"Editing row {row_num}")
            prefill_data = df.iloc[selected_row_index].to_dict()

    # Form with improved text handling for long content
    form_values = {}
    with st.form(key="entry_form"):
        # Organize fields in columns
        num_fields = len(form_data["fields"])
        cols_per_row = 2
        
        for i in range(0, num_fields, cols_per_row):
            cols = st.columns(cols_per_row)
            
            for j in range(cols_per_row):
                field_idx = i + j
                if field_idx < num_fields:
                    field = form_data["fields"][field_idx]
                    current_value = prefill_data.get(field, "")
                    
                    with cols[j]:
                        # Use text_area for potentially long content or fields that typically contain long text
                        if (len(str(current_value)) > 100 or 
                            any(keyword in field.lower() for keyword in 
                                ['description', 'details', 'notes', 'remarks', 'comment', 'address', 'specification', 'procedure'])):
                            
                            form_values[field] = st.text_area(
                                field, 
                                value=str(current_value),
                                height=120,
                                help=f"Characters: {len(str(current_value))}"
                            )
                        else:
                            form_values[field] = st.text_input(field, value=str(current_value))

        # Signatures section
        signature_values = {}
        if form_data.get("signatures"):
            st.markdown("---")
            st.markdown("### ✍️ Signatures & Approvals")
            
            sig_cols = st.columns(min(4, len(form_data["signatures"])))
            
            for i, signer in enumerate(form_data["signatures"]):
                default_checked = prefill_data.get(signer, "") == "✔️ Yes"
                with sig_cols[i % len(sig_cols)]:
                    signature_values[signer] = st.checkbox(signer, value=default_checked)

        # Submit button
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            submitted = st.form_submit_button("💾 Submit Entry", use_container_width=True, type="primary")

    if submitted:
        try:
            api_rate_limit()
            
            # Add timestamp
            form_values["Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Add signatures
            for signer, signed in signature_values.items():
                form_values[signer] = "✔️ Yes" if signed else "❌ No"

            # Ensure all data is preserved - handle long content properly
            row = []
            for col in headers:
                value = form_values.get(col, "")
                # Convert to string and preserve all content
                row.append(str(value) if value is not None else "")

            sheet_client = get_gsheet_client(GOOGLE_SHEET_ID)
            if sheet_client:
                worksheet = sheet_client.worksheet(selected_form)
                
                if edit_mode and selected_row_index is not None:
                    # Update entire row to preserve all data
                    worksheet.update(f"A{selected_row_index + 2}:{chr(ord('A') + len(headers) - 1)}{selected_row_index + 2}", [row])
                    st.success(f"✅ Row {selected_row_index + 2} updated successfully.")
                else:
                    worksheet.append_row(row)
                    st.success("✅ New entry submitted successfully.")
                
                st.cache_data.clear()
                st.rerun()
            
        except Exception as e:
            st.error(f"❌ Error writing to Google Sheet: {e}")
            st.error("Please check your connection and try again.")

# -----------------------------
# TAB 2: DATA VIEW
# -----------------------------
with tab2:
    st.markdown("## 📊 Data View & Management")
    
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        if st.button("🔄 Refresh Data"):
            st.cache_data.clear()
            st.rerun()
    
    with col2:
        view_mode = st.selectbox("View Mode", ["Enhanced Table View", "Card View"])
    
    with col3:
        available_sheets = [name for name in get_all_sheet_names(GOOGLE_SHEET_ID) if name in form_configs]
        view_sheet = st.selectbox("📄 Select Sheet to View", available_sheets)

    if view_sheet:
        view_headers, view_records = get_sheet_data(view_sheet, GOOGLE_SHEET_ID)
        view_df = pd.DataFrame(view_records) if view_records else pd.DataFrame()
        
        if view_df.empty:
            st.warning("No data available in the selected sheet.")
        else:
            st.info(f"📈 Total Records: {len(view_df)} | Last Updated: {datetime.now().strftime('%H:%M:%S')}")
            
            # Search and filter functionality
            search_term = st.text_input("🔍 Search in data:", placeholder="Enter search term...")
            
            if search_term:
                mask = view_df.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
                filtered_df = view_df[mask]
                st.info(f"Found {len(filtered_df)} records matching '{search_term}'")
            else:
                filtered_df = view_df
            
            if view_mode == "Enhanced Table View":
                display_df = filtered_df.copy()
                display_df.index = range(2, len(display_df) + 2)
                
                # Get form config for better organization
                sheet_config = form_configs.get(view_sheet, {})
                signature_columns = sheet_config.get("signatures", [])
                form_fields = sheet_config.get("fields", [])
                
                # Column selection for better visibility
                st.markdown("### 📋 Column Selection")
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    regular_fields = [col for col in display_df.columns if col in form_fields]
                    show_regular = st.checkbox("Show All Form Fields", value=True)
                    if not show_regular and regular_fields:
                        selected_regular = st.multiselect("Select Fields:", regular_fields, default=regular_fields[:5])
                    else:
                        selected_regular = regular_fields
                
                with col2:
                    signature_fields = [col for col in display_df.columns if col in signature_columns]
                    show_signatures = st.checkbox("Show All Signatures", value=True)
                    if not show_signatures and signature_fields:
                        selected_signatures = st.multiselect("Select Signatures:", signature_fields, default=signature_fields)
                    else:
                        selected_signatures = signature_fields
                
                with col3:
                    other_fields = [col for col in display_df.columns if col not in form_fields and col not in signature_columns]
                    show_timestamp = st.checkbox("Show Timestamp", value=True)
                    selected_other = ["Timestamp"] if show_timestamp and "Timestamp" in other_fields else []
                
                # Combine selected columns
                columns_to_show = selected_regular + selected_signatures + selected_other
                
                if columns_to_show:
                    display_filtered = display_df[columns_to_show]
                    
                    # Style the dataframe to highlight signatures
                    def style_signatures(val):
                        val_str = str(val)
                        if "✔️" in val_str or "Yes" in val_str:
                            return 'background-color: #d4edda; color: #155724; font-weight: bold'
                        elif "❌" in val_str or "No" in val_str:
                            return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
                        return ''
                    
                    available_sig_cols = [col for col in selected_signatures if col in display_filtered.columns]
                    if available_sig_cols:
                        styled_df = display_filtered.style.applymap(style_signatures, subset=available_sig_cols)
                        st.dataframe(styled_df, use_container_width=True, height=500)
                    else:
                        st.dataframe(display_filtered, use_container_width=True, height=500)
                    
                    st.caption(f"Showing {len(columns_to_show)} columns out of {len(display_df.columns)} total")
                else:
                    st.warning("Please select at least one column to display.")
            
            else:  # Card View
                st.markdown("### 📋 Card View")
                
                # Pagination for card view
                cards_per_page = 10
                total_cards = len(filtered_df)
                total_pages = (total_cards - 1) // cards_per_page + 1 if total_cards > 0 else 1
                
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    current_page = st.selectbox(f"Page (showing {cards_per_page} cards per page)", 
                                              range(1, total_pages + 1), 
                                              format_func=lambda x: f"Page {x} of {total_pages}")
                
                start_idx = (current_page - 1) * cards_per_page
                end_idx = min(start_idx + cards_per_page, total_cards)
                
                for idx in range(start_idx, end_idx):
                    row = filtered_df.iloc[idx]
                    actual_row_num = filtered_df.index[idx] + 2
                    
                    with st.container():
                        st.markdown(f"### 📄 Row {actual_row_num}")
                        col1, col2 = st.columns(2)
                        
                        fields = list(row.items())
                        mid_point = len(fields) // 2
                        
                        with col1:
                            for field, value in fields[:mid_point]:
                                if value and str(value).strip():
                                    if "✔️" in str(value) or "Yes" in str(value):
                                        st.markdown(f"• **{field}:** :green[{value}]")
                                    elif "❌" in str(value) or "No" in str(value):
                                        st.markdown(f"• **{field}:** :red[{value}]")
                                    else:
                                        # Handle long content in card view
                                        if len(str(value)) > 100:
                                            st.markdown(f"• **{field}:** {str(value)[:100]}...")
                                            with st.expander(f"Show full {field}"):
                                                st.write(value)
                                        else:
                                            st.write(f"• **{field}:** {value}")
                        
                        with col2:
                            for field, value in fields[mid_point:]:
                                if value and str(value).strip():
                                    if "✔️" in str(value) or "Yes" in str(value):
                                        st.markdown(f"• **{field}:** :green[{value}]")
                                    elif "❌" in str(value) or "No" in str(value):
                                        st.markdown(f"• **{field}:** :red[{value}]")
                                    else:
                                        # Handle long content in card view
                                        if len(str(value)) > 100:
                                            st.markdown(f"• **{field}:** {str(value)[:100]}...")
                                            with st.expander(f"Show full {field}"):
                                                st.write(value)
                                        else:
                                            st.write(f"• **{field}:** {value}")
                        
                        st.markdown("---")

# -----------------------------
# TAB 3: ENHANCED PDF EXPORT
# -----------------------------
with tab3:
    st.markdown("## 📝 Enhanced PDF Export System")
    
    # Create sub-tabs for different PDF types
    pdf_tab1, pdf_tab2 = st.tabs(["📄 Individual Row PDFs", "📊 Table PDF Export"])
    
    with pdf_tab1:
        st.markdown("### Individual Row PDF Generation")
        
        pdf_sheet = st.selectbox("📄 Select Sheet", available_sheets, key="pdf_individual_sheet")
        
        if pdf_sheet:
            pdf_headers, pdf_records = get_sheet_data(pdf_sheet, GOOGLE_SHEET_ID)
            pdf_df = pd.DataFrame(pdf_records) if pdf_records else pd.DataFrame()
            
            if not pdf_df.empty:
                selected_rows = st.multiselect("Select Row(s) for Individual PDFs", pdf_df.index + 2)
                
                if selected_rows:
                    config = form_configs.get(pdf_sheet, {})
                    
                    for row_num in selected_rows:
                        actual_index = row_num - 2
                        if 0 <= actual_index < len(pdf_df):
                            entry_data = pdf_df.iloc[actual_index].to_dict()
                            
                            with st.expander(f"📋 Preview Row {row_num}", expanded=False):
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    st.write("**Form Fields:**")
                                    for field in config.get("fields", []):
                                        value = entry_data.get(field, "")
                                        if len(str(value)) > 50:
                                            st.write(f"• **{field}:** {str(value)[:50]}...")
                                        else:
                                            st.write(f"• **{field}:** {value or '-'}")
                                
                                with col2:
                                    if config.get("signatures"):
                                        st.write("**Signatures:**")
                                        for signer in config["signatures"]:
                                            status = entry_data.get(signer, "❌ No")
                                            color = "green" if "✔️" in status or "Yes" in status else "red"
                                            st.markdown(f"• **{signer}:** :{color}[{status}]")
                            
                            if st.button(f"📄 Generate PDF for Row {row_num}", key=f"gen_individual_{row_num}"):
                                # Enhanced individual PDF template
                                individual_template = """
                                <!DOCTYPE html>
                                <html>
                                <head>
                                    <meta charset="utf-8">
                                    <style>
                                        body { 
                                            font-family: Arial, sans-serif; 
                                            margin: 20px; 
                                            color: #333; 
                                            line-height: 1.6;
                                        }
                                        .header { 
                                            text-align: center; 
                                            margin-bottom: 30px; 
                                            border-bottom: 3px solid #333; 
                                            padding-bottom: 15px; 
                                        }
                                        .form-title { 
                                            font-size: 28px; 
                                            font-weight: bold; 
                                            margin-bottom: 10px; 
                                            color: #2c3e50;
                                        }
                                        .export-info { 
                                            font-size: 12px; 
                                            color: #7f8c8d; 
                                            margin-top: 5px; 
                                        }
                                        .content-section {
                                            margin-bottom: 25px;
                                        }
                                        .section-title {
                                            font-size: 18px;
                                            font-weight: bold;
                                            color: #34495e;
                                            margin-bottom: 15px;
                                            border-bottom: 1px solid #bdc3c7;
                                            padding-bottom: 5px;
                                        }
                                        .field-container {
                                            margin-bottom: 15px;
                                            padding: 10px;
                                            background-color: #f8f9fa;
                                            border-left: 4px solid #3498db;
                                        }
                                        .field-name { 
                                            font-weight: bold; 
                                            color: #2c3e50;
                                            margin-bottom: 5px;
                                        }
                                        .field-value {
                                            color: #555;
                                            word-wrap: break-word;
                                        }
                                        .signatures-section { 
                                            margin-top: 30px; 
                                            border-top: 2px solid #ecf0f1; 
                                            padding-top: 20px; 
                                        }
                                        .signature-item { 
                                            margin-bottom: 12px; 
                                            padding: 10px; 
                                            border: 1px solid #ddd; 
                                            background-color: #fff;
                                            border-radius: 4px;
                                        }
                                        .signature-name { 
                                            font-weight: bold; 
                                            display: inline-block; 
                                            width: 250px; 
                                        }
                                        .signature-status { 
                                            font-size: 16px; 
                                            font-weight: bold;
                                        }
                                        .signed { color: #27ae60; }
                                        .not-signed { color: #e74c3c; }
                                        .footer { 
                                            margin-top: 40px; 
                                            text-align: center; 
                                            font-size: 10px; 
                                            color: #95a5a6; 
                                            border-top: 1px solid #ecf0f1; 
                                            padding-top: 15px; 
                                        }
                                    </style>
                                </head>
                                <body>
                                    <div class="header">
                                        <div class="form-title">{{ title }}</div>
                                        <div class="export-info">Row {{ row_number }} - Generated on {{ export_date }}</div>
                                    </div>
                                    
                                    <div class="content-section">
                                        <div class="section-title">📝 Form Data</div>
                                        {% for field, value in form_data.items() %}
                                        <div class="field-container">
                                            <div class="field-name">{{ field }}</div>
                                            <div class="field-value">{{ value if value and value.strip() else '-' }}</div>
                                        </div>
                                        {% endfor %}
                                    </div>
                                    
                                    {% if signatures %}
                                    <div class="signatures-section">
                                        <div class="section-title">✍️ Signatures & Approvals</div>
                                        {% for signer, status in signatures.items() %}
                                        <div class="signature-item">
                                            <span class="signature-name">{{ signer }}:</span>
                                            <span class="signature-status {% if '✔️' in status or 'Yes' in status %}signed{% else %}not-signed{% endif %}">{{ status }}</span>
                                        </div>
                                        {% endfor %}
                                    </div>
                                    {% endif %}
                                    
                                    <div class="footer">
                                        Generated from IMS Form Entry System | 
                                        {{ export_date }}
                                    </div>
                                </body>
                                </html>
                                """
                                
                                try:
                                    template = Template(individual_template)
                                    
                                    # Prepare form data
                                    form_data_for_pdf = {field: entry_data.get(field, "") for field in config.get("fields", [])}
                                    signature_data_for_pdf = {signer: entry_data.get(signer, "❌ No") for signer in config.get("signatures", [])}
                                    
                                    rendered_html = template.render(
                                        title=config.get("title", pdf_sheet),
                                        row_number=row_num,
                                        export_date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        form_data=form_data_for_pdf,
                                        signatures=signature_data_for_pdf
                                    )
                                    
                                    pdf_buffer = BytesIO()
                                    pisa_status = pisa.CreatePDF(BytesIO(rendered_html.encode("utf-8")), dest=pdf_buffer)
                                    
                                    if not pisa_status.err:
                                        pdf_buffer.seek(0)
                                        st.success("✅ Individual PDF generated successfully!")
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
                                    st.error(f"❌ Error generating PDF: {str(e)}")
    
    with pdf_tab2:
        st.markdown("### 📊 Enhanced Table PDF Export")
        
        table_pdf_sheet = st.selectbox("📄 Select Sheet", available_sheets, key="pdf_table_sheet")
        
        if table_pdf_sheet:
            table_headers, table_records = get_sheet_data(table_pdf_sheet, GOOGLE_SHEET_ID)
            table_df = pd.DataFrame(table_records) if table_records else pd.DataFrame()
            
            if not table_df.empty:
                sheet_config = form_configs.get(table_pdf_sheet, {})
                signature_columns = sheet_config.get("signatures", [])
                form_fields = sheet_config.get("fields", [])
                
                # Column selection for PDF
                st.markdown("### 📋 Select Columns for PDF Export")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown("**📝 Form Fields:**")
                    available_form_fields = [col for col in table_df.columns if col in form_fields]
                    selected_form_fields = st.multiselect(
                        "Select Form Fields:", 
                        available_form_fields, 
                        default=available_form_fields[:5],  # Default to first 5 to avoid overcrowding
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
                
                if all_selected_columns:
                    # Row selection
                    st.markdown("### 📊 Select Rows")
                    row_selection = st.radio("Row Selection:", ["All Rows", "Select Specific Rows"])
                    
                    if row_selection == "Select Specific Rows":
                        selected_rows = st.multiselect(
                            "Select specific rows:", 
                            table_df.index + 2,
                            key="table_pdf_rows"
                        )
                        rows_to_include = [row - 2 for row in selected_rows] if selected_rows else []
                    else:
                        rows_to_include = list(range(len(table_df)))
                        st.info(f"Will include all {len(table_df)} rows")
                    
                    # PDF options
                    st.markdown("### 🎨 PDF Layout Options")
                    col1, col2 = st.columns(2)
                    with col1:
                        page_orientation = st.selectbox("Page Orientation", ["Landscape", "Portrait"])
                    with col2:
                        include_summary = st.checkbox("Include Summary Statistics", value=True)
                    
                    if rows_to_include and st.button("📊 Generate Enhanced Table PDF", key="generate_table_pdf"):
                        try:
                            config = form_configs.get(table_pdf_sheet, {})
                            title = config.get("title", table_pdf_sheet)
                            
                            # Prepare data
                            filtered_df = table_df.iloc[rows_to_include][all_selected_columns]
                            
                            # Calculate summary statistics
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
                            
                            # Enhanced table PDF template
                            table_template = """
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
                                        font-size: 10px; 
                                    }
                                    .header { 
                                        text-align: center; 
                                        margin-bottom: 20px; 
                                        border-bottom: 2px solid #333; 
                                        padding-bottom: 10px; 
                                    }
                                    .form-title { 
                                        font-size: 20px; 
                                        font-weight: bold; 
                                        margin-bottom: 5px; 
                                    }
                                    .export-info { 
                                        font-size: 11px; 
                                        color: #666; 
                                        margin-top: 5px; 
                                    }
                                    .summary-section {
                                        margin-bottom: 20px;
                                        padding: 15px;
                                        background-color: #f8f9fa;
                                        border: 1px solid #dee2e6;
                                        border-radius: 5px;
                                    }
                                    .summary-title {
                                        font-size: 14px;
                                        font-weight: bold;
                                        margin-bottom: 10px;
                                        color: #495057;
                                    }
                                    .summary-grid {
                                        display: grid;
                                        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
                                        gap: 15px;
                                    }
                                    .summary-item {
                                        font-size: 11px;
                                        padding: 8px;
                                        background: white;
                                        border-radius: 3px;
                                    }
                                    .summary-name {
                                        font-weight: bold;
                                        color: #495057;
                                    }
                                    .data-table { 
                                        width: 100%; 
                                        border-collapse: collapse; 
                                        margin-bottom: 20px; 
                                        font-size: 9px; 
                                    }
                                    .data-table th { 
                                        background-color: #343a40; 
                                        color: white;
                                        border: 1px solid #495057; 
                                        padding: 8px 4px; 
                                        text-align: left; 
                                        font-weight: bold; 
                                        font-size: 9px;
                                        word-wrap: break-word;
                                    }
                                    .data-table td { 
                                        border: 1px solid #dee2e6; 
                                        padding: 6px 4px; 
                                        text-align: left; 
                                        word-wrap: break-word; 
                                        font-size: 8px;
                                        max-width: 150px;
                                        overflow: hidden;
                                    }
                                    .row-num { 
                                        background-color: #f8f9fa; 
                                        font-weight: bold; 
                                        text-align: center; 
                                        min-width: 30px;
                                        max-width: 40px;
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
                                    .column-header-form { background-color: #17a2b8 !important; }
                                    .column-header-signature { background-color: #28a745 !important; }
                                    .column-header-other { background-color: #6c757d !important; }
                                    .footer { 
                                        margin-top: 20px; 
                                        text-align: center; 
                                        font-size: 9px; 
                                        color: #6c757d; 
                                        border-top: 1px solid #dee2e6; 
                                        padding-top: 10px; 
                                    }
                                </style>
                            </head>
                            <body>
                                <div class="header">
                                    <div class="form-title">{{ title }} - Enhanced Table View</div>
                                    <div class="export-info">
                                        Generated on {{ export_date }} | 
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
                                            <td class="{% if column in signature_fields and ('✔️' in row_data.data.get(column, '') or 'Yes' in row_data.data.get(column, '')) %}signed{% elif column in signature_fields and ('❌' in row_data.data.get(column, '') or 'No' in row_data.data.get(column, '')) %}not-signed{% endif %}">
                                                {{ row_data.data.get(column, '') if row_data.data.get(column, '').strip() else '-' }}
                                            </td>
                                            {% endfor %}
                                        </tr>
                                        {% endfor %}
                                    </tbody>
                                </table>
                                
                                <div class="footer">
                                    Generated from IMS Form Entry System | 
                                    Page Orientation: {{ orientation }} | 
                                    Total Columns: {{ column_count }} | 
                                    Total Rows: {{ total_records }}
                                </div>
                            </body>
                            </html>
                            """
                            
                            # Prepare template data
                            table_data = []
                            for idx, (original_idx, row) in enumerate(filtered_df.iterrows()):
                                table_data.append({
                                    'row_number': original_idx + 2,
                                    'data': row.to_dict()
                                })
                            
                            template = Template(table_template)
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
                                with st.expander("📋 PDF Content Summary", expanded=False):
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.markdown("**📝 Form Fields:**")
                                        for field in selected_form_fields:
                                            st.write(f"• {field}")
                                    with col2:
                                        st.markdown("**✍️ Signatures:**")
                                        for field in selected_signature_fields:
                                            st.write(f"• {field}")
                                    with col3:
                                        st.markdown("**⏰ Other Fields:**")
                                        for field in selected_other_fields:
                                            st.write(f"• {field}")
                                
                                filename = f"{table_pdf_sheet}_Table_{len(filtered_df)}rows_{len(all_selected_columns)}cols_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                                
                                st.download_button(
                                    label=f"⬇️ Download Table PDF ({len(filtered_df)} rows, {len(all_selected_columns)} columns)",
                                    data=pdf_buffer.getvalue(),
                                    file_name=filename,
                                    mime="application/pdf",
                                    key="dl_table_pdf"
                                )
                            else:
                                st.error(f"❌ PDF generation error: {pisa_status.err}")
                                
                        except Exception as e:
                            st.error(f"❌ Error generating table PDF: {str(e)}")
                            st.error("Please check your data and try again.")
                
                else:
                    st.warning("Please select at least one column for the PDF export.")

# -----------------------------
# TAB 4: SHEET MANAGEMENT
# -----------------------------
with tab4:
    st.markdown("## ⚙️ Sheet Management")
    
    current_sheets = get_all_sheet_names(GOOGLE_SHEET_ID)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### ➕ Create New Sheet")
        
        new_sheet_name = st.text_input("New Sheet Name", help="Enter a unique name for the new worksheet")
        
        # Template selection
        template_form = st.selectbox(
            "Use template from:", 
            ["Custom Headers"] + list(form_configs.keys()),
            help="Choose an existing form configuration as template or create custom headers"
        )
        
        if template_form == "Custom Headers":
            headers_input = st.text_area(
                "Headers (comma-separated)", 
                placeholder="Timestamp, Field1, Field2, Signature1, Signature2",
                help="Enter column headers separated by commas"
            )
            headers = [h.strip() for h in headers_input.split(",") if h.strip()] if headers_input else []
            
            if headers:
                st.info(f"Will create sheet with {len(headers)} columns")
                with st.expander("Preview Headers"):
                    for i, header in enumerate(headers, 1):
                        st.write(f"{i}. {header}")
        else:
            config = form_configs[template_form]
            headers = ["Timestamp"] + config["fields"] + config.get("signatures", [])
            st.info(f"Using template from '{template_form}' ({len(headers)} columns)")
            
            with st.expander("Preview Template Headers"):
                st.write("**Form Fields:**")
                for field in config["fields"]:
                    st.write(f"• {field}")
                if config.get("signatures"):
                    st.write("**Signatures:**")
                    for sig in config["signatures"]:
                        st.write(f"• {sig}")
        
        if st.button("✅ Create New Sheet", type="primary") and new_sheet_name and headers:
            if new_sheet_name not in current_sheets:
                if create_new_worksheet(new_sheet_name, headers):
                    st.balloons()
                    st.rerun()
            else:
                st.error("❌ Sheet name already exists! Please choose a different name.")
    
    with col2:
        st.markdown("### 🗑️ Delete Existing Sheet")
        
        if current_sheets:
            sheets_not_in_config = [sheet for sheet in current_sheets if sheet not in form_configs]
            deletable_sheets = current_sheets if sheets_not_in_config else current_sheets
            
            sheet_to_delete = st.selectbox(
                "Select sheet to delete", 
                deletable_sheets,
                help="Select a worksheet to permanently delete"
            )
            
            # Show sheet info
            if sheet_to_delete:
                try:
                    sheet_headers, sheet_records = get_sheet_data(sheet_to_delete, GOOGLE_SHEET_ID)
                    if sheet_records:
                        st.info(f"📊 Sheet contains {len(sheet_records)} records with {len(sheet_headers)} columns")
                    else:
                        st.info("📊 Sheet is empty")
                    
                    config_status = "✅ Has form configuration" if sheet_to_delete in form_configs else "⚠️ No form configuration"
                    st.write(f"Status: {config_status}")
                        
                except Exception:
                    st.warning("Could not read sheet information")
            
            st.error("⚠️ **WARNING:** This action cannot be undone!")
            confirm_delete = st.checkbox(f"I understand this will permanently delete '{sheet_to_delete}' and all its data")
            
            if st.button("🗑️ Delete Sheet", type="secondary") and confirm_delete and sheet_to_delete:
                if delete_worksheet(sheet_to_delete):
                    st.success(f"Sheet '{sheet_to_delete}' has been deleted")
                    st.rerun()
        else:
            st.info("No sheets available to delete")
    
    # Current sheets overview
    st.markdown("---")
    st.markdown("### 📋 Current Worksheets Overview")
    
    if current_sheets:
        # Create a summary dataframe
        sheet_summary = []
        for sheet in current_sheets:
            try:
                headers, records = get_sheet_data(sheet, GOOGLE_SHEET_ID)
                has_config = "✅ Yes" if sheet in form_configs else "❌ No"
                record_count = len(records) if records else 0
                column_count = len(headers) if headers else 0
                
                sheet_summary.append({
                    "Sheet Name": sheet,
                    "Records": record_count,
                    "Columns": column_count,
                    "Has Config": has_config
                })
            except:
                sheet_summary.append({
                    "Sheet Name": sheet,
                    "Records": "Error",
                    "Columns": "Error", 
                    "Has Config": "❌ No"
                })
        
        summary_df = pd.DataFrame(sheet_summary)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
        
        # Statistics
        total_sheets = len(current_sheets)
        configured_sheets = len([s for s in current_sheets if s in form_configs])
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Sheets", total_sheets)
        with col2:
            st.metric("Configured Sheets", configured_sheets)
        with col3:
            st.metric("Unconfigured Sheets", total_sheets - configured_sheets)
            
    else:
        st.info("No worksheets found in the selected Google Sheet")

# -----------------------------
# SIDEBAR INFORMATION
# -----------------------------
st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Current Status")
st.sidebar.metric("Sheet Type", sheet_choice)
st.sidebar.metric("Available Forms", len(form_configs))
st.sidebar.metric("Total Worksheets", len(get_all_sheet_names(GOOGLE_SHEET_ID)))

# Connection status
connection_status = get_gsheet_client(GOOGLE_SHEET_ID) is not None
st.sidebar.markdown("### 🔗 Connection Status")
if connection_status:
    st.sidebar.success("✅ Google Sheets Connected")
else:
    st.sidebar.error("❌ Connection Failed")

# Quick actions
st.sidebar.markdown("### ⚡ Quick Actions")
if st.sidebar.button("🔄 Refresh All Data"):
    st.cache_data.clear()
    st.cache_resource.clear()
    st.success("All data refreshed!")
    st.rerun()

# Help section
with st.sidebar.expander("❓ Help & Tips"):
    st.markdown("""
    **Form Entry Tips:**
    - Use edit mode to modify existing records
    - Long text fields automatically expand
    - All data is preserved during edits
    
    **Data View Tips:**
    - Use search to find specific records
    - Card view shows full content for long fields
    - Signatures are color-coded (green/red)
    
    **PDF Export Tips:**
    - Individual PDFs for detailed records
    - Table PDFs for overview reports
    - Select specific columns for cleaner output
    
    **Sheet Management:**
    - Create sheets from existing templates
    - Delete unused or test sheets safely
    - Monitor sheet statistics
    """)

# Footer information
st.markdown("---")
st.markdown("**IMS Form Entry System** | Streamlined Version | Enhanced Data Handling")

col1, col2, col3 = st.columns(3)
with col1:
    st.caption("✅ Enhanced form entry with improved text handling")
with col2:
    st.caption("✅ Complete PDF export with advanced formatting")  
with col3:
    st.caption("✅ Sheet management with create/delete capabilities")