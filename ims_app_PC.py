import streamlit as st
st.set_page_config(page_title="IMS Form Entry", layout="wide")

import gspread
import json
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import time
from io import BytesIO
from xhtml2pdf import pisa
from jinja2 import Template
import hashlib

# =========================
# NEW: utility for column letters (safe beyond Z)
# =========================
def _col_letter(col_index_1based: int) -> str:
    """Convert 1-based column index to A1 letter(s)."""
    s = ""
    n = col_index_1based
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

# =========================
# NEW: retry helper for 429 quota/rate-limits
# =========================
def _retry_gspread(callable_fn, *args, tries=5, base_delay=0.75, **kwargs):
    """
    Retries gspread API calls on 429 / quota exceeded errors with exponential backoff.
    Works both locally and on Streamlit Cloud.
    """
    for attempt in range(tries):
        try:
            return callable_fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            # gspread wraps googleapiclient errors; inspect status and message
            msg = str(e).lower()
            code = getattr(getattr(e, "response", None), "status_code", None)
            # 429 or "quota exceeded" or "rate limit"
            if code == 429 or "quota exceeded" in msg or "rate limit" in msg:
                time.sleep(base_delay * (2 ** attempt))
                continue
            raise

# LOGIN SYSTEM - Updated permissions
DEFAULT_USERS = {
    "cso": {
        "password": "cso@2024",
        "role": "Chief Security Officer",
        "permissions": ["read", "write", "export"]  # Added write permission
    },
    "supervisor": {
        "password": "super@2024",
        "role": "Supervisor",
        "permissions": ["read", "write", "export", "manage"]
    },
    "controlling_officer": {
        "password": "control@2024",
        "role": "Controlling Officer", 
        "permissions": ["read", "write", "export"]  # Added write permission
    }
}

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(password, hashed):
    return hash_password(password) == hashed

def initialize_users():
    if 'users_initialized' not in st.session_state:
        if "IMS_USERS" in st.secrets:
            st.session_state.users = json.loads(st.secrets["IMS_USERS"])
        else:
            st.session_state.users = {}
            for username, user_data in DEFAULT_USERS.items():
                st.session_state.users[username] = {
                    "password_hash": hash_password(user_data["password"]),
                    "role": user_data["role"],
                    "permissions": user_data["permissions"]
                }
        st.session_state.users_initialized = True

def login_form():
    st.markdown("""
    <div style='text-align: center; padding: 2rem; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 10px; margin-bottom: 2rem;'>
        <h1 style='color: white; margin-bottom: 1rem;'>🔐 IMS Form Entry System</h1>
        <h3 style='color: #f0f0f0; font-weight: 300;'>Secure Access Portal</h3>
    </div>
    """, unsafe_allow_html=True)
    
    # Display default credentials for testing
    st.markdown("""
    <div style='background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px; padding: 15px; margin-bottom: 20px;'>
        <h4 style='color: #856404; margin-top: 0;'>🔑 Default Login Credentials (For Testing)</h4>
        <div style='display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px;'>
            <div style='background: white; padding: 10px; border-radius: 5px; border-left: 4px solid #17a2b8;'>
                <strong style='color: #17a2b8;'>Chief Security Officer</strong><br>
                Username: <code>cso</code><br>
                Password: <code>cso@2024</code><br>
                <small>Permissions: Read, Write, Export</small>
            </div>
            <div style='background: white; padding: 10px; border-radius: 5px; border-left: 4px solid #28a745;'>
                <strong style='color: #28a745;'>Supervisor</strong><br>
                Username: <code>supervisor</code><br>
                Password: <code>super@2024</code><br>
                <small>Permissions: Full Access</small>
            </div>
            <div style='background: white; padding: 10px; border-radius: 5px; border-left: 4px solid #dc3545;'>
                <strong style='color: #dc3545;'>Controlling Officer</strong><br>
                Username: <code>controlling_officer</code><br>
                Password: <code>control@2024</code><br>
                <small>Permissions: Read, Write, Export</small>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("### 👤 User Login")
        
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Enter your username")
            password = st.text_input("Password", type="password", placeholder="Enter your password")
            
            col_a, col_b, col_c = st.columns([1, 2, 1])
            with col_b:
                login_button = st.form_submit_button("🚀 Login", use_container_width=True, type="primary")
            
            if login_button:
                if authenticate_user(username, password):
                    st.success("✅ Login successful!")
                    st.balloons()
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("❌ Invalid username or password!")
                    st.warning("Please check your credentials and try again.")

def authenticate_user(username, password):
    initialize_users()
    
    if username in st.session_state.users:
        user_data = st.session_state.users[username]
        if verify_password(password, user_data["password_hash"]):
            st.session_state.logged_in = True
            st.session_state.current_user = username
            st.session_state.user_role = user_data["role"]
            st.session_state.user_permissions = user_data["permissions"]
            st.session_state.login_time = datetime.now()
            return True
    return False

def logout():
    for key in ['logged_in', 'current_user', 'user_role', 'user_permissions', 'login_time']:
        if key in st.session_state:
            del st.session_state[key]
    st.rerun()

def check_permission(required_permission):
    if not st.session_state.get('logged_in', False):
        return False
    return required_permission in st.session_state.get('user_permissions', [])

def require_permission(permission):
    if not check_permission(permission):
        st.error(f"❌ Access Denied: You need '{permission}' permission to access this feature.")
        st.info(f"Your current role: {st.session_state.get('user_role', 'Unknown')}")
        st.stop()

# Check login
initialize_users()

if not st.session_state.get('logged_in', False):
    login_form()
    st.stop()

# Sidebar user info - Updated permission display
st.sidebar.markdown("---")
st.sidebar.markdown("### 👤 Current User")
st.sidebar.success(f"**{st.session_state.get('user_role', 'Unknown')}**")
st.sidebar.info(f"Logged in as: **{st.session_state.get('current_user', 'Unknown')}**")

login_time = st.session_state.get('login_time', datetime.now())
session_duration = datetime.now() - login_time
st.sidebar.caption(f"Session: {int(session_duration.total_seconds()/60)} minutes")

if st.sidebar.button("🚪 Logout", type="secondary"):
    st.sidebar.success("Logged out successfully!")
    logout()

st.sidebar.markdown("### 🔑 Your Permissions")
permissions = st.session_state.get('user_permissions', [])
for perm in permissions:
    st.sidebar.text(f"✅ {perm.title()}")

# CONFIGURATION
require_permission("read")

sheet_choice = st.sidebar.selectbox("Choose File Type", ["LW FILES", "M&PR FILES"])

SHEET_IDS = {
    "LW FILES": "1wxntHZp4xEQWCmLAt2TVF8ohG6uHuvV_3QbaK7wSwGw",
    "M&PR FILES": "17KL-cKMJNGncAJngla_eX6aVdzvoihbdpgOJqOiGycc"
}

GOOGLE_SHEET_ID = SHEET_IDS.get(sheet_choice)

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

if "IMS_CREDENTIALS_JSON" in st.secrets:
    creds = json.loads(st.secrets["IMS_CREDENTIALS_JSON"])
    with open("temp_creds.json", "w") as f:
        json.dump(creds, f)
    CREDENTIAL_FILE = "temp_creds.json"
else:
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

# =========================
# API QUOTA MANAGEMENT (kept; used for writes/manual refresh UI only)
# =========================
class APIQuotaManager:
    def __init__(self, max_calls_per_minute=50):
        self.max_calls = max_calls_per_minute
        if 'api_calls' not in st.session_state:
            st.session_state.api_calls = []
    
    def can_make_call(self):
        now = datetime.now()
        st.session_state.api_calls = [
            call_time for call_time in st.session_state.api_calls 
            if (now - call_time).seconds < 60
        ]
        return len(st.session_state.api_calls) < self.max_calls
    
    def record_call(self):
        st.session_state.api_calls.append(datetime.now())
    
    def wait_time(self):
        if not st.session_state.api_calls:
            return 0
        oldest_call = min(st.session_state.api_calls)
        wait_seconds = 60 - (datetime.now() - oldest_call).seconds
        return max(0, wait_seconds)

quota_manager = APIQuotaManager()

# =========================
# UPDATED: single ranged read + caching + retry
# =========================
@st.cache_data(ttl=300, show_spinner=False)
def get_sheet_data(sheet_name, sheet_id):
    sheet_client = get_gsheet_client(sheet_id)
    if not sheet_client:
        return [], []

    # Only the column range here (no sheet prefix)
    RANGE_OVERRIDES = {
        "LW4 01A": "A1:L",   # 12 columns: Section..Signed by AWM
    }
    range_only = RANGE_OVERRIDES.get(sheet_name, "A1:Z")

    ws = sheet_client.worksheet(sheet_name)

    # ✅ Pass ONLY the range (no "sheet!" prefix)
    values = _retry_gspread(ws.get, range_only)
    if not values:
        return [], []

    headers = values[0]
    rows = values[1:]

    records = []
    for r in rows:
        if len(r) < len(headers):
            r = r + [""] * (len(headers) - len(r))
        records.append(dict(zip(headers, r)))

    return headers, records
# =========================
# UPDATED: sheet names fetch with retry
# =========================
@st.cache_data(ttl=600, show_spinner=False)
def get_all_sheet_names(sheet_id):
    sheet_client = get_gsheet_client(sheet_id)
    if not sheet_client:
        return []
    worksheets = _retry_gspread(sheet_client.worksheets)
    return [ws.title for ws in worksheets]

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
    # keep this for write paths; reads are now cached and retried separately
    if not quota_manager.can_make_call():
        wait_time = quota_manager.wait_time()
        st.warning(f"⏱️ API quota limit reached. Waiting {wait_time} seconds...")
        time.sleep(wait_time)
    
    quota_manager.record_call()
    
    if 'last_api_call' not in st.session_state:
        st.session_state.last_api_call = 0
    
    current_time = time.time()
    time_since_last = current_time - st.session_state.last_api_call
    
    if time_since_last < 1:
        time.sleep(1 - time_since_last)
    
    st.session_state.last_api_call = time.time()

# SHEET MANAGEMENT
def create_new_worksheet(sheet_name, headers):
    require_permission("manage")
    
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
    require_permission("manage")
    
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

form_configs = load_form_configs_for_sheet(sheet_choice)

# MAIN APPLICATION
st.title("📋 IMS Form Entry System")
st.markdown(f"**Active Sheet Type:** `{sheet_choice}` | **Available Forms:** {len(form_configs)} | **User:** {st.session_state.get('user_role', 'Unknown')}")

if not form_configs:
    st.warning(f"⚠️ No form configurations found for '{sheet_choice}'.")
    st.stop()

# Create tabs based on permissions
available_tabs = ["📊 Data View"]

if check_permission("write"):
    available_tabs.insert(0, "📝 Form Entry")

if check_permission("export"):
    available_tabs.append("📄 PDF Export")

if check_permission("manage"):
    available_tabs.append("⚙️ Sheet Management")

# Create dynamic tabs
if len(available_tabs) == 4:
    tab1, tab2, tab3, tab4 = st.tabs(available_tabs)
    tabs = [tab1, tab2, tab3, tab4]
elif len(available_tabs) == 3:
    tab1, tab2, tab3 = st.tabs(available_tabs)
    tabs = [tab1, tab2, tab3, None]
elif len(available_tabs) == 2:
    tab1, tab2 = st.tabs(available_tabs)
    tabs = [tab1, tab2, None, None]
else:
    tab1 = st.tabs(available_tabs)[0]
    tabs = [tab1, None, None, None]

# Map tabs
tab_mapping = {}
for i, tab_name in enumerate(available_tabs):
    if "Form Entry" in tab_name:
        tab_mapping["form_entry"] = tabs[i]
    elif "Data View" in tab_name:
        tab_mapping["data_view"] = tabs[i]
    elif "PDF Export" in tab_name:
        tab_mapping["pdf_export"] = tabs[i]
    elif "Sheet Management" in tab_name:
        tab_mapping["sheet_management"] = tabs[i]

# FORM ENTRY TAB
if "form_entry" in tab_mapping and tab_mapping["form_entry"] is not None:
    with tab_mapping["form_entry"]:
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

        # Form
        form_values = {}
        with st.form(key="entry_form"):
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

            # Signatures
            signature_values = {}
            if form_data.get("signatures"):
                st.markdown("---")
                st.markdown("### ✍️ Signatures & Approvals")
                
                sig_cols = st.columns(min(4, len(form_data["signatures"])))
                
                for i, signer in enumerate(form_data["signatures"]):
                    default_checked = prefill_data.get(signer, "") == "✔️ Yes"
                    with sig_cols[i % len(sig_cols)]:
                        signature_values[signer] = st.checkbox(signer, value=default_checked)

            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                submitted = st.form_submit_button("💾 Submit Entry", use_container_width=True, type="primary")

        if submitted:
            try:
                api_rate_limit()
                
                # NOTE: LW4 01A has no Timestamp/Submitted By columns in your sheet.
                # We will only write columns that exist in headers (matching your sheet).
                for signer, signed in signature_values.items():
                    form_values[signer] = "✔️ Yes" if signed else "❌ No"

                row = []
                for col in headers:
                    value = form_values.get(col, "")
                    row.append(str(value) if value is not None else "")

                sheet_client = get_gsheet_client(GOOGLE_SHEET_ID)
                if sheet_client:
                    worksheet = sheet_client.worksheet(selected_form)
                    
                    if edit_mode and selected_row_index is not None:
                        # Use safe col end letter
                        end_col_letter = _col_letter(len(headers))
                        a1 = f"A{selected_row_index + 2}:{end_col_letter}{selected_row_index + 2}"
                        _retry_gspread(worksheet.update, a1, [row])
                        st.success(f"✅ Row {selected_row_index + 2} updated successfully.")
                    else:
                        _retry_gspread(worksheet.append_row, row)
                        st.success("✅ New entry submitted successfully.")
                    
                    st.cache_data.clear()
                    st.rerun()
                
            except Exception as e:
                st.error(f"❌ Error writing to Google Sheet: {e}")

# DATA VIEW TAB - PRESERVE GOOGLE SHEETS COLUMN ORDER
if "data_view" in tab_mapping and tab_mapping["data_view"] is not None:
    with tab_mapping["data_view"]:
        st.markdown("## 📊 Data View & Management")
        
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            if st.button("🔄 Refresh Data"):
                if quota_manager.can_make_call():
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.warning("⏱️ Please wait before refreshing")
        
        with col2:
            view_mode = st.selectbox("View Mode", ["Enhanced Table View", "Card View"])
        
        with col3:
            available_sheets = [name for name in get_all_sheet_names(GOOGLE_SHEET_ID) if name in form_configs]
            view_sheet = st.selectbox("📄 Select Sheet to View", available_sheets)

        if view_sheet:
            view_headers, view_records = get_sheet_data(view_sheet, GOOGLE_SHEET_ID)
            view_df = pd.DataFrame(view_records) if view_records else pd.DataFrame()
            
            if view_df.empty:
                st.warning("⚠️ No data available in the selected sheet.")
                st.info("💡 This could be because:")
                st.write("• The sheet is empty")
                st.write("• There's a connection issue with Google Sheets") 
                st.write("• The sheet name doesn't match exactly")
                st.write("• API quota has been exceeded")
            else:
                st.info(f"📈 Total Records: {len(view_df)} | Columns: {len(view_df.columns)} | Last Updated: {datetime.now().strftime('%H:%M:%S')}")
                
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
                    
                    # exact order from Google Sheet
                    columns_in_sheet_order = view_headers if view_headers else list(display_df.columns)
                    
                    with st.expander("🔍 Column Order Debug Info", expanded=False):
                        st.write("**Google Sheets Headers (in order):**")
                        for i, header in enumerate(columns_in_sheet_order, 1):
                            st.write(f"{i}. `{header}`")
                        
                        st.write(f"\n**DataFrame Columns (in order):**")
                        for i, col in enumerate(list(display_df.columns), 1):
                            st.write(f"{i}. `{col}`")
                        
                        df_cols = list(display_df.columns)
                        if columns_in_sheet_order == df_cols:
                            st.success("✅ Headers and DataFrame columns match perfectly!")
                        else:
                            st.warning("⚠️ Headers and DataFrame columns don't match exactly")
                    
                    # reorder using headers
                    try:
                        valid_columns = [col for col in columns_in_sheet_order if col in display_df.columns]
                        display_ordered = display_df[valid_columns]
                    except:
                        display_ordered = display_df
                        st.warning("Using DataFrame's natural column order as fallback")
                    
                    # style signature fields if present
                    sheet_config = form_configs.get(view_sheet, {})
                    configured_signature_fields = sheet_config.get("signatures", [])
                    
                    def style_signatures(val):
                        val_str = str(val)
                        if "✔️" in val_str or "Yes" in val_str:
                            return 'background-color: #d4edda; color: #155724; font-weight: bold'
                        elif "❌" in val_str or "No" in val_str:
                            return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
                        return ''
                    
                    signature_cols_in_df = [col for col in configured_signature_fields if col in display_ordered.columns]
                    if signature_cols_in_df:
                        styled_df = display_ordered.style.applymap(style_signatures, subset=signature_cols_in_df)
                        st.dataframe(styled_df, use_container_width=True, height=500)
                    else:
                        st.dataframe(display_ordered, use_container_width=True, height=500)
                    
                    st.info(f"📋 Displaying columns in Google Sheets order: {len(display_ordered.columns)} columns")
                    
                    with st.expander("📊 Column Details"):
                        st.write("**Columns being displayed (in current order):**")
                        for i, col in enumerate(display_ordered.columns, 1):
                            col_type = "Unknown"
                            if col in ["Timestamp", "Submitted By"]:
                                col_type = "System"
                            elif col in sheet_config.get("fields", []):
                                col_type = "Form Field"
                            elif col in configured_signature_fields:
                                col_type = "Signature"
                            
                            st.write(f"{i}. **{col}** ({col_type})")
                
                else:  # Card View
                    st.markdown("### 📋 Card View")
                    
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
                    
                    columns_in_sheet_order = view_headers if view_headers else list(filtered_df.columns)
                    
                    for idx in range(start_idx, end_idx):
                        row = filtered_df.iloc[idx]
                        actual_row_num = filtered_df.index[idx] + 2
                        
                        with st.container():
                            st.markdown(f"### 📄 Row {actual_row_num}")
                            col1, col2 = st.columns(2)
                            
                            ordered_items = []
                            for field in columns_in_sheet_order:
                                if field in row.index:
                                    ordered_items.append((field, row[field]))
                            
                            mid_point = len(ordered_items) // 2
                            
                            with col1:
                                for field, value in ordered_items[:mid_point]:
                                    if value and str(value).strip():
                                        if "✔️" in str(value) or "Yes" in str(value):
                                            st.markdown(f"• **{field}:** :green[{value}]")
                                        elif "❌" in str(value) or "No" in str(value):
                                            st.markdown(f"• **{field}:** :red[{value}]")
                                        else:
                                            if len(str(value)) > 100:
                                                st.markdown(f"• **{field}:** {str(value)[:100]}...")
                                                with st.expander(f"Show full {field}"):
                                                    st.write(value)
                                            else:
                                                st.write(f"• **{field}:** {value}")
                            
                            with col2:
                                for field, value in ordered_items[mid_point:]:
                                    if value and str(value).strip():
                                        if "✔️" in str(value) or "Yes" in str(value):
                                            st.markdown(f"• **{field}:** :green[{value}]")
                                        elif "❌" in str(value) or "No" in str(value):
                                            st.markdown(f"• **{field}:** :red[{value}]")
                                        else:
                                            if len(str(value)) > 100:
                                                st.markdown(f"• **{field}:** {str(value)[:100]}...")
                                                with st.expander(f"Show full {field}"):
                                                    st.write(value)
                                            else:
                                                st.write(f"• **{field}:** {value}")
                            
                            st.markdown("---")

# PDF EXPORT TAB
if "pdf_export" in tab_mapping and tab_mapping["pdf_export"] is not None:
    with tab_mapping["pdf_export"]:
        st.markdown("## 📝 Enhanced PDF Export System")
        
        pdf_tab1, pdf_tab2 = st.tabs(["📄 Individual Row PDFs", "📊 Table PDF Export"])
        
        with pdf_tab1:
            st.markdown("### Individual Row PDF Generation")
            
            available_sheets = [name for name in get_all_sheet_names(GOOGLE_SHEET_ID) if name in form_configs]
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
                                    individual_template = """
                                    <!DOCTYPE html>
                                    <html>
                                    <head>
                                        <meta charset="utf-8">
                                        <style>
                                            body { font-family: Arial, sans-serif; margin: 20px; color: #333; line-height: 1.6; }
                                            .header { text-align: center; margin-bottom: 30px; border-bottom: 3px solid #333; padding-bottom: 15px; }
                                            .form-title { font-size: 28px; font-weight: bold; margin-bottom: 10px; color: #2c3e50; }
                                            .export-info { font-size: 12px; color: #7f8c8d; margin-top: 5px; }
                                            .content-section { margin-bottom: 25px; }
                                            .section-title { font-size: 18px; font-weight: bold; color: #34495e; margin-bottom: 15px; border-bottom: 1px solid #bdc3c7; padding-bottom: 5px; }
                                            .field-container { margin-bottom: 15px; padding: 10px; background-color: #f8f9fa; border-left: 4px solid #3498db; }
                                            .field-name { font-weight: bold; color: #2c3e50; margin-bottom: 5px; }
                                            .field-value { color: #555; word-wrap: break-word; }
                                            .signatures-section { margin-top: 30px; border-top: 2px solid #ecf0f1; padding-top: 20px; }
                                            .signature-item { margin-bottom: 12px; padding: 10px; border: 1px solid #ddd; background-color: #fff; border-radius: 4px; }
                                            .signature-name { font-weight: bold; display: inline-block; width: 250px; }
                                            .signature-status { font-size: 16px; font-weight: bold; }
                                            .signed { color: #27ae60; }
                                            .not-signed { color: #e74c3c; }
                                            .footer { margin-top: 40px; text-align: center; font-size: 10px; color: #95a5a6; border-top: 1px solid #ecf0f1; padding-top: 15px; }
                                            .user-info { font-size: 11px; color: #6c757d; margin-top: 10px; }
                                        </style>
                                    </head>
                                    <body>
                                        <div class="header">
                                            <div class="form-title">{{ title }}</div>
                                            <div class="export-info">Row {{ row_number }} - Generated on {{ export_date }}</div>
                                            <div class="user-info">Exported by: {{ user_info }}</div>
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
                                            {{ export_date }} | 
                                            Access Level: {{ user_role }}
                                        </div>
                                    </body>
                                    </html>
                                    """
                                    
                                    try:
                                        template = Template(individual_template)
                                        
                                        form_data_for_pdf = {field: entry_data.get(field, "") for field in config.get("fields", [])}
                                        signature_data_for_pdf = {signer: entry_data.get(signer, "❌ No") for signer in config.get("signatures", [])}
                                        
                                        rendered_html = template.render(
                                            title=config.get("title", pdf_sheet),
                                            row_number=row_num,
                                            export_date=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                            user_info=f"{st.session_state.get('current_user', 'Unknown')} ({st.session_state.get('user_role', 'Unknown')})",
                                            user_role=st.session_state.get('user_role', 'Unknown'),
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
            
            available_sheets = [name for name in get_all_sheet_names(GOOGLE_SHEET_ID) if name in form_configs]
            table_pdf_sheet = st.selectbox("📄 Select Sheet", available_sheets, key="pdf_table_sheet")
            
            if table_pdf_sheet:
                table_headers, table_records = get_sheet_data(table_pdf_sheet, GOOGLE_SHEET_ID)
                table_df = pd.DataFrame(table_records) if table_records else pd.DataFrame()
                
                if not table_df.empty:
                    sheet_config = form_configs.get(table_pdf_sheet, {})
                    signature_columns = sheet_config.get("signatures", [])
                    form_fields = sheet_config.get("fields", [])
                    
                    st.markdown("### 📋 Select Columns for PDF Export")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown("**📝 Form Fields:**")
                        available_form_fields = [col for col in table_df.columns if col in form_fields]
                        selected_form_fields = st.multiselect(
                            "Select Form Fields:", 
                            available_form_fields, 
                            default=available_form_fields[:5],
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
                            default=["Timestamp", "Submitted By"] if all(field in other_fields for field in ["Timestamp", "Submitted By"]) else ["Timestamp"] if "Timestamp" in other_fields else [],
                            key="pdf_other_fields"
                        )
                    
                    all_selected_columns = selected_form_fields + selected_signature_fields + selected_other_fields
                    
                    if all_selected_columns:
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
                                
                                filtered_df = table_df.iloc[rows_to_include][all_selected_columns]
                                
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
                                
                                table_template = """
                                <!DOCTYPE html>
                                <html>
                                <head>
                                    <meta charset="utf-8">
                                    <style>
                                        @page { size: {{ 'landscape' if orientation == 'Landscape' else 'portrait' }}; margin: 0.5in; }
                                        body { font-family: Arial, sans-serif; margin: 0; color: #333; font-size: 10px; }
                                        .header { text-align: center; margin-bottom: 20px; border-bottom: 2px solid #333; padding-bottom: 10px; }
                                        .form-title { font-size: 20px; font-weight: bold; margin-bottom: 5px; }
                                        .export-info { font-size: 11px; color: #666; margin-top: 5px; }
                                        .user-info { font-size: 10px; color: #888; margin-top: 5px; }
                                        .summary-section { margin-bottom: 20px; padding: 15px; background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 5px; }
                                        .summary-title { font-size: 14px; font-weight: bold; margin-bottom: 10px; color: #495057; }
                                        .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }
                                        .summary-item { font-size: 11px; padding: 8px; background: white; border-radius: 3px; }
                                        .summary-name { font-weight: bold; color: #495057; }
                                        .data-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 9px; }
                                        .data-table th { background-color: #343a40; color: white; border: 1px solid #495057; padding: 8px 4px; text-align: left; font-weight: bold; font-size: 9px; word-wrap: break-word; }
                                        .data-table td { border: 1px solid #dee2e6; padding: 6px 4px; text-align: left; word-wrap: break-word; font-size: 8px; max-width: 150px; overflow: hidden; }
                                        .row-num { background-color: #f8f9fa; font-weight: bold; text-align: center; min-width: 30px; max-width: 40px; }
                                        .signed { background-color: #d4edda; color: #155724; font-weight: bold; text-align: center; }
                                        .not-signed { background-color: #f8d7da; color: #721c24; text-align: center; }
                                        .column-header-form { background-color: #17a2b8 !important; }
                                        .column-header-signature { background-color: #28a745 !important; }
                                        .column-header-other { background-color: #6c757d !important; }
                                        .footer { margin-top: 20px; text-align: center; font-size: 9px; color: #6c757d; border-top: 1px solid #dee2e6; padding-top: 10px; }
                                    </style>
                                </head>
                                <body>
                                    <div class="header">
                                        <div class="form-title">{{ title }} - Enhanced Table View</div>
                                        <div class="export-info">Generated on {{ export_date }} | Records: {{ total_records }} | Columns: {{ column_count }}</div>
                                        <div class="user-info">Exported by: {{ user_info }} | Access Level: {{ user_role }}</div>
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
                                        Generated from IMS Form Entry System | Page Orientation: {{ orientation }} | Total Columns: {{ column_count }} | Total Rows: {{ total_records }} | Access Level: {{ user_role }}
                                    </div>
                                </body>
                                </html>
                                """
                                
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
                                    user_info=f"{st.session_state.get('current_user', 'Unknown')} ({st.session_state.get('user_role', 'Unknown')})",
                                    user_role=st.session_state.get('user_role', 'Unknown'),
                                    total_records=len(filtered_df),
                                    column_count=len(all_selected_columns),
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
                    
                    else:
                        st.warning("Please select at least one column for the PDF export.")

# SHEET MANAGEMENT TAB
if "sheet_management" in tab_mapping and tab_mapping["sheet_management"] is not None:
    with tab_mapping["sheet_management"]:
        st.markdown("## ⚙️ Sheet Management")
        
        current_sheets = get_all_sheet_names(GOOGLE_SHEET_ID)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### ➕ Create New Sheet")
            
            new_sheet_name = st.text_input("New Sheet Name", help="Enter a unique name for the new worksheet")
            
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
                headers = ["Timestamp"] + config["fields"] + config.get("signatures", []) + ["Submitted By"]
                st.info(f"Using template from '{template_form}' ({len(headers)} columns)")
                
                with st.expander("Preview Template Headers"):
                    st.write("**Form Fields:**")
                    for field in config["fields"]:
                        st.write(f"• {field}")
                    if config.get("signatures"):
                        st.write("**Signatures:**")
                        for sig in config["signatures"]:
                            st.write(f"• {sig}")
                    st.write("**System Fields:**")
                    st.write("• Timestamp")
                    st.write("• Submitted By")
            
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
                sheet_to_delete = st.selectbox(
                    "Select sheet to delete", 
                    current_sheets,
                    help="Select a worksheet to permanently delete"
                )
                
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
        
        st.markdown("---")
        st.markdown("### 📋 Current Worksheets Overview")
        
        if current_sheets:
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

# SIDEBAR STATUS - Updated access control display
st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Current Status")
st.sidebar.metric("Sheet Type", sheet_choice)
st.sidebar.metric("Available Forms", len(form_configs))
st.sidebar.metric("Total Worksheets", len(get_all_sheet_names(GOOGLE_SHEET_ID)))

st.sidebar.markdown("### 📈 API Usage")
current_calls = len(st.session_state.get('api_calls', []))
max_calls = quota_manager.max_calls
quota_percentage = (current_calls / max_calls) * 100

if quota_percentage > 80:
    st.sidebar.error(f"⚠️ API Usage: {current_calls}/{max_calls} ({quota_percentage:.1f}%)")
elif quota_percentage > 60:
    st.sidebar.warning(f"🔶 API Usage: {current_calls}/{max_calls} ({quota_percentage:.1f}%)")
else:
    st.sidebar.success(f"✅ API Usage: {current_calls}/{max_calls} ({quota_percentage:.1f}%)")

if quota_percentage > 90:
    wait_time = quota_manager.wait_time()
    if wait_time > 0:
        st.sidebar.info(f"⏱️ Cool-down: {wait_time}s remaining")

connection_status = get_gsheet_client(GOOGLE_SHEET_ID) is not None
st.sidebar.markdown("### 🔗 Connection Status")
if connection_status:
    st.sidebar.success("✅ Google Sheets Connected")
else:
    st.sidebar.error("❌ Connection Failed")

st.sidebar.markdown("### 🔐 Access Control")
user_role = st.session_state.get('user_role', 'Unknown')
if user_role == "Supervisor":
    st.sidebar.success("👑 Full Access")
elif user_role in ["Chief Security Officer", "Controlling Officer"]:
    st.sidebar.success("✍️ Read/Write/Export Access")
