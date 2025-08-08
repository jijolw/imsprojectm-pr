import streamlit as st
st.set_page_config(page_title="IMS Form Entry", layout="wide")

import gspread, json, pandas as pd, time, hashlib, functools, random, re, unicodedata
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from io import BytesIO
from jinja2 import Template
from xhtml2pdf import pisa
from gspread.utils import rowcol_to_a1

# =========================
# GLOBAL STYLE (beautify)
# =========================
st.markdown("""
<style>
/* softer base */
:root { --card-bg: #ffffff; --muted:#6b7280; --accent:#4f46e5; --border:#e5e7eb; }
.block-container{padding-top:1.2rem; padding-bottom:2rem;}
/* page header */
.page-header{background:linear-gradient(135deg,#4f46e5 0%,#7c3aed 100%); color:white; padding:18px 20px; border-radius:14px; margin-bottom:12px}
.page-header h1{margin:0;font-size:1.4rem}
.page-header .meta{opacity:.9; font-size:.9rem}
/* card */
.card{background:var(--card-bg); border:1px solid var(--border); border-radius:14px; padding:16px; margin-bottom:14px}
.card h3{margin-top:0}
.badge{display:inline-block; padding:2px 8px; border-radius:999px; font-size:.8rem; border:1px solid rgba(255,255,255,.4); background:rgba(255,255,255,.2); color:white}
.kv{color:var(--muted); font-size:.9rem}
.section-title{font-weight:700; margin:.2rem 0 .8rem 0}
hr{border:none; border-top:1px solid var(--border); margin:1rem 0}
</style>
<div class="page-header">
  <h1>📋 IMS Form Entry System</h1>
  <div class="meta">Fast • Clean • Quieter API usage</div>
</div>
""", unsafe_allow_html=True)

# =========================
# DEBUG MODE (bypass login)
# =========================
DEBUG_MODE = str(st.secrets.get("IMS_DEBUG_MODE", "false")).lower() == "true"

# =========================
# LOGIN SYSTEM
# =========================
DEFAULT_USERS = {
    "cso": {"password": "cso@2024", "role": "Chief Security Officer", "permissions": ["read", "write", "export"]},
    "supervisor": {"password": "super@2024", "role": "Supervisor", "permissions": ["read", "write", "export", "manage"]},
    "controlling_officer": {"password": "control@2024", "role": "Controlling Officer", "permissions": ["read", "write", "export"]},
}

def hash_password(password): return hashlib.sha256(password.encode()).hexdigest()
def verify_password(password, hashed): return hash_password(password) == hashed

def initialize_users():
    if 'users_initialized' not in st.session_state:
        if "IMS_USERS" in st.secrets:
            st.session_state.users = json.loads(st.secrets["IMS_USERS"])
        else:
            st.session_state.users = {u: {"password_hash": hash_password(d["password"]), "role": d["role"], "permissions": d["permissions"]} for u,d in DEFAULT_USERS.items()}
        st.session_state.users_initialized = True

def set_debug_session():
    st.session_state.logged_in = True
    st.session_state.current_user = "debug"
    st.session_state.user_role = "Supervisor"
    st.session_state.user_permissions = ["read", "write", "export", "manage"]
    st.session_state.login_time = datetime.now()

def login_form():
    st.markdown('<div class="card"><h3>🔐 Secure Login</h3>', unsafe_allow_html=True)
    if DEBUG_MODE:
        st.warning("DEBUG MODE is ON — login is bypassed.")
        if st.button("Continue in Debug Mode", type="primary", use_container_width=True):
            set_debug_session(); st.rerun()
        st.stop()
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.form_submit_button("🚀 Login", type="primary", use_container_width=True):
            if authenticate_user(username, password):
                st.success("✅ Login successful!"); st.balloons(); time.sleep(0.6); st.rerun()
            else:
                st.error("❌ Invalid username or password!")
    st.markdown('</div>', unsafe_allow_html=True)

def authenticate_user(username, password):
    initialize_users()
    user = st.session_state.users.get(username)
    if user and verify_password(password, user["password_hash"]):
        st.session_state.logged_in = True
        st.session_state.current_user = username
        st.session_state.user_role = user["role"]
        st.session_state.user_permissions = user["permissions"]
        st.session_state.login_time = datetime.now()
        return True
    return False

def logout():
    for k in ['logged_in','current_user','user_role','user_permissions','login_time']:
        st.session_state.pop(k, None)
    st.rerun()

def check_permission(p): return st.session_state.get('logged_in', False) and (p in st.session_state.get('user_permissions', []))
def require_permission(p):
    if not check_permission(p):
        st.error(f"❌ Access Denied: You need '{p}' permission."); st.stop()

# =========================
# LOGIN CHECK
# =========================
initialize_users()
if DEBUG_MODE and not st.session_state.get('logged_in', False): set_debug_session()
if not st.session_state.get('logged_in', False): login_form(); st.stop()

# Sidebar (compact, no extra API pings)
with st.sidebar:
    st.markdown("### 👤 Current User")
    st.success(f"{st.session_state.get('user_role','Unknown')}")
    st.caption(f"Logged in as: **{st.session_state.get('current_user','Unknown')}**")
    if st.button("🚪 Logout", use_container_width=True): logout()

# =========================
# CONFIGURATION
# =========================
require_permission("read")

sheet_choice = st.sidebar.selectbox("Choose File Type", ["LW FILES", "M&PR FILES"])
SHEET_IDS = {
    "LW FILES": "1wxntHZp4xEQWCmLAt2TVF8ohG6uHuvV_3QbaK7wSwGw",
    "M&PR FILES": "17KL-cKMJNGncAJngla_eX6aVdzvoihbdpgOJqOiGycc",
}
GOOGLE_SHEET_ID = SHEET_IDS[sheet_choice]

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

if "IMS_CREDENTIALS_JSON" in st.secrets:
    with open("temp_creds.json","w") as f: json.dump(json.loads(st.secrets["IMS_CREDENTIALS_JSON"]), f)
    CREDENTIAL_FILE = "temp_creds.json"
else:
    CREDENTIAL_FILE = "imscredentials.json"

# Keep UI state (prevents jumping)
if "active_section" not in st.session_state:
    st.session_state.active_section = "📝 Form Entry" if check_permission("write") else "📊 Data View"

SECTION_ORDER = []
if check_permission("write"): SECTION_ORDER.append("📝 Form Entry")
SECTION_ORDER += ["📊 Data View", "🧪 Diagnostics"]
if check_permission("export"): SECTION_ORDER.append("📄 PDF Export")
if check_permission("manage"): SECTION_ORDER.append("⚙️ Sheet Management")

st.session_state.active_section = st.radio(
    "Go to", SECTION_ORDER,
    index=SECTION_ORDER.index(st.session_state.active_section),
    horizontal=True, label_visibility="collapsed", key="nav_section_radio"
)

def stay_on(section_name: str):
    st.session_state.active_section = section_name

# =========================
# Backoff + Quota
# =========================
def with_backoff(max_retries=4, base=1.0, mult=2.0, max_sleep=6.0):
    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*a, **k):
            attempt = 0
            while True:
                try: return fn(*a, **k)
                except Exception as e:
                    msg = str(e)
                    if any(t in msg for t in ["429","Quota exceeded","rate limit"]):
                        if attempt >= max_retries: raise
                        sleep = min(max_sleep, base*(mult**attempt)) + random.uniform(0,0.3)
                        st.toast(f"⏳ Google API busy. Retrying in {sleep:.1f}s…", icon="⏳")
                        time.sleep(sleep); attempt += 1; continue
                    raise
        return wrapper
    return deco

class APIQuotaManager:
    def __init__(self, max_calls_per_minute=40):
        self.max_calls = max_calls_per_minute
        st.session_state.setdefault('api_calls', [])
    def can_make_call(self):
        now = datetime.now()
        st.session_state.api_calls = [t for t in st.session_state.api_calls if (now - t).seconds < 60]
        return len(st.session_state.api_calls) < self.max_calls
    def record_call(self): st.session_state.api_calls.append(datetime.now())
    def wait_time(self):
        if not st.session_state.api_calls: return 0
        oldest = min(st.session_state.api_calls); return max(0, 60 - (datetime.now() - oldest).seconds)

quota_manager = APIQuotaManager()

def api_rate_limit():
    if not quota_manager.can_make_call():
        wt = quota_manager.wait_time(); st.info(f"⏱️ Waiting {wt}s (quota)…"); time.sleep(wt)
    quota_manager.record_call()
    last = st.session_state.get('last_api_call', 0); now = time.time()
    if now - last < 0.8: time.sleep(0.8 - (now - last))
    st.session_state['last_api_call'] = time.time()

# =========================
# Google Sheets helpers
# =========================
@st.cache_resource
@with_backoff()
def get_gsheet_client(sheet_id):
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIAL_FILE, SCOPE)
    return gspread.authorize(creds).open_by_key(sheet_id)

@st.cache_data(ttl=900, show_spinner=False)
@with_backoff()
def list_worksheets(sheet_id):  # fewer calls, longer cache
    return [ws.title for ws in get_gsheet_client(sheet_id).worksheets()]

def get_all_sheet_names(sheet_id):
    try: return list_worksheets(sheet_id)
    except Exception as e: st.error(f"Error fetching sheet names: {e}"); return []

@st.cache_data(ttl=900, show_spinner=False)  # longer cache to reduce API reads
@with_backoff()
def fetch_sheet_all_values(sheet_name, sheet_id):
    return get_gsheet_client(sheet_id).worksheet(sheet_name).get_all_values()

def get_sheet_data(sheet_name, sheet_id):
    if not quota_manager.can_make_call():
        st.info(f"⏱️ Please wait {quota_manager.wait_time()}s…"); return [], []
    try:
        quota_manager.record_call()
        all_values = fetch_sheet_all_values(sheet_name, sheet_id)
        if not all_values: return [], []
        headers = all_values[0]
        # build dict rows without extra processing
        if len(all_values) > 1:
            width = len(headers)
            records = [dict(zip(headers, row + [""]*(width-len(row)))) for row in all_values[1:]]
        else:
            records = []
        return headers, records
    except Exception as e:
        st.error(f"Error fetching data from '{sheet_name}': {e}")
        return [], []

@st.cache_data(ttl=3600, show_spinner=False)
def load_form_configs_for_sheet(sheet_type):
    try:
        cfg_file = "form_configs.json" if sheet_type == "LW FILES" else "forms_mpr_configs.json"
        with open(cfg_file, "r", encoding="utf-8") as f: return json.load(f)
    except FileNotFoundError:
        st.error(f"Config file not found for {sheet_type}"); return {}
    except json.JSONDecodeError as e:
        st.error(f"Error parsing config: {e}"); return {}

# =========================
# Header Normalization
# =========================
def _to_ascii_equiv(s: str) -> str:
    repl = {'\u2018':"'", '\u2019':"'", '\u201C':'"', '\u201D':'"', '\u2013':'-', '\u2014':'-', '\u00A0':' '}
    return ''.join(repl.get(ch, ch) for ch in s)

def normalize_header(h: str) -> str:
    if not isinstance(h, str): return ""
    s = _to_ascii_equiv(h.strip())
    s = unicodedata.normalize('NFKC', s)
    s = re.sub(r'\s+', ' ', s)
    return s.lower()

def expected_headers_from_config(cfg: dict):
    return list(cfg.get("fields", [])) + list(cfg.get("signatures", []))

def build_header_mapping(actual_headers, expected_headers):
    norm_to_actual, problems = {}, []
    for a in actual_headers:
        na = normalize_header(a)
        if na in norm_to_actual and norm_to_actual[na] != a:
            problems.append(f"Collision after normalization: '{norm_to_actual[na]}' vs '{a}'")
        norm_to_actual[na] = a
    norm_expected = [normalize_header(e) for e in expected_headers]
    only_after_norm = []
    for e_raw, e_norm in zip(expected_headers, norm_expected):
        if e_norm in norm_to_actual and e_raw != norm_to_actual[e_norm]:
            only_after_norm.append((e_raw, norm_to_actual[e_norm]))
    if only_after_norm:
        problems.append("Matched after normalization: " + ", ".join([f"'{a}' ↔ '{b}'" for a,b in only_after_norm]))
    return norm_to_actual, norm_expected, problems

def diff_config_vs_sheet(config, headers):
    expected = expected_headers_from_config(config)
    headers = headers or []
    norm_to_actual, norm_expected, problems = build_header_mapping(headers, expected)
    norm_actual_order = [normalize_header(h) for h in headers]
    missing_norm = [ne for ne in norm_expected if ne not in norm_to_actual]
    extra_norm   = [na for na in norm_actual_order if na not in norm_expected]
    missing_display = [expected[norm_expected.index(ne)] for ne in missing_norm]
    extra_display = [next((h for h in headers if normalize_header(h) == na), na) for na in extra_norm]
    order_match = (norm_actual_order == norm_expected)
    return {
        "missing_in_sheet": missing_display,
        "extra_in_sheet": extra_display,
        "order_match": order_match,
        "expected": expected,
        "actual": headers,
        "notes": problems,
    }

# =========================
# DATA (load once per selection)
# =========================
form_configs = load_form_configs_for_sheet(sheet_choice)
all_sheet_names = get_all_sheet_names(GOOGLE_SHEET_ID)

# =========================
# RENDER SECTIONS (radio instead of tabs)
# =========================
render_form = st.session_state.active_section == "📝 Form Entry"
render_view = st.session_state.active_section == "📊 Data View"
render_diag = st.session_state.active_section == "🧪 Diagnostics"
render_pdf  = st.session_state.active_section == "📄 PDF Export"
render_mgmt = st.session_state.active_section == "⚙️ Sheet Management"

# -------------------------
# FORM ENTRY
# -------------------------
if render_form:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">✏️ Form Entry & Edit</div>', unsafe_allow_html=True)

    selected_form = st.selectbox(
        "Select a form",
        list(form_configs.keys()),
        format_func=lambda x: f"{x} — {form_configs[x]['title']}"
    )

    form_cfg = form_configs[selected_form]
    headers, records = get_sheet_data(selected_form, GOOGLE_SHEET_ID)
    df = pd.DataFrame(records) if records else pd.DataFrame()

    # On-demand sheet match check (no auto spam)
    if st.button("✅ Check Sheet Match", key="check_match_form", on_click=stay_on, args=("📝 Form Entry",)):
        res = diff_config_vs_sheet(form_cfg, headers)
        if res["missing_in_sheet"] or res["extra_in_sheet"] or not res["order_match"]:
            st.error("Header mismatch between CONFIG and SHEET.")
            if res["notes"]: st.caption("Notes: " + "; ".join(res["notes"]))
            with st.expander("Details", expanded=False):
                st.write("Expected:", res["expected"])
                st.write("Actual:", res["actual"])
                st.write("Missing:", res["missing_in_sheet"])
                st.write("Extra:", res["extra_in_sheet"])
        else:
            st.success("Perfect match ✅")

    st.markdown(f"**{form_cfg['title']}**")
    edit_mode = st.checkbox("Enable Edit Mode")
    selected_row_index, prefill_data = None, {}

    if edit_mode and not df.empty:
        st.info("Select a row to edit.")
        df_display = df.copy(); df_display.index += 2
        st.dataframe(df_display, use_container_width=True, hide_index=False)
        row_num = st.number_input("Enter row number to edit:", min_value=2, max_value=len(df)+1, step=1)
        selected_row_index = row_num - 2
        if 0 <= selected_row_index < len(df):
            st.success(f"Editing row {row_num}")
            prefill_data = df.iloc[selected_row_index].to_dict()

    # Build form controls
    form_values, signature_values = {}, {}
    with st.form("entry_form"):
        fields, sigs = form_cfg.get("fields", []), form_cfg.get("signatures", [])
        cols_per_row = 2
        for i in range(0, len(fields), cols_per_row):
            cols = st.columns(cols_per_row)
            for j in range(cols_per_row):
                idx = i + j
                if idx < len(fields):
                    field = fields[idx]
                    curr = prefill_data.get(field, "")
                    with cols[j]:
                        if (len(str(curr)) > 100 or any(k in field.lower() for k in ['description','details','notes','remarks','comment','address','specification','procedure'])):
                            form_values[field] = st.text_area(field, value=str(curr), height=120)
                        else:
                            form_values[field] = st.text_input(field, value=str(curr))

        if sigs:
            st.markdown("---"); st.markdown("### ✍️ Signatures")
            sig_cols = st.columns(min(4, len(sigs)))
            for i, signer in enumerate(sigs):
                default_checked = (str(prefill_data.get(signer, "")).strip().lower() in ["✔️ yes","✔️","yes"])
                with sig_cols[i % len(sig_cols)]:
                    signature_values[signer] = st.checkbox(signer, value=default_checked)

        submitted = st.form_submit_button("💾 Submit Entry", use_container_width=True, type="primary")

    if submitted:
        try:
            api_rate_limit()
            # Build payload (fields + signatures)
            payload = {}
            for f in form_cfg.get("fields", []): payload[f] = form_values.get(f, "")
            for s in form_cfg.get("signatures", []): payload[s] = "✔️ Yes" if signature_values.get(s, False) else "❌ No"

            # Map by normalized names to actual headers
            expected = expected_headers_from_config(form_cfg)
            nmap, n_expected, _ = build_header_mapping(headers, expected)
            payload_norm = {normalize_header(k): v for k,v in payload.items()}

            row = []
            for h in headers:
                nh = normalize_header(h)
                val = payload_norm.get(nh, "")
                row.append("" if val is None else str(val))

            ws = get_gsheet_client(GOOGLE_SHEET_ID).worksheet(selected_form)
            if edit_mode and selected_row_index is not None:
                start = rowcol_to_a1(selected_row_index + 2, 1)
                end = rowcol_to_a1(selected_row_index + 2, len(headers))
                ws.update(f"{start}:{end}", [row])
                st.success(f"✅ Row {selected_row_index + 2} updated.")
            else:
                ws.append_row(row)
                st.success("✅ New entry submitted.")
            st.cache_data.clear(); st.rerun()
        except Exception as e:
            st.error(f"❌ Error writing to Google Sheet: {e}")

    st.markdown('</div>', unsafe_allow_html=True)

# -------------------------
# DATA VIEW
# -------------------------
if render_view:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📊 Data View</div>', unsafe_allow_html=True)

    available_sheets = [n for n in all_sheet_names if n in form_configs]
    view_sheet = st.selectbox("Select sheet to view", available_sheets, key="view_sheet_select")

    col1, col2 = st.columns([1,2])
    with col1:
        if st.button("🔄 Refresh Data", on_click=stay_on, args=("📊 Data View",)):
            st.cache_data.clear(); st.rerun()
    with col2:
        if st.button("✅ Check Sheet Match", key="check_match_view", on_click=stay_on, args=("📊 Data View",)):
            vh, _ = get_sheet_data(view_sheet, GOOGLE_SHEET_ID)
            cfg = form_configs.get(view_sheet, {})
            resv = diff_config_vs_sheet(cfg, vh)
            if resv["missing_in_sheet"] or resv["extra_in_sheet"] or not resv["order_match"]:
                st.error("Header mismatch between CONFIG and SHEET.")
            else:
                st.success("Perfect match ✅")

    if view_sheet:
        view_headers, view_records = get_sheet_data(view_sheet, GOOGLE_SHEET_ID)
        view_df = pd.DataFrame(view_records) if view_records else pd.DataFrame()
        if view_df.empty:
            st.warning("No data available.")
        else:
            st.caption(f"Rows: {len(view_df)} • Cols: {len(view_df.columns)} • Updated: {datetime.now().strftime('%H:%M:%S')}")
            q = st.text_input("🔍 Search…")
            filtered_df = view_df if not q else view_df[view_df.astype(str).apply(lambda x: x.str.contains(q, case=False, na=False)).any(axis=1)]
            st.dataframe(filtered_df, use_container_width=True, height=520)

    st.markdown('</div>', unsafe_allow_html=True)

# -------------------------
# PDF EXPORT
# -------------------------
if render_pdf:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📝 PDF Export</div>', unsafe_allow_html=True)

    available = [n for n in all_sheet_names if n in form_configs]
    pdf_sheet = st.selectbox("Select Sheet", available, key="pdf_sheet")
    if pdf_sheet:
        headers, records = get_sheet_data(pdf_sheet, GOOGLE_SHEET_ID)
        df = pd.DataFrame(records) if records else pd.DataFrame()
        cfg = form_configs.get(pdf_sheet, {})
        if not df.empty:
            st.markdown("**Individual PDFs**")
            rows = st.multiselect("Select Row(s)", df.index + 2, key="pdf_rows")
            if rows:
                for rownum in rows:
                    idx = rownum - 2
                    entry = df.iloc[idx].to_dict()
                    if st.button(f"📄 Generate PDF for Row {rownum}", key=f"gen_{rownum}", on_click=stay_on, args=("📄 PDF Export",)):
                        html = """
                        <html><head><meta charset="utf-8">
                        <style>
                        body{font-family:Arial;margin:20px;color:#333;line-height:1.6}
                        .h{border-bottom:2px solid #333;padding-bottom:6px;text-align:center;margin-bottom:12px}
                        .t{font-size:20px;font-weight:bold}
                        .s{font-size:14px;font-weight:bold;margin-top:10px}
                        .y{color:#155724;font-weight:bold}.n{color:#721c24;font-weight:bold}
                        </style></head>
                        <body>
                          <div class="h"><div class="t">{{ title }}</div><div>Row {{ row }} • {{ dt }}</div></div>
                          <div class="s">📝 Form Data</div>
                          {% for k,v in form_data.items() %}<div><b>{{ k }}</b>: {{ v if v else '-' }}</div>{% endfor %}
                          {% if signatures %}
                            <div class="s">✍️ Signatures</div>
                            {% for k,v in signatures.items() %}
                              <div><b>{{ k }}</b>: <span class="{{ 'y' if ('✔️' in v or v|lower=='yes') else 'n' }}">{{ v }}</span></div>
                            {% endfor %}
                          {% endif %}
                        </body></html>
                        """
                        pdf_buffer = BytesIO()
                        rendered = Template(html).render(
                            title=cfg.get("title", pdf_sheet),
                            row=rownum, dt=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            form_data={f: entry.get(f,"") for f in cfg.get("fields", [])},
                            signatures={s: entry.get(s,"❌ No") for s in cfg.get("signatures", [])}
                        )
                        pisa_status = pisa.CreatePDF(BytesIO(rendered.encode("utf-8")), dest=pdf_buffer)
                        if not pisa_status.err:
                            pdf_buffer.seek(0)
                            st.success("✅ PDF ready")
                            st.download_button(
                                f"⬇️ Download Row {rownum} PDF",
                                data=pdf_buffer.getvalue(),
                                file_name=f"{pdf_sheet}_Row_{rownum}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                                mime="application/pdf",
                                key=f"dl_{rownum}",
                                on_click=stay_on,
                                args=("📄 PDF Export",)
                            )
                        else:
                            st.error(f"❌ PDF generation error: {pisa_status.err}")

            st.markdown("---")
            st.markdown("**Table PDF**")
            sig_cols, fields = cfg.get("signatures", []), cfg.get("fields", [])
            colA, colB, colC = st.columns(3)
            with colA:
                avail_fields = [c for c in df.columns if c in fields]
                sel_fields = st.multiselect("📝 Form Fields", avail_fields, default=avail_fields[:5], key="p_fields")
            with colB:
                avail_sigs = [c for c in df.columns if c in sig_cols]
                sel_sigs = st.multiselect("✍️ Signature Fields", avail_sigs, default=avail_sigs, key="p_sigs")
            with colC:
                others = [c for c in df.columns if c not in fields and c not in sig_cols]
                sel_others = st.multiselect("⏰ Other Fields", others, default=[], key="p_other")

            cols = sel_fields + sel_sigs + sel_others
            if cols:
                mode = st.radio("Rows", ["All", "Select"], horizontal=True, key="row_mode")
                rows = list(range(len(df))) if mode=="All" else [r-2 for r in st.multiselect("Pick rows (sheet row numbers):", df.index + 2, key="p_rows")]
                orient = st.selectbox("Page Orientation", ["Landscape","Portrait"], key="page_orient")
                include_summary = st.checkbox("Include Signature Summary", value=True, key="incl_sum")
                if rows and st.button("📊 Generate Table PDF", on_click=stay_on, args=("📄 PDF Export",)):
                    try:
                        filtered = df.iloc[rows][cols]
                        summary = {}
                        if include_summary:
                            for s in sel_sigs:
                                if s in filtered.columns:
                                    total = len(filtered)
                                    s_yes = filtered[s].astype(str).str.contains('✔️|Yes', case=False, na=False).sum()
                                    s_no = filtered[s].astype(str).str.contains('❌|No', case=False, na=False).sum()
                                    summary[s] = {"signed": s_yes, "not_signed": s_no, "pct": (s_yes/total*100) if total else 0}

                        html = """
                        <html><head><meta charset="utf-8">
                        <style>
                        @page { size: {{ 'landscape' if orientation=='Landscape' else 'portrait' }}; margin:0.5in; }
                        body{font-family:Arial;margin:0;color:#333;font-size:10px}
                        .h{text-align:center;margin:10px 0;border-bottom:2px solid #333;padding-bottom:6px}
                        .t{font-size:18px;font-weight:bold}
                        .sum{margin:10px;padding:8px;background:#f8f9fa;border:1px solid #dee2e6}
                        table{width:100%;border-collapse:collapse;margin-top:6px}
                        th{background:#343a40;color:#fff;border:1px solid #495057;padding:6px 4px;font-size:9px;text-align:left}
                        td{border:1px solid #dee2e6;padding:6px 4px;font-size:8px}
                        .signed{background:#d4edda;color:#155724;font-weight:bold;text-align:center}
                        .not-signed{background:#f8d7da;color:#721c24;text-align:center}
                        </style></head>
                        <body>
                          <div class="h"><div class="t">{{ title }} - Table Export</div>
                          <div>Generated {{ dt }} • Records: {{ n }} • Columns: {{ m }}</div></div>
                          {% if include_summary and summary %}
                            <div class="sum"><b>Signature Summary</b><br/>
                              {% for k,v in summary.items() %}{{ k }}: ✔️ {{ v.signed }} ({{ "%.1f"|format(v.pct) }}%) • ❌ {{ v.not_signed }}<br/>{% endfor %}
                            </div>
                          {% endif %}
                          <table><thead><tr>{% for c in cols %}<th>{{ c }}</th>{% endfor %}</tr></thead>
                          <tbody>
                            {% for row in data %}
                              <tr>
                                {% for c in cols %}
                                  {% set cell = row.get(c, '') %}
                                  {% if c in sigs %}
                                    <td class="{{ 'signed' if ('✔️' in (cell|string) or (cell|string)|lower=='yes') else 'not-signed' if ('❌' in (cell|string) or (cell|string)|lower=='no') else '' }}">{{ cell if cell else '-' }}</td>
                                  {% else %}
                                    <td>{{ cell if cell else '-' }}</td>
                                  {% endif %}
                                {% endfor %}
                              </tr>
                            {% endfor %}
                          </tbody></table>
                        </body></html>
                        """
                        pdf_buffer = BytesIO()
                        rendered = Template(html).render(
                            title=cfg.get("title", pdf_sheet),
                            dt=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            n=len(filtered), m=len(cols), cols=cols, sigs=sel_sigs,
                            include_summary=include_summary, summary=summary,
                            data=filtered.to_dict(orient="records"),
                            orientation=orient
                        )
                        pisa_status = pisa.CreatePDF(BytesIO(rendered.encode("utf-8")), dest=pdf_buffer)
                        if not pisa_status.err:
                            pdf_buffer.seek(0)
                            st.success("✅ Table PDF ready")
                            st.download_button(
                                "⬇️ Download Table PDF",
                                data=pdf_buffer.getvalue(),
                                file_name=f"{pdf_sheet}_Table_{len(filtered)}rows_{len(cols)}cols_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                                mime="application/pdf",
                                key="dl_table_pdf",
                                on_click=stay_on,
                                args=("📄 PDF Export",)
                            )
                        else:
                            st.error(f"❌ PDF generation error: {pisa_status.err}")
                    except Exception as e:
                        st.error(f"❌ Error generating table PDF: {e}")

    st.markdown('</div>', unsafe_allow_html=True)

# -------------------------
# SHEET MANAGEMENT
# -------------------------
if render_mgmt:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">⚙️ Sheet Management</div>', unsafe_allow_html=True)

    current = all_sheet_names
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**➕ Create New Sheet**")
        new_name = st.text_input("New Sheet Name")
        template_form = st.selectbox("Use template from:", ["Custom Headers"] + list(form_configs.keys()))
        if template_form == "Custom Headers":
            headers_input = st.text_area("Headers (comma-separated)")
            headers_new = [h.strip() for h in headers_input.split(",") if h.strip()] if headers_input else []
        else:
            cfg_tmp = form_configs[template_form]
            headers_new = expected_headers_from_config(cfg_tmp)
            with st.expander("Preview Template Headers", expanded=False):
                st.write("Fields:"); [st.write("• "+f) for f in cfg_tmp.get("fields", [])]
                st.write("Signatures:"); [st.write("• "+s) for s in cfg_tmp.get("signatures", [])]
        if st.button("✅ Create New Sheet", type="primary") and new_name and headers_new:
            if new_name not in current:
                if create_new_worksheet(new_name, headers_new): st.balloons(); st.rerun()
            else:
                st.error("❌ Sheet name already exists!")

    with col2:
        st.markdown("**🗑️ Delete Existing Sheet**")
        if current:
            to_del = st.selectbox("Select sheet to delete", current)
            if to_del:
                h, r = get_sheet_data(to_del, GOOGLE_SHEET_ID)
                st.caption(f"Rows: {len(r) if r else 0} • Columns: {len(h) if h else 0} • Has Config: {'✅' if to_del in form_configs else '❌'}")
            st.error("⚠️ This action cannot be undone.")
            if st.checkbox(f"I understand and want to delete '{to_del}'") and st.button("🗑️ Delete Sheet", type="secondary"):
                if delete_worksheet(to_del): st.success("Sheet deleted."); st.rerun()
        else:
            st.info("No worksheets available.")

    st.markdown('</div>', unsafe_allow_html=True)

# -------------------------
# DIAGNOSTICS (quiet)
# -------------------------
if render_diag:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🧪 Diagnostics (on-demand)</div>', unsafe_allow_html=True)

    selectable = [s for s in all_sheet_names if s in form_configs]
    diag_sheet = st.selectbox("Select sheet", selectable)
    if st.button("Run Check", on_click=stay_on, args=("🧪 Diagnostics",)):
        h, _ = get_sheet_data(diag_sheet, GOOGLE_SHEET_ID)
        cfg = form_configs.get(diag_sheet, {})
        r = diff_config_vs_sheet(cfg, h)
        if r["missing_in_sheet"] or r["extra_in_sheet"] or not r["order_match"]:
            st.error("Mismatch found.")
            with st.expander("Details"):
                st.write("Expected:", r["expected"])
                st.write("Actual:", r["actual"])
                st.write("Missing:", r["missing_in_sheet"])
                st.write("Extra:", r["extra_in_sheet"])
                if r["notes"]: st.caption("Notes: " + "; ".join(r["notes"]))
        else:
            st.success("Perfect match ✅")

    if st.button("Scan ALL", on_click=stay_on, args=("🧪 Diagnostics",)):
        rows = []
        for name in selectable:
            hh, _ = get_sheet_data(name, GOOGLE_SHEET_ID)
            cfg = form_configs.get(name, {})
            r = diff_config_vs_sheet(cfg, hh)
            rows.append({
                "Sheet": name,
                "Missing": len(r["missing_in_sheet"]),
                "Extra": len(r["extra_in_sheet"]),
                "Order OK": "Yes" if r["order_match"] else "No",
            })
        if rows:
            st.dataframe(pd.DataFrame(rows).sort_values(["Missing","Extra","Order OK"], ascending=[False, False, True]),
                         use_container_width=True, hide_index=True)
        else:
            st.info("No configured sheets found.")

    st.markdown('</div>', unsafe_allow_html=True)

# -------------------------
# SIDEBAR STATUS (keep)
# -------------------------
st.sidebar.markdown("---")
st.sidebar.metric("Sheet Type", sheet_choice)
st.sidebar.metric("Available Forms", len(form_configs))
st.sidebar.metric("Total Worksheets", len(all_sheet_names))
current_calls = len(st.session_state.get('api_calls', []))
qp = (current_calls / quota_manager.max_calls) * 100 if quota_manager.max_calls else 0
if qp > 80: st.sidebar.error(f"⚠️ API Usage: {current_calls}/{quota_manager.max_calls} ({qp:.1f}%)")
elif qp > 60: st.sidebar.warning(f"🔶 API Usage: {current_calls}/{quota_manager.max_calls} ({qp:.1f}%)")
else: st.sidebar.success(f"✅ API Usage: {current_calls}/{quota_manager.max_calls} ({qp:.1f}%)")
