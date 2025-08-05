import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# === 1. Google Sheets Setup ===
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name("imscredentials.json", scope)
client = gspread.authorize(creds)

# === 2. Open M&PR Google Sheet ===
mpr_sheet_id = "17KL-cKMJNGncAJngla_eX6aVdzvoihbdpgOJqOiGycc"
spreadsheet = client.open_by_key(mpr_sheet_id)

# === 3. Load M&PR Config ===
with open("forms_mpr_configs.json", "r", encoding="utf-8") as f:
    form_configs = json.load(f)

# === 4. Create Sheets with Correct Headers ===
created = []
skipped = []

for sheet_name, config in form_configs.items():
    try:
        if sheet_name not in [ws.title for ws in spreadsheet.worksheets()]:
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="30")
            headers = config.get("fields", []) + config.get("signatures", [])
            worksheet.append_row(headers)
            created.append(sheet_name)
        else:
            skipped.append(sheet_name)
    except Exception as e:
        skipped.append(f"{sheet_name} (Error: {str(e)})")

print("✅ Created sheets:", created)
print("⏩ Skipped sheets (already exist or errored):", skipped)
