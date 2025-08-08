import gspread
from oauth2client.service_account import ServiceAccountCredentials

# === Setup ===
SHEET_ID = "1wxntHZp4xEQWCmLAt2TVF8ohG6uHuvV_3QbaK7wSwGw"
CREDENTIALS_FILE = "imscredentials.json"

# === Mapping: Old Sheet Name => New Sheet Name ===
RENAME_MAP = {
    "LW 437": "LW4 37",
    "LW 438": "LW4 38",
    "LW 439": "LW4 39",
    "LW 440": "LW4 40",
    "LW 441": "LW4 41",
    "LW 442": "LW4 42",
    "LW 443": "LW4 43",
    "LW 444": "LW4 44",
    "LW 445": "LW4 45",
    "LW 446": "LW4 46",
    "LW 447": "LW4 47",
    "LW4 36/A": "LW4 36A"
}

def rename_worksheets():
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)

        sheet = client.open_by_key(SHEET_ID)
        worksheets = sheet.worksheets()
        worksheet_titles = [ws.title for ws in worksheets]

        print("🔍 Checking worksheets...")
        for old_name, new_name in RENAME_MAP.items():
            if old_name in worksheet_titles:
                ws = sheet.worksheet(old_name)
                ws.update_title(new_name)
                print(f"✅ Renamed: '{old_name}' → '{new_name}'")
            else:
                print(f"❌ Skipped: '{old_name}' not found")

        print("\n🎉 Rename process completed.")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    rename_worksheets()
