import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Setup
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name("imscredentials.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open("LW FILES")

updated_sheets = []

# Loop through all sheets
for worksheet in spreadsheet.worksheets():
    try:
        headers = worksheet.row_values(1)
        if "Signed by Officer" in headers:
            index = headers.index("Signed by Officer") + 1  # gspread is 1-indexed
            worksheet.update_cell(1, index, "Signed by Controlling Officer")
            updated_sheets.append(worksheet.title)
    except Exception as e:
        print(f"Error in {worksheet.title}: {e}")

# Report
if updated_sheets:
    print("✅ Updated sheets:")
    for name in updated_sheets:
        print(f" - {name}")
else:
    print("ℹ️ No changes needed. All sheets are already correct.")
