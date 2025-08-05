import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name("imscredentials.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open("LW FILES")

signature_columns_by_sheet = {}

for worksheet in spreadsheet.worksheets():
    try:
        headers = worksheet.row_values(1)
        signature_cols = [col for col in headers if "sign" in col.lower()]
        if signature_cols:
            signature_columns_by_sheet[worksheet.title] = signature_cols
    except Exception as e:
        print(f"Error reading {worksheet.title}: {e}")

# Display result
print("Signature Fields Detected:")
for sheet, columns in signature_columns_by_sheet.items():
    print(f"- {sheet}: {columns}")
