import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials

# === 1. SETUP AUTHENTICATION ===
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name("imscredentials.json", scope)
client = gspread.authorize(creds)

# === 2. OPEN THE GOOGLE SHEET ===
sheet_name = "LW FILES"
spreadsheet = client.open(sheet_name)

# === 3. DEFINE SHEET STRUCTURES ===
sheet_definitions = {
    "GF A": ["Document Title", "Issue No", "Issue Date", "Rev. No", "Rev. Date", "Prepared By", "Approved By"],
    "GF B": ["Title of Document", "Nature of Document", "Issue Status", "Copy Number", "Issued To", "Issued By", "Date"],
    "GF C": ["Sl. No", "Subject", "Record No", "No of Pages"],
    "GF D": ["Sl. No", "Chapter No & Page No", "Old Issue", "New Issue", "Summary of Change", "Rev. No", "Date"],
    "LW4 01": ["Section", "Activity", "Impact", "ID", "Condition", "S", "O", "D", "Index", "Legal"],
    "LW4 02": ["Objective/Aspect/Hazard", "Action Planned", "Responsibility", "Present Status", "Target", "Revised Target", "Remarks", "Completion Date", "Signed by SSE"],
    "LW4 03": ["Date", "Time", "Location", "Nature of Incident", "Equipments Involved", "Observations", "NC Noticed", "Correction", "Corrective Action", "Signed by SSE", "Signed by CSO"],
    "LW4 04": ["Item", "DBR/Chelan No & Date", "Received Qty", "Used Qty", "Date of Consumption", "Balance Qty", "Signed by SSE", "Remarks"],
    "LW4 05": ["Date", "Type of Waste", "Qty Generated", "Qty Disposed", "DS8 No. & Date", "Date of Disposal", "Balance for Disposal", "Signed by SSE"],
    "LW4 06": ["Date", "Complaint/Hazard", "Affected Area/Person", "Apparent Cause", "Investigation Details", "Corrective Action", "Remarks", "Signed by SSE"],
    "LW4 07": ["Date of Incident", "Staff Involved", "WCA Status", "Incident Details", "Injury", "Loss Assessment", "Signed by SSE", "Signed by Controlling Officer", "Signed by CSO"],
    "LW4 10": ["Instrument Name", "ID No", "Range", "Frequency", "Least Count", "Acceptance Criteria", "Date of Calibration", "Due Date", "Calibrated By", "Error Identified", "Status", "Signed by SSE", "Remarks", "Signed by Officer"],
    "LW4 11": ["ID No", "Drawing No", "Description", "Purpose", "Frequency", "Acceptance Criteria", "Verification Date", "Due Date", "Parameter", "Actual", "Deviation", "Checked By", "Remarks", "Verified By", "Reviewed By"],
    "LW4 13": ["Audit ID", "Standard Ref", "Section", "NCR No", "Violation", "Non-Conformity", "Auditor", "Auditee", "Root Cause", "Correction", "Corrective Action", "Target Date", "Follow-up", "Signed by Officer"],
    "LW4 14": ["Employee Name", "PF No", "Ticket No", "Appointment Date", "Designation", "Qualification", "Course", "Duration", "Location", "Assessment", "Signed by SSE", "Remarks"],
    "LW4 15": ["Initiator Name", "Date", "Designation", "Shop", "Document Name", "Chapter", "Page No", "Revision No", "Details of Change", "Reason", "Signed by Initiator", "Signed by SSE", "Approved by Officer"],
    "LW4 16": ["Sl.No", "Source", "Nonconformity", "Corrective Action", "Target Date", "Implemented", "Actual Date", "Remarks", "Signed by SSE", "Signed by Officer"],
    "LW4 18": ["ID No", "Location", "Refilled Date", "Condition", "Next Due", "Signed by SSE"],
    "LW4 20": ["Loco/Coach", "Work Order", "Component", "Qty", "NC Observed", "Date", "Disposition", "Signed by JE", "Reinspection", "Signed by SSE", "Signed by Officer", "Remarks"],
    "LW4 21": ["NC/Complaint", "Ref No", "Date", "Root Cause", "Corrective Action", "Further Analysis", "Remarks", "Signed by SSE", "Signed by Officer"],
    "LW4 22": ["Sl.No", "Procedure Ref", "Area", "Deviation Details", "Period From", "To", "Signed by SSE", "Authorized Period", "Signed by Officer", "Closed Date", "Reasons"],
    "LW4 23": ["Material", "PL No", "Qty", "Urgency Reason", "Signed by SSE", "Approved by Officer", "Released by SMM", "Qty Released", "Material Identified", "Remarks"],
    "LW4 25": ["Customer Name", "Period", "Item", "Performance", "Rating", "Remarks", "Promptness", "Complaints", "Suggestions", "Signed by Customer"],
    "LW4 26": ["Complaint No", "Date", "Customer", "Complaint", "Investigation", "Observation", "Corrective Action", "Target Date", "Signed by Officer"],
    "LW4 27": ["Month", "Complaints Received", "Section", "Cumulative", "Corrective Done", "Pending", "Remarks", "Signed by SSE", "Signed by Officer"],
    "LW4 29": ["Component", "Drawing No", "Sub-Assembly", "Material PL", "Route", "WO No", "Qty per Loco", "Total Qty", "Stock Status", "Previous WO", "Approval by Officer"],
    "LW4 30": ["Date", "Time", "Machine No", "Unified Code", "Shop", "Description", "Out of Order Date", "Apparent Defect", "Initiated By", "Forwarded To"],
    "LW4 31": ["W.O. No", "Component Description", "Qty", "Material Req.", "RM Source", "Raw Available", "Part Supply", "Balance", "Customer", "Remarks"],
    "LW4 32": ["Date", "NC Description", "Area", "Detected By", "Impact", "Corrective Action", "Remarks", "Signed by SSE", "Signed by Officer"],
    "LW4 33": ["Record No", "Title", "Type", "Location", "Retention Period"],
    "LW435": ["Document Code", "Version", "Document Name", "Copy Holder", "Copy No", "Issue Details", "Rev Details", "Withdrawal Status"],
    "LW4 36": ["Employee Name", "Ticket No", "Shop", "Date/Time", "Age", "DOA", "Address", "Injury", "Activity at Accident", "Start Time", "First Aid", "Controlling Officer", "Accident Summary", "Witness 1/2", "Signatures"],
    "LW4 36/A": ["Employee Name", "Ticket No", "Sex", "Shop", "Date", "Service Years", "Injury", "Accident Description", "Cause", "Agency", "PPE Used", "PPE Issued", "Trained", "Transport", "Control", "Avoidable", "Suggestions", "Signed by CSO"],
    "LW 437": ["Shop", "Energy Source", "Annual Consumption", "Variables", "Deviation Criteria", "OCP No"],
    "LW 438": ["Energy Source", "2021-22", "2020-21", "2019-20"],
    "LW 439": ["Variable", "Units Consumed", "Remarks"],
    "LW 440": ["Substation", "Shop", "Device", "Rating", "Performance", "Improvement Opportunities"],
    "LW 441": ["Opportunity", "Energy Management Plan"],
    "LW 442": ["Variable", "Units Consumed", "Remarks"],
    "LW 443": ["Description", "Qty", "Shop"],
    "LW 444": ["Machine Name", "Shop", "Connected Load", "Consumption", "Date", "Action", "Owner"],
    "LW 445": ["Date", "Description", "Vehicle No", "Qty Issued", "Ticket No", "Signed by Emp"],
    "LW 446": ["Date", "Cylinder No", "Shop", "Company", "Returned Date", "Qty (Kgs)"],
    "LW 447": ["Date", "Cylinder No", "Shop", "Company", "Returned Date", "Qty (cu.m)"]
}

# === 4. CREATE SHEETS WITH DELAY ===
for tab_name, headers in sheet_definitions.items():
    try:
        worksheet = spreadsheet.add_worksheet(title=tab_name, rows="100", cols=str(len(headers)))
        worksheet.append_row(headers)
        print(f"✅ Created sheet: {tab_name}")
        time.sleep(1.5)  # Wait to avoid API rate limit
    except Exception as e:
        print(f"⚠️ Sheet {tab_name} may already exist or error: {e}")
