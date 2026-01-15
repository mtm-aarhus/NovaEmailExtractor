from GetKmdAcessToken import GetKMDToken
import requests
import os
import uuid
from datetime import datetime, timedelta
import json
import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
orchestrator_connection = OrchestratorConnection("AktbobGenererAktindsigter", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None,None)
# ---- Henter assests og credentials -----
KMDNovaURL = orchestrator_connection.get_constant("KMDNovaURL").value

  # ---- Henter access tokens ----
KMD_access_token = GetKMDToken(orchestrator_connection)



# ---- Henter Sagsnummer og Sagsbeskrivelse ---- 
TransactionID = str(uuid.uuid4())
CurrentDate = datetime.now().strftime("%Y-%m-%dT00:00:00")
FromDate = (datetime.now() - timedelta(days=14)).strftime("%Y-%m-%dT00:00:00")


payload = {
    "common": {"transactionId": TransactionID},
    "paging": {"startRow": 1, "numberOfRows": 500},
"buildingCase":{
    "buildingCaseAttributes":{
        "buildingCaseClassName":"Forhåndsdialog",
        "applicationStatusDates":{
            "fromCloseDate": FromDate ,
            "toCloseDate": CurrentDate
        }
}
},
"state": {
    "progressState": "Afsluttet"
},
  "caseGetOutput": {
    "caseAttributes": {
      "userFriendlyCaseNumber": True,
      "title": True,
    },
    "buildingCase": {
      "buildingCaseAttributes": {
        "buildingCaseClassName": True,
        "buildingCaseClassId": True,
        "bbrCaseId": True,
  }}}}



# Define headers
headers = {
    "Authorization": f"Bearer {KMD_access_token}",
    "Content-Type": "application/json"
}

# Define the API endpoint
url = f"{KMDNovaURL}/Case/GetList?api-version=2.0-Case"
# Make the HTTP request
try:
    response = requests.put(url, headers=headers, json=payload)
    response.raise_for_status()  # Raise an error for non-2xx responses
    data = response.json()
    print("Success:", response.status_code)
    
    # Extract and print number of rows
    number_of_rows = data.get("pagingInformation", {}).get("numberOfRows", 0)
    print("Number of Rows:", number_of_rows)

        # Initialize an empty list to store queue items
    Cases = []
    # Iterate through cases
    for case in data.get("cases", []):
        case_uuid = case.get("common", {}).get("uuid", "Unknown")
        case_number = case.get("caseAttributes", {}).get("userFriendlyCaseNumber", "Unknown")
        case_title = case.get("caseAttributes", {}).get("title")

        if case_uuid:
            Cases.append({
                "CaseUuid": case_uuid,
                "CaseNumber": case_number,
                "CaseTitle": case_title
            })

except requests.exceptions.RequestException as e:
    print("Request Failed:", e)


EMAIL_REGEX = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
all_emails = []

# Now iterate through the casesUuid's and get the emails: 

for case in Cases:
    case_uuid = case["CaseUuid"]
    transaction_id = str(uuid.uuid4())

    payload_case_party = {
        "common": {
            "transactionId": transaction_id,
            "uuid": case_uuid
        },
        "paging": {
            "startRow": 1,
            "numberOfRows": 500
        },
        "caseParty": {
            "partyRole": "IND",
            "partyRoleName": "Indsender"
        },
        "caseGetOutput": {
            "caseAttributes": {
                "userFriendlyCaseNumber": True,
                "title": True,
                "caseDate": True
            },
            "caseParty": {
                "partyRole": True,
                "partyRoleName": True,
                "name": True,
                "participantRole": True,
                "participantRemark": True,
                "participantContactInformation": True
            }
        }
    }

    try:
        response = requests.put(url, headers=headers, json=payload_case_party)
        response.raise_for_status()
        data = response.json()

        for case_item in data.get("cases", []):
            for party in case_item.get("caseParties", []):
                if party.get("partyRole") == "IND":
                    contact_info = party.get("participantContactInformation", "")

                    emails = EMAIL_REGEX.findall(contact_info)

                    if emails:
                                email = emails[0]  # <-- take only the first email

                                print(f"CaseNumber: {case.get('CaseNumber')} | Email: {email}")

                                all_emails.append({
                                    "CaseUuid": case_uuid,
                                    "CaseNumber": case.get("CaseNumber"),
                                    "Email": email
                                })

                                email_found = True
                                break  # stop iterating parties

                    if email_found:
                        break  # stop iterating cases (defensive)

    except requests.exceptions.RequestException as e:
        print(f"Failed for case {case.get('CaseNumber')} ({case_uuid}): {e}")

# Create workbook and worksheet
wb = Workbook()
ws = wb.active
ws.title = "Indsender emails"

# Styles
header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
header_alignment = Alignment(horizontal="center", vertical="center")
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

# Headers
headers = ["Case number", "Email"]

for col, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_alignment
    cell.border = thin_border

# Data rows
for row_idx, item in enumerate(all_emails, start=2):
    ws.cell(row=row_idx, column=1, value=item["CaseNumber"]).border = thin_border
    ws.cell(row=row_idx, column=2, value=item["Email"]).border = thin_border

# Column widths
ws.column_dimensions["A"].width = 20
ws.column_dimensions["B"].width = 35

# Freeze header row
ws.freeze_panes = "A2"

# Save file
output_path = "Indsender_emails.xlsx"
wb.save(output_path)

print(f"Excel file created: {output_path}")


