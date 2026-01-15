from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from GetKmdAcessToken import GetKMDToken
import requests
import os
import uuid
from datetime import datetime, timedelta
import json
import pandas as pd
import re
orchestrator_connection = OrchestratorConnection("Henter Assets", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
# ---- Henter assests og credentials -----
KMDNovaURL = orchestrator_connection.get_constant("KMDNovaURL").value

  # ---- Henter access tokens ----
KMD_access_token = GetKMDToken()



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
    
    # Iterate through task list
    for case in data.get("cases", []):
        CaseUuid = case.get("uuid", "Unknown")
        print(CaseUuid)
        

        Cases.append({
            "CaseUuid": CaseUuid  # Assuming case_number provides a unique reference
        })


except requests.exceptions.RequestException as e:
    print("Request Failed:", e)