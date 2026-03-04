"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
from GetKmdAcessToken import GetKMDToken
import requests
import os
import uuid
from datetime import datetime, timedelta
import json
import pandas as pd
import re
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import time
from requests.auth import HTTPBasicAuth

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    
    orchestrator_connection.log_trace("Running process.")

    # ---- Run only on odd ISO weeks (ulige uger) ----
    current_week = datetime.now().isocalendar().week

    if current_week % 2 == 0:
        orchestrator_connection.log_info(
            f"Skipping run - week {current_week} is an even week."
        )
        return

    orchestrator_connection.log_info(
        f"Running - week {current_week} is an odd week."
    )

    # ---- Henter assests og credentials -----
    KMDNovaURL = orchestrator_connection.get_constant("KMDNovaURL").value
    SurveyXact = orchestrator_connection.get_credential("SurveyXact")
    SurveyXact_password = SurveyXact.password
    SurveyXact_unsername = SurveyXact.username

    # ---- Henter access tokens ----
    KMD_access_token = GetKMDToken(orchestrator_connection)


    # ---- Henter Sagsnummer og Sagsbeskrivelse ---- 
    TransactionID = str(uuid.uuid4())
    EndDate = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%dT00:00:00")
    #FromDate = (datetime.now() - timedelta(days=14)).strftime("%Y-%m-%dT00:00:00")
    FromDate = orchestrator_connection.get_constant('NovaEmailExtrator_Timestamp').value
    print(FromDate)

    payload = {
        "common": {"transactionId": TransactionID},
        "paging": {"startRow": 1, "numberOfRows": 500},
    "state": {
        "progressState": "Afsluttet"
    },
    "states": {
    "startFromDate": FromDate,
    "endFromDate": EndDate,
    "states": [
        {
        "progressState": "Afsluttet"
        }
    ]
    },

    "buildingCase":{
        "buildingCaseAttributes":{
            "buildingCaseClassName":"Forhåndsdialog"
    }
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

    REMOVE_TITLE_REGEX = re.compile(
        r"(Fejloprettet|Afsluttet\s+mangler\s+fuldmagt)",
        re.IGNORECASE
    )

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

        for case in data.get("cases", []):
            case_uuid = case.get("common", {}).get("uuid")
            case_number = case.get("caseAttributes", {}).get("userFriendlyCaseNumber", "Unknown")
            case_title = case.get("caseAttributes", {}).get("title", "")

            # Skip unwanted titles (case-insensitive)
            if case_title and REMOVE_TITLE_REGEX.search(case_title):
                print(f"Removing casenumber {case_number} - because of title: {case_title}")
                continue

            if case_uuid:
                Cases.append({
                    "CaseUuid": case_uuid,
                    "CaseNumber": case_number,
                    "CaseTitle": case_title
                })
    except requests.exceptions.RequestException as e:
        print("Request Failed:", e)


    EMAIL_REGEX = re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b")
    all_emails = []

    # Now iterate through the casesUuid's and get the emails: 

    for case in Cases:
        case_uuid = case["CaseUuid"]
        transaction_id = str(uuid.uuid4())
        email_found = False  
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
                        IndsenderName = party.get("name")
                        contact_info = re.sub(r"\[n\]|\n|\r", " ", contact_info)
                        emails = EMAIL_REGEX.findall(contact_info)

                        if emails:
                            email = emails[0]  # take only the first email
            
                            orchestrator_connection.log_info(f"CaseNumber: {case.get('CaseNumber')} | Email: {email}| Name: {IndsenderName} | Titel: {case.get("CaseTitle")}")

                            all_emails.append({
                                "CaseUuid": case_uuid,
                                "CaseNumber": case.get("CaseNumber"),
                                "CaseTitle": case.get("CaseTitle"),
                                "Email": email,
                                "Name":IndsenderName
                            })

                            email_found = True
                            break  # stop iterating parties

                if email_found:
                    break  # stop iterating cases (defensive)

            # THIS goes here — outside both loops
            if not email_found:
                print(f"CaseNumber: {case.get('CaseNumber')} | Email NOT FOUND")

        except requests.exceptions.RequestException as e:
            print(f"Failed for case {case.get('CaseNumber')} ({case_uuid}): {e}")

    print(f"\nTotal emails collected: {len(all_emails)}")
    # ---- TEST EMAILS (override production list) ----
    
    # ----- Opretter timestamps ------ 
    current_ts = int(time.time())      # current epoch time in seconds
    ts_plus_1_week = current_ts + 7 * 24 * 60 * 60
    distributionTs = 1


    # --- endpoint ---
    survey_url = "https://rest.survey-xact.dk/rest/surveys/1792115/respondents"

    params = {
        "distributionTs": distributionTs,
        "reminder1Ts": ts_plus_1_week
    }

    for item in all_emails:
        email = item["Email"]
        Name = item["Name"]
        Case_Number = item["CaseNumber"]
        Case_Title = item["CaseTitle"]
        payload = {
            "email": email,
            "b_1": Name,
            "b_2": Case_Number,
            "b_3": Case_Title
        }

        try:
            response = requests.post(
                survey_url,
                params=params,
                data=payload,
                auth=HTTPBasicAuth(SurveyXact_unsername, SurveyXact_password),
                timeout=10
            )

            if response.status_code in (200, 201):
                print(f" Survey sent to {email} (Case {Case_Number})")
            else:
                print(f"Failed for {email} | Status: {response.status_code}")
                print(response.text)

        except requests.exceptions.RequestException as e:
            print(f"Error sending survey to {email}: {e}")

    #Opdaterer timestamp: 
    orchestrator_connection.update_constant("NovaEmailExtrator_Timestamp",datetime.now().strftime("%Y-%m-%dT%H:%M:%S"))