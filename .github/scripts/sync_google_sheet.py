import json
import os

import google.oauth2.service_account
from googleapiclient.discovery import build


def main() -> None:
    key = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_KEY"])
    creds = google.oauth2.service_account.Credentials.from_service_account_info(
        key,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    service = build("sheets", "v4", credentials=creds)
    sheets = service.spreadsheets()

    spreadsheet_id = os.environ["SPREADSHEET_ID"]
    sheet_name = os.getenv("SHEET_NAME", "Sheet1")
    pr_number = os.environ["PR_NUMBER"]
    student_name = os.environ["STUDENT_NAME"]
    github_username = os.environ["GITHUB_USERNAME"]
    itsc_email = os.environ["ITSC_EMAIL"]
    student_id = os.environ["STUDENT_ID"]

    result = sheets.values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A:E",
    ).execute()
    rows = result.get("values", [])

    match_row = None
    for index, row in enumerate(rows[1:], start=2):
        if row and row[0] == pr_number:
            match_row = index
            break

    new_values = [[pr_number, itsc_email, student_id, student_name, github_username]]

    if match_row:
        sheets.values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A{match_row}:E{match_row}",
            valueInputOption="RAW",
            body={"values": new_values},
        ).execute()
        print(f"Updated row {match_row} for PR #{pr_number}")
        return

    next_row = len(rows) + 1
    sheets.values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A{next_row}:E{next_row}",
        valueInputOption="RAW",
        body={"values": new_values},
    ).execute()
    print(f"Inserted PR #{pr_number} at row {next_row}")


if __name__ == "__main__":
    main()
