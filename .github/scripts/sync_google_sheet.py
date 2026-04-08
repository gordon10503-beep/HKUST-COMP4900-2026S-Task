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
    github_username = os.environ["GITHUB_USERNAME"]
    hash_value = os.environ["HASH_VALUE"]

    result = sheets.values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A:C",
    ).execute()
    rows = result.get("values", [])

    match_row = None
    for index, row in enumerate(rows[1:], start=2):
        if row and row[0] == pr_number:
            match_row = index
            break

    new_values = [[pr_number, github_username, hash_value]]

    if match_row:
        sheets.values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A{match_row}:C{match_row}",
            valueInputOption="RAW",
            body={"values": new_values},
        ).execute()
        print(f"Updated row {match_row} for PR #{pr_number}")
        return

    next_row = len(rows) + 1
    sheets.values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A{next_row}:C{next_row}",
        valueInputOption="RAW",
        body={"values": new_values},
    ).execute()
    print(f"Inserted PR #{pr_number} at row {next_row}")


if __name__ == "__main__":
    main()
