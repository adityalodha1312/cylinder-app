import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

try:
    creds = Credentials.from_service_account_file(
        "credentials.json",
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
    doc = client.open("Cylinder Tracking")
    sheet1 = doc.worksheet("Sheet1")
    rows = sheet1.get_all_values()
    print("Total rows in Sheet1:", len(rows))
    if len(rows) > 1:
        print("Sample rows in Sheet1:")
        for r in rows[:10]:
            print(r)
        
        # Let's inspect unique gas types, actions, etc.
        # Wait, what are the columns in Sheet1?
        # Column A: Date, B: Time, C: Driver, D: Action, E: Cylinder UID, F: Customer
        # Wait, where is the gas type in Sheet1?
        # Is there a gas type column in Sheet1?
        # In Sheet1 there is only UID. The Gas Type is linked from the registry.
        
except Exception as e:
    print("Error:", e)
