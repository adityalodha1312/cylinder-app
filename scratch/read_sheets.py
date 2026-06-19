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
    print("Worksheets:")
    for ws in doc.worksheets():
        print(f"- {ws.title}")
        
    print("\nReading sample from Customers:")
    cust_ws = doc.worksheet("Customers")
    print(cust_ws.get_all_values())
    
    try:
        cyl_ws = doc.worksheet("Cylinders")
        print("\nReading sample from Cylinders:")
        print(cyl_ws.get_all_values()[:10])
    except Exception as e:
        print("\nNo Cylinders worksheet yet:", e)
        
    try:
        maint_ws = doc.worksheet("Cylinder Maintenance")
        print("\nReading sample from Cylinder Maintenance:")
        print(maint_ws.get_all_values()[:10])
    except Exception as e:
        print("\nNo Cylinder Maintenance worksheet yet:", e)

except Exception as e:
    print("Error:", e)
