import os
import sys
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

# Ensure we can import app modules
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

load_dotenv()

from app import app, db
from models import Customer

def run_test():
    print("Starting database-first read/write verification test...")
    
    # 1. Initialize Google Sheets client to check sheets content directly
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(
        "credentials.json",
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
    doc = client.open("Cylinder Tracking")
    customer_ws = doc.worksheet("Customers")

    test_customer_name = "Test DB First Customer"
    test_email = "dbfirst@test.com"
    test_phone = "1234567890"
    test_address = "Database Street"

    with app.app_context():
        # Check if customer already exists in DB/Sheets and clean up
        existing_cust = Customer.query.filter_by(name=test_customer_name).first()
        if existing_cust:
            print(f"Cleaning up existing customer '{test_customer_name}' from DB...")
            db.session.delete(existing_cust)
            db.session.commit()

        # Clean up from Sheets
        rows = customer_ws.get_all_values()
        for idx, r in enumerate(rows):
            if idx == 0: continue
            if len(r) > 1 and r[1].strip() == test_customer_name:
                print(f"Cleaning up existing customer '{test_customer_name}' from Sheets row {idx+1}...")
                customer_ws.update(f'B{idx+1}:E{idx+1}', [["", "", "", ""]])

        # 2. Test Customer Adding via database first
        # We will mimic the admin_customers_add route behavior
        all_info = []
        customers = Customer.query.all()
        for c in customers:
            all_info.append({'id': c.customer_id, 'name': c.name})

        max_id = 0
        for c in all_info:
            id_str = c.get('id', '')
            if id_str.startswith('C'):
                try:
                    val = int(id_str[1:])
                    if val > max_id: max_id = val
                except ValueError: pass
        new_id = f"C{str(max_id + 1).zfill(3)}"
        print(f"Generated new ID: {new_id}")

        # Write to PostgreSQL first
        print("Writing customer to PostgreSQL DB...")
        cust_db = Customer(
            customer_id=new_id,
            name=test_customer_name,
            email=test_email,
            phone=test_phone,
            address=test_address
        )
        db.session.add(cust_db)
        db.session.commit()
        print("Successfully committed to DB.")

        # Mirror to Google Sheets
        print("Mirroring customer to Google Sheets...")
        all_rows = customer_ws.get_all_values()
        empty_row_idx = None
        for idx, r in enumerate(all_rows):
            if idx == 0: continue
            if len(r) < 2 or not r[1].strip():
                empty_row_idx = idx + 1
                break

        if empty_row_idx:
            existing_row = all_rows[empty_row_idx - 1]
            existing_id = existing_row[0].strip() if len(existing_row) > 0 else ""
            new_id_sheets = existing_id if existing_id else new_id
            customer_ws.update(f'A{empty_row_idx}:E{empty_row_idx}', [[new_id_sheets, test_customer_name, test_email, test_phone, test_address]])
            sheets_row = empty_row_idx
            print(f"Updated empty row {empty_row_idx} in Sheets.")
        else:
            customer_ws.append_row([new_id, test_customer_name, test_email, test_phone, test_address])
            sheets_row = len(all_rows) + 1
            print(f"Appended new row {sheets_row} to Sheets.")

        # Verify DB read
        db_cust = Customer.query.filter_by(customer_id=new_id).first()
        assert db_cust is not None, "Error: Customer not found in DB!"
        assert db_cust.name == test_customer_name, "Error: Customer name mismatch in DB!"
        print("Verification: Customer successfully verified in DB.")

        # Verify Sheets read
        sheets_rows = customer_ws.get_all_values()
        sheet_cust_row = sheets_rows[sheets_row - 1]
        assert sheet_cust_row[1].strip() == test_customer_name, "Error: Customer name mismatch in Sheets!"
        assert sheet_cust_row[2].strip() == test_email, "Error: Customer email mismatch in Sheets!"
        print("Verification: Customer successfully verified in Google Sheets.")

        # 3. Test Customer Deleting via database first
        print("Deleting customer from PostgreSQL DB...")
        db.session.delete(db_cust)
        db.session.commit()
        print("Deleted from DB.")

        # Mirror delete to Sheets
        print("Mirroring deletion to Google Sheets...")
        customer_ws.update(f'B{sheets_row}:E{sheets_row}', [["", "", "", ""]])
        print("Cleared row in Sheets.")

        # Verify DB delete
        db_cust_deleted = Customer.query.filter_by(customer_id=new_id).first()
        assert db_cust_deleted is None, "Error: Customer still exists in DB after deletion!"
        print("Verification: Deletion successfully verified in DB.")

        # Verify Sheets delete
        sheets_rows_after = customer_ws.get_all_values()
        sheet_cust_row_after = sheets_rows_after[sheets_row - 1]
        assert sheet_cust_row_after[1].strip() == "", "Error: Customer still exists in Sheets after deletion!"
        print("Verification: Deletion successfully verified in Google Sheets.")

        print("\nAll DB-first and mirroring tests PASSED successfully!")

if __name__ == "__main__":
    run_test()
