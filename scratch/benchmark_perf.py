import os
import sys
import time
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

# Ensure we can import app modules
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

load_dotenv()

from app import app, db
from models import Customer

def benchmark():
    print("Initializing Google Sheets client...")
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

    print("\n--- Google Sheets Read Performance ---")
    sheets_times = []
    for i in range(5):
        start = time.time()
        # Fetching all values mimics what the sheets-only code did
        values = customer_ws.get_all_values()
        duration = time.time() - start
        sheets_times.append(duration)
        print(f"Sheets read {i+1}: {duration:.4f} seconds (returned {len(values)} rows)")

    avg_sheets = sum(sheets_times) / len(sheets_times)
    print(f"Average Google Sheets Read Time: {avg_sheets:.4f} seconds")

    print("\n--- Supabase PostgreSQL Read Performance ---")
    db_times = []
    with app.app_context():
        # Clear SQLAlchemy query cache by recreating the query session if needed
        # (Running multiple queries to get a realistic production average)
        for i in range(5):
            start = time.time()
            customers = Customer.query.all()
            duration = time.time() - start
            db_times.append(duration)
            print(f"PostgreSQL read {i+1}: {duration:.4f} seconds (returned {len(customers)} rows)")

    avg_db = sum(db_times) / len(db_times)
    print(f"Average PostgreSQL Read Time: {avg_db:.4f} seconds")

    speedup = avg_sheets / avg_db if avg_db > 0 else 0
    print(f"\n==================================================")
    print(f"PostgreSQL is {speedup:.1f}x FASTER than Google Sheets!")
    print(f"==================================================")

if __name__ == "__main__":
    benchmark()
