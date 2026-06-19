import os
import sys
from dotenv import load_dotenv

# Ensure we can import app modules
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Set broken DATABASE_URL to simulate a connection error
os.environ['DATABASE_URL'] = "postgresql://postgres:wrongpassword@localhost:5432/nonexistentdb"

from app import app, get_all_customer_info

def run_fallback_test():
    print("Starting database fallback verification test...")
    print(f"DATABASE_URL in env: {os.environ.get('DATABASE_URL')}")
    
    with app.app_context():
        # Call get_all_customer_info, which should fail database query and fall back to Google Sheets
        print("Calling get_all_customer_info() with broken database connection...")
        customers = get_all_customer_info()
        print(f"Found {len(customers)} customers from Sheets fallback:")
        for c in customers:
            print(f"- {c['id']}: {c['name']} ({c['email']})")
            
        assert len(customers) > 0, "Error: Fallback returned no customers from Sheets!"
        print("\nFallback verification test PASSED successfully!")

if __name__ == "__main__":
    run_fallback_test()
