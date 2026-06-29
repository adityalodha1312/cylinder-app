import os
import sys

# Ensure workspace root is in path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ['RUN_MIGRATION'] = '1'

from app import app, db
from models import Cylinder

def test_submit():
    # Setup test client
    with app.test_client() as client:
        # Mock session
        with client.session_transaction() as sess:
            sess['user'] = {'username': 'admin', 'role': 'manager', 'name': 'Admin'}
        
        # We need to make sure we don't hit validation error
        # Let's see if we have some cylinders in the DB or if we can mock them
        # Let's just create a test request payload
        payload = {
            "driver": "Test Driver",
            "customer": "Test Customer",
            "scans": [
                {"uid": "TEST-UID-001", "action": "Delivery"},
                {"uid": "TEST-UID-002", "action": "Collection"}
            ]
        }
        
        print("Sending POST /submit request...")
        try:
            # We can mock the sheets write using a mock or monkeypatch if we want,
            # but let's see what happens if we just run it.
            # To avoid writing to Google Sheets in our test, we can mock the functions
            # sheets_write_with_retry and async_sheets_write.
            import app as app_module
            app_module.async_sheets_write = lambda fn, *args, **kwargs: print("[Mock] async_sheets_write called")
            
            # Run within app context
            with app.app_context():
                # Ensure cylinders exist to bypass validation if needed
                # Wait, validation: "Check if we are collecting a cylinder that is already in stock/empty at the Depot"
                # For collection, if it doesn't exist in the DB, it doesn't fail.
                # Let's check:
                # c_db = Cylinder.query.filter(Cylinder.uid.ilike(scan_uid)).first()
                # if c_db: if c_db.status in ['Empty', 'Filled'] ...
                # If c_db is None, it doesn't fail. So using "TEST-UID-002" is fine.
                
                # Mock database session commit to not actually write to database during test if we want,
                # or we can let it write to sqlite memory / dev db.
                # Let's run it.
                response = client.post('/submit', json=payload)
                print("Status Code:", response.status_code)
                print("Response Data:", response.data.decode('utf-8'))
        except Exception as e:
            import traceback
            traceback.print_exc()

if __name__ == '__main__':
    test_submit()
