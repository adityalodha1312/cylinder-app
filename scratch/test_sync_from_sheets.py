import os
import sys

# Ensure workspace root is in path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ['RUN_MIGRATION'] = '1'

from app import app, db

def test_sync_from_sheets():
    with app.test_client() as client:
        # Mock session to bypass login and admin_required
        with client.session_transaction() as sess:
            sess['user'] = {'username': 'admin', 'role': 'manager', 'name': 'Admin'}

        # Mock sync_sheets_to_db to avoid making real network requests during testing
        import app as app_module
        import sync
        sync.sync_sheets_to_db = lambda doc: (print("[Mock] sync_sheets_to_db called!"), True)[1]

        print("Testing POST /admin/sync_from_sheets...")
        response = client.post('/admin/sync_from_sheets')
        print("Status Code:", response.status_code)
        print("Redirect Location:", response.location)

if __name__ == '__main__':
    test_sync_from_sheets()
