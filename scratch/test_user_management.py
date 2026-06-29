import os
import sys

# Ensure workspace root is in path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ['RUN_MIGRATION'] = '1'

from app import app, db
from models import User

def test_user_management():
    with app.test_client() as client:
        # Mock session to bypass login and admin_required
        with client.session_transaction() as sess:
            sess['user'] = {'username': 'admin', 'role': 'manager', 'name': 'Admin'}

        # Mock async_sheets_write
        import app as app_module
        app_module.async_sheets_write = lambda fn, *args, **kwargs: print("[Mock] async_sheets_write called!")

        with app.app_context():
            # Clean up existing test user if present
            User.query.filter_by(username="test_driver").delete()
            db.session.commit()

            print("--- Testing List Users ---")
            res = client.get('/admin/users')
            print("GET /admin/users Status:", res.status_code)

            print("\n--- Testing Add User ---")
            payload = {
                "name": "Test Driver User",
                "username": "test_driver",
                "password": "driverpassword",
                "role": "driver"
            }
            res = client.post('/admin/users/add', data=payload)
            print("POST /admin/users/add Status:", res.status_code)
            
            # Check user was created
            created_user = User.query.filter_by(username="test_driver").first()
            if created_user:
                print("Successfully created test user in database.")
                user_id = created_user.id
            else:
                print("ERROR: User was not created.")
                return

            print("\n--- Testing Edit User ---")
            edit_payload = {
                "name": "Test Driver User Updated",
                "username": "test_driver_up",
                "password": "", # blank password (no update)
                "role": "manager"
            }
            res = client.post(f'/admin/users/edit/{user_id}', data=edit_payload)
            print("POST /admin/users/edit Status:", res.status_code)

            db.session.refresh(created_user)
            print("Updated User Name:", created_user.name)
            print("Updated User Username:", created_user.username)
            print("Updated User Role:", created_user.role)

            print("\n--- Testing Delete User ---")
            res = client.post(f'/admin/users/delete/{user_id}')
            print("POST /admin/users/delete Status:", res.status_code)

            # Check deleted
            deleted_user = User.query.get(user_id)
            if not deleted_user:
                print("Successfully deleted test user from database.")
            else:
                print("ERROR: User was not deleted.")

if __name__ == '__main__':
    test_user_management()
