import os
import sys

# Ensure workspace root is in path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ['RUN_MIGRATION'] = '1'

from app import app, db
from models import Scan, Cylinder, CustomerMap

def test_delete_activity():
    with app.test_client() as client:
        # Mock session to bypass login and admin_required
        with client.session_transaction() as sess:
            sess['user'] = {'username': 'admin', 'role': 'manager', 'name': 'Admin'}

        import app as app_module
        app_module.async_sheets_write = lambda fn, *args, **kwargs: print("[Mock] async_sheets_write called")

        # Run within app context
        with app.app_context():
            # Setup database states
            # Create a cylinder
            cyl = Cylinder.query.filter_by(uid="TEST-CYL-99").first()
            if not cyl:
                cyl = Cylinder(uid="TEST-CYL-99", gas_type="Oxygen", status="Active", location="Depot")
                db.session.add(cyl)
                db.session.commit()
            
            # Reset cylinder state
            cyl.status = "Delivered"
            cyl.location = "Test Customer"
            db.session.commit()

            # Create scan log to delete
            scan1 = Scan(
                scan_date="26-06-2026",
                scan_time="12:00:00",
                driver="John Doe",
                action="Delivery",
                cylinder_uid="TEST-CYL-99",
                customer="Test Customer"
            )
            # Create another older scan log (for reverting status)
            scan0 = Scan(
                scan_date="25-06-2026",
                scan_time="10:00:00",
                driver="Jane Doe",
                action="Collection",
                cylinder_uid="TEST-CYL-99",
                customer=""
            )
            db.session.add(scan1)
            db.session.add(scan0)
            db.session.commit()

            # Verify current state before delete
            print("Before deletion:")
            print("Cylinder location:", cyl.location, "status:", cyl.status)
            print("Scan log count for TEST-CYL-99:", Scan.query.filter_by(cylinder_uid="TEST-CYL-99").count())

            # Perform deletion post request
            payload = {
                "date": "26-06-2026",
                "time": "12:00:00",
                "driver": "John Doe",
                "action": "Delivery",
                "customer": "Test Customer"
            }
            response = client.post('/admin/activity/delete', data=payload)
            print("POST Status Code:", response.status_code)

            # Re-fetch cylinder state
            db.session.refresh(cyl)
            print("After deletion:")
            print("Cylinder location:", cyl.location, "status:", cyl.status)
            print("Scan log count for TEST-CYL-99:", Scan.query.filter_by(cylinder_uid="TEST-CYL-99").count())
            
            # Clean up
            Scan.query.filter_by(cylinder_uid="TEST-CYL-99").delete()
            db.session.delete(cyl)
            db.session.commit()

if __name__ == '__main__':
    test_delete_activity()
