import os
import sys
import io

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ['RUN_MIGRATION'] = '1'

from app import app, db
from models import Cylinder

csv_data = """Sr. No.,Cylinder ID,Gas Type,Water Capacity,Invoice No.
1,TEST-10001,ARGON,46.7 LTR,INV-01
2,TEST-10002,ARGON,46.7 LTR,INV-02
3,TEST-10003,OXY,46.7 LTR,INV-03
"""

def test_upload():
    with app.test_client() as client:
        with app.app_context():
            # mock session for admin role
            with client.session_transaction() as sess:
                sess['user'] = {'username': 'admin', 'role': 'manager', 'name': 'Admin'}
            
            data = {
                'file': (io.BytesIO(csv_data.encode('utf-8')), 'test_upload.csv')
            }
            response = client.post('/admin/cylinders/upload', data=data, content_type='multipart/form-data', follow_redirects=True)
            
            print(f"Response status: {response.status_code}")
            print(f"Data in response: {'Successfully added' in response.get_data(as_text=True)}")
            
            cyls = Cylinder.query.filter(Cylinder.uid.like('TEST-%')).all()
            print(f"Inserted cylinders count: {len(cyls)}")
            for c in cyls:
                print(f" - {c.uid}, {c.gas_type}, {c.cylinder_type}")

if __name__ == '__main__':
    test_upload()
