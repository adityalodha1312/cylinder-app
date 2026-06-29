import os
import sys
import io
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ['RUN_MIGRATION'] = '1'

from app import app, db
from models import Cylinder

def test_upload_excel():
    df = pd.DataFrame({
        'Sr. No.': [1, 2],
        'Cylinder ID': ['TEST-EXCEL-1', 'TEST-EXCEL-2'],
        'Gas Type': ['ARGON', 'OXY'],
        'Water Capacity': ['46.7 LTR', '46.7 LTR']
    })
    
    excel_buffer = io.BytesIO()
    df.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)
    
    with app.test_client() as client:
        with app.app_context():
            # mock session for admin role
            with client.session_transaction() as sess:
                sess['user'] = {'username': 'admin', 'role': 'manager', 'name': 'Admin'}
            
            data = {
                'file': (excel_buffer, 'test_upload.xlsx')
            }
            try:
                response = client.post('/admin/cylinders/upload', data=data, content_type='multipart/form-data', follow_redirects=True)
                
                print(f"Response status: {response.status_code}")
                print(f"Response text contains Successfully added: {'Successfully added' in response.get_data(as_text=True)}")
                print(f"Response text contains Error: {'Error' in response.get_data(as_text=True)}")
                if response.status_code == 500:
                    print(response.get_data(as_text=True))
            except Exception as e:
                import traceback
                traceback.print_exc()

if __name__ == '__main__':
    test_upload_excel()
