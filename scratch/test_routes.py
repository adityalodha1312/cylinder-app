import os
import sys

# Ensure workspace root is in path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ['RUN_MIGRATION'] = '1'

from app import app, db

def test_pages():
    with app.test_client() as client:
        # Mock session to bypass admin_required decorator
        with client.session_transaction() as sess:
            sess['user'] = {'username': 'admin', 'role': 'manager', 'name': 'Admin'}

        print("Testing admin routes...")

        routes = [
            '/admin/cylinders',
            '/admin/activity',
            '/admin/rotation',
            '/admin/products'
        ]

        for r in routes:
            print(f"\n--- GET {r} ---")
            try:
                response = client.get(r)
                print("Status:", response.status_code)
                if response.status_code != 200:
                    print("ERROR Response:")
                    print(response.data.decode('utf-8')[:1500])
            except Exception as e:
                import traceback
                print("EXCEPTION:")
                traceback.print_exc()

if __name__ == '__main__':
    test_pages()
