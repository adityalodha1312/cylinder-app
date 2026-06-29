import os
from dotenv import load_dotenv
from psycopg2 import connect

load_dotenv()
db_url = os.environ.get('DATABASE_URL')
print("DATABASE_URL:", db_url)

if not db_url:
    print("Error: DATABASE_URL not set")
    exit(1)

try:
    conn = connect(db_url)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM scans ORDER BY id DESC LIMIT 10;")
    rows = cursor.fetchall()
    print("Last 10 scans in Supabase DB:")
    for r in rows:
        print(r)
    cursor.close()
    conn.close()
except Exception as e:
    print("Connection error:", e)
