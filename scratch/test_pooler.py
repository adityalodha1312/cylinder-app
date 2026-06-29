import psycopg2

pooler_url = "postgresql://postgres.spktqthqawmagwxqpybc:nobleairgas%40123@aws-0-ap-south-1.pooler.supabase.com:6543/postgres"
print("Testing connection to pooler URL...")

try:
    conn = psycopg2.connect(pooler_url)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM customers LIMIT 1;")
    row = cursor.fetchone()
    print("Pooler connection SUCCESS! Sample customer:", row)
    cursor.close()
    conn.close()
except Exception as e:
    print("Pooler connection FAILED:", e)
