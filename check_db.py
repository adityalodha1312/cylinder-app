import os
from sqlalchemy import create_engine, inspect, text
from dotenv import load_dotenv

load_dotenv()
db_url = os.environ.get('DATABASE_URL')
engine = create_engine(db_url)

with engine.connect() as conn:
    result = conn.execute(text("SELECT count(*) FROM users;"))
    count = result.scalar()
    print(f"Number of users: {count}")
    
    result2 = conn.execute(text("SELECT username FROM users;"))
    users = [row[0] for row in result2]
    print(f"Users: {users}")
