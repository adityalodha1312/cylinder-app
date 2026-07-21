import os
from sqlalchemy import create_engine, inspect, text
from dotenv import load_dotenv

load_dotenv()
db_url = os.environ.get('DATABASE_URL')
engine = create_engine(db_url)
inspector = inspect(engine)
tables = inspector.get_table_names()

with engine.connect() as conn:
    print("Database Table Row Counts:")
    for table in tables:
        try:
            result = conn.execute(text(f"SELECT count(*) FROM {table};"))
            count = result.scalar()
            print(f"- {table}: {count} rows")
        except Exception as e:
            print(f"- {table}: ERROR ({e})")
