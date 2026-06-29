"""
Run this script ONCE to create any missing tables (like dura_gas_history) in Supabase.
Usage:  python scratch/create_missing_tables.py
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from dotenv import load_dotenv
load_dotenv()

from db import db
from models import DuraGasHistory, Cylinder, Customer  # import all models so SQLAlchemy sees them
from app import app

with app.app_context():
    db.create_all()
    print("[OK] All tables ensured (created if missing).")

    # Verify dura_gas_history exists and show row count
    count = DuraGasHistory.query.count()
    print(f"[INFO] dura_gas_history currently has {count} rows.")
