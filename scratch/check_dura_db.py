"""Check current DB state of TEST001 and dura_gas_history."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
os.environ['RUN_MIGRATION'] = '1'
from dotenv import load_dotenv
load_dotenv()
from app import app
from db import db
from models import Cylinder, DuraGasHistory

with app.app_context():
    cyl = Cylinder.query.filter_by(uid='TEST001').first()
    print(f"TEST001 in DB: gas_type = {cyl.gas_type if cyl else 'NOT FOUND'}")

    rows = DuraGasHistory.query.order_by(DuraGasHistory.id.desc()).limit(5).all()
    print(f"\ndura_gas_history last {len(rows)} rows:")
    for r in rows:
        print(f"  [{r.id}] {r.cylinder_uid}: {r.previous_gas} -> {r.gas_filled}  op={r.operator}  date={r.fill_date} {r.fill_time}")
