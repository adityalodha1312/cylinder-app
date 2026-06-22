"""
Backfill script to populate scans.gas_type in both Supabase database and Google Sheets (Sheet1, Column G).
"""
import os
import sys
from datetime import datetime

# Add root folder to sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv
load_dotenv()

# Force migrations / db context
os.environ['RUN_MIGRATION'] = '1'

from app import app, db, get_all_cylinders, get_scan_rows, sheets_write_with_retry, scan_ws
from models import Scan, Cylinder

def backfill():
    with app.app_context():
        print("Starting backfill...")

        # 1. Ensure the scans table has the column
        try:
            from sqlalchemy import text
            db.session.execute(text("ALTER TABLE scans ADD COLUMN IF NOT EXISTS gas_type VARCHAR(50);"))
            db.session.commit()
            print("Successfully verified/added gas_type column in database scans table.")
        except Exception as e:
            print("Alter scans table check failed:", e)

        # 2. Build Cylinder UID -> Gas Type mapping
        print("Fetching cylinders from database...")
        cyls_db = Cylinder.query.all()
        gas_map = {c.uid.strip().upper(): c.gas_type for c in cyls_db if c.gas_type}

        # Fallback to Sheet cylinders
        print("Fetching cylinders from Google Sheets...")
        try:
            cyls_sheet = get_all_cylinders()
            for c in cyls_sheet:
                uid = c['uid'].strip().upper()
                if uid not in gas_map and c.get('gas_type'):
                    gas_map[uid] = c['gas_type']
        except Exception as se:
            print("Failed to load Sheets cylinders for mapping:", se)

        print(f"Loaded mapping for {len(gas_map)} cylinders.")

        # 3. Backfill database scans
        scans_updated_db = 0
        if os.environ.get('DATABASE_URL'):
            try:
                scans_to_fix = Scan.query.filter((Scan.gas_type == None) | (Scan.gas_type == '')).all()
                print(f"Found {len(scans_to_fix)} database scans without gas_type.")
                for s in scans_to_fix:
                    uid = s.cylinder_uid.strip().upper()
                    # Resolve gas type
                    resolved_gas = gas_map.get(uid, '')
                    if not resolved_gas:
                        # Fallback: parse from UID prefix
                        if '-' in uid:
                            prefix = uid.split('-')[0].upper()
                            if prefix in ['ARG', 'ACM', 'CO2', 'N2', 'OXY', 'HELIUM', 'DA', 'AHM']:
                                resolved_gas = prefix
                    
                    s.gas_type = resolved_gas
                    scans_updated_db += 1
                
                if scans_updated_db > 0:
                    db.session.commit()
                    print(f"Successfully backfilled {scans_updated_db} scans in Supabase database.")
                else:
                    print("No database scans needed backfilling.")
            except Exception as dbe:
                db.session.rollback()
                print("Error updating database scans:", dbe)

        # 4. Backfill Google Sheets (Sheet1)
        if scan_ws is not None:
            try:
                print("Fetching Sheet1 rows...")
                rows = scan_ws.get_all_values()
                if len(rows) > 0:
                    headers = rows[0]
                    # Ensure header exists at index 6 (Column G)
                    if len(headers) < 7:
                        scan_ws.update_cell(1, 7, "Gas Type")
                        print("Added 'Gas Type' column header in Column G.")
                        # Reload rows to align
                        rows = scan_ws.get_all_values()
                        headers = rows[0]
                    
                    updates = []
                    scans_updated_sheet = 0
                    for idx, r in enumerate(rows[1:], start=2): # 1-based index, skipping header
                        uid = r[4].strip().upper() if len(r) > 4 else ''
                        if not uid:
                            continue
                        
                        current_gas_val = r[6].strip() if len(r) > 6 else ''
                        if not current_gas_val:
                            resolved_gas = gas_map.get(uid, '')
                            if not resolved_gas:
                                if '-' in uid:
                                    prefix = uid.split('-')[0].upper()
                                    if prefix in ['ARG', 'ACM', 'CO2', 'N2', 'OXY', 'HELIUM', 'DA', 'AHM']:
                                        resolved_gas = prefix
                            
                            # Add update range
                            # Column G (7)
                            updates.append({
                                'range': f'G{idx}',
                                'values': [[resolved_gas]]
                            })
                            scans_updated_sheet += 1
                    
                    if updates:
                        print(f"Batch updating {len(updates)} rows in Google Sheets...")
                        # Run batch update to prevent API rate limiting
                        sheets_write_with_retry(scan_ws.batch_update, updates)
                        print(f"Successfully backfilled {scans_updated_sheet} rows in Google Sheets.")
                    else:
                        print("No Google Sheets rows needed backfilling.")
            except Exception as se:
                print("Error backfilling Google Sheets:", se)

        print("Backfill process finished!")

if __name__ == '__main__':
    backfill()
