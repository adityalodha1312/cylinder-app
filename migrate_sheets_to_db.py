import os
import sys
import time
from datetime import datetime
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

# Load env variables
load_dotenv()
os.environ['RUN_MIGRATION'] = '1'
from app import app, doc
from db import db
from models import User, Customer, Cylinder, CylinderMaintenance, Scan, CustomerMap, BulkTank, Product

def parse_float(val):
    if not val:
        return 0.0
    try:
        return float(str(val).strip())
    except ValueError:
        return 0.0

def parse_int(val):
    if not val:
        return 0
    try:
        return int(str(val).strip())
    except ValueError:
        return 0

def parse_bool(val):
    if not val:
        return False
    return str(val).strip().upper() in ('TRUE', 'YES', '1')

def migrate():
    print("Starting Google Sheets to PostgreSQL Migration...")
    
    if not os.environ.get('DATABASE_URL'):
        print("Error: DATABASE_URL environment variable is not set!")
        sys.exit(1)
        
    with app.app_context():
        # Create all tables in PostgreSQL
        db.create_all()
        print("Database tables ensured.")

        if not doc:
            print("Error: Google Sheets document could not be opened. Check credentials.json or SPREADSHEET_NAME.")
            sys.exit(1)

        # 1. Migrate Users
        try:
            users_ws = doc.worksheet("Users")
            records = users_ws.get_all_records()
            print(f"Fetched {len(records)} users from sheet.")
            migrated_count = 0
            for r in records:
                username = str(r.get('Username', '')).strip()
                if not username:
                    continue
                # Check if user already exists
                existing = User.query.filter_by(username=username).first()
                if not existing:
                    user = User(
                        username=username,
                        password=str(r.get('Password', '')).strip(),
                        role=str(r.get('Role', 'driver')).strip().lower(),
                        name=str(r.get('Name', username)).strip()
                    )
                    db.session.add(user)
                    migrated_count += 1
            db.session.commit()
            print(f"Migrated {migrated_count} new users.")
        except Exception as e:
            print("Error migrating users:", e)

        # 2. Migrate Customers
        try:
            customer_ws = doc.worksheet("Customers")
            rows = customer_ws.get_all_values()
            print(f"Fetched {len(rows) - 1 if len(rows) > 0 else 0} customers from sheet.")
            migrated_count = 0
            if len(rows) > 1:
                for r in rows[1:]:
                    if len(r) < 2 or not r[1].strip():
                        continue
                    cust_id = r[0].strip()
                    name = r[1].strip()
                    existing = Customer.query.filter_by(name=name).first()
                    if not existing:
                        customer = Customer(
                            customer_id=cust_id,
                            name=name,
                            email=r[2].strip() if len(r) > 2 else '',
                            phone=r[3].strip() if len(r) > 3 else '',
                            address=r[4].strip() if len(r) > 4 else ''
                        )
                        db.session.add(customer)
                        migrated_count += 1
                db.session.commit()
                print(f"Migrated {migrated_count} new customers.")
        except Exception as e:
            print("Error migrating customers:", e)

        # 3. Migrate Cylinders
        try:
            cyl_ws = doc.worksheet("Cylinders")
            rows = cyl_ws.get_all_values()
            print(f"Fetched {len(rows) - 1 if len(rows) > 0 else 0} cylinders from sheet.")
            migrated_count = 0
            if len(rows) > 1:
                for r in rows[1:]:
                    if not r or not r[0].strip():
                        continue
                    uid = r[0].strip()
                    existing = Cylinder.query.filter_by(uid=uid).first()
                    if not existing:
                        cylinder = Cylinder(
                            uid=uid,
                            gas_type=r[1].strip() if len(r) > 1 else '',
                            cylinder_type=r[2].strip() if len(r) > 2 else '',
                            owner=r[3].strip() if len(r) > 3 else 'Depot',
                            status=r[4].strip() if len(r) > 4 else 'Active',
                            location=r[5].strip() if len(r) > 5 else 'Depot',
                            last_activity_date=r[6].strip() if len(r) > 6 else ''
                        )
                        db.session.add(cylinder)
                        migrated_count += 1
                db.session.commit()
                print(f"Migrated {migrated_count} new cylinders.")
        except Exception as e:
            print("Error migrating cylinders:", e)

        # 4. Migrate Cylinder Maintenance
        try:
            cyl_maint_ws = doc.worksheet("Cylinder Maintenance")
            rows = cyl_maint_ws.get_all_values()
            print(f"Fetched {len(rows) - 1 if len(rows) > 0 else 0} cylinder maintenance rows from sheet.")
            migrated_count = 0
            if len(rows) > 1:
                for r in rows[1:]:
                    if not r or not r[0].strip():
                        continue
                    uid = r[0].strip()
                    existing = CylinderMaintenance.query.filter_by(cylinder_uid=uid).first()
                    if not existing:
                        maint = CylinderMaintenance(
                            cylinder_uid=uid,
                            water_capacity=r[1].strip() if len(r) > 1 else '',
                            fill_pressure=r[2].strip() if len(r) > 2 else '',
                            gas_capacity=r[3].strip() if len(r) > 3 else '',
                            unit=r[4].strip() if len(r) > 4 else '',
                            is_mixture=r[5].strip() if len(r) > 5 else 'No',
                            mix_ratio=r[6].strip() if len(r) > 6 else '',
                            manufacture_date=r[7].strip() if len(r) > 7 else '',
                            last_hydro_date=r[8].strip() if len(r) > 8 else '',
                            next_hydro_due=r[9].strip() if len(r) > 9 else '',
                            hydro_test_status=r[10].strip() if len(r) > 10 else '',
                            cert_no=r[11].strip() if len(r) > 11 else '',
                            is_uhp='Yes' if (len(r) > 12 and r[12].strip().upper() == 'YES') else 'No'
                        )
                        db.session.add(maint)
                        migrated_count += 1
                db.session.commit()
                print(f"Migrated {migrated_count} new maintenance rows.")
        except Exception as e:
            print("Error migrating cylinder maintenance:", e)

        # 5. Migrate Scans (Sheet1)
        try:
            scan_ws = doc.worksheet("Sheet1")
            rows = scan_ws.get_all_values()
            print(f"Fetched {len(rows) - 1 if len(rows) > 0 else 0} scans from sheet.")
            migrated_count = 0
            if len(rows) > 1:
                for r in rows[1:]:
                    if len(r) < 5 or not r[4].strip():
                        continue
                    s_date = r[0].strip()
                    s_time = r[1].strip()
                    driver = r[2].strip()
                    action = r[3].strip()
                    uid = r[4].strip()
                    customer = r[5].strip() if len(r) > 5 else ''
                    
                    # Deduplicate scans by checking exact same fields
                    existing = Scan.query.filter_by(
                        scan_date=s_date,
                        scan_time=s_time,
                        driver=driver,
                        action=action,
                        cylinder_uid=uid,
                        customer=customer
                    ).first()
                    
                    if not existing:
                        scan = Scan(
                            scan_date=s_date,
                            scan_time=s_time,
                            driver=driver,
                            action=action,
                            cylinder_uid=uid,
                            customer=customer
                        )
                        db.session.add(scan)
                        migrated_count += 1
                db.session.commit()
                print(f"Migrated {migrated_count} new scans.")
        except Exception as e:
            print("Error migrating scans:", e)

        # 6. Migrate Customer Map
        try:
            map_ws = doc.worksheet("Customer Map")
            rows = map_ws.get_all_values()
            print(f"Fetched {len(rows) - 1 if len(rows) > 0 else 0} customer map rows from sheet.")
            migrated_count = 0
            if len(rows) > 1:
                for r in rows[1:]:
                    if len(r) < 7 or not r[6].strip():
                        continue
                    s_date = r[0].strip()
                    s_time = r[1].strip()
                    driver = r[2].strip()
                    action = r[3].strip()
                    count = parse_int(r[4])
                    uids = r[5].strip()
                    customer = r[6].strip()
                    send_receipt = parse_bool(r[7]) if len(r) > 7 else False
                    receipt_status = r[8].strip() if len(r) > 8 else ''
                    
                    existing = CustomerMap.query.filter_by(
                        scan_date=s_date,
                        scan_time=s_time,
                        driver=driver,
                        action=action,
                        customer=customer
                    ).first()
                    
                    if not existing:
                        cmap = CustomerMap(
                            scan_date=s_date,
                            scan_time=s_time,
                            driver=driver,
                            action=action,
                            count=count,
                            uids=uids,
                            customer=customer,
                            send_receipt=send_receipt,
                            receipt_status=receipt_status
                        )
                        db.session.add(cmap)
                        migrated_count += 1
                db.session.commit()
                print(f"Migrated {migrated_count} new customer map rows.")
        except Exception as e:
            print("Error migrating customer map:", e)

        # 7. Migrate Bulk Tanks
        try:
            bulk_tanks_ws = doc.worksheet("Bulk Tanks")
            rows = bulk_tanks_ws.get_all_values()
            print(f"Fetched {len(rows) - 1 if len(rows) > 0 else 0} bulk tank entries from sheet.")
            migrated_count = 0
            if len(rows) > 1:
                for r in rows[1:]:
                    if len(r) < 6 or not r[0].strip():
                        continue
                    b_date = r[0].strip()
                    gas = r[1].strip()
                    opening = parse_float(r[2])
                    dead_volume = parse_float(r[3])
                    capacity = parse_float(r[4])
                    unit = r[5].strip()
                    
                    existing = BulkTank.query.filter_by(
                        date=b_date,
                        gas=gas
                    ).first()
                    
                    if not existing:
                        btank = BulkTank(
                            date=b_date,
                            gas=gas,
                            opening=opening,
                            dead_volume=dead_volume,
                            capacity=capacity,
                            unit=unit
                        )
                        db.session.add(btank)
                        migrated_count += 1
                db.session.commit()
                print(f"Migrated {migrated_count} new bulk tank entries.")
        except Exception as e:
            print("Error migrating bulk tanks:", e)

        # 8. Migrate Products
        try:
            products_ws = doc.worksheet("Products")
            rows = products_ws.get_all_values()
            print(f"Fetched {len(rows) - 1 if len(rows) > 0 else 0} products from sheet.")
            migrated_count = 0
            if len(rows) > 1:
                for r in rows[1:]:
                    if len(r) < 6 or not r[0].strip():
                        continue
                    prod_id = r[0].strip()
                    existing = Product.query.filter_by(product_id=prod_id).first()
                    if not existing:
                        prod = Product(
                            product_id=prod_id,
                            name=r[1].strip(),
                            gas_type=r[2].strip().upper(),
                            cylinder_type=r[3].strip().capitalize(),
                            gas_per_cyl=parse_float(r[4]),
                            unit=r[5].strip(),
                            is_virtual=parse_bool(r[6]) if len(r) > 6 else False
                        )
                        db.session.add(prod)
                        migrated_count += 1
                db.session.commit()
                print(f"Migrated {migrated_count} new products.")
        except Exception as e:
            print("Error migrating products:", e)

    print("Migration completed successfully!")

if __name__ == "__main__":
    migrate()
