import os
import sys
import time
from datetime import datetime
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
import gspread

# Load env variables
load_dotenv()

from db import db
from models import User, Customer, Cylinder, CylinderMaintenance, Scan, CustomerMap, BulkTank, Product, SystemSetting
from werkzeug.security import generate_password_hash

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

def sync_sheets_to_db(doc):
    """
    Reads all Google Sheets and upserts their content into the PostgreSQL database.
    Allows manual edits in Google Sheets to flow back to PostgreSQL.
    """
    if not doc:
        print("[sync] Error: doc is None, cannot sync.")
        return False

    print(f"[sync] Starting periodic sync from Google Sheets to DB at {datetime.now()}...")
    
    # 1. Sync Users
    try:
        users_ws = doc.worksheet("Users")
        records = users_ws.get_all_records()
        existing_users = {u.username: u for u in User.query.all()}
        
        sheet_usernames = set()
        for r in records:
            username = str(r.get('Username', '')).strip()
            if not username:
                continue
            sheet_usernames.add(username)
            password = str(r.get('Password', '')).strip()
            role = str(r.get('Role', 'driver')).strip().lower()
            name = str(r.get('Name', username)).strip()
            
            # Hash plain text password from Google Sheets
            db_password = password
            if password and not (password.startswith('pbkdf2:') or password.startswith('scrypt:')):
                db_password = generate_password_hash(password)
            
            if username in existing_users:
                user = existing_users[username]
                # Compare db_password with user.password
                if user.password != db_password or user.role != role or user.name != name:
                    user.password = db_password
                    user.role = role
                    user.name = name
                    print(f"[sync] Updated user: {username}")
            else:
                user = User(username=username, password=db_password, role=role, name=name)
                db.session.add(user)
                print(f"[sync] Added user: {username}")
                
        # Deletions
        for username, user_obj in existing_users.items():
            if username not in sheet_usernames:
                print(f"[sync] Deleting user from DB: {username}")
                db.session.delete(user_obj)
                
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Users sheet:", e)

    # 2. Sync Customers
    try:
        customer_ws = doc.worksheet("Customers")
        rows = customer_ws.get_all_values()
        existing_customers = {c.name: c for c in Customer.query.all()}
        
        sheet_customer_names = set()
        if len(rows) > 1:
            for r in rows[1:]:
                if len(r) < 2 or not r[1].strip():
                    continue
                cust_id = r[0].strip()
                name = r[1].strip()
                sheet_customer_names.add(name)
                email = r[2].strip() if len(r) > 2 else ''
                phone = r[3].strip() if len(r) > 3 else ''
                address = r[4].strip() if len(r) > 4 else ''
                
                if name in existing_customers:
                    c = existing_customers[name]
                    if c.customer_id != cust_id or c.email != email or c.phone != phone or c.address != address:
                        c.customer_id = cust_id
                        c.email = email
                        c.phone = phone
                        c.address = address
                        print(f"[sync] Updated customer: {name}")
                else:
                    customer = Customer(customer_id=cust_id, name=name, email=email, phone=phone, address=address)
                    db.session.add(customer)
                    print(f"[sync] Added customer: {name}")
                    
        # Deletions
        for name, cust_obj in existing_customers.items():
            if name not in sheet_customer_names:
                print(f"[sync] Deleting customer from DB: {name}")
                db.session.delete(cust_obj)
                
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Customers sheet:", e)

    # 3. Sync Cylinders
    try:
        cyl_ws = doc.worksheet("Cylinders")
        rows = cyl_ws.get_all_values()
        existing_cylinders = {cyl.uid: cyl for cyl in Cylinder.query.all()}
        
        sheet_uids = set()
        if len(rows) > 1:
            for r in rows[1:]:
                if not r or not r[0].strip():
                    continue
                uid = r[0].strip()
                sheet_uids.add(uid)
                gas_type = r[1].strip() if len(r) > 1 else ''
                cyl_type = r[2].strip() if len(r) > 2 else ''
                owner = r[3].strip() if len(r) > 3 else 'Depot'
                status = r[4].strip() if len(r) > 4 else 'Active'
                location = r[5].strip() if len(r) > 5 else 'Depot'
                last_act = r[6].strip() if len(r) > 6 else ''
                
                if uid in existing_cylinders:
                    c = existing_cylinders[uid]
                    if (c.gas_type != gas_type or c.cylinder_type != cyl_type or 
                        c.owner != owner or c.status != status or 
                        c.location != location or c.last_activity_date != last_act):
                        c.gas_type = gas_type
                        c.cylinder_type = cyl_type
                        c.owner = owner
                        c.status = status
                        c.location = location
                        c.last_activity_date = last_act
                        print(f"[sync] Updated cylinder: {uid}")
                else:
                    cylinder = Cylinder(
                        uid=uid, gas_type=gas_type, cylinder_type=cyl_type,
                        owner=owner, status=status, location=location, last_activity_date=last_act
                    )
                    db.session.add(cylinder)
                    print(f"[sync] Added cylinder: {uid}")
                    
        # Deletions
        for uid, cyl_obj in existing_cylinders.items():
            if uid not in sheet_uids:
                print(f"[sync] Deleting cylinder from DB: {uid}")
                db.session.delete(cyl_obj)
                
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Cylinders sheet:", e)

    # 4. Sync Cylinder Maintenance
    try:
        cyl_maint_ws = doc.worksheet("Cylinder Maintenance")
        rows = cyl_maint_ws.get_all_values()
        existing_maint = {m.cylinder_uid: m for m in CylinderMaintenance.query.all()}
        
        sheet_maint_uids = set()
        if len(rows) > 1:
            for r in rows[1:]:
                if not r or not r[0].strip():
                    continue
                uid = r[0].strip()
                sheet_maint_uids.add(uid)
                w_cap = r[1].strip() if len(r) > 1 else ''
                f_pres = r[2].strip() if len(r) > 2 else ''
                g_cap = r[3].strip() if len(r) > 3 else ''
                unit = r[4].strip() if len(r) > 4 else ''
                is_mix = r[5].strip() if len(r) > 5 else 'No'
                mix_ratio = r[6].strip() if len(r) > 6 else ''
                mfg_date = r[7].strip() if len(r) > 7 else ''
                last_hydro = r[8].strip() if len(r) > 8 else ''
                next_hydro = r[9].strip() if len(r) > 9 else ''
                hydro_status = r[10].strip() if len(r) > 10 else ''
                cert_no = r[11].strip() if len(r) > 11 else ''
                is_uhp = 'Yes' if (len(r) > 12 and r[12].strip().upper() == 'YES') else 'No'
                
                if uid in existing_maint:
                    m = existing_maint[uid]
                    if (m.water_capacity != w_cap or m.fill_pressure != f_pres or 
                        m.gas_capacity != g_cap or m.unit != unit or 
                        m.is_mixture != is_mix or m.mix_ratio != mix_ratio or 
                        m.manufacture_date != mfg_date or m.last_hydro_date != last_hydro or 
                        m.next_hydro_due != next_hydro or m.hydro_test_status != hydro_status or 
                        m.cert_no != cert_no or m.is_uhp != is_uhp):
                        m.water_capacity = w_cap
                        m.fill_pressure = f_pres
                        m.gas_capacity = g_cap
                        m.unit = unit
                        m.is_mixture = is_mix
                        m.mix_ratio = mix_ratio
                        m.manufacture_date = mfg_date
                        m.last_hydro_date = last_hydro
                        m.next_hydro_due = next_hydro
                        m.hydro_test_status = hydro_status
                        m.cert_no = cert_no
                        m.is_uhp = is_uhp
                        print(f"[sync] Updated maintenance for: {uid}")
                else:
                    maint = CylinderMaintenance(
                        cylinder_uid=uid, water_capacity=w_cap, fill_pressure=f_pres,
                        gas_capacity=g_cap, unit=unit, is_mixture=is_mix, mix_ratio=mix_ratio,
                        manufacture_date=mfg_date, last_hydro_date=last_hydro, next_hydro_due=next_hydro,
                        hydro_test_status=hydro_status, cert_no=cert_no, is_uhp=is_uhp
                    )
                    db.session.add(maint)
                    print(f"[sync] Added maintenance for: {uid}")
                    
        # Deletions
        for uid, maint_obj in existing_maint.items():
            if uid not in sheet_maint_uids:
                print(f"[sync] Deleting maintenance from DB: {uid}")
                db.session.delete(maint_obj)
                
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Cylinder Maintenance sheet:", e)

    # 5. Sync Scans (Sheet1)
    try:
        scan_ws = doc.worksheet("Sheet1")
        rows = scan_ws.get_all_values()
        existing_scans = {
            (s.scan_date, s.scan_time, s.driver, s.action, s.cylinder_uid, s.customer): s 
            for s in Scan.query.all()
        }
        
        sheet_keys = set()
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
                
                key = (s_date, s_time, driver, action, uid, customer)
                sheet_keys.add(key)
                if key not in existing_scans:
                    scan = Scan(
                        scan_date=s_date, scan_time=s_time, driver=driver,
                        action=action, cylinder_uid=uid, customer=customer
                    )
                    db.session.add(scan)
                    print(f"[sync] Added scan event: {uid} at {s_date} {s_time}")
                    
        # Deletions
        for key, scan_obj in existing_scans.items():
            if key not in sheet_keys:
                print(f"[sync] Deleting scan event from DB: {key[4]} at {key[0]} {key[1]}")
                db.session.delete(scan_obj)
                
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing scans:", e)

    # 6. Sync Customer Map
    try:
        map_ws = doc.worksheet("Customer Map")
        rows = map_ws.get_all_values()
        existing_maps = {
            (m.scan_date, m.scan_time, m.driver, m.action, m.customer): m 
            for m in CustomerMap.query.all()
        }
        
        sheet_map_keys = set()
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
                send_rec = parse_bool(r[7]) if len(r) > 7 else False
                rec_status = r[8].strip() if len(r) > 8 else ''
                
                key = (s_date, s_time, driver, action, customer)
                sheet_map_keys.add(key)
                if key in existing_maps:
                    cmap = existing_maps[key]
                    if cmap.count != count or cmap.uids != uids or cmap.send_receipt != send_rec or cmap.receipt_status != rec_status:
                        cmap.count = count
                        cmap.uids = uids
                        cmap.send_receipt = send_rec
                        cmap.receipt_status = rec_status
                        print(f"[sync] Updated Customer Map row: {customer} - {s_date}")
                else:
                    cmap = CustomerMap(
                        scan_date=s_date, scan_time=s_time, driver=driver, action=action,
                        count=count, uids=uids, customer=customer, send_receipt=send_rec,
                        receipt_status=rec_status
                    )
                    db.session.add(cmap)
                    print(f"[sync] Added Customer Map row: {customer} - {s_date}")
                    
        # Deletions
        for key, map_obj in existing_maps.items():
            if key not in sheet_map_keys:
                print(f"[sync] Deleting Customer Map row from DB: {key[4]} - {key[0]}")
                db.session.delete(map_obj)
                
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Customer Map sheet:", e)

    # 7. Sync Bulk Tanks
    try:
        bulk_tanks_ws = doc.worksheet("Bulk Tanks")
        rows = bulk_tanks_ws.get_all_values()
        existing_tanks = {(bt.date, bt.gas): bt for bt in BulkTank.query.all()}
        
        sheet_tank_keys = set()
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
                
                key = (b_date, gas)
                sheet_tank_keys.add(key)
                if key in existing_tanks:
                    bt = existing_tanks[key]
                    if bt.opening != opening or bt.dead_volume != dead_volume or bt.capacity != capacity or bt.unit != unit:
                        bt.opening = opening
                        bt.dead_volume = dead_volume
                        bt.capacity = capacity
                        bt.unit = unit
                        print(f"[sync] Updated Bulk Tank stock for: {gas} - {b_date}")
                else:
                    bt = BulkTank(
                        date=b_date, gas=gas, opening=opening,
                        dead_volume=dead_volume, capacity=capacity, unit=unit
                    )
                    db.session.add(bt)
                    print(f"[sync] Added Bulk Tank stock for: {gas} - {b_date}")
                    
        # Deletions
        for key, tank_obj in existing_tanks.items():
            if key not in sheet_tank_keys:
                print(f"[sync] Deleting Bulk Tank stock from DB: {key[1]} - {key[0]}")
                db.session.delete(tank_obj)
                
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Bulk Tanks sheet:", e)

    # 8. Sync Products
    try:
        products_ws = doc.worksheet("Products")
        rows = products_ws.get_all_values()
        existing_products = {p.product_id: p for p in Product.query.all()}
        
        sheet_prod_ids = set()
        if len(rows) > 1:
            for r in rows[1:]:
                if len(r) < 6 or not r[0].strip():
                    continue
                prod_id = r[0].strip()
                sheet_prod_ids.add(prod_id)
                name = r[1].strip()
                gas_type = r[2].strip().upper()
                cyl_type = r[3].strip().capitalize()
                gas_per = parse_float(r[4])
                unit = r[5].strip()
                is_virt = parse_bool(r[6]) if len(r) > 6 else False
                
                if prod_id in existing_products:
                    p = existing_products[prod_id]
                    if (p.name != name or p.gas_type != gas_type or 
                        p.cylinder_type != cyl_type or p.gas_per_cyl != gas_per or 
                        p.unit != unit or p.is_virtual != is_virt):
                        p.name = name
                        p.gas_type = gas_type
                        p.cylinder_type = cyl_type
                        p.gas_per_cyl = gas_per
                        p.unit = unit
                        p.is_virtual = is_virt
                        print(f"[sync] Updated Product config: {prod_id}")
                else:
                    prod = Product(
                        product_id=prod_id, name=name, gas_type=gas_type,
                        cylinder_type=cyl_type, gas_per_cyl=gas_per, unit=unit, is_virtual=is_virt
                    )
                    db.session.add(prod)
                    print(f"[sync] Added Product config: {prod_id}")
                    
        # Deletions
        for prod_id, prod_obj in existing_products.items():
            if prod_id not in sheet_prod_ids:
                print(f"[sync] Deleting Product config from DB: {prod_id}")
                db.session.delete(prod_obj)
                
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Products sheet:", e)

    # 9. Sync Settings
    try:
        settings_ws = doc.worksheet("Settings")
        rows = settings_ws.get_all_values()
        existing_settings = {s.key: s for s in SystemSetting.query.all()}
        
        sheet_keys = set()
        if len(rows) > 1:
            for r in rows[1:]:
                if not r or not r[0].strip():
                    continue
                key = r[0].strip()
                sheet_keys.add(key)
                val = r[1].strip() if len(r) > 1 else ''
                
                if key in existing_settings:
                    s = existing_settings[key]
                    if s.value != val:
                        s.value = val
                        print(f"[sync] Updated setting: {key} -> {val}")
                else:
                    s = SystemSetting(key=key, value=val)
                    db.session.add(s)
                    print(f"[sync] Added setting: {key} -> {val}")
                    
        db.session.commit()
    except Exception as e:
        # Settings sheet might not exist, ignore silently
        pass

    print("[sync] Periodic sync finished.")
    return True


def sync_db_to_sheets(doc):
    """
    Mirrors the PostgreSQL DB → Google Sheets.
    Runs automatically every 5 minutes via the background scheduler.
    Sheets become a live read-only backup of the DB.
    """
    if not doc:
        print("[sync_db→sheets] doc is None, skipping.")
        return False

    print(f"[sync_db→sheets] Starting DB→Sheets mirror at {datetime.now()}...")

    CYLINDER_SHEET_NAME   = "Cylinders"
    CYLINDER_MAINT_NAME   = "Cylinder Maintenance"
    USERS_SHEET_NAME      = "Users"
    SCANS_SHEET_NAME      = "Sheet1"
    CUSTOMERS_SHEET_NAME  = "Customers"

    # 1. Cylinders
    try:
        from models import Cylinder, CylinderMaintenance, User, Scan, Customer
        cylinders = Cylinder.query.all()
        maints    = {m.cylinder_uid: m for m in CylinderMaintenance.query.all()}

        cyl_rows   = []
        maint_rows = []
        for c in cylinders:
            cyl_rows.append([
                c.uid, c.gas_type or '', c.cylinder_type or '',
                c.owner or '', c.status or '', c.location or '',
                c.last_activity_date or ''
            ])
            m = maints.get(c.uid)
            if m:
                maint_rows.append([
                    c.uid, m.water_capacity or '', m.fill_pressure or '',
                    m.gas_capacity or '', m.unit or '', m.is_mixture or 'No',
                    m.mix_ratio or '', m.manufacture_date or '',
                    m.last_hydro_date or '', m.next_hydro_due or '',
                    m.hydro_test_status or '', m.cert_no or '', m.is_uhp or 'No'
                ])

        cyl_ws = doc.worksheet(CYLINDER_SHEET_NAME)
        cyl_ws.batch_clear(["A2:Z100000"])
        if cyl_rows:
            cyl_ws.update(cyl_rows, "A2")
        print(f"[sync_db→sheets] Cylinders: wrote {len(cyl_rows)} rows.")

        maint_ws = doc.worksheet(CYLINDER_MAINT_NAME)
        maint_ws.batch_clear(["A2:Z100000"])
        if maint_rows:
            maint_ws.update(maint_rows, "A2")
        print(f"[sync_db→sheets] Cylinder Maintenance: wrote {len(maint_rows)} rows.")
    except Exception as e:
        print("[sync_db→sheets] Error syncing Cylinders/Maintenance:", e)

    # 2. Users
    try:
        users     = User.query.all()
        user_rows = [[u.username, u.password, u.role, u.name or u.username] for u in users]
        users_ws  = doc.worksheet(USERS_SHEET_NAME)
        users_ws.batch_clear(["A2:Z100000"])
        if user_rows:
            users_ws.update(user_rows, "A2")
        print(f"[sync_db→sheets] Users: wrote {len(user_rows)} rows.")
    except Exception as e:
        print("[sync_db→sheets] Error syncing Users:", e)

    # 3. Scans (Sheet1)
    try:
        scans     = Scan.query.order_by(Scan.id.asc()).all()
        scan_rows = [
            [s.scan_date or '', s.scan_time or '', s.driver or '',
             s.action or '', s.cylinder_uid or '', s.customer or '']
            for s in scans
        ]
        scan_ws = doc.worksheet(SCANS_SHEET_NAME)
        scan_ws.batch_clear(["A2:Z100000"])
        if scan_rows:
            scan_ws.update(scan_rows, "A2")
        print(f"[sync_db→sheets] Scans: wrote {len(scan_rows)} rows.")
    except Exception as e:
        print("[sync_db→sheets] Error syncing Scans:", e)

    # 4. Customers
    try:
        customers  = Customer.query.all()
        cust_rows  = [
            [c.customer_id or '', c.name or '', c.email or '',
             c.phone or '', c.address or '']
            for c in customers
        ]
        cust_ws = doc.worksheet(CUSTOMERS_SHEET_NAME)
        cust_ws.batch_clear(["A2:Z100000"])
        if cust_rows:
            cust_ws.update(cust_rows, "A2")
        print(f"[sync_db→sheets] Customers: wrote {len(cust_rows)} rows.")
    except Exception as e:
        print("[sync_db→sheets] Error syncing Customers:", e)

    print("[sync_db→sheets] DB→Sheets mirror finished.")
    return True
