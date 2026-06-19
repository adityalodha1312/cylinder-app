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
from models import User, Customer, Cylinder, CylinderMaintenance, Scan, CustomerMap, BulkTank, Product
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
        
        for r in records:
            username = str(r.get('Username', '')).strip()
            if not username:
                continue
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
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Users sheet:", e)

    # 2. Sync Customers
    try:
        customer_ws = doc.worksheet("Customers")
        rows = customer_ws.get_all_values()
        existing_customers = {c.name: c for c in Customer.query.all()}
        
        if len(rows) > 1:
            for r in rows[1:]:
                if len(r) < 2 or not r[1].strip():
                    continue
                cust_id = r[0].strip()
                name = r[1].strip()
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
            db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Customers sheet:", e)

    # 3. Sync Cylinders
    try:
        cyl_ws = doc.worksheet("Cylinders")
        rows = cyl_ws.get_all_values()
        existing_cylinders = {cyl.uid: cyl for cyl in Cylinder.query.all()}
        
        if len(rows) > 1:
            for r in rows[1:]:
                if not r or not r[0].strip():
                    continue
                uid = r[0].strip()
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
            db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Cylinders sheet:", e)

    # 4. Sync Cylinder Maintenance
    try:
        cyl_maint_ws = doc.worksheet("Cylinder Maintenance")
        rows = cyl_maint_ws.get_all_values()
        existing_maint = {m.cylinder_uid: m for m in CylinderMaintenance.query.all()}
        
        if len(rows) > 1:
            for r in rows[1:]:
                if not r or not r[0].strip():
                    continue
                uid = r[0].strip()
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
                if key not in existing_scans:
                    scan = Scan(
                        scan_date=s_date, scan_time=s_time, driver=driver,
                        action=action, cylinder_uid=uid, customer=customer
                    )
                    db.session.add(scan)
                    print(f"[sync] Added scan event: {uid} at {s_date} {s_time}")
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
            db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Customer Map sheet:", e)

    # 7. Sync Bulk Tanks
    try:
        bulk_tanks_ws = doc.worksheet("Bulk Tanks")
        rows = bulk_tanks_ws.get_all_values()
        existing_tanks = {(bt.date, bt.gas): bt for bt in BulkTank.query.all()}
        
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
            db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Bulk Tanks sheet:", e)

    # 8. Sync Products
    try:
        products_ws = doc.worksheet("Products")
        rows = products_ws.get_all_values()
        existing_products = {p.product_id: p for p in Product.query.all()}
        
        if len(rows) > 1:
            for r in rows[1:]:
                if len(r) < 6 or not r[0].strip():
                    continue
                prod_id = r[0].strip()
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
            db.session.commit()
    except Exception as e:
        db.session.rollback()
        print("[sync] Error syncing Products sheet:", e)

    print("[sync] Periodic sync finished.")
    return True
