from flask import Flask, render_template, request, redirect, session, jsonify
from openpyxl import load_workbook
from datetime import datetime, date, timedelta
from functools import wraps
import gspread
from google.oauth2.service_account import Credentials
import time
from dotenv import load_dotenv
import os
import re


load_dotenv()

app = Flask(__name__)
app.secret_key = 'cyl-tracker-secret-2026'

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_file(
    "credentials.json",
    scopes=SCOPES
)
client = gspread.authorize(creds)

SPREADSHEET_NAME    = "Cylinder Tracking"
SCAN_SHEET_NAME     = "Sheet1"
MAP_SHEET_NAME      = "Customer Map"
USERS_SHEET_NAME    = "Users"
CUSTOMER_SHEET_NAME = "Customers"
CYLINDER_SHEET_NAME = "Cylinders"
CYLINDER_MAINT_NAME = "Cylinder Maintenance"
BULK_TANKS_NAME     = "Bulk Tanks"

# Cache worksheet objects at startup to avoid roundtrip sheet lookup calls
try:
    doc = client.open(SPREADSHEET_NAME)
    scan_ws = doc.worksheet(SCAN_SHEET_NAME)
    map_ws = doc.worksheet(MAP_SHEET_NAME)
    users_ws = doc.worksheet(USERS_SHEET_NAME)
    customer_ws = doc.worksheet(CUSTOMER_SHEET_NAME)
    sheet = scan_ws
    try:
        cyl_ws = doc.worksheet(CYLINDER_SHEET_NAME)
    except Exception:
        cyl_ws = None
    try:
        cyl_maint_ws = doc.worksheet(CYLINDER_MAINT_NAME)
    except Exception:
        cyl_maint_ws = None
    try:
        bulk_tanks_ws = doc.worksheet(BULK_TANKS_NAME)
    except Exception:
        bulk_tanks_ws = None
except Exception as e:
    print("Error caching Google Sheets worksheets:", e)
    doc = None
    scan_ws = None
    map_ws = None
    users_ws = None
    customer_ws = None
    sheet = None
    cyl_ws = None
    cyl_maint_ws = None
    bulk_tanks_ws = None

# Helper functions to fetch customer details from Google Sheets
def get_customer_names():
    try:
        if customer_ws is None:
            return []
        values = customer_ws.get_all_values()
        if len(values) < 2:
            return []
        # Column B is "Name" (index 1)
        names = [row[1].strip() for row in values[1:] if len(row) > 1 and row[1].strip()]
        return sorted(list(set(names)))
    except Exception as e:
        print("Error getting customer names from sheet:", e)
        return []

def get_customer_emails():
    """Returns a dict of {customer_name: email}"""
    try:
        if customer_ws is None:
            return {}
        values = customer_ws.get_all_values()
        if len(values) < 2:
            return {}
        # Name is Column B (index 1), Email is Column C (index 2)
        out = {}
        for row in values[1:]:
            if len(row) > 1 and row[1].strip():
                name = row[1].strip()
                email = row[2].strip() if len(row) > 2 else ''
                out[name] = email
        return out
    except Exception as e:
        print("Error getting customer emails from sheet:", e)
        return {}


# ── Cylinder Registry helpers ──────────────────────────────────────────────
def get_all_cylinders():
    """Returns list of dicts from Cylinders sheet."""
    global cyl_ws
    try:
        if cyl_ws is None:
            if doc:
                try:
                    cyl_ws = doc.worksheet(CYLINDER_SHEET_NAME)
                except Exception:
                    return []
            else:
                return []
        rows = cyl_ws.get_all_values()
        if len(rows) < 2:
            return []
        out = []
        for r in rows[1:]:
            if len(r) >= 1 and r[0].strip():
                out.append({
                    'uid'           : r[0].strip() if len(r) > 0 else '',
                    'gas_type'      : r[1].strip() if len(r) > 1 else '',
                    'cylinder_type' : r[2].strip() if len(r) > 2 else '',
                    'owner'         : r[3].strip() if len(r) > 3 else '',
                    'status'        : r[4].strip() if len(r) > 4 else 'Active',
                    'location'      : r[5].strip() if len(r) > 5 else 'Depot',
                    'last_activity' : r[6].strip() if len(r) > 6 else '',
                })
        return out
    except Exception as e:
        print("Error getting cylinders:", e)
        return []

def get_all_maintenance():
    """Returns dict of {uid: maintenance_dict} from Cylinder Maintenance sheet."""
    global cyl_maint_ws
    try:
        if cyl_maint_ws is None:
            if doc:
                try:
                    cyl_maint_ws = doc.worksheet(CYLINDER_MAINT_NAME)
                except Exception:
                    return {}
            else:
                return {}
        rows = cyl_maint_ws.get_all_values()
        if len(rows) < 2:
            return {}
        out = {}
        for r in rows[1:]:
            uid = r[0].strip() if len(r) > 0 else ''
            if not uid:
                continue
            out[uid] = {
                'uid'              : uid,
                'water_capacity'   : r[1].strip() if len(r) > 1 else '',
                'fill_pressure'    : r[2].strip() if len(r) > 2 else '',
                'gas_capacity'     : r[3].strip() if len(r) > 3 else '',
                'unit'             : r[4].strip() if len(r) > 4 else '',
                'is_mixture'       : r[5].strip() if len(r) > 5 else 'No',
                'mix_ratio'        : r[6].strip() if len(r) > 6 else '',
                'manufacture_date' : r[7].strip() if len(r) > 7 else '',
                'last_hydro_date'  : r[8].strip() if len(r) > 8 else '',
                'next_hydro_due'   : r[9].strip() if len(r) > 9 else '',
                'hydro_test_status': r[10].strip() if len(r) > 10 else '',
                'cert_no'          : r[11].strip() if len(r) > 11 else '',
            }
        return out
    except Exception as e:
        print("Error getting maintenance data:", e)
        return {}

def compute_hydro_badge(next_hydro_due_str):
    """Returns ('OK'|'Due Soon'|'Overdue'|'Not Set') based on next_hydro_due date string."""
    if not next_hydro_due_str:
        return 'Not Set'
    d = parse_date(next_hydro_due_str)
    if not d:
        return 'Not Set'
    today = date.today()
    delta = (d - today).days
    if delta < 0:
        return 'Overdue'
    elif delta <= 30:
        return 'Due Soon'
    else:
        return 'OK'

def merge_cylinder_data():
    """Merges Cylinders + Cylinder Maintenance sheets, computes hydro badge."""
    cylinders   = get_all_cylinders()
    maintenance = get_all_maintenance()
    merged = []
    for cyl in cylinders:
        maint = maintenance.get(cyl['uid'], {})
        hydro_badge = compute_hydro_badge(maint.get('next_hydro_due', ''))
        merged.append({**cyl, **maint, 'hydro_badge': hydro_badge})
    return merged

def find_cylinder_rows(uid):
    """Returns (cyl_row_1indexed, maint_row_1indexed) for a given UID. Returns (None, None) if not found."""
    global cyl_ws, cyl_maint_ws
    cyl_row = None
    maint_row = None
    try:
        if cyl_ws:
            rows = cyl_ws.get_all_values()
            for i, r in enumerate(rows):
                if i == 0:
                    continue
                if r and r[0].strip().upper() == uid.strip().upper():
                    cyl_row = i + 1
                    break
    except Exception as e:
        print("Error finding cylinder row:", e)
    try:
        if cyl_maint_ws:
            rows = cyl_maint_ws.get_all_values()
            for i, r in enumerate(rows):
                if i == 0:
                    continue
                if r and r[0].strip().upper() == uid.strip().upper():
                    maint_row = i + 1
                    break
    except Exception as e:
        print("Error finding maintenance row:", e)
    return cyl_row, maint_row


# In-memory TTL data cache configurations
_data_cache = {
    'scans': None,
    'scans_time': 0,
    'map': None,
    'map_time': 0
}
CACHE_TTL = 10  # Duration in seconds

def clear_cache():
    _data_cache['scans'] = None
    _data_cache['scans_time'] = 0
    _data_cache['map'] = None
    _data_cache['map_time'] = 0



# ================================================================
#  AUTH HELPERS
# ================================================================

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return redirect('/login')
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return redirect('/login')
        if session['user']['role'] not in ['manager', 'owner']:
            return redirect('/login')
        return f(*args, **kwargs)
    return decorated

def check_login(username, password):
    try:
        if users_ws is None:
            return None
        records  = users_ws.get_all_records()
        for r in records:
            u = str(r.get('Username', '')).strip()
            p = str(r.get('Password', '')).strip()
            if u.lower() == username.lower() and p == password:
                return {
                    'username': u,
                    'role'    : str(r.get('Role', 'driver')).strip().lower(),
                    'name'    : str(r.get('Name', u)).strip()
                }
    except Exception as e:
        print("Login check error:", e)
    return None


# ================================================================
#  DATE HELPERS
# ================================================================

def parse_date(value):
    if not value:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    for fmt in ('%d-%m-%Y', '%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None

def fmt_date(value):
    if not value:
        return ''
    if isinstance(value, datetime):
        return value.strftime('%d-%m-%Y')
    if isinstance(value, date):
        return value.strftime('%d-%m-%Y')
    return str(value).strip()


# ================================================================
#  DATA HELPERS  (all read from Google Sheets)
# ================================================================

def get_scan_rows():
    """List of dicts from Sheet1: date, time, driver, action, uid (cached)"""
    now = time.time()
    if _data_cache['scans'] is not None and (now - _data_cache['scans_time']) < CACHE_TTL:
        return _data_cache['scans']
    try:
        if scan_ws is None:
            return []
        rows = scan_ws.get_all_values()
        out  = []
        for r in rows[1:]:
            if len(r) >= 5 and r[4].strip():
                out.append({
                    'date'  : r[0].strip(),
                    'time'  : r[1].strip(),
                    'driver': r[2].strip(),
                    'action': r[3].strip(),
                    'uid'   : r[4].strip()
                })
        _data_cache['scans'] = out
        _data_cache['scans_time'] = now
        return out
    except Exception as e:
        print("Error reading scans:", e)
        if _data_cache['scans'] is not None:
            return _data_cache['scans']
        return []

def get_map_rows():
    """List of dicts from Customer Map: date, time, driver, action, customer (cached)"""
    now = time.time()
    if _data_cache['map'] is not None and (now - _data_cache['map_time']) < CACHE_TTL:
        return _data_cache['map']
    try:
        if map_ws is None:
            return []
        rows = map_ws.get_all_values()
        out  = []
        for r in rows[1:]:
            if len(r) >= 7 and r[6].strip():
                out.append({
                    'date'    : r[0].strip(),
                    'time'    : r[1].strip(),
                    'driver'  : r[2].strip(),
                    'action'  : r[3].strip(),
                    'customer': r[6].strip()
                })
        _data_cache['map'] = out
        _data_cache['map_time'] = now
        return out
    except Exception as e:
        print("Error reading map:", e)
        if _data_cache['map'] is not None:
            return _data_cache['map']
        return []

def build_batch_map():
    """dict: 'date||time||driver||action' → customer"""
    result = {}
    for r in get_map_rows():
        key = f"{r['date']}||{r['time']}||{r['driver']}||{r['action']}"
        result[key] = r['customer']
    return result

def build_events():
    """Sorted list of scan events enriched with customer name"""
    batch_map = build_batch_map()
    events    = []
    for r in get_scan_rows():
        key      = f"{r['date']}||{r['time']}||{r['driver']}||{r['action']}"
        customer = batch_map.get(key)
        if not customer:
            continue
        events.append({**r, 'customer': customer, 'date_obj': parse_date(r['date'])})
    events.sort(key=lambda x: (x['date_obj'] or date.min, x['time']))
    return events

def build_outstanding():
    """Outstanding cylinders per customer"""
    events          = build_events()
    cylinder_owner  = {}
    customer_stats  = {}

    for ev in events:
        c = ev['customer']
        if c not in customer_stats:
            customer_stats[c] = {'total_delivered': 0, 'total_collected': 0, 'last_activity': ev['date']}
        customer_stats[c]['last_activity'] = ev['date']

        if ev['action'] == 'Delivery':
            cylinder_owner[ev['uid']] = c
            customer_stats[c]['total_delivered'] += 1
        elif ev['action'] == 'Collection':
            if cylinder_owner.get(ev['uid']) == c:
                del cylinder_owner[ev['uid']]
            customer_stats[c]['total_collected'] += 1

    customer_outstanding = {}
    for uid, cust in cylinder_owner.items():
        customer_outstanding.setdefault(cust, []).append(uid)

    result = []
    for cust, stats in customer_stats.items():
        uid_list = customer_outstanding.get(cust, [])
        result.append({
            'customer'       : cust,
            'total_delivered': stats['total_delivered'],
            'total_collected': stats['total_collected'],
            'outstanding'    : len(uid_list),
            'cylinder_uids'  : uid_list,              # list — for template iteration
            'uids'           : ', '.join(uid_list) if uid_list else '—',  # string fallback
            'last_activity'  : stats['last_activity']
        })

    result.sort(key=lambda x: x['outstanding'], reverse=True)
    return result

def build_aging():
    """Days outstanding per cylinder currently with a customer"""
    events                 = build_events()
    cylinder_owner         = {}
    cylinder_delivery_date = {}

    for ev in events:
        if ev['action'] == 'Delivery':
            cylinder_owner[ev['uid']]         = ev['customer']
            cylinder_delivery_date[ev['uid']] = ev['date_obj']
        elif ev['action'] == 'Collection':
            if cylinder_owner.get(ev['uid']) == ev['customer']:
                cylinder_owner.pop(ev['uid'], None)
                cylinder_delivery_date.pop(ev['uid'], None)

    today  = date.today()
    result = []
    for uid, cust in cylinder_owner.items():
        d_date   = cylinder_delivery_date.get(uid)
        days_out = (today - d_date).days if d_date else None
        status   = (
            'overdue' if days_out and days_out > 30 else
            'warning' if days_out and days_out > 7  else
            'ok'
        )
        result.append({
            'cylinder_uid' : uid,          # matches aging.html template
            'uid'          : uid,          # kept for backward compat
            'customer'     : cust,
            'delivered_on' : fmt_date(d_date),   # matches aging.html template
            'delivery_date': fmt_date(d_date),   # kept for backward compat
            'days_out'     : days_out,
            'status'       : status
        })

    result.sort(key=lambda x: (x['days_out'] or 0), reverse=True)
    return result

def build_driver_stats():
    """Delivery / collection stats per driver"""
    scan_rows   = get_scan_rows()
    today       = date.today()
    week_start  = today - timedelta(days=today.weekday())
    month_start = today.replace(day=1)
    stats       = {}

    for r in scan_rows:
        d = r['driver']
        if d not in stats:
            stats[d] = {
                'driver'           : d,
                'total_deliveries' : 0,
                'total_collections': 0,
                'this_week'        : 0,
                'this_month'       : 0,
                'last_active'      : r['date']
            }
        stats[d]['last_active'] = r['date']
        rd = parse_date(r['date'])

        if r['action'] == 'Delivery':
            stats[d]['total_deliveries'] += 1
            if rd and rd >= week_start : stats[d]['this_week']  += 1
            if rd and rd >= month_start: stats[d]['this_month'] += 1
        elif r['action'] == 'Collection':
            stats[d]['total_collections'] += 1

    result = list(stats.values())
    result.sort(key=lambda x: x['total_deliveries'] + x['total_collections'], reverse=True)
    return result

def build_daily_movement():
    """Cylinders delivered vs collected per day (last 30 days)"""
    scan_rows = get_scan_rows()
    daily     = {}

    for r in scan_rows:
        d = r['date']
        if d not in daily:
            daily[d] = {'date': d, 'delivered': 0, 'collected': 0, 'date_obj': parse_date(d)}
        if   r['action'] == 'Delivery'  : daily[d]['delivered'] += 1  # matches movement.html
        elif r['action'] == 'Collection': daily[d]['collected'] += 1

    result = sorted(daily.values(), key=lambda x: x['date_obj'] or date.min)
    return result[-30:]

def get_cylinder_history(uid):
    """Full event history for a single cylinder UID"""
    batch_map = build_batch_map()
    history   = []

    for r in get_scan_rows():
        if r['uid'].upper() == uid.strip().upper():
            key      = f"{r['date']}||{r['time']}||{r['driver']}||{r['action']}"
            customer = batch_map.get(key, '(Not mapped yet)')
            history.append({**r, 'customer': customer})

    history.sort(key=lambda x: (parse_date(x['date']) or date.min, x['time']))
    return history

def get_cylinder_status(uid):
    uid_upper = uid.strip().upper()
    scan_rows = get_scan_rows()
    history = [r for r in scan_rows if r['uid'].strip().upper() == uid_upper]
    if not history:
        return {'status': 'Empty', 'owner': None, 'date': None}
    
    history.sort(key=lambda x: (parse_date(x['date']) or date.min, x['time']))
    last_event = history[-1]
    action = last_event['action']
    
    if action == 'Filling':
        return {'status': 'Filled', 'owner': 'Depot', 'date': last_event['date']}
    elif action == 'Delivery':
        batch_map = build_batch_map()
        key = f"{last_event['date']}||{last_event['time']}||{last_event['driver']}||{last_event['action']}"
        customer = batch_map.get(key, '(Unknown Customer)')
        return {'status': 'Delivered', 'owner': customer, 'date': last_event['date']}
    elif action == 'Collection':
        return {'status': 'Empty', 'owner': 'Depot', 'date': last_event['date']}
        
    return {'status': 'Empty', 'owner': None, 'date': None}

@app.route('/api/cylinder_status/<uid>')
def api_cylinder_status(uid):
    status_data = get_cylinder_status(uid)
    return jsonify(status_data)


# ================================================================
#  AUTH ROUTES
# ================================================================

@app.route('/login', methods=['GET', 'POST'])
@app.route('/admin/login', methods=['GET', 'POST'])   # alias — login.html posts here
def login():
    if 'user' in session:
        if session['user']['role'] in ['manager', 'owner']:
            return redirect('/admin')
        else:
            return redirect('/')

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        user     = check_login(username, password)
        if user:
            session['user'] = user
            if user['role'] in ['manager', 'owner']:
                return redirect('/admin')
            else:
                return redirect('/')
        return render_template('login.html', error='Invalid username or password')
    return render_template('login.html', error=None)

@app.route('/logout')
@app.route('/admin/logout')
def logout():
    session.clear()
    return redirect('/login')


# ================================================================
#  ADMIN ROUTES
# ================================================================

@app.route('/admin')
@admin_required
def admin():
    return redirect('/admin/dashboard')

@app.route('/admin/dashboard')
@admin_required
def admin_dashboard():
    outstanding  = build_outstanding()
    total_out    = sum(c['outstanding'] for c in outstanding)
    cust_count   = len([c for c in outstanding if c['outstanding'] > 0])
    high_count   = len([c for c in outstanding if c['outstanding'] > 10])

    today_str    = date.today().strftime('%d-%m-%Y')
    scan_rows    = get_scan_rows()
    today_scans  = len([r for r in scan_rows if r['date'] == today_str])

    top_customers = [c for c in outstanding if c['outstanding'] > 0][:5]

    return render_template('dashboard.html',
        user          = session['user'],
        total_out     = total_out,
        cust_count    = cust_count,
        high_count    = high_count,
        today_scans   = today_scans,
        top_customers = top_customers
    )

@app.route('/admin/outstanding')
@admin_required
def admin_outstanding():
    data = build_outstanding()
    return render_template('outstanding.html',
        user = session['user'],
        data = data,
        total_out = sum(c['outstanding'] for c in data)
    )

@app.route('/admin/aging')
@admin_required
def admin_aging():
    data     = build_aging()
    overdue  = [r for r in data if r['status'] == 'overdue']
    warning  = [r for r in data if r['status'] == 'warning']
    ok       = [r for r in data if r['status'] == 'ok']
    return render_template('aging.html',
        user    = session['user'],
        data    = data,
        overdue = len(overdue),
        warning = len(warning),
        ok_count= len(ok)
    )

@app.route('/admin/drivers')
@admin_required
def admin_drivers():
    data = build_driver_stats()
    return render_template('drivers.html',
        user = session['user'],
        data = data
    )

@app.route('/admin/movement')
@admin_required
def admin_movement():
    data        = build_daily_movement()
    labels      = [r['date'] for r in data]
    delivered   = [r['delivered'] for r in data]   # fixed: was 'out'
    collected   = [r['collected'] for r in data]
    return render_template('movement.html',
        user      = session['user'],
        data      = data,
        labels    = labels,
        delivered = delivered,
        collected = collected
    )

@app.route('/admin/search')
@admin_required
def admin_search():
    uid     = request.args.get('uid', '').strip()
    history = get_cylinder_history(uid) if uid else []
    current = None
    if history:
        # Current status = last delivery not yet collected
        events = build_events()
        cyl_owner = {}
        for ev in events:
            if ev['uid'].upper() == uid.upper():
                if ev['action'] == 'Delivery':
                    cyl_owner['owner'] = ev['customer']
                    cyl_owner['since'] = ev['date']
                elif ev['action'] == 'Collection':
                    cyl_owner = {}
        current = cyl_owner if cyl_owner else None

    return render_template('search.html',
        user    = session['user'],
        uid     = uid,
        history = history,
        current = current
    )
def get_customer_map_batches():
    if map_ws is None:
        return []
    try:
        rows = map_ws.get_all_values()
        if len(rows) < 2:
            return []
        out = []
        for idx, r in enumerate(rows[1:]):
            # Columns: Date(0), Time(1), Driver(2), Action(3), Count(4), UIDs(5), Customer(6), Send Receipt?(7), Receipt Status(8)
            if len(r) >= 4:
                out.append({
                    'row_num': idx + 2,
                    'date': r[0].strip(),
                    'time': r[1].strip(),
                    'driver': r[2].strip(),
                    'action': r[3].strip(),
                    'count': r[4].strip() if len(r) > 4 else '0',
                    'uids': r[5].strip() if len(r) > 5 else '',
                    'customer': r[6].strip() if len(r) > 6 else '',
                    'send_receipt': r[7].strip() if len(r) > 7 else 'FALSE',
                    'status': r[8].strip() if len(r) > 8 else ''
                })
        # Show newest batches first (reverse order)
        out.reverse()
        return out
    except Exception as e:
        print("Error reading Customer Map:", e)
        return []

@app.route('/admin/receipts')
@admin_required
def admin_receipts():
    batches = get_customer_map_batches()
    customers = get_customer_names()
    emails = get_customer_emails()
    
    # Calculate counts in Python to avoid template filter errors
    sent_count = sum(1 for b in batches if str(b.get('status', '')).startswith('Sent'))
    pending_count = sum(1 for b in batches if b.get('customer') and not str(b.get('status', '')).startswith('Sent'))
    unmapped_count = sum(1 for b in batches if not b.get('customer'))
    
    return render_template('receipts.html',
        user=session['user'],
        batches=batches,
        customers=customers,
        emails=emails,
        sent_count=sent_count,
        pending_count=pending_count,
        unmapped_count=unmapped_count
    )

@app.route('/admin/update_mapping', methods=['POST'])
@admin_required
def admin_update_mapping():
    data = request.get_json() or {}
    date_val = data.get('date', '').strip()
    time_val = data.get('time', '').strip()
    driver_val = data.get('driver', '').strip()
    action_val = data.get('action', '').strip()
    customer_val = data.get('customer', '').strip()

    if not (date_val and time_val and driver_val and action_val):
        return jsonify({'status': 'Error', 'message': 'Missing batch identifiers'})

    try:
        if map_ws is None:
            return jsonify({'status': 'Error', 'message': 'Sheet offline'})
        
        rows = map_ws.get_all_values()
        for idx, r in enumerate(rows):
            if idx == 0:
                continue
            if len(r) >= 4:
                if (r[0].strip() == date_val and 
                    r[1].strip() == time_val and 
                    r[2].strip() == driver_val and 
                    r[3].strip() == action_val):
                    map_ws.update_cell(idx + 1, 7, customer_val) # Column G is 7
                    clear_cache()
                    return jsonify({'status': 'Success', 'message': 'Customer mapped successfully'})
        return jsonify({'status': 'Error', 'message': 'Batch row not found'})
    except Exception as e:
        return jsonify({'status': 'Error', 'message': str(e)})

@app.route('/admin/send_receipt', methods=['POST'])
@admin_required
def admin_send_receipt():
    data = request.get_json() or {}
    date_val = data.get('date', '').strip()
    time_val = data.get('time', '').strip()
    driver_val = data.get('driver', '').strip()
    action_val = data.get('action', '').strip()

    if not (date_val and time_val and driver_val and action_val):
        return jsonify({'status': 'Error', 'message': 'Missing batch identifiers'})

    try:
        if map_ws is None:
            return jsonify({'status': 'Error', 'message': 'Sheet offline'})
        
        rows = map_ws.get_all_values()
        for idx, r in enumerate(rows):
            if idx == 0:
                continue
            if len(r) >= 4:
                if (r[0].strip() == date_val and 
                    r[1].strip() == time_val and 
                    r[2].strip() == driver_val and 
                    r[3].strip() == action_val):
                    
                    customer_name = r[6].strip() if len(r) > 6 else ''
                    if not customer_name:
                        return jsonify({'status': 'Error', 'message': 'Please map a customer first'})
                    
                    # Validate if email exists for this customer
                    emails = get_customer_emails()
                    email = emails.get(customer_name, '').strip()
                    if not email or '@' not in email:
                        return jsonify({'status': 'Error', 'message': f"Missing Email: Please enter an email address for {customer_name} in the Customers page first."})
                    
                    # Batch update Column H (Send Receipt? = TRUE) and Column I (Status = Sending...)
                    # Column H is 8, Column I is 9. This single call executes instantly.
                    map_ws.update(f"H{idx + 1}:I{idx + 1}", [[True, "Sending..."]], value_input_option='USER_ENTERED')
                    clear_cache()
                    return jsonify({'status': 'Success', 'message': 'Receipt trigger sent'})
        return jsonify({'status': 'Error', 'message': 'Batch row not found'})
    except Exception as e:
        return jsonify({'status': 'Error', 'message': str(e)})

@app.route('/admin/receipt_status')
@admin_required
def admin_receipt_status():
    date_val = request.args.get('date', '').strip()
    time_val = request.args.get('time', '').strip()
    driver_val = request.args.get('driver', '').strip()
    action_val = request.args.get('action', '').strip()

    if not (date_val and time_val and driver_val and action_val):
        return jsonify({'status': 'Error', 'message': 'Missing batch identifiers'})

    try:
        if map_ws is None:
            return jsonify({'status': 'Error', 'message': 'Sheet offline'})
        
        rows = map_ws.get_all_values()
        for r in rows[1:]:
            if len(r) >= 4:
                if (r[0].strip() == date_val and 
                    r[1].strip() == time_val and 
                    r[2].strip() == driver_val and 
                    r[3].strip() == action_val):
                    send_receipt = r[7].strip() if len(r) > 7 else 'FALSE'
                    status = r[8].strip() if len(r) > 8 else ''
                    return jsonify({
                        'status': 'Success',
                        'send_receipt': send_receipt,
                        'receipt_status': status
                    })
        return jsonify({'status': 'Error', 'message': 'Batch row not found'})
    except Exception as e:
        return jsonify({'status': 'Error', 'message': str(e)})


# ── Inventory & Bulk Tanks Calculations ─────────────────────────

def parse_mix_ratio(mix_ratio_str, gas_type):
    """Parses mix ratio percentages from a string."""
    if not mix_ratio_str:
        if gas_type == 'ACM':
            return {'Argon': 0.9073, 'CO2': 0.1928}  # default ACM 90/10 ratio
        elif gas_type == 'AHM':
            return {'Argon': 0.9886}  # default AHM
        return {}
    
    mix_ratio_str = str(mix_ratio_str).upper()
    arg_match = re.search(r'(\d+)%\s*(?:ARG|ARGON)', mix_ratio_str)
    co2_match = re.search(r'(\d+)%\s*(?:CO2|CARBON)', mix_ratio_str)
    n2_match = re.search(r'(\d+)%\s*(?:N2|NITROGEN)', mix_ratio_str)
    oxy_match = re.search(r'(\d+)%\s*(?:O2|OXY|OXYGEN)', mix_ratio_str)
    
    parts = {}
    if arg_match: parts['Argon'] = float(arg_match.group(1)) / 100.0
    if co2_match: parts['CO2'] = float(co2_match.group(1)) / 100.0
    if n2_match: parts['N2'] = float(n2_match.group(1)) / 100.0
    if oxy_match: parts['Oxygen'] = float(oxy_match.group(1)) / 100.0
    
    if not parts:
        slash_match = re.search(r'(\d+)\s*/\s*(\d+)', mix_ratio_str)
        if slash_match:
            val1 = float(slash_match.group(1))
            val2 = float(slash_match.group(2))
            if gas_type == 'ACM':
                if val1 == 90 and val2 == 10:
                    return {'Argon': 0.9073, 'CO2': 0.1928}
                elif val1 == 80 and val2 == 20:
                    return {'Argon': 0.80, 'CO2': 0.20}
                else:
                    total = val1 + val2
                    return {'Argon': val1 / total, 'CO2': val2 / total}
            elif gas_type == 'AHM':
                total = val1 + val2
                return {'Argon': val1 / total}
                
    if not parts:
        if gas_type == 'ACM':
            return {'Argon': 0.9073, 'CO2': 0.1928}
        elif gas_type == 'AHM':
            return {'Argon': 0.9886}
            
    return parts

def get_tank_data_for_date(target_date_str):
    """Fetches opening stock, dead volume, capacity, and unit for bulk tanks."""
    global bulk_tanks_ws
    defaults = {
        'Argon': {'opening': 5664.03, 'dead_volume': 500.0, 'capacity': 10000.0, 'unit': 'Cum'},
        'CO2': {'opening': 10941.05, 'dead_volume': 200.0, 'capacity': 15000.0, 'unit': 'KG'},
        'N2': {'opening': 4271.20, 'dead_volume': 300.0, 'capacity': 8000.0, 'unit': 'Cum'},
        'Oxygen': {'opening': 9215.30, 'dead_volume': 400.0, 'capacity': 12000.0, 'unit': 'Cum'}
    }
    try:
        if bulk_tanks_ws is None and doc:
            try: bulk_tanks_ws = doc.worksheet(BULK_TANKS_NAME)
            except Exception: pass
        if bulk_tanks_ws is None:
            return defaults
        
        rows = bulk_tanks_ws.get_all_values()
        if len(rows) < 2:
            return defaults
            
        records = []
        for r in rows[1:]:
            if len(r) >= 6 and r[0].strip():
                try:
                    records.append({
                        'date': r[0].strip(),
                        'date_obj': parse_date(r[0]),
                        'gas': r[1].strip(),
                        'opening': float(r[2].strip() or 0.0),
                        'dead_volume': float(r[3].strip() or 0.0),
                        'capacity': float(r[4].strip() or 0.0),
                        'unit': r[5].strip()
                    })
                except ValueError:
                    continue
                    
        target_records = [rec for rec in records if rec['date'] == target_date_str]
        if target_records:
            out = {}
            for rec in target_records:
                out[rec['gas']] = {
                    'opening': rec['opening'],
                    'dead_volume': rec['dead_volume'],
                    'capacity': rec['capacity'],
                    'unit': rec['unit']
                }
            for g in defaults:
                if g not in out:
                    out[g] = defaults[g]
            return out
            
        # Rollover closing stock of previous date
        sorted_records = sorted([r for r in records if r['date_obj'] is not None], key=lambda x: x['date_obj'])
        if not sorted_records:
            return defaults
            
        target_date_obj = parse_date(target_date_str)
        prev_records = []
        if target_date_obj:
            prev_records = [r for r in sorted_records if r['date_obj'] < target_date_obj]
            
        if prev_records:
            latest_prev_date = prev_records[-1]['date']
            prev_date_records = [r for r in prev_records if r['date'] == latest_prev_date]
        else:
            latest_date = sorted_records[-1]['date']
            prev_date_records = [r for r in sorted_records if r['date'] == latest_date]
            latest_prev_date = latest_date
            
        prev_used = calculate_used_today(latest_prev_date)
        out = {}
        for rec in prev_date_records:
            used = prev_used.get(rec['gas'], 0.0)
            closing = max(0.0, rec['opening'] - used)
            out[rec['gas']] = {
                'opening': round(closing, 2),
                'dead_volume': rec['dead_volume'],
                'capacity': rec['capacity'],
                'unit': rec['unit']
            }
        for g in defaults:
            if g not in out:
                out[g] = defaults[g]
        return out
        
    except Exception as e:
        print("Error fetching bulk tank data:", e)
        return defaults

def calculate_used_today(target_date_str):
    """Calculates used gas today from fillings logged in scan log."""
    global scan_ws
    used = {'Argon': 0.0, 'CO2': 0.0, 'N2': 0.0, 'Oxygen': 0.0}
    try:
        if scan_ws is None and doc:
            try: scan_ws = doc.worksheet(SCAN_SHEET_NAME)
            except Exception: pass
        if scan_ws is None:
            return used
            
        rows = scan_ws.get_all_values()
        if len(rows) < 2:
            return used
            
        fillings = []
        for r in rows[1:]:
            if len(r) >= 5:
                # Col A: Date, Col D: Action, Col E: UID
                if r[0].strip() == target_date_str and r[3].strip() == 'Filling':
                    fillings.append(r[4].strip().upper())
                    
        if not fillings:
            return used
            
        cylinders = get_all_cylinders()
        maintenance = get_all_maintenance()
        cyl_map = {c['uid'].upper(): c for c in cylinders}
        
        for uid in fillings:
            cyl = cyl_map.get(uid)
            if cyl:
                gas_type = cyl['gas_type'].upper()
                cyl_type = cyl['cylinder_type'].capitalize()
            else:
                cyl_type = 'Standard'
                if uid.startswith('ARG'): gas_type = 'ARG'
                elif uid.startswith('CO2'): gas_type = 'CO2'
                elif uid.startswith('N2'): gas_type = 'N2'
                elif uid.startswith('OXY'): gas_type = 'OXY'
                elif uid.startswith('ACM'): gas_type = 'ACM'
                else: continue
                
            maint = maintenance.get(uid, {})
            try:
                capacity = float(maint.get('gas_capacity') or 0.0) if maint.get('gas_capacity') else None
            except ValueError:
                capacity = None
                
            mix_ratio = maint.get('mix_ratio', '')
            
            if capacity is None:
                if gas_type == 'ARG':
                    capacity = 7.0 if cyl_type == 'Standard' else 0.88
                elif gas_type == 'CO2':
                    capacity = 30.0
                elif gas_type == 'N2':
                    capacity = 7.0 if cyl_type == 'Standard' else 0.88
                elif gas_type == 'OXY':
                    capacity = 7.0 if cyl_type == 'Standard' else 0.88
                elif gas_type == 'ACM':
                    capacity = 7.0
                elif gas_type == 'AHM':
                    capacity = 7.0
                else:
                    capacity = 0.0
                    
            if gas_type == 'ARG':
                used['Argon'] += capacity
            elif gas_type == 'CO2':
                used['CO2'] += capacity
            elif gas_type == 'N2':
                used['N2'] += capacity
            elif gas_type == 'OXY':
                used['Oxygen'] += capacity
            elif gas_type == 'ACM':
                ratios = parse_mix_ratio(mix_ratio, 'ACM')
                used['Argon'] += capacity * ratios.get('Argon', 0.80)
                used['CO2'] += capacity * ratios.get('CO2', 0.20)
            elif gas_type == 'AHM':
                ratios = parse_mix_ratio(mix_ratio, 'AHM')
                used['Argon'] += capacity * ratios.get('Argon', 0.92)
                
        for k in used:
            used[k] = round(used[k], 2)
        return used
    except Exception as e:
        print("Error calculating used today:", e)
        return used

def calculate_fleet_sizes():
    """Calculates fleet size (total cylinder count) by gas type."""
    cylinders = get_all_cylinders()
    fleet = {'Argon': 0, 'CO2': 0, 'N2': 0, 'Oxygen': 0}
    for c in cylinders:
        gas = c.get('gas_type', '').upper()
        if gas in ('ARG', 'ACM', 'AHM'):
            fleet['Argon'] += 1
        elif gas == 'CO2':
            fleet['CO2'] += 1
        elif gas in ('N2', 'N2D'):
            fleet['N2'] += 1
        elif gas == 'OXY':
            fleet['Oxygen'] += 1
    return fleet

def calculate_table1_filled_inventory():
    """Calculates Table 1 — Filled Cylinder Inventory."""
    cylinders = get_all_cylinders()
    maintenance = get_all_maintenance()
    
    products_config = [
        {'id': 'arg_pura', 'name': 'ARG Pura', 'gas_type': 'ARG', 'cylinder_type': 'Standard', 'gas_per_cyl': 7.0, 'unit': 'Cum'},
        {'id': 'acm_90_10', 'name': 'ACM (90.10)_', 'gas_type': 'ACM', 'cylinder_type': 'Standard', 'gas_per_cyl': 6.3512, 'unit': 'Cum'},
        {'id': 'co2_90_10', 'name': 'Co2 (90.10)_', 'gas_type': 'ACM', 'cylinder_type': 'Standard', 'gas_per_cyl': 1.35, 'unit': 'KG', 'is_virtual': True},
        {'id': 'co2_pure', 'name': 'Co2', 'gas_type': 'CO2', 'cylinder_type': 'Standard', 'gas_per_cyl': 30.0, 'unit': 'KG'},
        {'id': 'n2_cyl', 'name': 'N2 Cyl', 'gas_type': 'N2', 'cylinder_type': 'Standard', 'gas_per_cyl': 7.0, 'unit': 'Cum'},
        {'id': 'oxygen_pure', 'name': 'OXYGEN', 'gas_type': 'OXY', 'cylinder_type': 'Standard', 'gas_per_cyl': 7.0, 'unit': 'Cum'},
        {'id': 'ahm_92_08', 'name': 'AHM(92.08)', 'gas_type': 'AHM', 'cylinder_type': 'Standard', 'gas_per_cyl': 6.92, 'unit': 'Cum'},
        {'id': 'ahm_98_02', 'name': 'AHM (98.02)', 'gas_type': 'AHM', 'cylinder_type': 'Standard', 'gas_per_cyl': 6.98, 'unit': 'Cum'},
        {'id': 'arg_dura', 'name': 'ARG Dura', 'gas_type': 'ARG', 'cylinder_type': 'Dura', 'gas_per_cyl': 0.0, 'unit': 'Cum'},
        {'id': 'n2_dura', 'name': 'N2Dura', 'gas_type': 'N2', 'cylinder_type': 'Dura', 'gas_per_cyl': 0.88, 'unit': 'Cum'},
        {'id': 'oxygen_dura', 'name': 'Oxygen Dura', 'gas_type': 'OXY', 'cylinder_type': 'Dura', 'gas_per_cyl': 0.0, 'unit': 'Cum'}
    ]
    
    t1_rows = {p['id']: {**p, 'filled_count': 0, 'total_gas': 0.0} for p in products_config}
    
    for c in cylinders:
        if c.get('status') != 'Filled':
            continue
            
        uid = c.get('uid', '').strip().upper()
        gas_type = c.get('gas_type', '').upper()
        cyl_type = c.get('cylinder_type', '').capitalize()
        maint = maintenance.get(uid, {})
        
        try:
            capacity = float(maint.get('gas_capacity') or 0.0) if maint.get('gas_capacity') else None
        except ValueError:
            capacity = None
            
        mix_ratio = maint.get('mix_ratio', '')
        
        pid = None
        if gas_type == 'ARG':
            pid = 'arg_pura' if cyl_type == 'Standard' else 'arg_dura'
        elif gas_type == 'CO2':
            pid = 'co2_pure'
        elif gas_type == 'N2':
            pid = 'n2_cyl' if cyl_type == 'Standard' else 'n2_dura'
        elif gas_type == 'OXY':
            pid = 'oxygen_pure' if cyl_type == 'Standard' else 'oxygen_dura'
        elif gas_type == 'ACM':
            pid = 'acm_90_10'
        elif gas_type == 'AHM':
            if '98' in mix_ratio:
                pid = 'ahm_98_02'
            else:
                pid = 'ahm_92_08'
                
        if pid:
            if pid == 'acm_90_10':
                t1_rows['acm_90_10']['filled_count'] += 1
                t1_rows['co2_90_10']['filled_count'] += 1
                cap_val = capacity if capacity is not None else 7.0
                t1_rows['acm_90_10']['total_gas'] += cap_val * 0.9073
                t1_rows['co2_90_10']['total_gas'] += cap_val * 0.1928
            else:
                t1_rows[pid]['filled_count'] += 1
                cap_val = capacity if capacity is not None else t1_rows[pid]['gas_per_cyl']
                t1_rows[pid]['total_gas'] += cap_val
                
    results = []
    total_physical_filled = 0
    total_cum = 0.0
    total_kg = 0.0
    
    for p in products_config:
        res = t1_rows[p['id']]
        res['total_gas'] = round(res['total_gas'], 3)
        results.append(res)
        
        if not res.get('is_virtual'):
            total_physical_filled += res['filled_count']
        if res['unit'] == 'Cum':
            total_cum += res['total_gas']
        elif res['unit'] == 'KG':
            total_kg += res['total_gas']
            
    return {
        'rows': results,
        'total_filled': total_physical_filled,
        'total_cum': round(total_cum, 3),
        'total_kg': round(total_kg, 3)
    }

def calculate_table2_bulk_inventory(target_date_str):
    """Calculates Table 2 — Bulk Tank Inventory."""
    tanks = get_tank_data_for_date(target_date_str)
    used = calculate_used_today(target_date_str)
    fleet = calculate_fleet_sizes()
    
    gases = [
        {'id': 'Argon', 'name': 'Argon'},
        {'id': 'CO2', 'name': 'CO2'},
        {'id': 'N2', 'name': 'N2'},
        {'id': 'Oxygen', 'name': 'Oxygen'}
    ]
    
    rows = []
    for g in gases:
        tank_data = tanks.get(g['id'], {'opening': 0.0, 'dead_volume': 0.0, 'capacity': 10000.0, 'unit': 'Cum'})
        used_today = used.get(g['id'], 0.0)
        closing = round(max(0.0, tank_data['opening'] - used_today), 2)
        usable = round(max(0.0, tank_data['opening'] - tank_data['dead_volume']), 2)
        
        rows.append({
            'gas': g['name'],
            'opening': tank_data['opening'],
            'used_today': used_today,
            'closing': closing,
            'usable': usable,
            'dead_volume': tank_data['dead_volume'],
            'capacity': tank_data['capacity'],
            'fleet': fleet.get(g['id'], 0),
            'unit': tank_data['unit']
        })
        
    return rows

def update_tank_opening_stock(date_str, gas_name, opening, capacity, dead_volume, unit):
    """Updates opening stock, capacity, and dead volume for a specific gas and date."""
    global bulk_tanks_ws
    try:
        if bulk_tanks_ws is None and doc:
            try: bulk_tanks_ws = doc.worksheet(BULK_TANKS_NAME)
            except Exception: pass
        if bulk_tanks_ws is None:
            return False
            
        rows = bulk_tanks_ws.get_all_values()
        row_num = None
        for idx, r in enumerate(rows):
            if idx == 0: continue
            if len(r) >= 2 and r[0].strip() == date_str and r[1].strip() == gas_name:
                row_num = idx + 1
                break
                
        row_data = [date_str, gas_name, float(opening), float(dead_volume), float(capacity), unit]
        if row_num:
            bulk_tanks_ws.update(f'A{row_num}:F{row_num}', [row_data])
        else:
            bulk_tanks_ws.append_row(row_data)
        return True
    except Exception as e:
        print("Error updating tank opening stock:", e)
        return False


# ================================================================
#  CYLINDER REGISTRY ROUTES
# ================================================================


@app.route('/admin/cylinders')
@admin_required
def admin_cylinders():
    cylinders = merge_cylinder_data()
    # Collect unique filter options
    gas_types = sorted(set(c['gas_type'] for c in cylinders if c.get('gas_type')))
    statuses  = ['Empty', 'Filled', 'Delivered']
    return render_template('cylinders.html',
        user       = session['user'],
        cylinders  = cylinders,
        gas_types  = gas_types,
        statuses   = statuses,
        total      = len(cylinders),
        active     = sum(1 for c in cylinders if c.get('status') in ('Active', 'Empty', 'Filled', 'Delivered')),
        overdue    = sum(1 for c in cylinders if c.get('hydro_badge') == 'Overdue'),
        due_soon   = sum(1 for c in cylinders if c.get('hydro_badge') == 'Due Soon'),
    )

@app.route('/admin/cylinders/add', methods=['GET', 'POST'])
@admin_required
def admin_cylinders_add():
    global cyl_ws, cyl_maint_ws
    if request.method == 'POST':
        data = request.form
        uid = data.get('uid', '').strip()
        if not uid:
            return render_template('cylinders_form.html',
                user=session['user'], mode='add',
                error='Cylinder UID is required.', form=data)
        # Check duplicate
        existing = get_all_cylinders()
        if any(c['uid'].upper() == uid.upper() for c in existing):
            return render_template('cylinders_form.html',
                user=session['user'], mode='add',
                error=f'Cylinder UID "{uid}" already exists.', form=data)
        try:
            # Refresh worksheet references if needed
            if cyl_ws is None and doc:
                try: cyl_ws = doc.worksheet(CYLINDER_SHEET_NAME)
                except Exception: pass
            if cyl_maint_ws is None and doc:
                try: cyl_maint_ws = doc.worksheet(CYLINDER_MAINT_NAME)
                except Exception: pass

            if cyl_ws is None:
                return render_template('cylinders_form.html',
                    user=session['user'], mode='add',
                    error='Cylinders sheet not found. Please run Setup Registry Sheets from Google Sheets menu first.', form=data)

            # Write to Cylinders sheet
            cyl_ws.append_row([
                uid,
                data.get('gas_type', '').strip(),
                data.get('cylinder_type', '').strip(),
                data.get('owner', 'Depot').strip(),
                data.get('status', 'Active').strip(),
                data.get('location', 'Depot').strip(),
                date.today().strftime('%d-%m-%Y'),
            ])
            # Write to Cylinder Maintenance sheet
            if cyl_maint_ws:
                cyl_maint_ws.append_row([
                    uid,
                    data.get('water_capacity', '').strip(),
                    data.get('fill_pressure', '').strip(),
                    data.get('gas_capacity', '').strip(),
                    data.get('unit', '').strip(),
                    data.get('is_mixture', 'No').strip(),
                    data.get('mix_ratio', '').strip(),
                    data.get('manufacture_date', '').strip(),
                    data.get('last_hydro_date', '').strip(),
                    data.get('next_hydro_due', '').strip(),
                    data.get('hydro_test_status', '').strip(),
                    data.get('cert_no', '').strip(),
                ])
            return redirect('/admin/cylinders')
        except Exception as e:
            return render_template('cylinders_form.html',
                user=session['user'], mode='add',
                error=str(e), form=data)
    return render_template('cylinders_form.html',
        user=session['user'], mode='add', error=None, form={})

@app.route('/admin/cylinders/<uid>/edit', methods=['GET', 'POST'])
@admin_required
def admin_cylinders_edit(uid):
    cylinders   = get_all_cylinders()
    maintenance = get_all_maintenance()
    cyl  = next((c for c in cylinders if c['uid'].upper() == uid.upper()), None)
    maint = maintenance.get(uid, maintenance.get(uid.upper(), {}))
    if not cyl:
        return redirect('/admin/cylinders')

    if request.method == 'POST':
        data = request.form
        try:
            cyl_row, maint_row = find_cylinder_rows(uid)
            if cyl_row and cyl_ws:
                cyl_ws.update(f'A{cyl_row}:G{cyl_row}', [[
                    uid,
                    data.get('gas_type', '').strip(),
                    data.get('cylinder_type', '').strip(),
                    data.get('owner', '').strip(),
                    data.get('status', 'Active').strip(),
                    data.get('location', '').strip(),
                    cyl.get('last_activity', ''),
                ]])
            if maint_row and cyl_maint_ws:
                cyl_maint_ws.update(f'A{maint_row}:L{maint_row}', [[
                    uid,
                    data.get('water_capacity', '').strip(),
                    data.get('fill_pressure', '').strip(),
                    data.get('gas_capacity', '').strip(),
                    data.get('unit', '').strip(),
                    data.get('is_mixture', 'No').strip(),
                    data.get('mix_ratio', '').strip(),
                    data.get('manufacture_date', '').strip(),
                    data.get('last_hydro_date', '').strip(),
                    data.get('next_hydro_due', '').strip(),
                    data.get('hydro_test_status', '').strip(),
                    data.get('cert_no', '').strip(),
                ]])
            return redirect('/admin/cylinders')
        except Exception as e:
            merged = {**cyl, **maint}
            return render_template('cylinders_form.html',
                user=session['user'], mode='edit',
                error=str(e), form=merged, uid=uid)

    merged = {**cyl, **maint}
    return render_template('cylinders_form.html',
        user=session['user'], mode='edit', error=None, form=merged, uid=uid)

@app.route('/admin/cylinders/<uid>')
@admin_required
def admin_cylinder_detail(uid):
    cylinders   = get_all_cylinders()
    maintenance = get_all_maintenance()
    cyl   = next((c for c in cylinders if c['uid'].upper() == uid.upper()), None)
    if not cyl:
        return redirect('/admin/cylinders')
    maint = maintenance.get(uid, maintenance.get(uid.upper(), {}))
    hydro_badge = compute_hydro_badge(maint.get('next_hydro_due', ''))
    history = get_cylinder_history(uid)
    return render_template('cylinder_detail.html',
        user        = session['user'],
        cyl         = cyl,
        maint       = maint,
        hydro_badge = hydro_badge,
        history     = history,
    )


# ================================================================
#  EXISTING ROUTES (unchanged)
# ================================================================

# ================================================================
#  CUSTOMER PROFILE ROUTES
# ================================================================

def get_all_customer_info():
    """Returns list of dicts from Customers sheet: {id, name, email, phone}"""
    try:
        if customer_ws is None:
            return []
        values = customer_ws.get_all_values()
        if len(values) < 2:
            return []
        out = []
        for i, row in enumerate(values[1:], start=1):
            if len(row) < 2 or not row[1].strip():
                continue
            out.append({
                'id'   : row[0].strip() if len(row) > 0 else f'C{str(i).zfill(3)}',
                'name' : row[1].strip() if len(row) > 1 else '',
                'email': row[2].strip() if len(row) > 2 else '',
                'phone': row[3].strip() if len(row) > 3 else '',
            })
        return out
    except Exception as e:
        print('Error getting customer info:', e)
        return []

def build_customer_outstanding_detail(customer_name):
    """
    Returns list of dicts for cylinders currently with this customer:
    {uid, gas_type, delivered_on, days_out, overdue}
    """
    events = build_events()
    cylinder_owner = {}
    cylinder_delivery_date = {}

    for ev in events:
        if ev['customer'].lower() != customer_name.lower():
            continue
        if ev['action'] == 'Delivery':
            cylinder_owner[ev['uid']] = ev['customer']
            cylinder_delivery_date[ev['uid']] = ev['date_obj']
        elif ev['action'] == 'Collection':
            cylinder_owner.pop(ev['uid'], None)
            cylinder_delivery_date.pop(ev['uid'], None)

    # Build gas type map from Cylinders registry (if available)
    gas_map = {}
    try:
        cyls = get_all_cylinders()
        gas_map = {c['uid'].upper(): c.get('gas_type', '') for c in cyls}
    except Exception:
        pass

    today = date.today()
    result = []
    for uid in cylinder_owner:
        d_date   = cylinder_delivery_date.get(uid)
        days_out = (today - d_date).days if d_date else None
        result.append({
            'uid'         : uid,
            'gas_type'    : gas_map.get(uid.upper(), '—'),
            'delivered_on': fmt_date(d_date),
            'days_out'    : days_out,
            'overdue'     : days_out is not None and days_out > 30,
        })

    result.sort(key=lambda x: x['days_out'] or 0, reverse=True)
    return result

def build_customer_history(customer_name):
    """
    All delivery/collection events for a customer, newest first.
    Returns list of dicts: {date, time, driver, action, uids, count}
    """
    batch_map = build_batch_map()
    scan_rows = get_scan_rows()

    # Collect all batches that belong to this customer
    matching_batches = set()
    for key, cust in batch_map.items():
        if cust.lower() == customer_name.lower():
            matching_batches.add(key)

    # Group scan rows by batch key
    batches = {}
    for r in scan_rows:
        key = f"{r['date']}||{r['time']}||{r['driver']}||{r['action']}"
        if key not in matching_batches:
            continue
        if key not in batches:
            batches[key] = {
                'date'    : r['date'],
                'time'    : r['time'],
                'driver'  : r['driver'],
                'action'  : r['action'],
                'uids'    : [],
                'date_obj': parse_date(r['date']),
            }
        batches[key]['uids'].append(r['uid'])

    result = []
    for b in batches.values():
        result.append({
            **b,
            'count': len(b['uids']),
            'uids_str': ', '.join(b['uids']),
        })

    result.sort(key=lambda x: (x['date_obj'] or date.min, x['time']), reverse=True)
    return result

def build_customer_statement(customer_name):
    """
    Running balance statement for a customer — like a bank statement for cylinders.
    Returns list of dicts: {date, time, driver, action, change, balance, uids_str}
    """
    history = build_customer_history(customer_name)
    # history is newest first — reverse for running balance
    history_asc = list(reversed(history))
    balance = 0
    statement = []
    for event in history_asc:
        if event['action'] == 'Delivery':
            change  = +event['count']
        elif event['action'] == 'Collection':
            change  = -event['count']
        else:
            change  = 0
        balance += change
        statement.append({
            **event,
            'change' : change,
            'balance': balance,
        })
    # Return newest first for display
    return list(reversed(statement))

def build_gas_type_breakdown(outstanding_detail):
    """
    Returns dict of {gas_type: count} from outstanding cylinders.
    """
    breakdown = {}
    for c in outstanding_detail:
        gt = c.get('gas_type') or '—'
        breakdown[gt] = breakdown.get(gt, 0) + 1
    return dict(sorted(breakdown.items()))


@app.route('/admin/customers')
@admin_required
def admin_customers():
    # All customers from Customers sheet
    all_info   = get_all_customer_info()
    info_map   = {c['name'].lower(): c for c in all_info}
    outstanding = build_outstanding()

    # Build aging data for overdue flags
    aging = build_aging()
    overdue_customers = set()
    for a in aging:
        if a['status'] == 'overdue':
            overdue_customers.add(a['customer'].lower())

    # Merge: customer sheet info + outstanding stats
    result = []
    seen = set()
    for o in outstanding:
        name  = o['customer']
        name_l = name.lower()
        seen.add(name_l)
        info = info_map.get(name_l, {})
        result.append({
            'name'            : name,
            'email'           : info.get('email', ''),
            'phone'           : info.get('phone', ''),
            'id'              : info.get('id', ''),
            'outstanding'     : o['outstanding'],
            'total_delivered' : o['total_delivered'],
            'total_collected' : o['total_collected'],
            'last_activity'   : o['last_activity'],
            'has_overdue'     : name_l in overdue_customers,
        })

    # Include customers from Customers sheet who have no scans yet
    for info in all_info:
        if info['name'].lower() not in seen:
            result.append({
                'name'           : info['name'],
                'email'          : info['email'],
                'phone'          : info['phone'],
                'id'             : info['id'],
                'outstanding'    : 0,
                'total_delivered': 0,
                'total_collected': 0,
                'last_activity'  : '—',
                'has_overdue'    : False,
            })

    return render_template('customers_list.html',
        user      = session['user'],
        customers = result,
        total     = len(result),
        with_outstanding = sum(1 for c in result if c['outstanding'] > 0),
        overdue_count    = len(overdue_customers),
    )


@app.route('/admin/customers/<path:customer_name>')
@admin_required
def admin_customer_profile(customer_name):
    # Decode URL-encoded name
    from urllib.parse import unquote
    customer_name = unquote(customer_name)

    # Basic info from Customers sheet
    all_info = get_all_customer_info()
    info = next((c for c in all_info if c['name'].lower() == customer_name.lower()), {})

    # Outstanding detail
    outstanding_detail = build_customer_outstanding_detail(customer_name)
    overdue_cyls  = [c for c in outstanding_detail if c['overdue']]

    # Gas type breakdown
    gas_breakdown = build_gas_type_breakdown(outstanding_detail)

    # History + statement
    history   = build_customer_history(customer_name)
    statement = build_customer_statement(customer_name)

    # Summary stats
    total_delivered = sum(e['count'] for e in history if e['action'] == 'Delivery')
    total_collected = sum(e['count'] for e in history if e['action'] == 'Collection')

    return render_template('customer_profile.html',
        user               = session['user'],
        info               = info,
        customer_name      = customer_name,
        outstanding_detail = outstanding_detail,
        outstanding_count  = len(outstanding_detail),
        overdue_cyls       = overdue_cyls,
        gas_breakdown      = gas_breakdown,
        history            = history,
        statement          = statement,
        total_delivered    = total_delivered,
        total_collected    = total_collected,
    )


@app.route('/')
@login_required
def home():
    customers = get_customer_names()
    return render_template('scan.html', user=session.get('user'), customers=customers)

@app.route('/submit', methods=['POST'])
def submit():
    now = datetime.now()
    valid_cylinders = []
    rows_to_append = []

    # Check and add "Customer" column in Sheet 1 if missing
    try:
        if scan_ws is not None:
            headers = scan_ws.row_values(1)
            if len(headers) < 6:
                # Add "Customer" as the 6th column header
                scan_ws.update_cell(1, 6, "Customer")
    except Exception as e:
        print("Error checking/updating Sheet1 headers:", e)

    # Support JSON payload (modern split-actions submission)
    if request.is_json:
        data = request.get_json()
        driver = data.get('driver', '').strip()
        customer = data.get('customer', '').strip()
        scans = data.get('scans', [])
        
        for s in scans:
            uid = s.get('uid', '').strip()
            action = s.get('action', '').strip()
            if uid and action:
                valid_cylinders.append(uid)
                # Only write customer for Delivery/Collection, leave blank for Filling
                cust_val = customer if action in ['Delivery', 'Collection'] else ''
                rows_to_append.append([
                    now.strftime('%d-%m-%Y'),
                    now.strftime('%H:%M:%S'),
                    driver,
                    action,
                    uid,
                    cust_val
                ])
    else:
        # Fallback to standard form-data
        action    = request.form['action']
        driver    = request.form['driver']
        customer  = request.form.get('customer', '').strip()
        cylinders = request.form.getlist('cylinders')
    
        for uid in cylinders:
            uid = uid.strip()
            if uid:
                valid_cylinders.append(uid)
                cust_val = customer if action in ['Delivery', 'Collection'] else ''
                rows_to_append.append([
                    now.strftime('%d-%m-%Y'),
                    now.strftime('%H:%M:%S'),
                    driver,
                    action,
                    uid,
                    cust_val
                ])
    
    if rows_to_append:
        sheet.append_rows(rows_to_append)

    # Auto-update Cylinders registry sheet for each scanned UID
    try:
        global cyl_ws
        if cyl_ws is None and doc:
            try:
                cyl_ws = doc.worksheet(CYLINDER_SHEET_NAME)
            except Exception:
                cyl_ws = None
        if cyl_ws:
            cyl_rows = cyl_ws.get_all_values()
            # Build UID → row index map
            uid_row_map = {}
            for idx, r in enumerate(cyl_rows):
                if idx == 0:
                    continue
                if r and r[0].strip():
                    uid_row_map[r[0].strip().upper()] = idx + 1

            today_str = datetime.now().strftime('%d-%m-%Y')
            for row_data in rows_to_append:
                scan_uid    = row_data[4].strip().upper()
                scan_action = row_data[3].strip()
                scan_cust   = row_data[5].strip() if len(row_data) > 5 else ''
                row_num = uid_row_map.get(scan_uid)
                if not row_num:
                    continue  # UID not in registry — skip silently
                if scan_action == 'Delivery':
                    new_status   = 'Delivered'
                    new_location = scan_cust or 'Customer'
                elif scan_action == 'Collection':
                    new_status   = 'Empty'
                    new_location = 'Depot'
                elif scan_action == 'Filling':
                    new_status   = 'Filled'
                    new_location = 'Depot'
                else:
                    continue
                cyl_ws.update(f'E{row_num}:G{row_num}',
                              [[new_status, new_location, today_str]])
    except Exception as e:
        print("Registry auto-update error (non-fatal):", e)

    clear_cache()
    return f"{len(valid_cylinders)} cylinders saved successfully"

@app.route('/manager')
def manager():
    wb    = load_workbook('data.xlsx')
    ws    = wb['Scan_Log']
    scans = list(ws.iter_rows(min_row=2, values_only=True))
    scans.reverse()
    return render_template('manager.html', scans=scans)

@app.route('/customers')
def customers():
    wb        = load_workbook('data.xlsx')
    ws        = wb['Customers']
    customers = list(ws.iter_rows(min_row=2, values_only=True))
    return render_template('customers.html', customers=customers)

@app.route('/add_customer', methods=['POST'])
def add_customer():
    customer = request.form['customer']
    email    = request.form['email']
    phone    = request.form['phone']
    wb       = load_workbook('data.xlsx')
    ws       = wb['Customers']
    ws.append([customer, email, phone])
    wb.save('data.xlsx')
    return redirect('/customers')


# ================================================================
#  INVENTORY & BULK TANKS ROUTES
# ================================================================

@app.route('/admin/inventory')
@admin_required
def admin_inventory():
    selected_date = request.args.get('date', '')
    parsed = parse_date(selected_date) if selected_date else date.today()
    target_date_str = parsed.strftime('%d-%m-%Y')
    target_date_iso = parsed.strftime('%Y-%m-%d')
    
    t1 = calculate_table1_filled_inventory()
    t2 = calculate_table2_bulk_inventory(target_date_str)
    
    hydro_alerts = [c for c in merge_cylinder_data() if c.get('hydro_badge') in ('Overdue', 'Due Soon')]
    hydro_alerts.sort(key=lambda x: parse_date(x.get('next_hydro_due')) or date.max)
    
    return render_template('inventory.html',
        user=session['user'],
        target_date=target_date_str,
        target_date_iso=target_date_iso,
        t1=t1,
        t2=t2,
        hydro_alerts=hydro_alerts
    )

@app.route('/admin/tanks/update', methods=['POST'])
@admin_required
def admin_tanks_update():
    target_date = request.form.get('date', date.today().strftime('%d-%m-%Y'))
    
    gases = ['Argon', 'CO2', 'N2', 'Oxygen']
    success = True
    for g in gases:
        opening = request.form.get(f'{g}_opening', '0').strip() or '0'
        capacity = request.form.get(f'{g}_capacity', '0').strip() or '0'
        dead_volume = request.form.get(f'{g}_dead', '0').strip() or '0'
        unit = 'KG' if g == 'CO2' else 'Cum'
        
        ok = update_tank_opening_stock(target_date, g, opening, capacity, dead_volume, unit)
        if not ok:
            success = False
            
    if success:
        clear_cache()
        return redirect(f'/admin/inventory?date={target_date}')
    else:
        return "Error updating some tank levels. Please verify your sheets connection.", 500

@app.route('/admin/inventory/export/excel')
@admin_required
def export_excel():
    target_date = request.args.get('date', date.today().strftime('%d-%m-%Y'))
    
    t1 = calculate_table1_filled_inventory()
    t2 = calculate_table2_bulk_inventory(target_date)
    
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Inventory Report"
    
    title_font = Font(name='Arial', size=16, bold=True, color='0F6E56')
    header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    sub_font = Font(name='Arial', size=11, bold=True, color='0F6E56')
    normal_font = Font(name='Arial', size=11)
    bold_font = Font(name='Arial', size=11, bold=True)
    
    green_fill = PatternFill(start_color='0F6E56', end_color='0F6E56', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFFCC', end_color='FFFFCC', fill_type='solid')
    orange_fill = PatternFill(start_color='C25E3B', end_color='C25E3B', fill_type='solid')
    
    border_thin = Border(
        left=Side(style='thin', color='DDDDDD'),
        right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'),
        bottom=Side(style='thin', color='DDDDDD')
    )
    
    ws1['A1'] = "Daily Dispatch & Tank Status Report"
    ws1['A1'].font = title_font
    ws1['A2'] = f"Report Date: {target_date}"
    ws1['A2'].font = Font(name='Arial', size=11, italic=True)
    
    # Table 1: Cyl Status Dispatch Stock
    ws1['A4'] = "Cyl Status Dispatch Stock (Filled Cylinder Inventory)"
    ws1['A4'].font = sub_font
    
    t1_headers = ["Product", "Full Feeling Cyl", "Gas per Cyl", "Total Gas", "Unit"]
    for col_idx, h in enumerate(t1_headers, 1):
        cell = ws1.cell(row=5, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = green_fill
        cell.alignment = Alignment(horizontal='center')
        
    row_idx = 6
    for row in t1['rows']:
        ws1.cell(row=row_idx, column=1, value=row['name']).font = normal_font
        ws1.cell(row=row_idx, column=2, value=row['filled_count']).font = normal_font
        ws1.cell(row=row_idx, column=3, value=row['gas_per_cyl']).font = normal_font
        ws1.cell(row=row_idx, column=4, value=row['total_gas']).font = normal_font
        ws1.cell(row=row_idx, column=5, value=row['unit']).font = normal_font
        
        for col_idx in range(1, 6):
            ws1.cell(row=row_idx, column=col_idx).border = border_thin
            
        row_idx += 1
        
    # Total row
    total_cell = ws1.cell(row=row_idx, column=1, value="TOTAL FILLED")
    total_cell.font = bold_font
    total_cell.fill = yellow_fill
    
    cnt_cell = ws1.cell(row=row_idx, column=2, value=t1['total_filled'])
    cnt_cell.font = bold_font
    cnt_cell.fill = yellow_fill
    
    ws1.cell(row=row_idx, column=3, value="").fill = yellow_fill
    
    gas_tot_cell = ws1.cell(row=row_idx, column=4, value=f"{t1['total_cum']} Cum + {t1['total_kg']} KG")
    gas_tot_cell.font = bold_font
    gas_tot_cell.fill = yellow_fill
    
    ws1.cell(row=row_idx, column=5, value="Mixed").font = bold_font
    ws1.cell(row=row_idx, column=5).fill = yellow_fill
    
    for col_idx in range(1, 6):
        ws1.cell(row=row_idx, column=col_idx).border = border_thin
        
    row_idx += 3
    
    # Table 2: Bulk Tanks
    ws1.cell(row=row_idx, column=1, value="Total USED Stock (Bulk Tank Gas Inventory)").font = sub_font
    row_idx += 1
    
    t2_headers = ["Gas", "Opening Stock", "Used Today", "Closing Stock", "Fleet Size", "Usable Stock", "Unit"]
    for col_idx, h in enumerate(t2_headers, 1):
        cell = ws1.cell(row=row_idx, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = orange_fill
        cell.alignment = Alignment(horizontal='center')
        
    row_idx += 1
    for row in t2:
        ws1.cell(row=row_idx, column=1, value=row['gas']).font = normal_font
        ws1.cell(row=row_idx, column=2, value=row['opening']).font = normal_font
        ws1.cell(row=row_idx, column=3, value=row['used_today']).font = normal_font
        ws1.cell(row=row_idx, column=4, value=row['closing']).font = normal_font
        ws1.cell(row=row_idx, column=5, value=row['fleet']).font = normal_font
        ws1.cell(row=row_idx, column=6, value=row['usable']).font = normal_font
        ws1.cell(row=row_idx, column=7, value=row['unit']).font = normal_font
        
        for col_idx in range(1, 8):
            ws1.cell(row=row_idx, column=col_idx).border = border_thin
            
        row_idx += 1
        
    # Auto-fit columns
    for col in ws1.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    from io import BytesIO
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    
    filename = f"Inventory_Report_{target_date}.xlsx"
    from flask import send_file
    return send_file(out, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route('/admin/inventory/export/pdf')
@admin_required
def export_pdf():
    target_date = request.args.get('date', date.today().strftime('%d-%m-%Y'))
    
    t1 = calculate_table1_filled_inventory()
    t2 = calculate_table2_bulk_inventory(target_date)
    
    from io import BytesIO
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    
    styles = getSampleStyleSheet()
    primary_color = colors.HexColor('#0F6E56')
    dark_text = colors.HexColor('#2C2C2A')
    orange_color = colors.HexColor('#C25E3B')
    
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=18,
        textColor=primary_color,
        spaceAfter=4
    )
    subtitle_style = ParagraphStyle(
        'DocSubTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=10,
        textColor=dark_text,
        spaceAfter=12
    )
    section_style = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=12,
        textColor=primary_color,
        spaceBefore=10,
        spaceAfter=6
    )
    
    story.append(Paragraph("Cylinder Tracker — Daily Inventory & Tanks Report", title_style))
    story.append(Paragraph(f"Report Generated for Date: {target_date}", subtitle_style))
    
    story.append(Paragraph("Cyl Status Dispatch Stock (Filled Cylinder Inventory)", section_style))
    
    t1_data = [["Product", "Full Feeling Cyl", "Gas/Cyl", "Total Gas", "Unit"]]
    for row in t1['rows']:
        t1_data.append([
            row['name'],
            str(row['filled_count']),
            f"{row['gas_per_cyl']:.4f}".rstrip('0').rstrip('.'),
            f"{row['total_gas']:.4f}".rstrip('0').rstrip('.'),
            row['unit']
        ])
    t1_data.append([
        "TOTAL FILLED",
        str(t1['total_filled']),
        "",
        f"{t1['total_cum']} Cum + {t1['total_kg']} KG",
        "Mixed"
    ])
    
    table1 = Table(t1_data, colWidths=[150, 100, 100, 120, 60])
    table1.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), primary_color),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('BOTTOMPADDING', (0,0), (-1,0), 5),
        ('TOPPADDING', (0,0), (-1,0), 5),
        ('ALIGN', (1,1), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,1), (-1,-2), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-2), 9),
        ('GRID', (0,0), (-1,-2), 0.5, colors.lightgrey),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#FFFFCC')),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,-1), (-1,-1), 9),
        ('GRID', (0,-1), (-1,-1), 0.5, colors.HexColor('#CCCCCC')),
        ('TOPPADDING', (0,-1), (-1,-1), 5),
        ('BOTTOMPADDING', (0,-1), (-1,-1), 5),
    ]))
    story.append(table1)
    
    story.append(Spacer(1, 15))
    
    story.append(Paragraph("Total USED Stock (Bulk Tank Gas Inventory)", section_style))
    
    t2_data = [["Gas", "Opening Stock", "Used Today", "Closing Stock", "Fleet Size", "Usable Stock", "Unit"]]
    for row in t2:
        t2_data.append([
            row['gas'],
            f"{row['opening']:.2f}",
            f"{row['used_today']:.2f}",
            f"{row['closing']:.2f}",
            str(row['fleet']),
            f"{row['usable']:.2f}",
            row['unit']
        ])
        
    table2 = Table(t2_data, colWidths=[90, 85, 75, 85, 65, 85, 45])
    table2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), orange_color),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('BOTTOMPADDING', (0,0), (-1,0), 5),
        ('TOPPADDING', (0,0), (-1,0), 5),
        ('ALIGN', (1,1), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,1), (-1,-1), 9),
        ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#FAF5F3')])
    ]))
    story.append(table2)
    
    doc.build(story)
    buffer.seek(0)
    
    filename = f"Inventory_Report_{target_date}.pdf"
    from flask import send_file
    return send_file(buffer, as_attachment=True, download_name=filename, mimetype="application/pdf")


if __name__ == '__main__':
    app.run(debug=True)