import yagmail
from flask import Flask, render_template, request, redirect, session, jsonify
from openpyxl import load_workbook
from datetime import datetime, date, timedelta
from functools import wraps
import gspread
from google.oauth2.service_account import Credentials
import time
from dotenv import load_dotenv
import os

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

SPREADSHEET_NAME = "Cylinder Tracking"
SCAN_SHEET_NAME  = "Sheet1"
MAP_SHEET_NAME   = "Customer Map"
USERS_SHEET_NAME = "Users"

# Cache worksheet objects at startup to avoid roundtrip sheet lookup calls
try:
    doc = client.open(SPREADSHEET_NAME)
    scan_ws = doc.worksheet(SCAN_SHEET_NAME)
    map_ws = doc.worksheet(MAP_SHEET_NAME)
    users_ws = doc.worksheet(USERS_SHEET_NAME)
    sheet = scan_ws
except Exception as e:
    print("Error caching Google Sheets worksheets:", e)
    doc = None
    scan_ws = None
    map_ws = None
    users_ws = None
    sheet = None

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


# ================================================================
#  EXISTING ROUTES (unchanged)
# ================================================================

@app.route('/')
@login_required
def home():
    return render_template('scan.html', user=session.get('user'))

@app.route('/submit', methods=['POST'])
def submit():
    action    = request.form['action']
    driver    = request.form['driver']
    cylinders = request.form.getlist('cylinders')
    now       = datetime.now()

    valid_cylinders = []
    for uid in cylinders:
        uid = uid.strip()
        if uid:
            valid_cylinders.append(uid)
            sheet.append_row([
                now.strftime('%d-%m-%Y'),
                now.strftime('%H:%M:%S'),
                driver,
                action,
                uid
            ])

    email_body = f"Driver: {driver}\n\nAction: {action}\n\nCylinders:\n\n"
    for uid in valid_cylinders:
        email_body += uid + "\n"
    email_body += f"\nTotal Cylinders: {len(valid_cylinders)}"

    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pass = os.environ.get("GMAIL_PASS")

    if gmail_user and gmail_pass:
        try:
            yag = yagmail.SMTP(gmail_user, gmail_pass)
            yag.send(
                to      = ["adityalodha26@gmail.com"],
                subject = f"{action} Scan Report",
                contents= email_body
            )
            print("Email sent successfully")
        except Exception as e:
            print("Email Error:", e)
    else:
        print("Email skipped: GMAIL_USER or GMAIL_PASS environment variables not set.")

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

if __name__ == '__main__':
    app.run(debug=True)