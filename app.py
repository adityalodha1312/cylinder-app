import threading
import yagmail
from flask import Flask, render_template, request, redirect, send_from_directory
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds = Credentials.from_service_account_file(
    "credentials.json",
    scopes=SCOPES
)

client = gspread.authorize(creds)
sheet = client.open("Cylinder Tracking").worksheet("Sheet1")


# ── PWA routes ────────────────────────────────────────────────
@app.route('/static/manifest.json')
def manifest():
    return send_from_directory('static', 'manifest.json',
                               mimetype='application/manifest+json')

@app.route('/static/sw.js')
def service_worker():
    response = send_from_directory('static', 'sw.js',
                                   mimetype='application/javascript')
    response.headers['Service-Worker-Allowed'] = '/'
    return response

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory('static', filename)
# ──────────────────────────────────────────────────────────────


def send_email(driver, action, valid_cylinders):
    try:
        email_body = f"Driver: {driver}\nAction: {action}\nCylinders:\n\n"
        for uid in valid_cylinders:
            email_body += uid + "\n"
        email_body += f"\nTotal Cylinders: {len(valid_cylinders)}"

        yag = yagmail.SMTP("lodhachaya7@gmail.com", "YOUR_APP_PASSWORD")
        yag.send(
            to="adityalodha26@gmail.com",
            subject=f"{action} Scan Report",
            contents=email_body
        )
        print("Email sent successfully")
    except Exception as e:
        print("Email Error:", e)


@app.route('/')
def home():
    return render_template('scan.html')


@app.route('/submit', methods=['POST'])
def submit():
    action = request.form['action']
    driver = request.form['driver']
    cylinders = request.form.getlist('cylinders')

    now = datetime.now()
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

    # Send email in background so it doesn't block or crash the response
    t = threading.Thread(
        target=send_email,
        args=(driver, action, valid_cylinders)
    )
    t.daemon = True
    t.start()

    return f"{len(valid_cylinders)} cylinders saved successfully"


@app.route('/manager')
def manager():
    records = sheet.get_all_values()
    scans = records[1:]
    scans.reverse()
    return render_template('manager.html', scans=scans)


@app.route('/customers')
def customers():
    customers_sheet = client.open("Cylinder Tracking").worksheet("Customers")
    records = customers_sheet.get_all_values()
    customers_list = records[1:]
    return render_template('customers.html', customers=customers_list)


@app.route('/add_customer', methods=['POST'])
def add_customer():
    customer = request.form['customer']
    email = request.form['email']
    phone = request.form['phone']

    customers_sheet = client.open("Cylinder Tracking").worksheet("Customers")
    customers_sheet.append_row([customer, email, phone])

    return redirect('/customers')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
