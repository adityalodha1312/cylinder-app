import yagmail
from flask import Flask, render_template, request
from openpyxl import load_workbook
from datetime import datetime
from flask import redirect
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

sheet = client.open("Cylinder Tracking").sheet1
@app.route('/manager')
def manager():

    wb = load_workbook('data.xlsx')

    sheet = wb['Scan_Log']

    scans = []

    for row in sheet.iter_rows(
        min_row=2,
        values_only=True
    ):

        scans.append(row)

    scans.reverse()

    return render_template(
        'manager.html',
        scans=scans
    )
@app.route('/customers')
def customers():

    wb = load_workbook('data.xlsx')

    sheet = wb['Customers']

    customers = []

    for row in sheet.iter_rows(
        min_row=2,
        values_only=True
    ):
        customers.append(row)

    return render_template(
        'customers.html',
        customers=customers
    )


@app.route('/add_customer', methods=['POST'])
def add_customer():

    customer = request.form['customer']
    email = request.form['email']
    phone = request.form['phone']

    wb = load_workbook('data.xlsx')

    sheet = wb['Customers']

    sheet.append([
        customer,
        email,
        phone
    ])

    wb.save('data.xlsx')

    return redirect('/customers')
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

    # Save directly to Google Sheets
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

    # Email body
    email_body = f"""
Driver: {driver}

Action: {action}

Cylinders:

"""

    for uid in valid_cylinders:
        email_body += uid + "\n"

    email_body += f"\nTotal Cylinders: {len(valid_cylinders)}"

    # Send email
    try:

        yag = yagmail.SMTP(
            "lodhachaya7@gmail.com",
            "qofq vjks imsq oyil"
        )

        yag.send(
            to="adityalodha26@gmail.com",
            subject=f"{action} Scan Report",
            contents=email_body
        )

        print("Email sent successfully")

    except Exception as e:

        print("Email Error:", e)

    return f"{len(valid_cylinders)} cylinders saved successfully"
if __name__ == '__main__':
    app.run(debug=True)