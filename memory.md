# Codebase Intelligence Memory Document

## 1. Project Overview
This repository contains a Cylinder Tracking and Management MVP (Minimum Viable Product) built for a gas cylinder supply and delivery company. The application serves two core users:
1. **Drivers** who use a mobile-friendly scan application to register cylinder deliveries, empty collections, and cylinder fillings.
2. **Managers/Owners** who use an administrative portal to view reports, track cylinder rotation and aging, manage cylinder specifications (hydro tests, water capacities, gas mixtures), generate commercial offers, and monitor inventory/tank stock levels.

The system relies on a dual-storage setup: a relational database (PostgreSQL via Supabase or SQLite locally) acts as the high-speed operational data store, while a Google Sheets workbook acts as the ground truth configuration and visual data editor. A periodic scheduler syncs sheet modifications back to the database.

---

## 2. Business Purpose
The primary business problem solved is tracking the lifecycle of high-value industrial gas cylinders (such as Argon, Oxygen, Nitrogen, CO2, etc.) to:
* **Prevent Cylinder Loss**: Keep an accurate outstanding ledger of how many cylinders are in the possession of each customer.
* **Track Aging & Utilization**: Monitor how long cylinders stay at customer sites before collection.
* **Maintain Safety Compliance**: Monitor next hydro test dates, pressure specifications, and water capacity limits, distinguishing between standard cylinders and ultra-high purity (UHP) or mixture cylinders.
* **Manage Gas Replacements**: Prevent contamination by tracking gas filling history, specifically for Duracylinders (Dura Gas Tracking) where purges are required when changing gas types.
* **Generate Billing Documentation**: Automate delivery challans, customer receipts, and commercial offers.

---

## 3. Tech Stack
* **Backend Framework**: Python (Flask) with `SQLAlchemy` ORM and `APScheduler` for background sync.
* **Frontend UI**: Server-side Flask templates (Jinja2) styled with Vanilla CSS, using the DM Sans typography system and curated HSL color schemes.
* **Client-Side Scripting**: Vanilla Javascript (supporting barcode scanning via camera, dynamic row additions, interactive tab filtering, and Chart.js visualizations).
* **Database**: PostgreSQL (operational/production) or SQLite (in-memory/local development).
* **External Integration**: Google Sheets API via `gspread` and service account credentials.
* **Reports Generation**: `openpyxl` (Excel export), `ReportLab` (PDF generation).
* **Automations**: Google Apps Script (bound to the Google Sheet) for sending daily emails, managing outstanding sheets, and triggering email receipts.

---

## 4. Repository Structure
```
e:\Cylinder_MVP\
├── .env                     # Environment configurations (DB URL, Sheet IDs)
├── .gitignore               # Git ignored paths
├── app.py                   # Main Flask application entry point and routes
├── models.py                # SQLAlchemy DB schema models
├── db.py                    # Database instance configuration
├── sync.py                  # Google Sheets to DB synchronization module
├── cylinder_full_script.gs  # Google Apps Script codebase bound to Google Sheet
├── requirements.txt         # Python dependencies
├── static/                  # Static assets (images, CSS styles, JS scripts)
├── templates/               # Jinja2 HTML templates for Web UI
└── scratch/                 # Developer scratchpad scripts, benchmarks, and tests
```

---

## 5. System Architecture
```
+--------------------------------------------------------------+
|                        Browser (Web UI)                      |
|      (Driver /scan SPA & Admin /admin Multi-Page Portal)     |
+------------------------------+-------------------------------+
                               |
                        HTTP / JSON APIs
                               v
+------------------------------+-------------------------------+
|                       Python Flask Server                    |
|       (app.py - Routing, Auth, Validation & Reports)         |
+---------+--------------------+---------------------+---------+
          |                    |                     |
     SQLAlchemy           gspread API             openpyxl
          v                    v                     v
+---------+----------+  +------+------+       +------+------+
|     PostgreSQL     |  | Google Sheets|----->| Excel / PDF |
|  (Supabase/Local)  |  |  (Workbook)  |      |   Exports   |
+--------------------+  +------+------+       +-------------+
                               ^
                               | Apps Script Triggers
                        +------+------+
                        | Gmail SMTP / |
                        | PDF Receipts |
                        +--------------+
```

---

## 6. Routing Map
The routing is split between driver operations, administrative dashboards, and APIs.

| Route | View File | Purpose | Auth Required |
| :--- | :--- | :--- | :--- |
| `/` | Redirects to `/scan` | Root entry point | Driver or Admin session |
| `/login` | `login.html` | User sign-in page | None |
| `/logout` | - | Destroys Flask session | None |
| `/scan` | `scan.html` | Driver mobile-friendly scan UI | Yes (`driver` or `manager`) |
| `/submit` | - | Processes scanned barcodes batch (POST) | Yes (`driver` or `manager`) |
| `/admin` | Redirects to `/admin/dashboard` | Admin root | Yes (`manager` or `owner`) |
| `/admin/dashboard` | `dashboard.html` | High-level business KPIs & top outstanding | Yes (`manager` or `owner`) |
| `/admin/activity` | `activity.html` | Operational activity log (scans timeline) | Yes (`manager` or `owner`) |
| `/admin/daily_summary` | `daily_summary.html` | Dispatch summary, net gas usage, and stats | Yes (`manager` or `owner`) |
| `/admin/outstanding` | `outstanding.html` | Customer outstanding cylinder ledger | Yes (`manager` or `owner`) |
| `/admin/aging` | `aging.html` | Overdue cylinder details (days out) | Yes (`manager` or `owner`) |
| `/admin/rotation` | `rotation.html` | Cylinder movement log and journey history | Yes (`manager` or `owner`) |
| `/admin/products` | `products_config.html` | Product size/gas type configuration | Yes (`manager` or `owner`) |
| `/admin/drivers` | `drivers.html` | Driver performance and activity stats | Yes (`manager` or `owner`) |
| `/admin/movement` | `movement.html` | Daily net movement chart and statistics | Yes (`manager` or `owner`) |
| `/admin/search` | `search.html` | Cylinder journey search tool by UID | Yes (`manager` or `owner`) |
| `/admin/receipts` | `receipts.html` | Send receipt logs and delivery status | Yes (`manager` or `owner`) |
| `/admin/cylinders` | `cylinders.html` | Cylinder Master Registry and Dura Tracking | Yes (`manager` or `owner`) |
| `/admin/cylinders/add`| `cylinders_form.html` | Add new cylinder to registry | Yes (`manager` or `owner`) |
| `/admin/cylinders/<uid>/edit` | `cylinders_form.html` | Edit cylinder specifications | Yes (`manager` or `owner`) |
| `/admin/customers` | `customers_list.html` | List of customers | Yes (`manager` or `owner`) |
| `/admin/customers/<name>` | `customer_profile.html` | Customer details, ledger, and profile | Yes (`manager` or `owner`) |
| `/admin/cold_calls` | `cold_call_checklist.html` | Sales checklist for customer calls | Yes (`manager` or `owner`) |
| `/admin/offer/new` | `offer_form.html` | Create custom commercial offers | Yes (`manager` or `owner`) |
| `/admin/orders/new` | `order_form.html` | Generate purchase order sheets | Yes (`manager` or `owner`) |

---

## 7. Frontend Architecture
The frontend consists of server-rendered HTML templates utilizing a centralized CSS framework defined in `templates/admin_base.html`.
* **Design System**: Built around the custom colors:
  * Primary Green: `#0F6E56`
  * Green Accent: `#1D9E75`
  * Typography: `DM Sans`
* **Gas Badge Styling**: Centralized badges corresponding to gas types:
  * `ARG` (Argon): Blue (`hsl(210, 85%, 90%)`, text `hsl(210, 85%, 25%)`)
  * `OXY` (Oxygen): Sky Blue (`hsl(190, 90%, 90%)`, text `hsl(190, 90%, 25%)`)
  * `N2` (Nitrogen): Deep Purple (`hsl(270, 75%, 92%)`, text `hsl(270, 75%, 30%)`)
  * `CO2` (Carbon Dioxide): Orange (`hsl(25, 85%, 90%)`, text `hsl(25, 85%, 30%)`)
  * `ACM` (Argon/CO2 Mix): Teal (`hsl(170, 70%, 90%)`, text `hsl(170, 70%, 25%)`)
  * `AHM` (Argon/Hydrogen Mix): Pink/Rose (`hsl(340, 80%, 92%)`, text `hsl(340, 80%, 30%)`)
  * `GEN` (General Fallback): Gray (`hsl(60, 10%, 90%)`, text `hsl(60, 10%, 30%)`)
* **Interactive Scripts**:
  * Camera barcode reading is enabled inside `scan.html` via HTML5 video stream and standard barcode APIs.
  * Search bars dynamically filter tables on the client-side without page refreshes.

---

## 8. Backend Architecture
The backend is structured into:
* **Controllers & Views**: Flask route handlers in `app.py` process requests, check permissions, query databases, and return HTML templates or JSON data.
* **Database Layer**: Defined in `models.py` using Flask-SQLAlchemy. Models represent local tables maps.
* **Google Sheets Sync**: Handled by the `sync.py` module. It queries the spreadsheet using a service account and updates local tables (`users`, `customers`, `cylinders`, `cylinder_maintenance`, `scans`, `customer_map`, `bulk_tanks`, `products`).
* **Cron/Task Schedulers**: Handled inside `app.py` via `BackgroundScheduler`, executing `sync_sheets_to_db` every 5 minutes.

---

## 9. Database Architecture
Tables correspond to Google Sheet names:

* **users**: Credentials and user permission roles.
* **customers**: Registered clients.
* **cylinders**: Master list of cylinders, current location, gas type, and owner.
* **cylinder_maintenance**: Technical specification fields (next hydro dates, manufacture dates, water capacities).
* **scans**: Individual transaction log rows.
* **customer_map**: Grouped scan transactions representing a receipt or dispatch batch.
* **bulk_tanks**: Historical bulk tank level logs.
* **products**: Valid product dimensions and standard gas capacities.
* **dura_gas_history**: Tracking gas filling transitions for Duracylinders.

---

## 10. Authentication Flow
Authentication is managed via Flask Session cookies:
1. User logs in at `/login` by submitting a username and password.
2. The password is verified (checking hash patterns starting with `pbkdf2:` or `scrypt:` first, and falling back to direct equality for migration purposes).
3. If valid, the session stores user details: `session['user'] = {'username': u.username, 'role': u.role, 'name': u.name}`.
4. Route decorators `@login_required` and `@admin_required` protect operational endpoints.

---

## 11. Environment Variables
Configuration parameters inside `.env` include:
* `DATABASE_URL`: SQLAlchemy PostgreSQL database connection URL.
* `GOOGLE_APPLICATION_CREDENTIALS`: Path to the service account JSON credential file.
* `SPREADSHEET_ID` or `SPREADSHEET_NAME`: Name or identifier of the Google Sheets workbook.
* `MAIL_SERVER`, `MAIL_PORT`, `MAIL_USERNAME`, `MAIL_PASSWORD`: Mail configurations for receipts and automated ledgers.

---

## 12. Known Technical Debt & Risks
* **Dual Writes Delay**: Google Sheets appends are synchronous and can be slow under load. A retry mechanism (`sheets_write_with_retry`) mitigates transient API failures.
* **Template Filter Dependencies**: Changing templates requires matching filter registrations (e.g., standardizing `|trim` instead of custom filters).
* **No Database Migrations Engine**: Database schemas are checked and modified programmatically on startup instead of using Alembic.
