# Codebase Dependency Graph

This document details how different files within the Cylinder Tracker project connect to each other.

---

## 1. Import Map & File Relationships

```
           ┌──────────────────────┐
           │        .env          │
           └──────────┬───────────┘
                      │ loads config
                      ▼
┌───────────────────────────────────────────┐
│                 models.py                 │
│  (Database tables & SQLAlchemy schema)    │
└──────────▲─────────────────────▲──────────┘
           │ imports             │ imports
           │                     │
┌──────────┴──────────┐   ┌──────┴──────────┐
│       app.py        │   │     sync.py     │
│ (Web routes, auth,  │◄──┼─────────────────┤
│  PDF/Excel exports) │   │ (Periodic sheet │
└─────────────────────┘   │  to DB syncer)  │
                          └─────────────────┘
```

* **models.py** depends on **db.py** (which instantiates the database object `db = SQLAlchemy()`).
* **sync.py** imports database models from **models.py** and uses `gspread` to query Sheets.
* **app.py** imports:
  * Database connection and models from **models.py**.
  * Synchronization function `sync_sheets_to_db` from **sync.py** to run within the Background Scheduler.

---

## 2. Core / High-Impact Files
These files form the backbone of the application and should not be modified without careful planning:

### 1. `app.py`
* **Impact**: Critical.
* **Role**: Handles Flask initialization, routing, authentication decorators (`@admin_required`), PDF report structures (ReportLab), Excel generation (openpyxl), scanning validation, and background syncer scheduling.
* **Modification Warning**: A minor syntax error or incorrect decorator placement can lock users out or crash background cron synchronization.

### 2. `models.py`
* **Impact**: High.
* **Role**: Defines the SQLAlchemy database model classes which map exactly to PostgreSQL tables and Google Sheet columns.
* **Modification Warning**: Modifying column names or types in this file will break downstream sync queries inside `sync.py` and table mapping logic inside `app.py` unless matching migrations are executed in both PostgreSQL and Google Sheets columns simultaneously.

### 3. `sync.py`
* **Impact**: High.
* **Role**: Governs periodic data replication between PostgreSQL and Google Sheets.
* **Modification Warning**: Small bugs here can lead to database desynchronization, duplicate entries, or unwanted deletion cascades if local records are mistakenly identified as deleted.

### 4. `cylinder_full_script.gs`
* **Impact**: Medium-High.
* **Role**: Runs directly inside Google Sheets to manage UI drop-downs, trigger receipt generation on edit, and send automated ledger emails.
* **Modification Warning**: Breaking changes in this script will cause Google Sheets triggers to fail silently, stopping customer email receipt workflows.
