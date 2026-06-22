# Database Schema Map

The Cylinder Tracker database uses a relational database schema mapped via Flask-SQLAlchemy in `models.py`.

---

## 1. Table Definitions

### `users`
Tracks usernames, roles, and password hashes for authentication.
* **Fields**:
  * `id` (`Integer`, Primary Key)
  * `username` (`String(100)`, Unique, Nullable=False)
  * `password` (`String(255)`, Nullable=False) - Hashed with PBKDF2/Scrypt.
  * `role` (`String(50)`) - Either `'driver'`, `'manager'`, or `'owner'`.
  * `name` (`String(100)`) - User's display name.
  * `created_at` (`DateTime`) - System creation timestamp.

---

### `customers`
Tracks customer contact details and checkmarks.
* **Fields**:
  * `id` (`Integer`, Primary Key)
  * `customer_id` (`String(20)`, Unique)
  * `name` (`String(255)`, Nullable=False) - Customer name (joined with Scans/Maps).
  * `email` (`String(255)`) - Primary target email for delivery receipts.
  * `phone` (`String(50)`)
  * `address` (`Text`)
  * `cold_call_done` (`Boolean`) - Tracks checklist status.
  * `created_at` (`DateTime`)

---

### `cylinders`
The Master Cylinder Registry tracking current allocation status and locations.
* **Fields**:
  * `id` (`Integer`, Primary Key)
  * `uid` (`String(100)`, Unique, Nullable=False) - The barcode key.
  * `gas_type` (`String(50)`) - Active gas filled (e.g. `'ARG'`, `'N2'`, `'OXY'`).
  * `cylinder_type` (`String(50)`) - Sizing classification (e.g. `'Standard'`, `'Dura'`).
  * `owner` (`String(100)`) - Company owning the cylinder asset.
  * `status` (`String(20)`) - Operational status (e.g. `'Active'`, `'Retired'`, `'Lost'`).
  * `location` (`String(255)`) - Current location (e.g. `'Depot'`, customer name).
  * `last_activity_date` (`String(50)`) - Formatted activity date.
  * `created_at` (`DateTime`)

---

### `cylinder_maintenance`
Technical parameters, hydrostatic tests, and certification values for cylinders.
* **Fields**:
  * `id` (`Integer`, Primary Key)
  * `cylinder_uid` (`String(100)`, Unique, Nullable=False) - Joins with `cylinders.uid`.
  * `water_capacity` (`String(50)`) - Volume in liters.
  * `fill_pressure` (`String(50)`) - Working pressure in bar.
  * `gas_capacity` (`String(50)`) - Normal gas content.
  * `unit` (`String(20)`) - Unit of gas content (e.g. `'Cum'`, `'KG'`).
  * `is_mixture` (`String(20)`) - `'Yes'` or `'No'`.
  * `mix_ratio` (`String(100)`) - Component details.
  * `manufacture_date` (`String(50)`)
  * `last_hydro_date` (`String(50)`)
  * `next_hydro_due` (`String(50)`) - Hydrostatic test due date (OK/Due Soon/Overdue badge).
  * `hydro_test_status` (`String(50)`)
  * `cert_no` (`String(100)`)
  * `is_uhp` (`String(20)`) - Ultra-High Purity flag (`'Yes'` or `'No'`).

---

### `scans`
Logs every barcode scan event submitted by drivers.
* **Fields**:
  * `id` (`Integer`, Primary Key)
  * `scan_date` (`String(50)`, Nullable=False)
  * `scan_time` (`String(50)`)
  * `driver` (`String(100)`)
  * `action` (`String(50)`, Nullable=False) - `'Delivery'`, `'Collection'`, or `'Filling'`.
  * `cylinder_uid` (`String(100)`, Nullable=False)
  * `customer` (`String(255)`) - Targeted customer or `'Depot'`.
  * `gas_type` (`String(50)`) - Snapshot of the gas type at scan time.
  * `created_at` (`DateTime`)

---

### `customer_map`
Logs transaction summaries (batches) mapped to customers for receipt tracking.
* **Fields**:
  * `id` (`Integer`, Primary Key)
  * `scan_date` (`String(50)`)
  * `scan_time` (`String(50)`)
  * `driver` (`String(100)`)
  * `action` (`String(50)`)
  * `count` (`Integer`) - Number of cylinders in this batch.
  * `uids` (`Text`) - Space or comma-delimited UIDs list.
  * `customer` (`String(255)`)
  * `send_receipt` (`Boolean`) - Email checkbox trigger.
  * `receipt_status` (`String(100)`) - Receipt SMTP delivery status (e.g. `'Sent'`, `'Pending'`).
  * `created_at` (`DateTime`)

---

### `bulk_tanks`
Tracks tank inventory opening levels.
* **Fields**:
  * `id` (`Integer`, Primary Key)
  * `date` (`String(50)`, Nullable=False)
  * `gas` (`String(50)`, Nullable=False)
  * `opening` (`Float`)
  * `dead_volume` (`Float`)
  * `capacity` (`Float`)
  * `unit` (`String(20)`)
  * `created_at` (`DateTime`)

---

### `products`
The configuration lookup of valid cylinder product sizes, gas contents, and dimensions.
* **Fields**:
  * `id` (`Integer`, Primary Key)
  * `product_id` (`String(50)`, Unique, Nullable=False)
  * `name` (`String(100)`)
  * `gas_type` (`String(50)`)
  * `cylinder_type` (`String(50)`)
  * `gas_per_cyl` (`Float`) - Standard gas capacity content.
  * `unit` (`String(20)`) - Content unit.
  * `is_virtual` (`Boolean`) - Virtual products flags.

---

### `dura_gas_history`
Tracks historical fill events for Duracylinders to prevent cross-contamination.
* **Fields**:
  * `id` (`Integer`, Primary Key)
  * `cylinder_uid` (`String(100)`, Nullable=False)
  * `gas_filled` (`String(50)`, Nullable=False)
  * `previous_gas` (`String(50)`)
  * `purge_required` (`Boolean`) - True if filled gas type does not match previous.
  * `purge_acknowledged` (`Boolean`) - Operator verification of purge completion.
  * `operator` (`String(100)`)
  * `fill_date` (`String(50)`)
  * `fill_time` (`String(50)`)
  * `created_at` (`DateTime`)

---

## 2. Entity Relationships
The relationships in this database are loose and managed operational-side (in code/queries) rather than strict foreign key constraints, facilitating smooth bidirectional synchronization with Google Sheets.

```
┌─────────────────┐           ┌───────────────────────┐
│    cylinders    │ ◄──1:1──► │  cylinder_maintenance │
│  (uid: PK)      │           │  (cylinder_uid: FK)   │
└────────┬────────┘           └───────────────────────┘
         │
        1:N
         │
         ▼
┌─────────────────┐           ┌───────────────────────┐
│     scans       │ ◄──N:1─── │      customers        │
│ (cylinder_uid)  │           │      (name: Key)      │
└─────────────────┘           └──────────┬────────────┘
                                         │
                                        1:N
                                         │
                                         ▼
                              ┌───────────────────────┐
                              │     customer_map      │
                              │     (customer: Key)   │
                              └───────────────────────┘
```
