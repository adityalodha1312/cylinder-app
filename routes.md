# Application Routing Table

The Flask backend manages routing in `app.py`. Routes are protected by `@login_required` or `@admin_required` decorators, which verify the role stored in the Flask session cookie.

## Web Portal and API Routes

| Route | Method | Backend Function | Description | Access Required |
| :--- | :--- | :--- | :--- | :--- |
| `/` | `GET` | `index()` | Redirects root traffic to driver scan app `/scan` | Login |
| `/login` | `GET`, `POST` | `login()` | Performs username and password checks | Open |
| `/logout` | `GET` | `logout()` | Destroys user session | Open |
| `/scan` | `GET` | `scan()` | Mobile-friendly cylinder scanner SPA for drivers | Driver/Admin |
| `/submit` | `POST` | `submit()` | Processes a scanned batch of UIDs (Delivery/Collection/Filling) | Driver/Admin |
| `/admin` | `GET` | `admin()` | Redirects to `/admin/dashboard` | Manager/Owner |
| `/admin/dashboard` | `GET` | `admin_dashboard()` | Business summary dashboard with KPIs and top customers | Manager/Owner |
| `/admin/activity` | `GET` | `admin_activity()` | Logs timeline of recent scan batches | Manager/Owner |
| `/admin/daily_summary` | `GET` | `admin_daily_summary()`| View summary of dispatch/collection statistics and net gas usage | Manager/Owner |
| `/admin/outstanding` | `GET` | `admin_outstanding()` | View customer outstanding cylinder accounts | Manager/Owner |
| `/admin/aging` | `GET` | `admin_aging()` | Detailed cylinder aging status report (days out from Depot) | Manager/Owner |
| `/admin/rotation` | `GET` | `admin_rotation()` | Logs of cylinder movements, fillings, and journey histories | Manager/Owner |
| `/admin/products` | `GET` | `admin_products()` | View product sizing and gas capacity lists | Manager/Owner |
| `/admin/products/save`| `POST` | `admin_products_save()`| Saves modified product capacities and sizes to DB/Sheet | Manager/Owner |
| `/admin/drivers` | `GET` | `admin_drivers()` | Driver performance stats (total deliveries vs collections) | Manager/Owner |
| `/admin/movement` | `GET` | `admin_movement()` | Daily movements line chart (delivered vs collected) | Manager/Owner |
| `/admin/search` | `GET` | `admin_search()` | Lookup journey history and current location of a cylinder | Manager/Owner |
| `/admin/receipts` | `GET` | `admin_receipts()` | Log of email receipts sent out to customers | Manager/Owner |
| `/admin/update_mapping`| `POST`| `admin_update_mapping()`| Triggers sync from database maps to Sheet 2 | Manager/Owner |
| `/admin/send_receipt` | `POST` | `admin_send_receipt()` | Requests Gmail SMTP receipt send for a batch scan | Manager/Owner |
| `/admin/receipt_status`| `GET` | `admin_receipt_status()`| API check for asynchronous email send results | Manager/Owner |
| `/admin/cylinders` | `GET` | `admin_cylinders()` | Master cylinder list showing gas, specification, and hydro test status | Manager/Owner |
| `/admin/cylinders/add`| `GET`, `POST`| `admin_cylinders_add()`| Create a new cylinder record and save to DB/Sheets | Manager/Owner |
| `/admin/cylinders/<uid>/edit`| `GET`, `POST`| `admin_cylinders_edit()`| Update technical maintenance specifications for a cylinder | Manager/Owner |
| `/admin/cylinders/<uid>`| `GET` | `admin_cylinder_detail()`| View specific details of a single cylinder | Manager/Owner |
| `/admin/cylinders/mark_collected`| `POST`| `admin_mark_collected()`| Mark a list of cylinders collected directly | Manager/Owner |
| `/admin/customers/bulk_collect`| `POST`| `admin_bulk_collect()`| Mark all cylinders of a customer collected at once | Manager/Owner |
| `/admin/customers` | `GET` | `admin_customers()` | View list of customers and details | Manager/Owner |
| `/admin/customers/<name>`| `GET` | `admin_customer_profile()`| Customer profile card and their ledger statistics | Manager/Owner |
| `/admin/customers/add`| `GET`, `POST`| `admin_customers_add()`| Add a new customer to the database and Sheet | Manager/Owner |
| `/admin/customers/<name>/edit`| `GET`, `POST`| `admin_customers_edit()`| Modify details of an existing customer | Manager/Owner |
| `/admin/customers/<name>/delete`| `POST`| `admin_customers_delete()`| Delete customer record | Manager/Owner |
| `/admin/cold_calls` | `GET` | `admin_cold_calls()` | Checklist dashboard of sales calls | Manager/Owner |
| `/admin/api/cold_call/toggle`| `POST`| `admin_cold_call_toggle()`| Toggles cold call checklist status in DB | Manager/Owner |
| `/admin/api/cold_call/reset_all`| `POST`| `admin_cold_call_reset()`| Clear all checklist statuses | Manager/Owner |
| `/admin/api/dura_fill` | `POST` | `admin_dura_fill()` | Record a Duracylinder filling (updates previous/current gas) | Manager/Owner |
| `/admin/customers/<name>/offer`| `GET`| `admin_customer_offer()`| Direct access to customer commercial offer builder | Manager/Owner |
| `/admin/offer/new` | `GET` | `admin_offer_new()` | Open commercial offer document builder | Manager/Owner |
| `/admin/offer/generate`| `POST` | `admin_offer_generate()`| Compiles PDF file of custom contract offer for download | Manager/Owner |
| `/admin/orders/new` | `GET` | `admin_order_new()` | Open order and delivery challan spreadsheet builder | Manager/Owner |
| `/admin/orders/generate`| `POST` | `admin_order_generate()`| Generates delivery challan sheets | Manager/Owner |
| `/admin/inventory` | `GET` | `admin_inventory()` | Stock levels, bulk tanks, and daily dispatch grid | Manager/Owner |
| `/admin/inventory/export/excel`| `GET`| `export_excel()` | Excel export of Daily Inventory & Tanks status | Manager/Owner |
| `/admin/inventory/export/pdf`| `GET` | `export_pdf()` | PDF export of Daily Inventory & Tanks status | Manager/Owner |
| `/api/cylinder_status/<uid>`| `GET` | `api_cylinder_status()`| API JSON output of a cylinder status for scans | Driver/Admin |
| `/admin/api/mapping_mismatches`| `GET`| `api_mapping_mismatches()`| Detect discrepancies between scans and batch maps | Manager/Owner |
