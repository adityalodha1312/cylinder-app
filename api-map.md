# Endpoint API Catalog

This document details the JSON endpoints and POST actions used for data operations, receipt delivery, and client-side page updates.

---

### 1. `/submit`
* **Method**: `POST`
* **Purpose**: Driver submits scanned cylinder transactions.
* **Format**: Supports both standard HTML `application/x-www-form-urlencoded` and JSON payload `application/json`.
* **Payload (JSON)**:
  ```json
  {
    "driver": "John Doe",
    "customer": "ABC Industrial Gases Ltd",
    "scans": [
      { "uid": "CYL101", "action": "Delivery" },
      { "uid": "CYL102", "action": "Collection" }
    ]
  }
  ```
* **Payload (Form Data)**:
  * `action`: `"Delivery" | "Collection" | "Filling"`
  * `driver`: `"John Doe"`
  * `customer`: `"ABC Industrial Gases Ltd"`
  * `cylinders`: `["CYL101", "CYL102"]`
* **Response**:
  * **Success (200)**: `"Scans logged successfully"` or redirection.
  * **Validation Error (400)**: `"Validation Error: Cylinder 'CYL101' is already at the Depot. Cannot collect twice."`
* **Used By**: Mobile Scan application (`templates/scan.html`).

---

### 2. `/admin/send_receipt`
* **Method**: `POST`
* **Purpose**: Request the system to build and email a receipt PDF to a customer.
* **Payload (Form Data)**:
  * `date`: `"22-06-2026"`
  * `time`: `"14:30:00"`
  * `driver`: `"John Doe"`
  * `action`: `"Delivery"`
  * `customer`: `"ABC Industrial Gases Ltd"`
* **Response (JSON)**:
  ```json
  {
    "status": "success | error",
    "message": "Receipt sent successfully | error reason"
  }
  ```
* **Used By**: Admin receipts list page (`templates/receipts.html`).

---

### 3. `/admin/api/cold_call/toggle`
* **Method**: `POST`
* **Purpose**: Update toggle status of a sales cold call task.
* **Payload (JSON)**:
  ```json
  {
    "customer_id": "CUST001",
    "field": "cold_call_done",
    "value": true
  }
  ```
* **Response (JSON)**:
  ```json
  {
    "status": "success"
  }
  ```
* **Used By**: Sales checklist dashboard (`templates/cold_call_checklist.html`).

---

### 4. `/admin/api/cold_call/reset_all`
* **Method**: `POST`
* **Purpose**: Clear all cold call statuses for a fresh campaign.
* **Payload**: None.
* **Response (JSON)**:
  ```json
  {
    "status": "success",
    "message": "All cold call checkmarks reset."
  }
  ```
* **Used By**: Sales checklist dashboard reset button.

---

### 5. `/admin/api/dura_fill`
* **Method**: `POST`
* **Purpose**: Record a new filling operation for a Duracylinder.
* **Payload (JSON)**:
  ```json
  {
    "uid": "DURA990",
    "gas": "ARG",
    "operator": "Dave Fill"
  }
  ```
* **Response (JSON)**:
  * **Success (200)**:
    ```json
    {
      "status": "success",
      "message": "Refill logged. Cylinder current gas set to ARG.",
      "purge_required": false
    }
    ```
  * **Purge Alert Required (200)**:
    ```json
    {
      "status": "success",
      "message": "Refill logged. PURGE REQUIRED: Gas type changed from OXY to ARG.",
      "purge_required": true,
      "previous_gas": "OXY",
      "new_gas": "ARG"
    }
    ```
* **Used By**: Cylinder tracking page refill log modal (`templates/cylinders.html`).

---

### 6. `/api/cylinder_status/<uid>`
* **Method**: `GET`
* **Purpose**: Verify a cylinder details (owner, status, gas type) in real-time during scans.
* **Payload**: None.
* **Response (JSON)**:
  ```json
  {
    "exists": true,
    "uid": "CYL887",
    "gas_type": "N2",
    "owner": "Depot",
    "status": "Empty",
    "location": "Depot"
  }
  ```
* **Used By**: Scan barcode verification scripting (`scan.html`).

---

### 7. `/admin/api/mapping_mismatches`
* **Method**: `GET`
* **Purpose**: Fetch mismatch logs between individual scans and batch map summaries.
* **Payload**: None.
* **Response (JSON)**:
  ```json
  {
    "mismatches": [
      {
        "date": "22-06-2026",
        "time": "10:15:00",
        "customer": "Acme Corp",
        "action": "Delivery",
        "map_count": 5,
        "actual_count": 4,
        "mismatch_reason": "Counts do not match"
      }
    ]
  }
  ```
* **Used By**: Outstanding or dashboard audit page widgets.
