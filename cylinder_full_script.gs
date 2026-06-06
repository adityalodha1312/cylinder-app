// ============================================================
//  CYLINDER TRACKER — Google Apps Script (Full Script)
//  Replace everything in Extensions > Apps Script > Code.gs
//  with this file and save.
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
const SCAN_SHEET_NAME     = 'Sheet1';             // Raw scan log (app writes here)
const MAP_SHEET_NAME      = 'Customer Map';        // Sheet 2 — manager maps batches
const LEDGER_SHEET_NAME   = 'Outstanding Ledger'; // Sheet 3 — auto-calculated
const CUSTOMER_SHEET_NAME = 'Customers';           // Customer list for dropdown

// ── EMAIL CONFIG ─────────────────────────────────────────────
const EMAIL_RECIPIENTS = [
  'adityalodha26@gmail.com',   // Manager
  // 'owner@gmail.com',        // Add owner email here
];
const EMAIL_SENDER_NAME = 'Cylinder Tracker';      // Shown as sender name in inbox
// ────────────────────────────────────────────────────────────


// ============================================================
//  MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛢 Cylinder Tracker')
    .addItem('📋 Setup Customer Map Sheet', 'setupCustomerMapSheet')
    .addItem('🔄 Refresh New Batches → Sheet 2', 'refreshBatches')
    .addSeparator()
    .addItem('📊 Refresh Outstanding Ledger → Sheet 3', 'refreshLedger')
    .addSeparator()
    .addItem('📧 Send Ledger Email Now', 'sendDailyLedgerEmail')
    .addItem('⏰ Setup Daily 8 AM Email Trigger', 'setupDailyTrigger')
    .addItem('🚫 Remove Daily Email Trigger', 'removeDailyTrigger')
    .addSeparator()
    .addItem('⚡ Enable Auto-Refresh for Sheet 2', 'setupAutoRefreshTrigger')
    .addItem('🚫 Disable Auto-Refresh for Sheet 2', 'removeAutoRefreshTrigger')
    .addSeparator()
    .addItem('🔑 Enable Send Receipt Checkbox Trigger', 'setupEditTrigger')
    .addItem('🚫 Disable Send Receipt Checkbox Trigger', 'removeEditTrigger')
    .addSeparator()
    .addItem('📧 Send Pending Scan Emails Now', 'onNewRow')
    .addItem('📧 Reset Scan Email Counter', 'resetEmailCounter')
    .addSeparator()
    .addItem('🔧 Fix All Dropdowns (run once)', 'fixAllDropdowns')
    .addToUi();
}


// ============================================================
//  AUTO-REFRESH — Triggers when manager fills Customer column
// ============================================================
function handleSheetEdit(e) {
  if (!e || !e.range) return;

  const sheet    = e.range.getSheet();
  const sheetName = sheet.getName();
  const col       = e.range.getColumn();
  const row       = e.range.getRow();
  const value     = String(e.value || '').trim();

  // Auto-refresh Sheet 3 when manager fills Customer column in Sheet 2
  if (sheetName === MAP_SHEET_NAME && col === 7 && value !== '') {
    refreshLedger(true);
  }

  // Send Customer Email Receipt when Checkbox in Column 8 (Send Receipt?) is checked
  if (sheetName === MAP_SHEET_NAME && col === 8 && row > 1 && value.toUpperCase() === 'TRUE') {
    sendCustomerReceipt(row);
  }

  // Auto-refresh Sheet 3 when manager changes date filter cells (B1 or D1)
  if (sheetName === LEDGER_SHEET_NAME && row === 1 && (col === 2 || col === 4)) {
    refreshLedger(true);
  }
}

// ── SETUP INSTALLABLE EDIT TRIGGER (to allow MailApp.sendEmail) ──

function setupEditTrigger() {
  // Clear any existing installable edit triggers to avoid duplicates
  removeEditTrigger(true);

  // Set up installable edit trigger
  ScriptApp.newTrigger('handleSheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Installable Edit Trigger Enabled!\n\n' +
    'The script now has the permission to send email receipts automatically when a checkbox is checked.'
  );
}

function removeEditTrigger(silent = false) {
  const triggers = ScriptApp.getProjectTriggers();
  let removed    = 0;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'handleSheetEdit') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      removed > 0
        ? '✅ Installable edit trigger removed successfully.'
        : 'ℹ️ No installable edit trigger was active.'
    );
  }
}


// ============================================================
//  AUTO DAILY EMAIL — Outstanding Ledger
// ============================================================

// Called by the daily trigger at 8 AM — or manually from menu
function sendDailyLedgerEmail() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const scanSheet = ss.getSheetByName(SCAN_SHEET_NAME);
  const mapSheet  = ss.getSheetByName(MAP_SHEET_NAME);

  if (!scanSheet || !mapSheet) return;

  // ── Build batch → customer lookup from Sheet 2 ─────────
  const mapData         = mapSheet.getDataRange().getValues();
  const batchToCustomer = {};
  for (let i = 1; i < mapData.length; i++) {
    const r        = mapData[i];
    const customer = String(r[6]).trim();
    if (!customer) continue;
    const key = `${formatDate(r[0])}||${formatTime(r[1])}||${String(r[2]).trim()}||${String(r[3]).trim()}`;
    batchToCustomer[key] = customer;
  }

  // ── Build events from Sheet 1 ──────────────────────────
  const scanData = scanSheet.getDataRange().getValues();
  const events   = [];
  for (let i = 1; i < scanData.length; i++) {
    const row    = scanData[i];
    const date   = formatDate(row[0]);
    const time   = formatTime(row[1]);
    const driver = String(row[2]).trim();
    const action = String(row[3]).trim();
    const uid    = String(row[4]).trim();
    if (!uid) continue;
    const key      = `${date}||${time}||${driver}||${action}`;
    const customer = batchToCustomer[key];
    if (!customer) continue;
    events.push({ date, time, action, uid, customer, dateObj: parseDateTime(date, time) });
  }

  events.sort((a, b) => a.dateObj - b.dateObj);

  // ── Track cylinder ownership ───────────────────────────
  const cylinderOwner  = {};
  const customerStats  = {};
  for (const ev of events) {
    const { action, uid, customer, date } = ev;
    if (!customerStats[customer]) {
      customerStats[customer] = { totalDelivered: 0, totalCollected: 0, lastActivity: date };
    }
    customerStats[customer].lastActivity = date;
    if (action === 'Delivery') {
      cylinderOwner[uid] = customer;
      customerStats[customer].totalDelivered++;
    } else if (action === 'Collection') {
      if (cylinderOwner[uid] === customer) delete cylinderOwner[uid];
      customerStats[customer].totalCollected++;
    }
  }

  // ── Build outstanding per customer ─────────────────────
  const customerOutstanding = {};
  for (const [uid, customer] of Object.entries(cylinderOwner)) {
    if (!customerOutstanding[customer]) customerOutstanding[customer] = [];
    customerOutstanding[customer].push(uid);
  }

  // ── Assemble rows (only customers with outstanding > 0) ─
  const rows = [];
  for (const [customer, stats] of Object.entries(customerStats)) {
    const outstanding = customerOutstanding[customer] || [];
    if (outstanding.length > 0) {
      rows.push({
        customer,
        delivered : stats.totalDelivered,
        collected : stats.totalCollected,
        outstanding: outstanding.length,
        uids      : outstanding.join(', '),
        lastActivity: stats.lastActivity
      });
    }
  }
  rows.sort((a, b) => b.outstanding - a.outstanding);

  const totalOut   = rows.reduce((s, r) => s + r.outstanding, 0);
  const today      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
  const dayName    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEEE');

  // ── Build HTML email ───────────────────────────────────
  const tableRows = rows.map((r, i) => `
    <tr style="background:${i % 2 === 0 ? '#f8fafc' : '#ffffff'}">
      <td style="padding:10px 14px;font-weight:600;color:#1a1a2e">${r.customer}</td>
      <td style="padding:10px 14px;text-align:center">${r.delivered}</td>
      <td style="padding:10px 14px;text-align:center">${r.collected}</td>
      <td style="padding:10px 14px;text-align:center">
        <span style="background:${r.outstanding > 10 ? '#f8d7da' : r.outstanding > 5 ? '#fff3cd' : '#d4edda'};
                     color:${r.outstanding > 10 ? '#721c24' : r.outstanding > 5 ? '#856404' : '#155724'};
                     padding:3px 10px;border-radius:12px;font-weight:700">
          ${r.outstanding}
        </span>
      </td>
      <td style="padding:10px 14px;font-size:12px;color:#64748b">${r.uids}</td>
    </tr>`).join('');

  const htmlBody = `
  <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:750px;margin:0 auto">

    <!-- Header -->
    <div style="background:#1a1a2e;padding:28px 32px;border-radius:12px 12px 0 0">
      <h1 style="margin:0;color:#ffffff;font-size:22px">🛢 Cylinder Outstanding Ledger</h1>
      <p style="margin:6px 0 0;color:#94a3b8;font-size:14px">${dayName}, ${today} &nbsp;·&nbsp; Daily Morning Report</p>
    </div>

    <!-- Summary Cards -->
    <div style="background:#f0f4ff;padding:20px 32px;display:flex;gap:16px">
      <div style="background:#fff;border-radius:10px;padding:16px 24px;flex:1;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
        <div style="font-size:32px;font-weight:800;color:#1a1a2e">${totalOut}</div>
        <div style="font-size:12px;color:#64748b;margin-top:4px">Total Cylinders Out</div>
      </div>
      <div style="background:#fff;border-radius:10px;padding:16px 24px;flex:1;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
        <div style="font-size:32px;font-weight:800;color:#0f766e">${rows.length}</div>
        <div style="font-size:12px;color:#64748b;margin-top:4px">Customers with Cylinders</div>
      </div>
      <div style="background:#fff;border-radius:10px;padding:16px 24px;flex:1;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,0.08)">
        <div style="font-size:32px;font-weight:800;color:#b45309">${rows.filter(r => r.outstanding > 10).length}</div>
        <div style="font-size:12px;color:#64748b;margin-top:4px">High Outstanding (>10)</div>
      </div>
    </div>

    <!-- Table -->
    <div style="padding:0 0 24px">
      ${rows.length > 0 ? `
      <table style="width:100%;border-collapse:collapse;font-size:14px">
        <thead>
          <tr style="background:#16213e">
            <th style="padding:12px 14px;color:#e2e8f0;text-align:left;font-weight:600">Customer</th>
            <th style="padding:12px 14px;color:#e2e8f0;text-align:center;font-weight:600">Delivered</th>
            <th style="padding:12px 14px;color:#e2e8f0;text-align:center;font-weight:600">Collected</th>
            <th style="padding:12px 14px;color:#e2e8f0;text-align:center;font-weight:600">Outstanding</th>
            <th style="padding:12px 14px;color:#e2e8f0;text-align:left;font-weight:600">Cylinder UIDs</th>
          </tr>
        </thead>
        <tbody>${tableRows}</tbody>
      </table>` : `
      <div style="text-align:center;padding:40px;color:#64748b;font-style:italic">
        ✅ No outstanding cylinders! All returned.
      </div>`}
    </div>

    <!-- Footer -->
    <div style="background:#f8fafc;padding:16px 32px;border-radius:0 0 12px 12px;border-top:1px solid #e2e8f0">
      <p style="margin:0;font-size:12px;color:#94a3b8;text-align:center">
        📊 Auto-generated by Cylinder Tracker &nbsp;·&nbsp; ${today} 8:00 AM
        &nbsp;·&nbsp; <a href="https://docs.google.com/spreadsheets" style="color:#1a73e8">Open Google Sheet</a>
      </p>
    </div>

  </div>`;

  // ── Send the email ─────────────────────────────────────
  MailApp.sendEmail({
    to      : EMAIL_RECIPIENTS.join(','),
    subject : `🛢 Cylinder Ledger — ${totalOut} cylinders out | ${today}`,
    htmlBody: htmlBody,
    name    : EMAIL_SENDER_NAME
  });

  Logger.log(`Daily ledger email sent to: ${EMAIL_RECIPIENTS.join(', ')}`);
}


// ── Sets up the daily 8 AM trigger (run this ONCE) ─────────
function setupDailyTrigger() {
  // Remove existing triggers with same name to avoid duplicates
  removeDailyTrigger(true);

  ScriptApp.newTrigger('sendDailyLedgerEmail')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Daily email trigger set!\n\n' +
    '📧 An email will be sent every morning at 8 AM to:\n' +
    EMAIL_RECIPIENTS.join('\n') + '\n\n' +
    'To remove it, click "🚫 Remove Daily Email Trigger" from the menu.'
  );
}


// ── Removes the daily trigger ───────────────────────────────
function removeDailyTrigger(silent = false) {
  const triggers = ScriptApp.getProjectTriggers();
  let removed    = 0;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendDailyLedgerEmail') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      removed > 0
        ? `✅ Daily email trigger removed. No more automatic emails.`
        : `ℹ️ No daily trigger was active.`
    );
  }
}


// ============================================================
//  SHEET 2 — CUSTOMER MAP
// ============================================================

function setupCustomerMapSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let mapSheet = ss.getSheetByName(MAP_SHEET_NAME);
  if (!mapSheet) {
    mapSheet = ss.insertSheet(MAP_SHEET_NAME);
  }

  mapSheet.clearContents();
  mapSheet.clearFormats();

  const headers = ['Date', 'Time', 'Driver', 'Action', 'Cyl. Count', 'Cylinder UIDs', 'Customer', 'Send Receipt?', 'Receipt Status'];
  mapSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  mapSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  mapSheet.setColumnWidth(1, 120);
  mapSheet.setColumnWidth(2, 100);
  mapSheet.setColumnWidth(3, 130);
  mapSheet.setColumnWidth(4, 110);
  mapSheet.setColumnWidth(5, 90);
  mapSheet.setColumnWidth(6, 360);
  mapSheet.setColumnWidth(7, 200);
  mapSheet.setColumnWidth(8, 120);
  mapSheet.setColumnWidth(9, 180);

  mapSheet.setFrozenRows(1);
  mapSheet.getRange(1, 7).setNote('⬅ Manager fills this column after driver calls in.');
  mapSheet.getRange(1, 8).setNote('⬅ Check this box to email a receipt to the customer.');

  SpreadsheetApp.getUi().alert('✅ Customer Map sheet is ready!\n\nNext: Click "Refresh New Batches → Sheet 2"');
}


// Pull new unassigned batches from Sheet 1 into Sheet 2
// Pull new unassigned batches from Sheet 1 into Sheet 2 (supports silent auto-refresh and bi-directional sync)
function refreshBatches(silent = false) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const scanSheet = ss.getSheetByName(SCAN_SHEET_NAME);
  const mapSheet  = ss.getSheetByName(MAP_SHEET_NAME);

  if (!scanSheet) {
    if (!silent) SpreadsheetApp.getUi().alert('❌ Sheet not found: ' + SCAN_SHEET_NAME);
    return;
  }
  if (!mapSheet) {
    if (!silent) SpreadsheetApp.getUi().alert('❌ Run "Setup Customer Map Sheet" first.');
    return;
  }

  // Auto-upgrade Sheet 2 from 7 columns to 9 columns if needed, preserving existing data
  const maxCols = mapSheet.getMaxColumns();
  if (maxCols < 9) {
    mapSheet.insertColumnsAfter(maxCols, 9 - maxCols);
    mapSheet.getRange(1, 8, 1, 2).setValues([['Send Receipt?', 'Receipt Status']]);
    mapSheet.getRange(1, 8, 1, 2)
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center');
    
    const lastRow = mapSheet.getLastRow();
    if (lastRow > 1) {
      mapSheet.getRange(2, 8, lastRow - 1, 1).insertCheckboxes();
    }
    mapSheet.setColumnWidth(8, 120);
    mapSheet.setColumnWidth(9, 180);
  }

  const scanData = scanSheet.getDataRange().getValues();
  const batches = {};
  for (let i = 1; i < scanData.length; i++) {
    const row    = scanData[i];
    const date   = formatDate(row[0]);
    const time   = formatTime(row[1]);
    const driver = String(row[2]).trim();
    const action = String(row[3]).trim();
    const uid    = String(row[4]).trim();
    const customer = (row.length >= 6) ? String(row[5]).trim() : '';

    if (!uid || !driver || action === 'Filling') continue;

    const key = `${date}||${time}||${driver}||${action}`;
    if (!batches[key]) {
      batches[key] = { date, time, driver, action, cylinders: [], customer: customer };
    }
    batches[key].cylinders.push(uid);
    if (customer && !batches[key].customer) {
      batches[key].customer = customer;
    }
  }

  // Find batches already in Sheet 2 & Sync Deletes/Updates
  const mapData    = mapSheet.getDataRange().getValues();
  const mappedKeys = new Set();
  const rowsToDelete = [];
  let updatedCount = 0;

  // Traverse bottom-to-top to delete rows safely without index shifting
  for (let i = mapData.length - 1; i >= 1; i--) {
    const r   = mapData[i];
    const key = `${formatDate(r[0])}||${formatTime(r[1])}||${String(r[2]).trim()}||${String(r[3]).trim()}`;
    const rowNum = i + 1;

    if (batches[key]) {
      if (mappedKeys.has(key)) {
        // Obsolete duplicate row — delete it!
        rowsToDelete.push(rowNum);
      } else {
        mappedKeys.add(key);
        const b = batches[key];
        const currentCount = parseInt(r[4]);
        const currentUids  = String(r[5]).trim();
        const currentCust  = String(r[6] || '').trim();
        const newCount     = b.cylinders.length;
        const newUids      = b.cylinders.join(', ');

        let needsUpdate = false;
        // If count or UIDs changed, sync them
        if (currentCount !== newCount || currentUids !== newUids) {
          mapSheet.getRange(rowNum, 5).setValue(newCount);
          mapSheet.getRange(rowNum, 6).setValue(newUids);
          needsUpdate = true;
        }
        // Sync customer if blank in Sheet 2 but driver set it
        if (!currentCust && b.customer) {
          mapSheet.getRange(rowNum, 7).setValue(b.customer);
          needsUpdate = true;
        }
        if (needsUpdate) {
          updatedCount++;
        }
      }
    } else {
      // Obsolete batch (all scans deleted in Sheet 1)
      rowsToDelete.push(rowNum);
    }
  }

  // Perform deletions
  for (let j = 0; j < rowsToDelete.length; j++) {
    mapSheet.deleteRow(rowsToDelete[j]);
  }

  // Collect only NEW batches
  const newRows = [];
  for (const [key, b] of Object.entries(batches)) {
    if (!mappedKeys.has(key)) {
      newRows.push([b.date, b.time, b.driver, b.action, b.cylinders.length, b.cylinders.join(', '), b.customer || '', false, '']);
    }
  }

  if (newRows.length > 0) {
    const startRow = mapSheet.getLastRow() + 1;
    mapSheet.getRange(startRow, 1, newRows.length, 9).setValues(newRows);

    for (let i = 0; i < newRows.length; i++) {
      const rowNum  = startRow + i;
      const bgColor = (rowNum % 2 === 0) ? '#f0f4ff' : '#ffffff';
      mapSheet.getRange(rowNum, 1, 1, 9).setBackground(bgColor).setVerticalAlignment('middle');

      const action     = newRows[i][3];
      const actionCell = mapSheet.getRange(rowNum, 4);
      if (action === 'Delivery') {
        actionCell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
      } else if (action === 'Collection') {
        actionCell.setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');
      }

      mapSheet.getRange(rowNum, 7)
        .setBackground('#fffde7')
        .setFontColor('#5d4037')
        .setFontStyle('italic');

      applyCustomerDropdown(mapSheet, rowNum);
      
      // Insert checkbox in Column 8 (Send Receipt?)
      mapSheet.getRange(rowNum, 8).insertCheckboxes();
    }
  }

  if (!silent) {
    let summary = '';
    if (newRows.length > 0) summary += `✅ Added ${newRows.length} new batch(es).\n`;
    if (rowsToDelete.length > 0) summary += `🗑 Deleted ${rowsToDelete.length} obsolete batch(es) from Sheet 2.\n`;
    if (updatedCount > 0) summary += `🔄 Updated cylinder list for ${updatedCount} batch(es).\n`;
    
    if (summary === '') {
      SpreadsheetApp.getUi().alert('✅ Sheet 2 is already fully in sync with Sheet 1.');
    } else {
      SpreadsheetApp.getUi().alert(summary + '\nSheet 2 is now up to date!');
    }
  }
}

// ── AUTO-REFRESH TRIGGER CONFIGURATION ───────────────────────

function setupAutoRefreshTrigger() {
  // Clear any existing duplicates
  removeAutoRefreshTrigger(true);

  // Set up installable onChange trigger
  ScriptApp.newTrigger('autoRefreshHandler')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Auto-Refresh Enabled!\n\n' +
    'Sheet 2 (Customer Map) will now automatically pull new scan batches whenever a driver submits scans.'
  );
}

function removeAutoRefreshTrigger(silent = false) {
  const triggers = ScriptApp.getProjectTriggers();
  let removed    = 0;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'autoRefreshHandler') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      removed > 0
        ? '✅ Auto-Refresh Disabled. Sheet 2 will no longer update automatically.'
        : 'ℹ️ Auto-Refresh was not active.'
    );
  }
}

function autoRefreshHandler(e) {
  // Trigger silent refresh and real-time emails when rows are appended by Flask API
  if (e && (e.changeType === 'INSERT_ROW' || e.changeType === 'EDIT' || e.changeType === 'OTHER')) {
    onNewRow();
    refreshBatches(true);
  }
}


// ── FIXED: reads Column B (Name) dynamically, NOT Column A (Customer ID) ─
function applyCustomerDropdown(mapSheet, rowNum) {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
  if (!customerSheet) return;

  const lastRow = customerSheet.getLastRow();
  if (lastRow < 2) return;

  // ✅ Link dynamically to Column B range (Customer Name)
  const customerRange = customerSheet.getRange(2, 2, lastRow - 1, 1);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(customerRange, true) // Dynamically linked!
    .setAllowInvalid(true)
    .build();

  mapSheet.getRange(rowNum, 7).setDataValidation(rule);
}


// ============================================================
//  SHEET 3 — OUTSTANDING LEDGER
// ============================================================

// silent = true when called from onEdit (no alert popup)
function refreshLedger(silent = false) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const scanSheet = ss.getSheetByName(SCAN_SHEET_NAME);
  const mapSheet  = ss.getSheetByName(MAP_SHEET_NAME);

  if (!scanSheet || !mapSheet) {
    if (!silent) SpreadsheetApp.getUi().alert('❌ Sheet 1 or Sheet 2 not found. Please set up first.');
    return;
  }

  let ledgerSheet = ss.getSheetByName(LEDGER_SHEET_NAME);
  if (!ledgerSheet) {
    ledgerSheet = ss.insertSheet(LEDGER_SHEET_NAME);
  }

  // ── Read existing date filter values BEFORE clearing ─────
  const existingFrom = ledgerSheet.getRange('B1').getValue();
  const existingTo   = ledgerSheet.getRange('D1').getValue();

  // Parse filter dates
  const fromDate = parseFilterDate(existingFrom);
  const toDate   = parseFilterDate(existingTo);

  // ── Build batch → customer lookup from Sheet 2 ────────────
  const mapData         = mapSheet.getDataRange().getValues();
  const batchToCustomer = {};

  for (let i = 1; i < mapData.length; i++) {
    const r        = mapData[i];
    const customer = String(r[6]).trim();
    if (!customer) continue;
    const key = `${formatDate(r[0])}||${formatTime(r[1])}||${String(r[2]).trim()}||${String(r[3]).trim()}`;
    batchToCustomer[key] = customer;
  }

  // ── Build events from Sheet 1 + customer info ─────────────
  const scanData = scanSheet.getDataRange().getValues();
  const events   = [];

  for (let i = 1; i < scanData.length; i++) {
    const row    = scanData[i];
    const date   = formatDate(row[0]);
    const time   = formatTime(row[1]);
    const driver = String(row[2]).trim();
    const action = String(row[3]).trim();
    const uid    = String(row[4]).trim();

    if (!uid) continue;

    const key      = `${date}||${time}||${driver}||${action}`;
    const customer = batchToCustomer[key];
    if (!customer) continue;

    const dateObj = parseDateTime(date, time);

    // ── Apply date filter ─────────────────────────────────
    if (fromDate && dateObj < fromDate) continue;
    if (toDate) {
      // Include full toDate day (up to 23:59:59)
      const toDateEnd = new Date(toDate);
      toDateEnd.setHours(23, 59, 59);
      if (dateObj > toDateEnd) continue;
    }

    events.push({ date, time, action, uid, customer, dateObj });
  }

  // ── Sort chronologically ──────────────────────────────────
  events.sort((a, b) => a.dateObj - b.dateObj);

  // ── Track cylinder ownership ──────────────────────────────
  const cylinderOwner = {};
  const customerStats = {};

  for (const ev of events) {
    const { action, uid, customer, date } = ev;

    if (!customerStats[customer]) {
      customerStats[customer] = { totalDelivered: 0, totalCollected: 0, lastActivity: date };
    }
    customerStats[customer].lastActivity = date;

    if (action === 'Delivery') {
      cylinderOwner[uid] = customer;
      customerStats[customer].totalDelivered++;
    } else if (action === 'Collection') {
      if (cylinderOwner[uid] === customer) delete cylinderOwner[uid];
      customerStats[customer].totalCollected++;
    }
  }

  // ── Build outstanding list per customer ───────────────────
  const customerOutstanding = {};
  for (const [uid, customer] of Object.entries(cylinderOwner)) {
    if (!customerOutstanding[customer]) customerOutstanding[customer] = [];
    customerOutstanding[customer].push(uid);
  }

  // ── Assemble rows ─────────────────────────────────────────
  const rows = [];
  for (const [customer, stats] of Object.entries(customerStats)) {
    const outstanding = customerOutstanding[customer] || [];
    rows.push([
      customer,
      stats.totalDelivered,
      stats.totalCollected,
      outstanding.length,
      outstanding.join(', ') || '—',
      stats.lastActivity
    ]);
  }
  rows.sort((a, b) => b[3] - a[3]);

  // ── Write Sheet 3 ─────────────────────────────────────────
  ledgerSheet.clearContents();
  ledgerSheet.clearFormats();

  // ── Row 1: Date Filter Bar ────────────────────────────────
  ledgerSheet.setRowHeight(1, 38);

  // Label: From
  ledgerSheet.getRange('A1')
    .setValue('📅 From:')
    .setBackground('#e8f0fe')
    .setFontColor('#1a73e8')
    .setFontWeight('bold')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  // From date input cell
  ledgerSheet.getRange('B1')
    .setBackground('#ffffff')
    .setFontColor('#1a1a2e')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('dd-mm-yyyy')
    .setBorder(true, true, true, true, false, false, '#1a73e8', SpreadsheetApp.BorderStyle.SOLID);

  // Label: To
  ledgerSheet.getRange('C1')
    .setValue('To:')
    .setBackground('#e8f0fe')
    .setFontColor('#1a73e8')
    .setFontWeight('bold')
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  // To date input cell
  ledgerSheet.getRange('D1')
    .setBackground('#ffffff')
    .setFontColor('#1a1a2e')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('dd-mm-yyyy')
    .setBorder(true, true, true, true, false, false, '#1a73e8', SpreadsheetApp.BorderStyle.SOLID);

  // Hint label
  const filterActive = fromDate || toDate;
  ledgerSheet.getRange('E1').merge();
  ledgerSheet.getRange('E1')
    .setValue(filterActive ? '🔍 Filter active — clear cells B1 & D1 to show all dates' : 'ℹ️ Leave blank to show all dates')
    .setBackground('#e8f0fe')
    .setFontColor(filterActive ? '#b45309' : '#5f6368')
    .setFontStyle('italic')
    .setVerticalAlignment('middle');

  // Restore filter values (preserved through clear)
  if (existingFrom) ledgerSheet.getRange('B1').setValue(existingFrom);
  if (existingTo)   ledgerSheet.getRange('D1').setValue(existingTo);

  // ── Row 2: Summary Banner ─────────────────────────────────
  const totalOutstanding = rows.reduce((sum, r) => sum + r[3], 0);
  const timestamp        = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
  const filterLabel      = fromDate || toDate
    ? ` | 📅 ${existingFrom ? formatDate(existingFrom) : '...'} → ${existingTo ? formatDate(existingTo) : 'Today'}`
    : '';

  ledgerSheet.getRange(2, 1, 1, 6).merge()
    .setValue(`🛢 Cylinder Outstanding Ledger${filterLabel}   |   Total Out: ${totalOutstanding}   |   Updated: ${timestamp}`)
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  ledgerSheet.setRowHeight(2, 38);

  // ── Row 3: Column Headers ─────────────────────────────────
  const headers = ['Customer', 'Total Delivered', 'Total Collected', 'Outstanding', 'Outstanding Cylinder UIDs', 'Last Activity'];
  ledgerSheet.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground('#16213e')
    .setFontColor('#e2e8f0')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  ledgerSheet.setRowHeight(3, 32);

  // ── Rows 4+: Data ─────────────────────────────────────────
  if (rows.length > 0) {
    ledgerSheet.getRange(4, 1, rows.length, headers.length).setValues(rows);

    for (let i = 0; i < rows.length; i++) {
      const rowNum      = i + 4;
      const outstanding = rows[i][3];
      const bgColor     = (i % 2 === 0) ? '#f8fafc' : '#ffffff';

      ledgerSheet.getRange(rowNum, 1, 1, headers.length)
        .setBackground(bgColor)
        .setVerticalAlignment('middle');
      ledgerSheet.setRowHeight(rowNum, 36);

      ledgerSheet.getRange(rowNum, 1).setFontWeight('bold').setFontSize(11);

      const outstandingCell = ledgerSheet.getRange(rowNum, 4);
      if (outstanding === 0) {
        outstandingCell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
      } else if (outstanding <= 5) {
        outstandingCell.setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');
      } else {
        outstandingCell.setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold');
      }

      ledgerSheet.getRange(rowNum, 5).setWrap(true).setFontSize(9).setFontColor('#475569');
    }
  } else {
    ledgerSheet.getRange(4, 1, 1, 6).merge()
      .setValue('No data found for the selected date range.')
      .setFontColor('#888888')
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
  }

  // ── Column widths ─────────────────────────────────────────
  ledgerSheet.setColumnWidth(1, 180);
  ledgerSheet.setColumnWidth(2, 130);
  ledgerSheet.setColumnWidth(3, 130);
  ledgerSheet.setColumnWidth(4, 110);
  ledgerSheet.setColumnWidth(5, 420);
  ledgerSheet.setColumnWidth(6, 120);

  ledgerSheet.setFrozenRows(3); // Freeze filter row + banner + headers

  if (!silent) {
    SpreadsheetApp.getUi().alert(`✅ Outstanding Ledger updated!\n\n📦 Total cylinders out: ${totalOutstanding}\n👥 Customers: ${rows.length}${filterActive ? '\n📅 Date filter is active' : ''}`);
  }
}


// ============================================================
//  FIX ALL DROPDOWNS — Run once from menu to fix existing rows
// ============================================================

function fixAllDropdowns() {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet      = ss.getSheetByName(MAP_SHEET_NAME);
  const customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);

  if (!mapSheet) {
    SpreadsheetApp.getUi().alert('❌ Customer Map sheet not found.');
    return;
  }
  if (!customerSheet) {
    SpreadsheetApp.getUi().alert('❌ Customers sheet not found.');
    return;
  }

  const lastCustomerRow = customerSheet.getLastRow();
  if (lastCustomerRow < 2) {
    SpreadsheetApp.getUi().alert('❌ No customers found in Customers sheet.');
    return;
  }

  // ✅ Link dynamically to Column B range (Customer Name)
  const customerRange = customerSheet.getRange(2, 2, lastCustomerRow - 1, 1);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(customerRange, true) // Dynamically linked!
    .setAllowInvalid(true)
    .build();

  const lastMapRow = mapSheet.getLastRow();
  if (lastMapRow < 2) {
    SpreadsheetApp.getUi().alert('No rows in Customer Map to fix.');
    return;
  }

  // Apply to entire Column G in one shot
  mapSheet.getRange(2, 7, lastMapRow - 1, 1).setDataValidation(rule);

  SpreadsheetApp.getUi().alert(
    '✅ Done!\n\nAll dropdowns linked dynamically to the Customers sheet.\n' +
    'Adding new customers will now automatically update all dropdown lists instantly.'
  );
}


// ============================================================
//  HELPERS
// ============================================================

function formatDate(value) {
  if (!value) return '';
  if (value instanceof Date) {
    const d = String(value.getDate()).padStart(2, '0');
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const y = value.getFullYear();
    return `${d}-${m}-${y}`;
  }
  return String(value).trim();
}

function formatTime(value) {
  if (!value) return '';
  if (value instanceof Date) {
    const h  = String(value.getHours()).padStart(2, '0');
    const mi = String(value.getMinutes()).padStart(2, '0');
    const s  = String(value.getSeconds()).padStart(2, '0');
    return `${h}:${mi}:${s}`;
  }
  return String(value).trim();
}

function parseDateTime(dateStr, timeStr) {
  try {
    const [d, m, y]   = dateStr.split('-').map(Number);
    const [h, mi, s]  = (timeStr || '00:00:00').split(':').map(Number);
    return new Date(y, m - 1, d, h, mi, s);
  } catch (e) {
    return new Date(0);
  }
}

// Parses a filter date cell value — handles Date objects and DD-MM-YYYY strings
function parseFilterDate(value) {
  if (!value || value === '') return null;
  if (value instanceof Date && !isNaN(value)) {
    // Set to start of day
    return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0);
  }
  if (typeof value === 'string') {
    const parts = value.trim().split('-');
    if (parts.length === 3) {
      const [d, m, y] = parts.map(Number);
      if (!isNaN(d) && !isNaN(m) && !isNaN(y)) {
        return new Date(y, m - 1, d, 0, 0, 0);
      }
    }
  }
  return null;
}

// ============================================================
//  REAL-TIME SCAN EMAILS
// ============================================================
var LAST_ROW_KEY = "lastEmailedRow";

function onNewRow() {
  const lock = LockService.getScriptLock();
  try {
    // Wait up to 15 seconds for lock to avoid concurrent execution issues
    lock.waitLock(15000);
  } catch (e) {
    Logger.log('Could not acquire lock: ' + e.toString());
    return;
  }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    if (!sheet) return;
    
    var lastRow = sheet.getLastRow();
    var props = PropertiesService.getScriptProperties();
    var lastEmailedVal = props.getProperty(LAST_ROW_KEY);
    
    var lastEmailed;
    if (lastEmailedVal === null) {
      // First run: initialize to lastRow - 1 so we only process the very latest row
      var startRow = lastRow > 1 ? lastRow - 1 : 1;
      props.setProperty(LAST_ROW_KEY, startRow.toString());
      lastEmailed = startRow;
    } else {
      lastEmailed = parseInt(lastEmailedVal, 10);
    }

    if (isNaN(lastEmailed) || lastEmailed < 1) {
      lastEmailed = 1;
    }

    if (lastRow <= lastEmailed) {
      return;
    }

    // Fetch all new rows in a single call to save network time
    var numRows = lastRow - lastEmailed;
    var range = sheet.getRange(lastEmailed + 1, 1, numRows, 5);
    var allValues = range.getValues();

    // Check if the last row is incomplete (still being written by API)
    var lastRowValues = allValues[allValues.length - 1];
    var isLastRowIncomplete = false;
    if (lastRowValues) {
      var d = String(lastRowValues[2] || '').trim();
      var a = String(lastRowValues[3] || '').trim();
      var c = String(lastRowValues[4] || '').trim();
      if ((d || a || c) && (!d || !a || !c)) {
        isLastRowIncomplete = true;
      }
    }

    // Wait 1.5 seconds if we suspect Google Sheets is still writing the cell values
    if (isLastRowIncomplete) {
      Utilities.sleep(1500);
      allValues = range.getValues();
    }

    // Group rows by Unique Batch (Date + Time + Driver + Action)
    var batches = {};
    var processedCount = 0;

    for (var i = 0; i < allValues.length; i++) {
      var rowValues = allValues[i];
      var date     = rowValues[0];
      var time     = rowValues[1];
      var driver   = String(rowValues[2] || '').trim();
      var action   = String(rowValues[3] || '').trim();
      var cylinder = String(rowValues[4] || '').trim();

      // If the row is completely empty, skip it to avoid getting stuck
      if (!date && !time && !driver && !action && !cylinder) {
        processedCount++;
        continue;
      }

      // If a row is partially incomplete (missing driver, action, or cylinder):
      if (!driver || !action || !cylinder) {
        // If it is the last row of the fetched range, break and leave it for the next run
        if (i === allValues.length - 1) {
          break;
        }
        // Otherwise, skip this malformed row so we don't get stuck
        processedCount++;
        continue;
      }

      // Format Date and Time values properly in case they are Date objects
      var dateStr = (date instanceof Date) ? Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MM-yyyy") : String(date).trim();
      var timeStr = (time instanceof Date) ? Utilities.formatDate(time, Session.getScriptTimeZone(), "HH:mm:ss") : String(time).trim();

      var batchKey = dateStr + "||" + timeStr + "||" + driver + "||" + action;

      if (!batches[batchKey]) {
        batches[batchKey] = {
          date: dateStr,
          time: timeStr,
          driver: driver,
          action: action,
          cylinders: []
        };
      }
      batches[batchKey].cylinders.push(cylinder);
      processedCount++;
    }

    if (processedCount === 0) {
      return;
    }

    // Send a separate email for each unique batch
    for (var key in batches) {
      var b = batches[key];
      var count = b.cylinders.length;

      // Skip sending emails for in-house Filling scans
      if (b.action === 'Filling') {
        Logger.log("Skipping email for Filling batch: " + key);
        continue;
      }

      var body = "Cylinder Tracking Report\n\n" +
                 "Driver : " + b.driver + "\n" +
                 "Action : " + b.action + "\n" +
                 "Total Cylinders : " + count + "\n" +
                 "Date : " + b.date + "\n" +
                 "Time : " + b.time + "\n\n" +
                 "Open Google Sheet : " + SpreadsheetApp.getActiveSpreadsheet().getUrl();

      try {
        MailApp.sendEmail(
          EMAIL_RECIPIENTS.join(','),
          b.action + " Report (" + count + " Cylinders)",
          body
        );
        Logger.log("Email sent for batch: " + key);
      } catch (err) {
        Logger.log("Failed to send email for batch: " + key + ". Error: " + err.toString());
      }
    }

    // Update the last emailed row marker only up to what we successfully processed
    var newLastEmailed = lastEmailed + processedCount;
    props.setProperty(LAST_ROW_KEY, newLastEmailed.toString());
  } finally {
    lock.releaseLock();
  }
}

function resetEmailCounter() {
  PropertiesService.getScriptProperties().deleteProperty(LAST_ROW_KEY);
  SpreadsheetApp.getUi().alert('✅ Scan email counter has been reset.\n\nThe next scan submission will initialize the counter and send an email.');
}

// ============================================================
//  CUSTOMER EMAIL RECEIPTS
// ============================================================

function sendCustomerReceipt(rowNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(MAP_SHEET_NAME);
  const customerSheet = ss.getSheetByName(CUSTOMER_SHEET_NAME);
  if (!mapSheet || !customerSheet) return;

  const rowData = mapSheet.getRange(rowNum, 1, 1, 9).getValues()[0];
  const dateVal = formatDate(rowData[0]);
  const timeVal = formatTime(rowData[1]);
  const driver = String(rowData[2]).trim();
  const action = String(rowData[3]).trim();
  const count = parseInt(rowData[4]);
  const uids = String(rowData[5]).trim();
  const customerName = String(rowData[6]).trim();

  if (!customerName) {
    ss.toast("❌ Please assign a customer first before sending the receipt.", "Receipt Error");
    mapSheet.getRange(rowNum, 8).setValue(false);
    return;
  }

  // Look up customer's email from Column C (assuming Col A is ID, Col B is Name, Col C is Email)
  const customerData = customerSheet.getDataRange().getValues();
  let email = '';
  for (let i = 1; i < customerData.length; i++) {
    const r = customerData[i];
    const name = String(r[1] || r[0]).trim(); // Column B (or fallback Col A)
    if (name.toLowerCase() === customerName.toLowerCase()) {
      email = String(r[2] || r[1]).trim(); // Column C (or fallback Col B)
      break;
    }
  }

  if (!email || email.indexOf('@') === -1) {
    ss.toast("❌ Email address not found or invalid for customer: " + customerName, "Receipt Error");
    mapSheet.getRange(rowNum, 8).setValue(false);
    mapSheet.getRange(rowNum, 9).setValue("Missing Email");
    return;
  }

  // Get live outstanding count and outstanding UIDs
  const outstandingData = getCustomerOutstandingData(customerName);
  const outstandingCount = outstandingData.count;
  const outstandingUidsList = outstandingData.uidsList;

  const subject = `Cylinder Receipt: ${action} Summary — ${customerName}`;
  
  // Branded HTML email body
  const htmlBody = `
  <div style="font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;max-width:600px;margin:0 auto;color:#2C2C2A;background:#f7f8f6;padding:20px;border-radius:12px;">
    <!-- Card Container -->
    <div style="background:#ffffff;border:1.5px solid #d8d9d4;border-radius:14px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.05);">
      
      <!-- Banner Header -->
      <div style="background:#0F6E56;padding:24px;text-align:center;color:#ffffff;">
        <div style="font-size:12px;text-transform:uppercase;letter-spacing:1px;font-weight:600;opacity:0.85;">Transaction Receipt</div>
        <h2 style="margin:6px 0 0;font-size:22px;font-weight:600;letter-spacing:-0.3px;">Cylinder ${action}</h2>
      </div>

      <div style="padding:28px 24px;">
        <p style="font-size:15px;margin-bottom:16px;line-height:1.4;">Dear <strong>${customerName}</strong>,</p>
        <p style="font-size:14px;color:#5F5E5A;margin-bottom:24px;line-height:1.5;">
          This is to confirm that a cylinder transaction has been logged for your account. Details of the transaction are below:
        </p>

        <!-- Transaction Details Table -->
        <table style="width:100%;border-collapse:collapse;margin-bottom:28px;font-size:14px;">
          <thead>
            <tr style="background:#E1F5EE;">
              <th style="padding:10px 12px;text-align:left;font-weight:600;color:#04342C;border-bottom:1.5px solid #1D9E75;">Detail</th>
              <th style="padding:10px 12px;text-align:left;font-weight:600;color:#04342C;border-bottom:1.5px solid #1D9E75;">Value</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#5F5E5A;">Transaction Type</td>
              <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;font-weight:600;color:${action === 'Delivery' ? '#0F6E56' : '#c2410c'};">${action}</td>
            </tr>
            <tr>
              <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#5F5E5A;">Date &amp; Time</td>
              <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;">${dateVal} &nbsp;·&nbsp; ${timeVal}</td>
            </tr>
            <tr>
              <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#5F5E5A;">Driver/Dispatcher</td>
              <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;">${driver}</td>
            </tr>
            <tr>
              <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#5F5E5A;">Cylinder Count</td>
              <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;font-weight:600;">${count}</td>
            </tr>
            <tr>
              <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#5F5E5A;vertical-align:top;">Cylinder UIDs</td>
              <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;font-size:12px;font-family:monospace;color:#2C2C2A;word-break:break-all;">${uids}</td>
            </tr>
          </tbody>
        </table>

        <!-- Outstanding Summary Card -->
        <div style="background:#E1F5EE;border:1.5px solid #1D9E75;border-radius:10px;padding:16px 20px;text-align:center;margin-bottom:28px;">
          <div style="font-size:12px;text-transform:uppercase;letter-spacing:0.8px;color:#0F6E56;font-weight:600;">Live Outstanding Balance</div>
          <div style="font-size:36px;font-weight:700;color:#04342C;margin:6px 0 2px;line-height:1;">${outstandingCount}</div>
          <div style="font-size:12px;color:#5F5E5A;">cylinders currently in your custody</div>
        </div>

        <!-- Outstanding UIDs List -->
        ${outstandingCount > 0 ? `
        <div style="margin-bottom:24px;">
          <h4 style="margin:0 0 8px;font-size:12px;color:#04342C;text-transform:uppercase;letter-spacing:0.6px;font-weight:600;">Cylinder UIDs in your custody:</h4>
          <div style="font-size:12.5px;color:#2C2C2A;background:#f7f8f6;border:1px solid #d8d9d4;padding:12px;border-radius:8px;font-family:monospace;word-break:break-all;line-height:1.4;">
            ${outstandingUidsList}
          </div>
        </div>
        ` : ''}

        <p style="font-size:13px;color:#5F5E5A;line-height:1.5;margin-bottom:0;">
          If you have any questions or notice any discrepancies, please reach out to our support team.
        </p>
      </div>

      <!-- Footer -->
      <div style="background:#f7f8f6;border-top:1px solid #d8d9d4;padding:16px;text-align:center;font-size:11px;color:#5F5E5A;">
        📊 Generated by <strong>Cylinder Tracker</strong> &nbsp;·&nbsp; Automated Transaction Log<br>
        Please do not reply directly to this email.
      </div>
    </div>
  </div>
  `;

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      name: EMAIL_SENDER_NAME
    });
    
    const nowTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
    mapSheet.getRange(rowNum, 8).setValue(false); // Auto-uncheck on success
    mapSheet.getRange(rowNum, 9).setValue("Sent @ " + nowTimestamp);
    ss.toast("✅ Receipt sent to " + email, "Success");
  } catch (err) {
    ss.toast("❌ Failed to send email: " + err.toString(), "Error");
    mapSheet.getRange(rowNum, 8).setValue(false);
    mapSheet.getRange(rowNum, 9).setValue("Error: " + err.toString().substring(0, 40));
  }
}

function getCustomerOutstandingCount(customerName) {
  return getCustomerOutstandingData(customerName).count;
}

function getCustomerOutstandingData(customerName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scanSheet = ss.getSheetByName(SCAN_SHEET_NAME);
  const mapSheet = ss.getSheetByName(MAP_SHEET_NAME);
  if (!scanSheet || !mapSheet) return { count: 0, uidsList: 'None' };

  const mapData = mapSheet.getDataRange().getValues();
  const batchToCustomer = {};
  for (let i = 1; i < mapData.length; i++) {
    const r = mapData[i];
    const customer = String(r[6]).trim();
    if (!customer) continue;
    const key = `${formatDate(r[0])}||${formatTime(r[1])}||${String(r[2]).trim()}||${String(r[3]).trim()}`;
    batchToCustomer[key] = customer;
  }

  const scanData = scanSheet.getDataRange().getValues();
  const events = [];
  for (let i = 1; i < scanData.length; i++) {
    const row = scanData[i];
    const date = formatDate(row[0]);
    const time = formatTime(row[1]);
    const driver = String(row[2]).trim();
    const action = String(row[3]).trim();
    const uid = String(row[4]).trim();
    if (!uid) continue;
    
    const key = `${date}||${time}||${driver}||${action}`;
    const customer = batchToCustomer[key];
    if (customer && customer.toLowerCase() === customerName.toLowerCase()) {
      events.push({ action: action, uid: uid, dateObj: parseDateTime(date, time) });
    }
  }

  events.sort((a, b) => a.dateObj - b.dateObj);

  const cylinderWithCustomer = {};
  for (const ev of events) {
    if (ev.action === 'Delivery') {
      cylinderWithCustomer[ev.uid] = true;
    } else if (ev.action === 'Collection') {
      delete cylinderWithCustomer[ev.uid];
    }
  }

  const uids = Object.keys(cylinderWithCustomer);
  uids.sort();
  
  return {
    count: uids.length,
    uidsList: uids.join(', ') || 'None'
  };
}
