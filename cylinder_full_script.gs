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
    .addItem('🔑 Enable Background Receipt Trigger', 'setupBackgroundReceiptTrigger')
    .addItem('🚫 Disable Background Receipt Trigger', 'removeBackgroundReceiptTrigger')
    .addItem('📧 Run Pending Receipts Process Now', 'processPendingReceipts')
    .addSeparator()
    .addItem('📧 Send Pending Scan Emails Now', 'onNewRow')
    .addItem('📧 Reset Scan Email Counter', 'resetEmailCounter')
    .addSeparator()
    .addItem('🔧 Fix All Dropdowns (run once)', 'fixAllDropdowns')
    .addSeparator()
    .addItem('🗄️ Setup Registry Sheets (Cylinders + Maintenance)', 'setupRegistrySheets')
    .addItem('📊 Setup Bulk Tanks Sheet', 'setupBulkTanksSheet')
    .addItem('📋 Setup Products Config Sheet', 'setupProductsSheet')
    .addItem('🔄 Rebuild Cylinder Statuses from Log', 'rebuildCylinderRegistry')
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

// ── BACKGROUND RECEIPT TRIGGER (Scans and sends pending receipts) ──

function setupBackgroundReceiptTrigger() {
  removeBackgroundReceiptTrigger(true);
  
  ScriptApp.newTrigger('processPendingReceipts')
    .timeBased()
    .everyMinutes(1)
    .create();
    
  SpreadsheetApp.getUi().alert(
    '⏰ Background Receipt Process Enabled!\n\n' +
    'The portal will now check for and send pending receipts every minute.'
  );
}

function removeBackgroundReceiptTrigger(silent = false) {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processPendingReceipts') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }
  if (!silent) {
    SpreadsheetApp.getUi().alert(
      removed > 0
        ? '✅ Background receipt trigger removed.'
        : 'ℹ️ No background receipt trigger was active.'
    );
  }
}

function processPendingReceipts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(MAP_SHEET_NAME);
  if (!mapSheet) return;
  
  const lastRow = mapSheet.getLastRow();
  if (lastRow < 2) return;
  
  const data = mapSheet.getRange(2, 8, lastRow - 1, 2).getValues(); // Read columns H (checkbox) and I (status)
  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    const sendReceipt = String(data[i][0]).toUpperCase();
    const status = String(data[i][1]).trim();
    
    if (sendReceipt === 'TRUE' && (status === 'Sending...' || status === 'Pending')) {
      try {
        sendCustomerReceipt(rowNum);
      } catch (err) {
        console.error("Error sending receipt for row " + rowNum + ": " + err);
        mapSheet.getRange(rowNum, 9).setValue("Error: " + err.message);
      }
    }
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
  // Automatically refresh Sheet 2 batches first to map the latest scans
  refreshBatches(true);

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
    } else if (action === 'Collection' || action === 'Filling') {
      delete cylinderOwner[uid];
      if (action === 'Collection') {
        customerStats[customer].totalCollected++;
      }
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

function getCustomerDailyTotals(customerName, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scanSheet = ss.getSheetByName(SCAN_SHEET_NAME);
  const mapSheet = ss.getSheetByName(MAP_SHEET_NAME);
  if (!scanSheet || !mapSheet) return { delivered: 0, collected: 0 };

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
  let delivered = 0;
  let collected = 0;

  for (let i = 1; i < scanData.length; i++) {
    const row = scanData[i];
    const date = formatDate(row[0]);
    if (date !== dateStr) continue;

    const time = formatTime(row[1]);
    const driver = String(row[2]).trim();
    const action = String(row[3]).trim();
    const uid = String(row[4]).trim();
    if (!uid) continue;

    const key = `${date}||${time}||${driver}||${action}`;
    const customer = batchToCustomer[key];
    if (customer && customer.toLowerCase() === customerName.toLowerCase()) {
      if (action === 'Delivery') {
        delivered++;
      } else if (action === 'Collection') {
        collected++;
      }
    }
  }

  return { delivered: delivered, collected: collected };
}

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

  // Look up other pending rows for the same customer on the same date to group them
  const lastRow = mapSheet.getLastRow();
  const allMapRows = mapSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const groupedRows = []; // Array of object: { rowNum, action, count, uids }
  
  // Include the current row first
  groupedRows.push({
    rowNum: rowNum,
    action: action,
    count: count,
    uids: uids
  });
  
  for (let i = 0; i < allMapRows.length; i++) {
    const curRowNum = i + 2;
    if (curRowNum === rowNum) continue;
    
    const rDate = formatDate(allMapRows[i][0]);
    const rAction = String(allMapRows[i][3]).trim();
    const rCount = parseInt(allMapRows[i][4]) || 0;
    const rUids = String(allMapRows[i][5]).trim();
    const rCust = String(allMapRows[i][6]).trim();
    const rSend = String(allMapRows[i][7]).toUpperCase();
    const rStatus = String(allMapRows[i][8]).trim();
    
    // Group if same customer, same date, and is pending/sending
    if (rCust.toLowerCase() === customerName.toLowerCase() && 
        rDate === dateVal && 
        rSend === 'TRUE' && 
        (rStatus === 'Sending...' || rStatus === 'Pending')) {
      groupedRows.push({
        rowNum: curRowNum,
        action: rAction,
        count: rCount,
        uids: rUids
      });
    }
  }

  // Aggregate grouped totals
  let totalDelivered = 0;
  let totalCollected = 0;
  let deliveryUids = [];
  let collectionUids = [];
  
  for (let g of groupedRows) {
    if (g.action === 'Delivery') {
      totalDelivered += g.count;
      if (g.uids) deliveryUids.push(g.uids);
    } else if (g.action === 'Collection') {
      totalCollected += g.count;
      if (g.uids) collectionUids.push(g.uids);
    }
  }
  
  const deliveryUidsStr = deliveryUids.join(', ');
  const collectionUidsStr = collectionUids.join(', ');

  // Get daily totals for this customer
  const dailyTotals = getCustomerDailyTotals(customerName, dateVal);
  const deliveredToday = dailyTotals.delivered;
  const collectedToday = dailyTotals.collected;

  // Get live outstanding count and outstanding UIDs
  const outstandingData = getCustomerOutstandingData(customerName);
  const outstandingCount = outstandingData.count;
  const outstandingUidsList = outstandingData.uidsList;

  // Determine email type and subject
  let emailActionType = action;
  let transactionRowsHtml = '';
  
  if (totalDelivered > 0 && totalCollected > 0) {
    emailActionType = 'Delivery & Collection';
    transactionRowsHtml = `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#5F5E5A;">Transaction Type</td>
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;font-weight:600;color:#0F6E56;">Delivery &amp; Collection</td>
      </tr>
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#5F5E5A;">Date &amp; Time</td>
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;">${dateVal} &nbsp;·&nbsp; ${timeVal}</td>
      </tr>
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#5F5E5A;">Driver/Dispatcher</td>
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;">${driver}</td>
      </tr>
      <tr style="background:#f0fbf6;">
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#0F6E56;font-weight:600;">Cylinders Delivered</td>
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;font-weight:600;color:#0F6E56;">${totalDelivered}</td>
      </tr>
      <tr style="background:#f0fbf6;">
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#5F5E5A;vertical-align:top;font-size:12px;padding-left:20px;">Delivered UIDs</td>
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;font-size:11px;font-family:monospace;color:#2C2C2A;word-break:break-all;">${deliveryUidsStr}</td>
      </tr>
      <tr style="background:#fffcf5;">
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#c2410c;font-weight:600;">Cylinders Collected</td>
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;font-weight:600;color:#c2410c;">${totalCollected}</td>
      </tr>
      <tr style="background:#fffcf5;">
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;color:#5F5E5A;vertical-align:top;font-size:12px;padding-left:20px;">Collected UIDs</td>
        <td style="padding:10px 12px;border-bottom:1px solid #eeeee9;font-size:11px;font-family:monospace;color:#2C2C2A;word-break:break-all;">${collectionUidsStr}</td>
      </tr>
    `;
  } else {
    transactionRowsHtml = `
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
    `;
  }

  const subject = `Cylinder Receipt: ${emailActionType} Summary — ${customerName}`;
  
  // Branded HTML email body
  const htmlBody = `
  <div style="font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;max-width:600px;margin:0 auto;color:#2C2C2A;background:#f7f8f6;padding:20px;border-radius:12px;">
    <!-- Card Container -->
    <div style="background:#ffffff;border:1.5px solid #d8d9d4;border-radius:14px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.05);">
      
      <!-- Banner Header -->
      <div style="background:#0F6E56;padding:24px;text-align:center;color:#ffffff;">
        <img src="cid:nobleLogo" style="max-height:50px; width:auto; display:block; margin:0 auto 12px auto;" alt="Noble Air Gases" />
        <div style="font-size:12px;text-transform:uppercase;letter-spacing:1px;font-weight:600;opacity:0.85;">Transaction Receipt</div>
        <h2 style="margin:6px 0 0;font-size:22px;font-weight:600;letter-spacing:-0.3px;">Cylinder ${emailActionType}</h2>
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
            ${transactionRowsHtml}
          </tbody>
        </table>

        <!-- Daily Summary & Outstanding Card -->
        <div style="background:#ffffff;border:1.5px solid #d8d9d4;border-radius:12px;padding:16px 14px;margin-bottom:28px;">
          <table style="width:100%;border-collapse:collapse;text-align:center;">
            <tr>
              <td style="width:33%;padding:8px;border-right:1px solid #eeeee9;vertical-align:top;">
                <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.6px;color:#5F5E5A;font-weight:600;margin-bottom:4px;">Delivered Today</div>
                <div style="font-size:24px;font-weight:700;color:#0F6E56;">${deliveredToday}</div>
              </td>
              <td style="width:33%;padding:8px;border-right:1px solid #eeeee9;vertical-align:top;">
                <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.6px;color:#5F5E5A;font-weight:600;margin-bottom:4px;">Collected Today</div>
                <div style="font-size:24px;font-weight:700;color:#c2410c;">${collectedToday}</div>
              </td>
              <td style="width:34%;padding:8px;vertical-align:top;">
                <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.6px;color:#0F6E56;font-weight:600;margin-bottom:4px;">Live Outstanding</div>
                <div style="font-size:24px;font-weight:700;color:#04342C;">${outstandingCount}</div>
              </td>
            </tr>
          </table>
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

  const logoBase64 = "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";
  const logoBlob = Utilities.newBlob(Utilities.base64Decode(logoBase64), "image/png", "nobleLogo");

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      name: EMAIL_SENDER_NAME,
      inlineImages: { nobleLogo: logoBlob }
    });
    
    const nowTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
    for (let g of groupedRows) {
      mapSheet.getRange(g.rowNum, 8).setValue(false); // Auto-uncheck on success
      mapSheet.getRange(g.rowNum, 9).setValue("Sent @ " + nowTimestamp);
    }
    ss.toast("✅ Receipt sent to " + email, "Success");
  } catch (err) {
    ss.toast("❌ Failed to send email: " + err.toString(), "Error");
    for (let g of groupedRows) {
      mapSheet.getRange(g.rowNum, 8).setValue(false);
      mapSheet.getRange(g.rowNum, 9).setValue("Error: " + err.toString().substring(0, 40));
    }
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

// ============================================================
//  CYLINDER MASTER REGISTRY
// ============================================================

const CYLINDER_SHEET_NAME = 'Cylinders';
const CYLINDER_MAINT_SHEET = 'Cylinder Maintenance';

// ── Create both registry sheets in one click ────────────────
function setupRegistrySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Cylinders sheet ────────────────────────────────────
  let cylSheet = ss.getSheetByName(CYLINDER_SHEET_NAME);
  if (!cylSheet) {
    cylSheet = ss.insertSheet(CYLINDER_SHEET_NAME);
  } else {
    cylSheet.clearContents();
    cylSheet.clearFormats();
  }

  const cylHeaders = [
    'Cylinder UID', 'Gas Type', 'Cylinder Type', 'Owner',
    'Current Status', 'Current Location', 'Last Activity Date'
  ];
  cylSheet.getRange(1, 1, 1, cylHeaders.length).setValues([cylHeaders])
    .setBackground('#0F6E56')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  cylSheet.setColumnWidth(1, 160);
  cylSheet.setColumnWidth(2, 130);
  cylSheet.setColumnWidth(3, 140);
  cylSheet.setColumnWidth(4, 130);
  cylSheet.setColumnWidth(5, 130);
  cylSheet.setColumnWidth(6, 180);
  cylSheet.setColumnWidth(7, 160);
  cylSheet.setFrozenRows(1);

  // Status dropdown validation for Column E
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Empty', 'Filled', 'Delivered'], true)
    .setAllowInvalid(false)
    .build();
  cylSheet.getRange(2, 5, 500, 1).setDataValidation(statusRule);

  // ── 2. Cylinder Maintenance sheet ─────────────────────────
  let maintSheet = ss.getSheetByName(CYLINDER_MAINT_SHEET);
  if (!maintSheet) {
    maintSheet = ss.insertSheet(CYLINDER_MAINT_SHEET);
  } else {
    maintSheet.clearContents();
    maintSheet.clearFormats();
  }

  const maintHeaders = [
    'Cylinder UID', 'Water Capacity (L)', 'Fill Pressure (bar)',
    'Gas Capacity', 'Unit', 'Is Mixture', 'Mix Ratio',
    'Manufacture Date', 'Last Hydro Test Date', 'Next Hydro Test Due',
    'Hydro Test Status', 'Test Certificate No.', 'Is UHP'
  ];
  maintSheet.getRange(1, 1, 1, maintHeaders.length).setValues([maintHeaders])
    .setBackground('#0F6E56')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  maintSheet.setColumnWidth(1, 160);
  maintSheet.setColumnWidth(2, 160);
  maintSheet.setColumnWidth(3, 160);
  maintSheet.setColumnWidth(4, 130);
  maintSheet.setColumnWidth(5, 100);
  maintSheet.setColumnWidth(6, 110);
  maintSheet.setColumnWidth(7, 200);
  maintSheet.setColumnWidth(8, 150);
  maintSheet.setColumnWidth(9, 175);
  maintSheet.setColumnWidth(10, 175);
  maintSheet.setColumnWidth(11, 160);
  maintSheet.setColumnWidth(12, 180);
  maintSheet.setColumnWidth(13, 110);
  maintSheet.setFrozenRows(1);

  // Is Mixture dropdown for Column F
  const mixtureRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();
  maintSheet.getRange(2, 6, 500, 1).setDataValidation(mixtureRule);

  // Is UHP dropdown for Column M
  const uhpRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true)
    .setAllowInvalid(false)
    .build();
  maintSheet.getRange(2, 13, 500, 1).setDataValidation(uhpRule);

  // Date format for date columns (H, I, J)
  maintSheet.getRange(2, 8, 500, 3).setNumberFormat('dd-mm-yyyy');

  SpreadsheetApp.getUi().alert(
    '✅ Registry Sheets Created!\n\n' +
    '• "Cylinders" sheet — 7 tracking fields\n' +
    '• "Cylinder Maintenance" sheet — 12 specification fields\n\n' +
    'You can now add cylinders from the Admin Portal at /admin/cylinders\n' +
    'or manually enter them directly in the sheets.'
  );
}

// ── Auto-update Cylinders sheet when a scan is submitted ─────
// Called from autoRefreshHandler after refreshBatches()
// Updates Status, Current Location, and Last Activity Date for each scanned UID.
function updateCylinderRegistryFromScans(scannedBatches) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cylSheet = ss.getSheetByName(CYLINDER_SHEET_NAME);
  if (!cylSheet) return; // Registry not set up yet — silently skip

  const lastRow = cylSheet.getLastRow();
  if (lastRow < 2) return;

  // Build a map of UID → row number from the Cylinders sheet
  const uidData = cylSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const uidRowMap = {};
  for (let i = 0; i < uidData.length; i++) {
    const uid = String(uidData[i][0]).trim().toUpperCase();
    if (uid) uidRowMap[uid] = i + 2; // 1-indexed row
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');

  for (const batch of scannedBatches) {
    const action = String(batch.action || '').trim();
    const customer = String(batch.customer || '').trim();

    for (const uid of (batch.cylinders || [])) {
      const uidUpper = String(uid).trim().toUpperCase();
      const rowNum = uidRowMap[uidUpper];
      if (!rowNum) continue; // UID not in registry — skip

      // Col E = Current Status, Col F = Current Location, Col G = Last Activity Date
      let newStatus = '';
      let newLocation = '';

      if (action === 'Delivery') {
        newStatus = 'Active';
        newLocation = customer || 'Customer';
      } else if (action === 'Collection') {
        newStatus = 'Active';
        newLocation = 'Depot';
      } else if (action === 'Filling') {
        newStatus = 'Active';
        newLocation = 'Depot';
      }

      if (newStatus) {
        cylSheet.getRange(rowNum, 5).setValue(newStatus);    // Status
        cylSheet.getRange(rowNum, 6).setValue(newLocation);  // Location
        cylSheet.getRange(rowNum, 7).setValue(today);        // Last Activity Date
      }
    }
  }
}

// ── Setup Bulk Tanks Sheet ──────────────────────────────────
const BULK_TANKS_SHEET_NAME = 'Bulk Tanks';

function setupBulkTanksSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tankSheet = ss.getSheetByName(BULK_TANKS_SHEET_NAME);
  if (!tankSheet) {
    tankSheet = ss.insertSheet(BULK_TANKS_SHEET_NAME);
  } else {
    if (tankSheet.getLastRow() < 1) {
      tankSheet.clearContents();
      tankSheet.clearFormats();
    }
  }

  if (tankSheet.getLastRow() < 1) {
    const headers = ['Date', 'Gas', 'Opening Stock', 'Dead Volume', 'Tank Capacity', 'Unit'];
    tankSheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#0F6E56')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center');

    tankSheet.setColumnWidth(1, 140);
    tankSheet.setColumnWidth(2, 130);
    tankSheet.setColumnWidth(3, 160);
    tankSheet.setColumnWidth(4, 160);
    tankSheet.setColumnWidth(5, 160);
    tankSheet.setColumnWidth(6, 100);
    tankSheet.setFrozenRows(1);

    // Format date column A
    tankSheet.getRange(2, 1, 1000, 1).setNumberFormat('dd-mm-yyyy');
    
    // Add some default rows for today so the manager has something to start with
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
    const defaultData = [
      [today, 'Argon', 5664.03, 500.0, 10000.0, 'Cum'],
      [today, 'CO2', 10941.05, 200.0, 15000.0, 'KG'],
      [today, 'N2', 4271.2, 300.0, 8000.0, 'Cum'],
      [today, 'Oxygen', 9215.3, 400.0, 12000.0, 'Cum']
    ];
    tankSheet.getRange(2, 1, defaultData.length, defaultData[0].length).setValues(defaultData);
  }

  SpreadsheetApp.getUi().alert(
    '✅ Bulk Tanks Sheet Created / Verified!\n\n' +
    '• "Bulk Tanks" sheet is ready.\n' +
    '• Date, Gas, Opening Stock, Dead Volume, Tank Capacity, and Unit columns are set up.\n' +
    '• Manager can update opening stocks via "/admin/inventory" on the Admin Portal.'
  );
}

// ── Rebuild Cylinder statuses and locations from raw Scan Log history ──
function rebuildCylinderRegistry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scanSheet = ss.getSheetByName(SCAN_SHEET_NAME);
  const cylSheet = ss.getSheetByName(CYLINDER_SHEET_NAME);
  if (!scanSheet || !cylSheet) {
    SpreadsheetApp.getUi().alert('❌ Sheets not found.');
    return;
  }
  
  const scanRows = scanSheet.getDataRange().getValues();
  const cylRows = cylSheet.getDataRange().getValues();
  if (cylRows.length < 2) {
    SpreadsheetApp.getUi().alert('ℹ️ No cylinders found in the registry sheet.');
    return;
  }
  
  // Build a map of cylinder UID -> row index (1-based) in Cylinders sheet
  const uidRowMap = {};
  for (let i = 1; i < cylRows.length; i++) {
    const uid = String(cylRows[i][0]).trim().toUpperCase();
    if (uid) uidRowMap[uid] = i + 1;
  }
  
  // Scan log columns: Date(0), Time(1), Driver(2), Action(3), UID(4), Customer(5)
  // Scan from top to bottom (oldest to newest) to get the latest status
  const latestCylState = {};
  for (let i = 1; i < scanRows.length; i++) {
    const row = scanRows[i];
    if (row.length < 5) continue;
    const date = formatDate(row[0]);
    const action = String(row[3]).trim();
    const uid = String(row[4]).trim().toUpperCase();
    const customer = row.length > 5 ? String(row[5]).trim() : '';
    
    if (!uid) continue;
    
    let status = '';
    let location = '';
    
    if (action === 'Delivery') {
      status = 'Delivered';
      location = customer || 'Customer';
    } else if (action === 'Collection') {
      status = 'Empty';
      location = 'Depot';
    } else if (action === 'Filling') {
      status = 'Filled';
      location = 'Depot';
    }
    
    if (status) {
      latestCylState[uid] = {
        status: status,
        location: location,
        date: date
      };
    }
  }
  
  // Write states back to Cylinders registry sheet
  for (const uid in uidRowMap) {
    const rowNum = uidRowMap[uid];
    const state = latestCylState[uid];
    
    if (state) {
      cylSheet.getRange(rowNum, 5).setValue(state.status);
      cylSheet.getRange(rowNum, 6).setValue(state.location);
      cylSheet.getRange(rowNum, 7).setValue(state.date);
    } else {
      // Revert to default empty depot status if never scanned
      cylSheet.getRange(rowNum, 5).setValue('Empty');
      cylSheet.getRange(rowNum, 6).setValue('Depot');
      cylSheet.getRange(rowNum, 7).setValue('—');
    }
  }
  
  SpreadsheetApp.getUi().alert('✅ Cylinder registry successfully rebuilt from scan log history!');
}


// ============================================================
//  PRODUCTS CONFIG SHEET SETUP
//  Creates / verifies the "Products" sheet that controls
//  which rows appear in Table 1 (Filled Cylinder Inventory)
//  on the Admin Portal's Inventory / Report page.
// ============================================================
function setupProductsSheet() {
  const PRODUCTS_SHEET = 'Products';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ps = ss.getSheetByName(PRODUCTS_SHEET);

  if (!ps) {
    ps = ss.insertSheet(PRODUCTS_SHEET);
  }

  // Only write if sheet is empty (don't overwrite manager edits)
  if (ps.getLastRow() > 0) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Products sheet already exists and has data.\n\n' +
      'No changes were made to preserve your existing configuration.\n\n' +
      'To reset, manually clear all rows in the "Products" sheet and run this again.'
    );
    return;
  }

  // ── Headers ──────────────────────────────────────────────────────────────
  const headers = [
    'Product ID',
    'Display Name',
    'Gas Type',
    'Cylinder Type',
    'Gas Per Cylinder',
    'Unit',
    'Is Virtual?'
  ];

  const headerRange = ps.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground('#0F6E56')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  // Column widths
  ps.setColumnWidth(1, 140);  // Product ID
  ps.setColumnWidth(2, 180);  // Display Name
  ps.setColumnWidth(3, 110);  // Gas Type
  ps.setColumnWidth(4, 130);  // Cylinder Type
  ps.setColumnWidth(5, 150);  // Gas Per Cylinder
  ps.setColumnWidth(6, 80);   // Unit
  ps.setColumnWidth(7, 110);  // Is Virtual?
  ps.setFrozenRows(1);

  // ── Default 11 product rows ──────────────────────────────────────────────
  // Matches exactly what is hardcoded in app.py DEFAULT_PRODUCTS_CONFIG
  const rows = [
    ['arg_pura',     'ARG Pura',       'ARG', 'Standard', 7.0,    'Cum', 'FALSE'],
    ['acm_90_10',   'ACM (90.10)_',   'ACM', 'Standard', 6.3512, 'Cum', 'FALSE'],
    ['co2_90_10',   'Co2 (90.10)_',   'ACM', 'Standard', 1.35,   'KG',  'TRUE'],
    ['co2_pure',    'Co2',            'CO2', 'Standard', 30.0,   'KG',  'FALSE'],
    ['n2_cyl',      'N2 Cyl',         'N2',  'Standard', 7.0,    'Cum', 'FALSE'],
    ['oxygen_pure', 'OXYGEN',         'OXY', 'Standard', 7.0,    'Cum', 'FALSE'],
    ['ahm_92_08',   'AHM(92.08)',     'AHM', 'Standard', 6.92,   'Cum', 'FALSE'],
    ['ahm_98_02',   'AHM (98.02)',    'AHM', 'Standard', 6.98,   'Cum', 'FALSE'],
    ['arg_dura',    'ARG Dura',       'ARG', 'Dura',     0.0,    'Cum', 'FALSE'],
    ['n2_dura',     'N2Dura',         'N2',  'Dura',     0.88,   'Cum', 'FALSE'],
    ['oxygen_dura', 'Oxygen Dura',    'OXY', 'Dura',     0.0,    'Cum', 'FALSE'],
  ];

  ps.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // ── Styling for data rows ─────────────────────────────────────────────────
  // Alternate row shading
  for (let i = 0; i < rows.length; i++) {
    const bg = i % 2 === 0 ? '#f8faf8' : '#ffffff';
    ps.getRange(i + 2, 1, 1, headers.length).setBackground(bg);
  }

  // Highlight the one virtual row (co2_90_10) in light orange so it stands out
  ps.getRange(4, 1, 1, headers.length).setBackground('#fff3cd'); // row 4 = co2_90_10

  // Is Virtual? column: dropdown validation TRUE / FALSE
  const dvRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true)
    .setAllowInvalid(false)
    .build();
  ps.getRange(2, 7, rows.length, 1).setDataValidation(dvRule);

  // Gas Type column: dropdown validation
  const gasRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ARG', 'CO2', 'N2', 'OXY', 'ACM', 'AHM'], true)
    .setAllowInvalid(true)
    .build();
  ps.getRange(2, 3, rows.length, 1).setDataValidation(gasRule);

  // Cylinder Type column: dropdown validation
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Standard', 'Dura'], true)
    .setAllowInvalid(true)
    .build();
  ps.getRange(2, 4, rows.length, 1).setDataValidation(typeRule);

  // Gas Per Cylinder: number format
  ps.getRange(2, 5, rows.length, 1).setNumberFormat('0.0000');

  // Add a note on the header explaining the columns
  ps.getRange(1, 1).setNote(
    'Product ID: unique key used by the app (do not change existing IDs)'
  );
  ps.getRange(1, 5).setNote(
    'Gas Per Cylinder: default volume/weight per cylinder used in the report.\n' +
    'If a cylinder has a Gas Capacity in the Maintenance sheet, that value is used instead.'
  );
  ps.getRange(1, 7).setNote(
    'Is Virtual? TRUE means this row shares the cylinder count of another row\n' +
    '(e.g. Co2 (90.10)_ borrows count from ACM cylinders).'
  );

  SpreadsheetApp.getUi().alert(
    '✅ Products Config Sheet Created!\n\n' +
    '• 11 default product rows have been pre-filled.\n' +
    '• You can now:\n' +
    '   – Rename any Display Name (Column B)\n' +
    '   – Adjust Gas Per Cylinder values (Column E)\n' +
    '   – Change Unit (Column F)\n' +
    '   – Add new rows for new products\n\n' +
    '• The Admin Portal Inventory page (Table 1) will automatically\n' +
    '  pick up your changes on the next page load.\n\n' +
    '⚠️ Do NOT change Product ID (Column A) for existing rows.'
  );
}
