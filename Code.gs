/*******************************************************
 * EVENTPASS LITE — Code.gs
 * Main backend: routing, auth, guest ops, check-in, logs
 *
 * MIGRATED to a JSON API (GitHub Pages migration).
 * This backend is now purely an API server — doGet/doPost
 * accept an "action" and return JSON. All the original
 * business-logic functions below are UNCHANGED from the
 * pre-migration version; only the entry points at the top
 * were rewritten.
 *
 * CORS NOTE: Apps Script does not support preflighted
 * (OPTIONS) requests reliably. To avoid triggering a
 * preflight, the frontend (js/api.js) sends POST requests
 * with Content-Type: text/plain;charset=utf-8 and a JSON
 * string body, which the browser treats as a "simple
 * request". Do NOT change the frontend to send
 * Content-Type: application/json — that will break CORS.
 *******************************************************/

const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_USERS = 'USERS';
const SHEET_GUESTS = 'GUESTS';
const SHEET_LOGS = 'LOGS';
const SHEET_SETTINGS = 'SETTINGS';

const EVENT_ID = 'EVT2026-001'; // Change per event/client. See README.

/*******************************************************
 * WEB APP ENTRY POINTS (JSON API)
 *******************************************************/

function doGet(e) {
  const params = (e && e.parameter) || {};

  // Legacy JSONP path used by the standalone scanner page.
  // Kept as-is for backward compatibility — safe to remove
  // once every client calls the unified action API below.
  if (params.page === 'scanner-api') {
    return scannerApi_(e);
  }

  if (params.action) {
    return apiResponse_(handleAction_(params.action, params));
  }

  // No action + no legacy page param: nothing to serve.
  // (HTML used to be rendered from here pre-migration; the
  // frontend now lives entirely on GitHub Pages.)
  return apiResponse_({
    success: false,
    error: 'Missing "action" parameter. This endpoint is a JSON API only.'
  });
}

function doPost(e) {
  let body = {};
  try {
    const raw = (e && e.postData && e.postData.contents) || '{}';
    body = JSON.parse(raw);
  } catch (err) {
    return apiResponse_({ success: false, error: 'Invalid JSON body.' });
  }

  return apiResponse_(handleAction_(body.action, body));
}

// Optional: some browsers/environments send an OPTIONS preflight even for
// "simple" requests in edge cases. Apps Script will route it here if so.
function doOptions(e) {
  return ContentService.createTextOutput('');
}

/**
 * Central action router. Maps a public "action" name to one of the
 * business-logic functions further down, coercing param types where the
 * original functions expect (e.g. numbers, booleans) since everything
 * arrives over the wire as strings (GET) or JSON (POST).
 */
function handleAction_(action, p) {
  p = p || {};

  try {
    switch (action) {
      case 'login':
        return login(p.email, p.password);

      case 'getGuests':
        return { success: true, data: getAllGuests() };

      case 'getDashboardStats':
        return { success: true, data: getDashboardStats() };

      case 'getStaff':
        return { success: true, data: getAllStaff() };

      case 'getLogs':
        return { success: true, data: getLogs() };

      case 'getSettings':
        return { success: true, data: getSettings() };

      case 'updateSettings':
        return updateSettings(p.settings || {});

      case 'exportAttendanceCsv':
        return { success: true, csv: exportAttendanceCSV() };

      case 'generateTicket':
        return generateTicket(p.name, p.phone);

      case 'generateBlankTickets':
        return generateBlankTickets(parseInt(p.count, 10));

      case 'editGuest':
        return editGuest(p.ticketId, p.name, p.phone);

      case 'deleteGuest':
        return deleteGuest(p.ticketId);

      case 'resetTicket':
        return resetTicket(p.ticketId, p.adminName);

      case 'importGuests':
        return importGuests(p.guests || []);

      case 'generateQrCache':
        return generateQrCacheForGuests();

      case 'addStaff':
        return addStaff(p.name, p.email, p.password);

      case 'deleteStaff':
        return deleteStaff(p.id);

      case 'toggleStaffActive':
        return toggleStaffActive(p.id, p.active === true || p.active === 'true');

      case 'resetStaffPassword':
        return resetStaffPassword(p.id, p.password);

      case 'searchGuest':
        return { success: true, data: searchGuest(p.query) };

      case 'checkIn':
        return checkInTicket(p.ticketId, p.staffName, p.gateLabel);

      case 'manualCheckIn':
        return manualCheckIn(p.ticketId, p.staffName);

      default:
        return { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    return {
      success: false,
      error: err && err.message ? err.message : 'Backend error.'
    };
  }
}

function apiResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Legacy JSONP endpoint for scanner.html (?page=scanner-api&action=checkIn&...).
// Left in place because scanner.html already works cross-origin via this path;
// no need to force a rewrite. New action-based requests use handleAction_ above.
function scannerApi_(e) {
  const params = (e && e.parameter) || {};
  const callback = String(params.callback || 'callback');
  const action = String(params.action || '');

  let result;

  try {
    if (action === 'checkIn') {
      result = checkInTicket(
        params.ticketId || '',
        params.staffName || 'Scanner',
        params.gateLabel || 'Scanner'
      );
    } else {
      result = { success: false, status: 'ERROR', error: 'Unknown action.' };
    }
  } catch (err) {
    result = {
      success: false,
      status: 'ERROR',
      error: err && err.message ? err.message : 'Backend error.'
    };
  }

  const body = `${callback}(${JSON.stringify(result)});`;
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/*******************************************************
 * UTILITIES
 *******************************************************/

function getSheet_(name) {
  const sh = SS.getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}

// Confirms row 1 of a sheet still has the expected column headers, in order.
// Manual editing (e.g. deleting rows near the top) can shift or damage row 1
// without an obvious error — this turns that into a clear message instead
// of a silent blank page.
function verifyHeaders_(sheet, expectedHeaders) {
  const actual = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0].map(v => String(v).trim());
  const mismatch = expectedHeaders.some((h, i) => actual[i] !== h);
  if (mismatch) {
    throw new Error(
      'The "' + sheet.getName() + '" tab\'s header row (row 1) doesn\'t match what the app expects.\n' +
      'It should read exactly: ' + expectedHeaders.join(' | ') + '\n' +
      'Currently reads: ' + actual.join(' | ')
    );
  }
}

function sheetToObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows
    .filter(r => r.join('') !== '') // skip fully blank rows
    .map((r, idx) => {
      const obj = { _row: idx + 2 }; // actual sheet row number
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
}

function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

// Simple, dependency-free hash (SHA-256, hex). Not bank-grade, but no
// plaintext passwords sitting in the Sheet.
function hashPassword_(plain) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, plain, Utilities.Charset.UTF_8);
  return raw.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function logAction_(staffName, action, ticketId) {
  const sh = getSheet_(SHEET_LOGS);
  sh.appendRow([nowStr_(), staffName, action, ticketId || '', EVENT_ID]);
}

/*******************************************************
 * AUTH
 *******************************************************/

/**
 * Logs a user in. Returns { success, role, name, id } or { success:false, error }
 */
function login(email, password) {
  email = String(email || '').trim().toLowerCase();
  password = String(password || '');

  if (!email || !password) {
    return { success: false, error: 'Email and password are required.' };
  }

  const users = sheetToObjects_(getSheet_(SHEET_USERS));
  const hashed = hashPassword_(password);

  const user = users.find(u => String(u.Email).trim().toLowerCase() === email);

  if (!user) return { success: false, error: 'Account not found.' };
  if (String(user.Active).toUpperCase() !== 'TRUE' && user.Active !== true) {
    return { success: false, error: 'This account has been disabled.' };
  }
  if (String(user.Password) !== hashed) {
    return { success: false, error: 'Incorrect password.' };
  }

  logAction_(user.Name, 'LOGIN', '');

  return {
    success: true,
    role: user.Role,
    name: user.Name,
    id: user.ID
  };
}

/*******************************************************
 * GUEST MANAGEMENT (Admin)
 *******************************************************/

function makeQrPayload_(ticketId, eventName) {
  const num = parseInt(String(ticketId).replace(/\D/g, ''), 10);

  let ticketType = 'General Ticket';
  if (num >= 1 && num <= 70) {
    ticketType = 'Male Ticket';
  } else if (num >= 71 && num <= 130) {
    ticketType = 'Female Ticket';
  }

  return `${eventName || 'Pool Party Event Pass'} | ${ticketType} | ${ticketId}`;
}

function generateQrCacheForGuests() {
  const sh = getSheet_(SHEET_GUESTS);
  const guests = sheetToObjects_(sh);
  if (!guests.length) return { success: true, count: 0 };

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  let qrCol = headers.indexOf('QR Payload') + 1;
  if (!qrCol) {
    qrCol = headers.length + 1;
    sh.getRange(1, qrCol).setValue('QR Payload');
  }

  const eventName = (getSettings().EventName || 'Pool Party Event Pass');
  let count = 0;
  guests.forEach(g => {
    const payload = makeQrPayload_(g['Ticket ID'], eventName);
    if (String(g['QR Payload'] || '') !== payload) {
      sh.getRange(g._row, qrCol).setValue(payload);
      count++;
    }
  });

  return { success: true, count };
}

function getAllGuests() {
  const sh = getSheet_(SHEET_GUESTS);
  const data = sheetToObjects_(sh);
  return JSON.parse(JSON.stringify(data));
}

// Generates a single blank ticket (no name required — sold physically later)
function generateTicket(name, phone) {
  const sh = getSheet_(SHEET_GUESTS);
  const settings = getSettings();
  const guests = sheetToObjects_(sh);

  const prefix = settings.TicketPrefix || 'EVT';
  const nextNum = guests.length + 1;
  const ticketId = prefix + String(nextNum).padStart(3, '0');

  sh.appendRow([ticketId, name || '', phone || '', 'Unused', makeQrPayload_(ticketId, settings.EventName), '', '', EVENT_ID]);
  return { success: true, ticketId: ticketId };
}

// Bulk-generate N blank tickets at once (e.g. 130 for a print run)
function generateBlankTickets(count) {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000);
  if (!gotLock) {
    return { success: false, error: 'System busy, please try again in a moment.' };
  }

  try {
    const sh = getSheet_(SHEET_GUESTS);
    const settings = getSettings();
    const prefix = settings.TicketPrefix || 'EVT';
    const existing = sheetToObjects_(sh);

    let counter = existing.length + 1;
    const rows = [];
    const ticketIds = [];

    for (let i = 0; i < count; i++) {
      const ticketId = prefix + String(counter).padStart(3, '0');
      rows.push([ticketId, '', '', 'Unused', makeQrPayload_(ticketId, settings.EventName), '', '', EVENT_ID]);
      ticketIds.push(ticketId);
      counter++;
    }

    if (rows.length > 0) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return { success: true, count: rows.length, ticketIds: ticketIds };
  } finally {
    lock.releaseLock();
  }
}

// Attach or update a name/phone on an existing ticket (used at point of sale
// or at check-in, since tickets are often sold physically without pre-registration)
function assignTicketHolder(ticketId, name, phone) {
  const sh = getSheet_(SHEET_GUESTS);
  const guests = sheetToObjects_(sh);
  const g = guests.find(x => String(x['Ticket ID']).trim() === String(ticketId).trim());
  if (!g) return { success: false, error: 'Ticket not found.' };

  sh.getRange(g._row, 2).setValue(name || '');
  sh.getRange(g._row, 3).setValue(phone || '');
  return { success: true };
}

function editGuest(ticketId, name, phone) {
  const sh = getSheet_(SHEET_GUESTS);
  const guests = sheetToObjects_(sh);
  const g = guests.find(x => x['Ticket ID'] === ticketId);
  if (!g) return { success: false, error: 'Ticket not found.' };

  sh.getRange(g._row, 2).setValue(name);  // Name column
  sh.getRange(g._row, 3).setValue(phone); // Phone column
  return { success: true };
}

function deleteGuest(ticketId) {
  const sh = getSheet_(SHEET_GUESTS);
  const guests = sheetToObjects_(sh);
  const g = guests.find(x => x['Ticket ID'] === ticketId);
  if (!g) return { success: false, error: 'Ticket not found.' };

  sh.deleteRow(g._row);
  return { success: true };
}

/**
 * Bulk import guests. Expects array of {name, phone}
 */
function importGuests(guestArray) {
  const sh = getSheet_(SHEET_GUESTS);
  const settings = getSettings();
  const prefix = settings.TicketPrefix || 'EVT';

  const existing = sheetToObjects_(sh);
  let counter = existing.length + 1;
  const rows = [];

  guestArray.forEach(g => {
    if (!g.name) return;
    const ticketId = prefix + String(counter).padStart(3, '0');
    rows.push([ticketId, g.name, g.phone || '', 'Unused', makeQrPayload_(ticketId, settings.EventName), '', '', EVENT_ID]);
    counter++;
  });

  if (rows.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return { success: true, count: rows.length };
}

/*******************************************************
 * CHECK-IN (Staff + Admin) — LOCKED to prevent double
 * check-ins from simultaneous scans at multiple gates
 *******************************************************/

function checkInTicket(ticketId, staffName, gateLabel) {
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(10000); // wait up to 10s

  if (!gotLock) {
    return { success: false, status: 'ERROR', error: 'System busy, try again in a second.' };
  }

  try {
    const sh = getSheet_(SHEET_GUESTS);
    const guests = sheetToObjects_(sh);
    const g = guests.find(x => String(x['Ticket ID']).trim() === String(ticketId).trim());

    if (!g) {
      logAction_(staffName, 'INVALID SCAN', ticketId);
      return { success: false, status: 'INVALID', error: 'Ticket not found.' };
    }

    if (String(g.Status).toLowerCase() === 'used') {
      return {
        success: false,
        status: 'ALREADY_USED',
        ticketId: g['Ticket ID'],
        checkedInTime: g['Check-in Time'],
        checkedInBy: g['Checked In By']
      };
    }

    // Mark as used
    const time = nowStr_();
    sh.getRange(g._row, 4).setValue('Used');                         // Status
    sh.getRange(g._row, 5).setValue((gateLabel || staffName));       // Checked In By
    sh.getRange(g._row, 6).setNumberFormat('@').setValue(time);   // Check-in Time (forced text — prevents Sheets auto-date conversion)

    logAction_(staffName, 'CHECKED IN' + (gateLabel ? ' (' + gateLabel + ')' : ''), ticketId);

    return {
      success: true,
      status: 'VALID',
      ticketId: g['Ticket ID'],
      checkedInTime: time
    };

  } finally {
    lock.releaseLock();
  }
}

function searchGuest(query) {
  const guests = sheetToObjects_(getSheet_(SHEET_GUESTS));
  query = String(query || '').trim().toLowerCase();
  if (!query) return [];

  return guests.filter(g =>
    String(g['Ticket ID']).toLowerCase().includes(query) ||
    String(g.Name).toLowerCase().includes(query)
  ).slice(0, 20);
}

function manualCheckIn(ticketId, staffName) {
  return checkInTicket(ticketId, staffName, 'Manual');
}

/*******************************************************
 * ADMIN: RESET TICKET
 *******************************************************/

function resetTicket(ticketId, adminName) {
  const sh = getSheet_(SHEET_GUESTS);
  const guests = sheetToObjects_(sh);
  const g = guests.find(x => x['Ticket ID'] === ticketId);
  if (!g) return { success: false, error: 'Ticket not found.' };

  sh.getRange(g._row, 4).setValue('Unused');
  sh.getRange(g._row, 5).setValue('');
  sh.getRange(g._row, 6).setValue('');

  logAction_(adminName, 'RESET TICKET', ticketId);
  return { success: true };
}

/*******************************************************
 * STAFF MANAGEMENT (Admin)
 *******************************************************/

function getAllStaff() {
  const users = sheetToObjects_(getSheet_(SHEET_USERS));
  return users.filter(u => u.Role === 'Staff');
}

function addStaff(name, email, password) {
  const sh = getSheet_(SHEET_USERS);
  const users = sheetToObjects_(sh);
  const id = 'U' + String(users.length + 1).padStart(3, '0');

  if (users.some(u => String(u.Email).toLowerCase() === email.toLowerCase())) {
    return { success: false, error: 'Email already in use.' };
  }

  sh.appendRow([id, name, email, hashPassword_(password), 'Staff', true]);
  return { success: true };
}

// Permanently removes a staff login. Cannot be undone — the UI should
// confirm with the admin before calling this.
function deleteStaff(userId) {
  const sh = getSheet_(SHEET_USERS);
  const users = sheetToObjects_(sh);
  const u = users.find(x => x.ID === userId);
  if (!u) return { success: false, error: 'User not found.' };
  if (u.Role !== 'Staff') return { success: false, error: 'Only staff accounts can be removed here.' };

  sh.deleteRow(u._row);
  return { success: true };
}

function toggleStaffActive(userId, active) {
  const sh = getSheet_(SHEET_USERS);
  const users = sheetToObjects_(sh);
  const u = users.find(x => x.ID === userId);
  if (!u) return { success: false, error: 'User not found.' };

  sh.getRange(u._row, 6).setValue(active);
  return { success: true };
}

function resetStaffPassword(userId, newPassword) {
  const sh = getSheet_(SHEET_USERS);
  const users = sheetToObjects_(sh);
  const u = users.find(x => x.ID === userId);
  if (!u) return { success: false, error: 'User not found.' };

  sh.getRange(u._row, 4).setValue(hashPassword_(newPassword));
  return { success: true };
}

/*******************************************************
 * SETTINGS
 *******************************************************/

function getSettings() {
  const sh = getSheet_(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  const settings = {};
  data.forEach(row => {
    if (row[0]) settings[row[0]] = row[1];
  });
  return settings;
}

function updateSettings(settingsObj) {
  const sh = getSheet_(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();

  Object.keys(settingsObj).forEach(key => {
    let found = false;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        sh.getRange(i + 1, 2).setValue(settingsObj[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sh.appendRow([key, settingsObj[key]]);
    }
  });

  return { success: true };
}

/*******************************************************
 * DASHBOARD DATA
 *******************************************************/

function getDashboardStats() {
  const guests = sheetToObjects_(getSheet_(SHEET_GUESTS));
  const logs = sheetToObjects_(getSheet_(SHEET_LOGS));

  const total = guests.length;
  const checkedIn = guests.filter(g => String(g.Status).toLowerCase() === 'used').length;
  const remaining = total - checkedIn;
  const invalidAttempts = logs.filter(l => l.Action === 'INVALID SCAN').length;

  // Normalizes a Check-in Time cell to a display string, whether it's stored
  // as plain text or was auto-converted to a Date object by Sheets.
  function normalizeTime_(raw) {
    if (!raw) return '';
    if (Object.prototype.toString.call(raw) === '[object Date]') {
      return Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }
    return String(raw);
  }

  // Attendance by hour
  const hourly = {};
  guests.forEach(g => {
    const t = normalizeTime_(g['Check-in Time']);
    if (t) {
      const hourMatch = t.match(/(\d{2}):\d{2}:\d{2}/);
      if (hourMatch) {
        const hour = hourMatch[1] + ':00';
        hourly[hour] = (hourly[hour] || 0) + 1;
      }
    }
  });

  // Recent activity (last 10 check-ins, most recent first)
  const recent = guests
    .map(g => ({
      name: g.Name || g['Ticket ID'],
      time: normalizeTime_(g['Check-in Time']),
      by: g['Checked In By']
    }))
    .filter(g => g.time)
    .sort((a, b) => new Date(b.time) - new Date(a.time))
    .slice(0, 10);

  return {
    total, checkedIn, remaining, invalidAttempts,
    hourly, recent
  };
}

function exportAttendanceCSV() {
  const guests = sheetToObjects_(getSheet_(SHEET_GUESTS));
  let csv = 'Ticket ID,Name,Phone,Status,Checked In By,Check-in Time\n';
  guests.forEach(g => {
    csv += [
      g['Ticket ID'], g.Name, g.Phone, g.Status,
      g['Checked In By'] || '', g['Check-in Time'] || ''
    ].map(v => '"' + String(v).replace(/"/g, '""') + '"').join(',') + '\n';
  });
  return csv;
}

function getLogs() {
  const sh = getSheet_(SHEET_LOGS);
  const data = sheetToObjects_(sh).reverse();
  return JSON.parse(JSON.stringify(data));
}
