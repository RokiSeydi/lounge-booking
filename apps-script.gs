// =============================================================
// Google Apps Script — paste this into script.google.com
// =============================================================
// After pasting:
// 1. Click Deploy > New deployment
// 2. Type: Web app
// 3. Execute as: Me
// 4. Who has access: Anyone
// 5. Copy the deployment URL and paste it into index.html
//
// IMPORTANT: Every time you edit this code, you must create a
// NEW deployment (Deploy > New deployment), not just save.
// Then update the URL in index.html if it changed.
// =============================================================

const SHEET_NAME = 'Bookings';

function doGet(e) {
  try {
    const action = e.parameter.action;
    const payload = e.parameter.data ? JSON.parse(e.parameter.data) : null;

    if (action === 'getBookings') {
      return respond(getBookings());
    }

    if (action === 'addBooking' && payload) {
      return respond(addBooking(payload));
    }

    if (action === 'deleteBooking' && payload) {
      return respond(deleteBooking(payload));
    }

    return respond({ error: 'Unknown action' });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Data helpers ----

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'name', 'flat', 'date', 'startTime', 'endTime', 'notes', 'createdAt']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  return sheet;
}

function getBookings() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { bookings: [] };

  const bookings = rows.slice(1).map(row => ({
    id: String(row[0]),
    name: String(row[1]),
    flat: String(row[2]),
    date: row[3] instanceof Date
      ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[3]),
    startTime: row[4] instanceof Date
      ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'HH:mm')
      : String(row[4]),
    endTime: row[5] instanceof Date
      ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'HH:mm')
      : String(row[5]),
    notes: String(row[6]),
    createdAt: String(row[7]),
  }));

  return { bookings };
}

function addBooking(data) {
  const sheet = getSheet();
  const id = Utilities.getUuid();
  const now = new Date().toISOString();

  // Check for overlapping bookings
  const existing = getBookings().bookings;
  const overlap = existing.find(b =>
    b.date === data.date &&
    data.startTime < b.endTime &&
    data.endTime > b.startTime
  );

  if (overlap) {
    return {
      error: `Time slot overlaps with ${overlap.name}'s booking (${overlap.startTime}–${overlap.endTime})`,
    };
  }

  sheet.appendRow([
    id,
    data.name,
    data.flat || '',
    data.date,
    data.startTime,
    data.endTime,
    data.notes || '',
    now,
  ]);

  return { success: true, id };
}

function deleteBooking(data) {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { error: 'Booking not found' };
}
