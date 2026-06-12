// =============================================================
// Google Apps Script — paste this into script.google.com
// =============================================================
// After pasting:
// 1. Click Deploy > New deployment
// 2. Type: Web app
// 3. Execute as: Me
// 4. Who has access: Anyone
// 5. Copy the deployment URL and paste it into index.html
// =============================================================

const SHEET_NAME = 'Bookings';

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json',
  };

  try {
    const action = e.parameter.action;

    if (action === 'getBookings') {
      return respond(getBookings(), headers);
    }

    if (action === 'addBooking') {
      const data = JSON.parse(e.postData.contents);
      return respond(addBooking(data), headers);
    }

    if (action === 'deleteBooking') {
      const data = JSON.parse(e.postData.contents);
      return respond(deleteBooking(data), headers);
    }

    return respond({ error: 'Unknown action' }, headers);
  } catch (err) {
    return respond({ error: err.message }, headers);
  }
}

function respond(data, headers) {
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
    id: row[0],
    name: row[1],
    flat: row[2],
    date: row[3],
    startTime: row[4],
    endTime: row[5],
    notes: row[6],
    createdAt: row[7],
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
