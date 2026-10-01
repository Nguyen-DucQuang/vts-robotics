const SHEET_NAME = 'Trang tính1';
const HEADER_LABEL = 'STT';

function doGet() {
  return json_({
    ok: true,
    message: 'VTS TECH workshop lead endpoint is running.'
  });
}

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    const required = ['parent_name', 'phone', 'workshop_date', 'child_age', 'consent'];
    const missing = required.filter(function(key) {
      return payload[key] === undefined || payload[key] === null || String(payload[key]).trim() === '';
    });

    if (missing.length) {
      return json_({ ok: false, error: 'Missing fields: ' + missing.join(', ') });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const headerRow = findHeaderRow_(sheet);
    const row = findNextDataRow_(sheet, headerRow + 1);
    const stt = row - headerRow;

    sheet.getRange(row, 1, 1, 7).setValues([[
      stt,
      clean_(payload.parent_name),
      clean_(payload.phone),
      clean_(payload.workshop_date),
      clean_(payload.child_age),
      clean_(payload.interest),
      payload.consent === true || payload.consent === 'true' || payload.consent === 'on' ? 'Đồng ý' : 'Chưa đồng ý'
    ]]);

    return json_({ ok: true, row: row });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function parsePayload_(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      return e.parameter || {};
    }
  }
  return e && e.parameter ? e.parameter : {};
}

function findHeaderRow_(sheet) {
  const values = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === HEADER_LABEL.toLowerCase()) {
      return i + 1;
    }
  }
  return 4;
}

function findNextDataRow_(sheet, firstDataRow) {
  const maxRows = sheet.getMaxRows();
  if (firstDataRow > maxRows) {
    sheet.insertRowsAfter(maxRows, firstDataRow - maxRows);
  }

  const height = Math.max(sheet.getMaxRows() - firstDataRow + 1, 1);
  const values = sheet.getRange(firstDataRow, 2, height, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!String(values[i][0]).trim()) {
      return firstDataRow + i;
    }
  }

  sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  return sheet.getMaxRows();
}

function clean_(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
