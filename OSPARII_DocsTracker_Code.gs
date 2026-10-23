const SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID_HERE";
const SHEET_NAME = "Records";

/**
 * Expected columns:
 * A ID | B Document | C Tracking Number | D Status | E Date | F Data (full JSON)
 */
function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "Document", "Tracking Number", "Status", "Date", "Data"]);
  } else if (sheet.getRange(1, 6).getValue() !== "Data") {
    sheet.getRange(1, 6).setValue("Data");
  }
  return sheet;
}

function doGet(e) {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values.shift() || [];
  const records = [];

  values.forEach(row => {
    if (!row.some(value => value !== "" && value !== null)) return;

    const raw = row[5];
    if (raw) {
      try {
        records.push(JSON.parse(raw));
        return;
      } catch (err) {
        // Fall back to legacy columns if Data is malformed.
      }
    }

    records.push({
      ID: row[0],
      Document: row[1],
      "Tracking Number": row[2],
      Status: row[3],
      Date: row[4]
    });
  });

  const result = JSON.stringify({ success: true, records: records });
  const callback = e && e.parameter ? e.parameter.callback : "";

  if (callback && /^[a-zA-Z_$][\w$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + "(" + result + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(result)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();

    if (payload.action === "delete") {
      const tracking = String(payload.tracking || "");
      if (!tracking) throw new Error("Missing tracking number.");

      if (lastRow >= 2) {
        const trackingValues = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
        for (let i = trackingValues.length - 1; i >= 0; i--) {
          if (String(trackingValues[i][0]) === tracking) {
            sheet.deleteRow(i + 2);
            break;
          }
        }
      }

      return jsonResponse({ success: true, action: "delete", tracking: tracking });
    }

    if (payload.action === "save") {
      const record = payload.record;
      if (!record || !record.tracking) {
        throw new Error("Missing record or tracking number.");
      }

      const tracking = String(record.tracking);
      const rowData = [
        tracking,
        String(record.title || ""),
        tracking,
        String(record.status || "Pending"),
        String(record.dateReceived || ""),
        JSON.stringify(record)
      ];

      let targetRow = 0;
      if (lastRow >= 2) {
        const trackingValues = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
        for (let i = 0; i < trackingValues.length; i++) {
          if (String(trackingValues[i][0]) === tracking) {
            targetRow = i + 2;
            break;
          }
        }
      }

      if (targetRow) {
        sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }

      return jsonResponse({ success: true, action: "save", tracking: tracking });
    }

    throw new Error("Unknown action. Use save or delete.");
  } finally {
    lock.releaseLock();
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
