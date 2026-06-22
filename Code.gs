// ===== CONFIG =====
const SHEET_ID      = '1cXWyZyVm5fwhbIVGnGStmrVU1jxBwin1iIq2qPdI3UQ';
const FOLDER_REPAIR = '1UXG1qpZ9_StTT48NQ4MGVwn0FYkxR4EH'; // Drive Folder ID สำหรับรูปแจ้งซ่อม
const FOLDER_STOCK  = '18fDCaJQaPW6xsjPoJ02Kg84-71IjlxgE'; // Drive Folder ID สำหรับรูปสต็อก

// ===== CORS helper =====
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== doGet =====
function doGet(e) {
  const action = (e.parameter || {}).action;
  try {
    if (action === 'history')  return json(getHistory());
    return json({ success: true, message: 'TRR999 GAS OK' });
  } catch (err) {
    return json({ success: false, error: err.message });
  }
}

// ===== doPost =====
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    switch (p.action) {
      case 'syncToSheet':   return json(syncToSheet(p));
      default: return json({ success: false, error: 'Unknown action: ' + p.action });
    }
  } catch (err) {
    return json({ success: false, error: err.message });
  }
}

// ===== syncToSheet — write row to Sheets + upload image + return success =====
// Frontend deletes the Firebase doc after this returns { success: true }
function syncToSheet(p) {
  const d = p.data;
  const isRepair = p.collection === 'repairs';
  const sheetName = isRepair ? 'แจ้งซ่อม' : 'เบิกอุปกรณ์';
  const sheet = getOrCreateSheet(sheetName);

  // Create a default header only if the sheet is brand new
  if (sheet.getLastRow() === 0) {
    const headers = isRepair
      ? ['เลขที่','วันที่-เวลา','ชื่อผู้แจ้ง','ทะเบียนรถ','เลขไมล์','อาการ','รูปถ่าย','สถานะ']
      : ['เลขที่','วันที่-เวลา','ชื่อผู้เบิก','ทะเบียนรถ','อ้างอิงใบซ่อม','รายการ','รูปถ่าย','หมายเหตุ','สถานะ'];
    sheet.appendRow(headers);
    formatHeader(sheet);
  }

  // Upload images to Drive — repair and requisition photos go to separate folders
  let photoUrl = '';
  const folderId = isRepair ? FOLDER_REPAIR : FOLDER_STOCK;
  if (p.photos && p.photos.length > 0) {
    const urls = [];
    p.photos.forEach(function(ph) {
      if (ph.base64 && ph.name) {
        const url = uploadToDrive(ph.base64, ph.name, folderId);
        if (url) urls.push(url);
      }
    });
    photoUrl = urls.join(', ');
  } else if (p.photoBase64 && p.photoName) {
    photoUrl = uploadToDrive(p.photoBase64, p.photoName, folderId);
  }

  // Write each value into the column matching its header — keeps the sheet's own
  // columns (เลขที่, อ้างอิงใบซ่อม, หมายเหตุ ...) intact instead of shifting everything.
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h); });
  const rowArr = new Array(header.length).fill('');
  const put = function(kw, val) { var i = colByHeader(header, kw); if (i >= 0) rowArr[i] = val; };

  put('วันที่', d.datetime);
  put('ทะเบียน', d.vehicle);
  if (isRepair) {
    put('ชื่อ', d.reporter);
    put('ไมล์', d.mileage);
    put('อาการ', d.problem);
    put('รูป', photoUrl);
    put('สถานะ', d.status || 'รอดำเนินการ');
  } else {
    put('ชื่อ', d.requester);
    put('รายการ', d.description || '');
    put('รูป', photoUrl);
    put('สถานะ', d.status || 'รออนุมัติ');
    // requisition sheet has no รูป column → keep the photo URL in หมายเหตุ instead
    if (photoUrl && colByHeader(header, 'รูป') < 0) put('หมายเหตุ', photoUrl);
  }

  sheet.appendRow(rowArr);
  return { success: true };
}

// ===== getHistory — read both sheets, mapping columns by HEADER NAME =====
// Robust to extra/reordered columns (เลขที่, อ้างอิงใบซ่อม, หมายเหตุ ...) that the app doesn't know about.
function getHistory() {
  const items = [];
  collectHistory('แจ้งซ่อม', 'repair', items);
  collectHistory('เบิกอุปกรณ์', 'requisition', items);
  return { success: true, items: items };
}

// Find a column index by a keyword in the header row (-1 if absent)
function colByHeader(header, kw) {
  for (var i = 0; i < header.length; i++) { if (header[i].indexOf(kw) >= 0) return i; }
  return -1;
}

function collectHistory(sheetName, type, items) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;

  const data = sheet.getDataRange().getValues();
  const header = data[0].map(function(h) { return String(h); });
  const ci = {
    datetime: colByHeader(header, 'วันที่'),
    name:     colByHeader(header, 'ชื่อ'),
    vehicle:  colByHeader(header, 'ทะเบียน'),
    mileage:  colByHeader(header, 'ไมล์'),
    detail:   type === 'repair' ? colByHeader(header, 'อาการ') : colByHeader(header, 'รายการ'),
    photo:    colByHeader(header, 'รูป'),
    status:   colByHeader(header, 'สถานะ')
  };
  const get = function(row, i) { return i >= 0 ? String(row[i] == null ? '' : row[i]) : ''; };

  for (var r = 1; r < data.length; r++) {
    const row = data[r];
    const dt = ci.datetime >= 0 ? row[ci.datetime] : '';
    const name = get(row, ci.name);
    if (!dt && !name) continue;  // skip blank rows

    var detail = get(row, ci.detail);
    if (type === 'repair' && ci.mileage >= 0 && row[ci.mileage]) {
      detail += (detail ? ' · ' : '') + 'ไมล์ ' + row[ci.mileage];
    }

    items.push({
      type: type,
      ts: rowTs(dt),
      datetime: dtDisplay(dt),
      name: name,
      vehicle: get(row, ci.vehicle),
      detail: detail,
      photoUrl: get(row, ci.photo),
      status: get(row, ci.status)
    });
  }
}

// Sortable epoch from a Date cell or a dd/MM/yyyy string (handles both Buddhist and Gregorian years)
function rowTs(v) {
  if (v instanceof Date) return v.getTime();
  var m = String(v).match(/(\d+)\/(\d+)\/(\d+)[,\s]+(\d+):(\d+)(?::(\d+))?/);
  if (m) {
    var y = +m[3];
    if (y >= 2500) y -= 543;  // Buddhist year → Gregorian
    return new Date(y, +m[2] - 1, +m[1], +m[4], +m[5], +(m[6] || 0)).getTime();
  }
  return 0;
}

// Short Thai (Buddhist) display string, normalized from a Date cell or a date string
function dtDisplay(v) {
  var ts = rowTs(v);
  if (ts) {
    var d = new Date(ts);
    return Utilities.formatDate(d, 'Asia/Bangkok', 'dd/MM/') + (d.getFullYear() + 543) +
           Utilities.formatDate(d, 'Asia/Bangkok', ' HH:mm');
  }
  return String(v);
}

// ===== uploadToDrive =====
function uploadToDrive(base64, filename, folderId) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), MimeType.JPEG, filename);
    const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch (err) {
    console.error('uploadToDrive error:', err);
    return '';
  }
}

// ===== Helpers =====
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function formatHeader(sheet) {
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setBackground('#1a3a6b')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
}
