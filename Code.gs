// ===== CONFIG =====
const SHEET_ID      = '1cXWyZyVm5fwhbIVGnGStmrVU1jxBwin1iIq2qPdI3UQ';
const FOLDER_REPAIR = '1UXG1qpZ9_StTT48NQ4MGVwn0FYkxR4EH'; // Drive Folder ID สำหรับรูปแจ้งซ่อม
const FOLDER_STOCK  = '18fDCaJQaPW6xsjPoJ02Kg84-71IjlxgE'; // Drive Folder ID สำหรับรูปสต็อก

const SHEET_STOCK = 'สต็อก';

// Vision API Key — ตั้งใน Script Properties:
// Project Settings → Script Properties → เพิ่ม VISION_API_KEY = <key>

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
    if (action === 'getStock') return json(getStock());
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
      case 'uploadImage':   return json(uploadImageHandler(p));
      case 'analyzeImage':  return json(analyzeImage(p));
      case 'addStock':      return json(addStock(p));
      case 'updateStock':   return json(updateStock(p));
      case 'deleteStock':   return json(deleteStock(p));
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

  // Create header row on first use
  if (sheet.getLastRow() === 0) {
    const headers = isRepair
      ? ['วันที่/เวลา','ผู้แจ้ง','ทะเบียนรถ','เลขไมล์ (กม.)','อาการที่พบ','URL รูปภาพ','สถานะ']
      : ['วันที่/เวลา','ผู้เบิก','ทะเบียนรถ','ชื่ออุปกรณ์','จำนวน','หน่วย','แผนก','วัตถุประสงค์','URL รูปภาพ','สถานะ'];
    sheet.appendRow(headers);
    formatHeader(sheet);
  }

  // Upload image to Drive if provided
  let photoUrl = '';
  if (p.photoBase64 && p.photoName) {
    const folderId = isRepair ? FOLDER_REPAIR : FOLDER_STOCK;
    photoUrl = uploadToDrive(p.photoBase64, p.photoName, folderId);
  }

  // Append data row
  if (isRepair) {
    sheet.appendRow([
      d.datetime, d.reporter, d.vehicle, d.mileage,
      d.problem, photoUrl, d.status || 'รอดำเนินการ'
    ]);
  } else {
    sheet.appendRow([
      d.datetime, d.requester, d.vehicle, d.equipName,
      d.qty, d.unit || '', d.department || '', d.purpose || '',
      photoUrl, d.status || 'รออนุมัติ'
    ]);
  }

  return { success: true };
}

// ===== uploadImageHandler =====
// รับ base64 → upload Drive → return URL
// (ถูกเรียกจาก background fetch หลัง submit สำเร็จแล้ว)
function uploadImageHandler(p) {
  if (!p.base64 || !p.filename) return { success: false, error: 'Missing base64 or filename' };
  const folderId = p.collection === 'requisitions' ? FOLDER_REPAIR : FOLDER_REPAIR;
  const url = uploadToDrive(p.base64, p.filename, folderId);
  return { success: true, url };
}

// ===== analyzeImage — Google Vision API =====
function analyzeImage(p) {
  if (!p.imageBase64) return { success: false, error: 'Missing imageBase64' };

  const apiKey = PropertiesService.getScriptProperties().getProperty('VISION_API_KEY');
  if (!apiKey) return { success: false, error: 'VISION_API_KEY not set in Script Properties' };

  const endpoint = 'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey;

  const body = {
    requests: [{
      image: { content: p.imageBase64 },
      features: [
        { type: 'OBJECT_LOCALIZATION', maxResults: 10 },
        { type: 'LABEL_DETECTION',     maxResults: 10 },
        { type: 'TEXT_DETECTION',      maxResults: 1  }
      ]
    }]
  };

  const response = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (result.error) return { success: false, error: result.error.message };

  const ann = result.responses[0];

  // Objects — localized items in image
  const objects = (ann.localizedObjectAnnotations || [])
    .map(o => ({ name: o.name, confidence: Math.round(o.score * 100) }))
    .filter(o => o.confidence >= 50);

  // Labels — general categories
  const labels = (ann.labelAnnotations || [])
    .map(l => ({ name: l.description, confidence: Math.round(l.score * 100) }))
    .filter(l => l.confidence >= 60)
    .slice(0, 6);

  // Full text
  const texts = ann.textAnnotations || [];
  const fullText = texts.length > 0 ? texts[0].description : '';

  // Parse quantities from text — หาตัวเลขตามด้วยหน่วย
  const quantities = parseQuantities(fullText);

  return {
    success: true,
    objects,
    labels,
    text: fullText,
    quantities
  };
}

// ===== parseQuantities — ดึงจำนวน+หน่วยจาก OCR text =====
function parseQuantities(text) {
  if (!text) return [];
  var units = 'ชิ้น|อัน|กล่อง|แกลลอน|ลิตร|กระปุก|ถุง|แผ่น|ม้วน|โหล|คู่|ชุด|หลอด|กระป๋อง|ml|ML|L|kg|KG|g|G|m|cm|mm';
  var pattern = new RegExp('(\\d+(?:\\.\\d+)?)\\s*(' + units + ')', 'g');
  var results = [];
  var match;
  while ((match = pattern.exec(text)) !== null) {
    results.push({ amount: match[1], unit: match[2] });
  }
  return results;
}

// ===== getStock =====
function getStock() {
  const sheet = getOrCreateSheet(SHEET_STOCK);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, items: [] };

  const rows = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const items = rows
    .filter(r => r[1])
    .map(r => ({
      id:        String(r[0]),
      name:      r[1],
      qty:       String(r[2]),
      unit:      r[3] || '',
      category:  r[4] || '',
      details:   r[5] || '',
      photoUrl:  r[6] || '',
      updatedAt: r[7] || ''
    }));

  return { success: true, items };
}

// ===== addStock =====
function addStock(p) {
  const sheet = getOrCreateSheet(SHEET_STOCK);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID','ชื่ออุปกรณ์','จำนวน','หน่วย','หมวดหมู่','รายละเอียด','URL รูปภาพ','อัพเดทล่าสุด']);
    formatHeader(sheet);
  }

  let photoUrl = '';
  if (p.photoBase64 && p.photoName) {
    photoUrl = uploadToDrive(p.photoBase64, p.photoName, FOLDER_STOCK);
  }

  const newId = 'STK-' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMddHHmmss');
  sheet.appendRow([newId, p.name, p.qty, p.unit || '', p.category || 'ทั่วไป', p.details || '', photoUrl, p.updatedAt]);

  return { success: true, id: newId };
}

// ===== updateStock =====
function updateStock(p) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_STOCK);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.id)) {
      if (p.qty       !== undefined) sheet.getRange(i + 1, 3).setValue(p.qty);
      if (p.name      !== undefined) sheet.getRange(i + 1, 2).setValue(p.name);
      if (p.details   !== undefined) sheet.getRange(i + 1, 6).setValue(p.details);
      if (p.updatedAt !== undefined) sheet.getRange(i + 1, 8).setValue(p.updatedAt);
      return { success: true };
    }
  }
  return { success: false, error: 'Item not found' };
}

// ===== deleteStock =====
function deleteStock(p) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_STOCK);
  if (!sheet) return { success: false, error: 'Sheet not found' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.id)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Item not found' };
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
