/**
 * ============================================================
 *  Universal Google Sheet logger (Apps Script Web App)
 * ------------------------------------------------------------
 *  รับข้อมูล JSON จากหน้าเว็บ (HTML) หลายไฟล์ แล้วบันทึกลง
 *  Google Sheet คนละแท็บ (sheet) ตามค่าฟิลด์ "sheet" ที่ส่งมา
 *
 *  ตัวอย่าง payload ที่ฝั่งเว็บส่งมา:
 *  {
 *    "sheet": "Sheet1",              // ชื่อแท็บปลายทาง (จำเป็น ถ้าไม่ส่งมาจะใช้ "Sheet1")
 *    "timestamp": "...",
 *    "scenario": "Lane Departure Warning",
 *    "user": "...",
 *    "grade": "...",
 *    "show": 100, "learn": 100,
 *    "learn_detail": {"task1":100,"task2":100,"task3":100},
 *    "apply": 100, "reflect": 100,
 *    "overall": 100, "level": 4, "levelName": "...",
 *    "timedOut": false
 *  }
 *
 *  รองรับไฟล์ใหม่ ๆ ในอนาคตโดยอัตโนมัติ:
 *  - ถ้า sheet ที่ระบุยังไม่มีในไฟล์ -> จะสร้างแท็บใหม่ให้เอง
 *  - ถ้า record มีฟิลด์ที่ยังไม่เคยมีในหัวตาราง -> จะเพิ่มคอลัมน์ใหม่ให้เอง
 *  - object ซ้อน (เช่น learn_detail) จะถูก "flatten" เป็นคอลัมน์ย่อย
 *    เช่น learn_detail.task1 -> คอลัมน์ "learn_detail_task1"
 *
 *  วิธีติดตั้ง:
 *  1. เปิด Google Sheet ที่ต้องการใช้เก็บข้อมูล
 *  2. เมนู Extensions > Apps Script วางโค้ดนี้ทับไฟล์ Code.gs เดิม
 *  3. Deploy > New deployment > เลือกประเภท "Web app"
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  4. คัดลอก URL ที่ได้ (…/exec) ไปตั้งเป็น DEFAULT_WEBHOOK_URL ในไฟล์ HTML
 * ============================================================
 */

// ชื่อ sheet เริ่มต้น กรณีฝั่งเว็บไม่ได้ส่งฟิลด์ "sheet" มาด้วย
var DEFAULT_SHEET_NAME = 'Sheet1';

function doPost(e) {
  try {
    var payload = parseIncomingData_(e);
    var sheetName = (payload && payload.sheet) ? String(payload.sheet).trim() : DEFAULT_SHEET_NAME;
    if (!sheetName) sheetName = DEFAULT_SHEET_NAME;

    // ไม่ต้องเก็บฟิลด์ "sheet" เป็นคอลัมน์ในตาราง เพราะใช้แค่บอกปลายทาง
    var data = {};
    Object.keys(payload || {}).forEach(function (k) {
      if (k !== 'sheet') data[k] = payload[k];
    });

    var row = appendRecordToSheet_(sheetName, data);

    return jsonOutput_({ result: 'success', sheet: sheetName, row: row });
  } catch (err) {
    return jsonOutput_({ result: 'error', error: String(err && err.message ? err.message : err) });
  }
}

// เผื่อเปิด URL ด้วยเบราว์เซอร์โดยตรงเพื่อเช็กว่า deploy สำเร็จ
function doGet(e) {
  return jsonOutput_({
    result: 'ok',
    message: 'Google Apps Script web app is running. ส่งข้อมูลด้วย POST (JSON) เท่านั้น',
    sheets: getExistingSheetNames_()
  });
}

/* ============================================================
   ฟังก์ชันหลัก: เพิ่มแถวข้อมูลลง sheet ที่ระบุ (สร้างให้เองถ้ายังไม่มี)
   ใช้ LockService กันการเขียนชนกันเวลามีคนส่งข้อมูลพร้อมกันหลายคน
   ============================================================ */
function appendRecordToSheet_(sheetName, data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    var flat = flattenObject_(data);

    // เติมเวลาบันทึกฝั่งเซิร์ฟเวอร์ไว้ด้วย เผื่อฝั่งเว็บไม่ได้ส่ง timestamp มา
    if (!('timestamp' in flat)) {
      flat['timestamp'] = new Date().toISOString();
    }
    flat['received_at'] = new Date();

    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var headers = [];

    if (lastRow === 0 || lastCol === 0) {
      // แท็บใหม่ ยังไม่มีหัวตาราง -> สร้างหัวตารางจากคีย์ของ record แรกที่ส่งเข้ามา
      headers = Object.keys(flat);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    } else {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    // ถ้ามีฟิลด์ใหม่ที่ยังไม่มีในหัวตาราง ให้เพิ่มคอลัมน์ใหม่ต่อท้ายอัตโนมัติ
    var newKeys = Object.keys(flat).filter(function (k) {
      return headers.indexOf(k) === -1;
    });
    if (newKeys.length > 0) {
      headers = headers.concat(newKeys);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    var rowValues = headers.map(function (h) {
      var v = flat[h];
      if (v === undefined || v === null) return '';
      if (v instanceof Date) return v;
      if (typeof v === 'object') return JSON.stringify(v);
      return v;
    });

    sheet.appendRow(rowValues);
    return sheet.getLastRow();
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
   แปลง object ที่ซ้อนกัน (nested) ให้แบนราบเป็น key เดียว
   เช่น {learn_detail:{task1:1}} -> {learn_detail_task1:1}
   ============================================================ */
function flattenObject_(obj, prefix, out) {
  prefix = prefix || '';
  out = out || {};
  Object.keys(obj || {}).forEach(function (key) {
    var value = obj[key];
    var newKey = prefix ? prefix + '_' + key : key;
    if (value !== null && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date)) {
      flattenObject_(value, newKey, out);
    } else {
      out[newKey] = value;
    }
  });
  return out;
}

/* ============================================================
   อ่านข้อมูลที่ส่งเข้ามา รองรับทั้ง JSON body ปกติ
   และกรณี Content-Type เป็น text/plain (ที่ฝั่งเว็บใช้ mode:no-cors)
   ============================================================ */
function parseIncomingData_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('ไม่พบข้อมูลที่ส่งมา (postData ว่าง)');
  }
  return JSON.parse(e.postData.contents);
}

function getExistingSheetNames_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (s) {
    return s.getName();
  });
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}