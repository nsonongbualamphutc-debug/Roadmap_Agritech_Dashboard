// ============================================================
// Google Apps Script — Roadmap เกษตรเพิ่มมูลค่า จ.หนองบัวลำภู
// ============================================================
// วิธีใช้:
// 1. ไปที่ https://script.google.com → สร้างโปรเจกต์ใหม่
// 2. วาง Code นี้ทั้งหมดแทนที่ Code.gs
// 3. กด "ทำให้ใช้งานได้" → เว็บแอป → ใครก็ได้สามารถเข้าถึง
// 4. คัดลอก URL ที่ได้ไปใส่ใน dashboard.html ตรง APPS_SCRIPT_URL
// ============================================================

// ===== ตั้งค่า Sheet =====
const SHEET_NAME_KPI = 'KPIData';          // ชีตเก็บข้อมูล KPI
const SHEET_NAME_USERS = 'Users';          // ชีตเก็บข้อมูลผู้ใช้
const SHEET_NAME_HISTORY = 'KPIHistory';   // ชีตเก็บประวัติการกรอก

// ===== สร้าง Sheet อัตโนมัติถ้ายังไม่มี =====
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1b5e20')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ===== Headers =====
const KPI_HEADERS = [
  'ยุทธศาสตร์', 'ชื่อยุทธศาสตร์', 'หน่วยงาน', 'ตัวชี้วัด',
  'เป้าหมาย', 'ฐานเดิม', 'ผลงานรอบเดือน', 'หน่วยนับ(เดือน)',
  'ผลงานสะสม', 'หน่วยนับ(สะสม)', '%เปลี่ยนแปลง', 'เดือนรายงาน',
  'ผู้รายงาน', 'วันที่บันทึก'
];

const HISTORY_HEADERS = [
  'Timestamp', 'ยุทธศาสตร์', 'ตัวชี้วัด', 'ฐานเดิม',
  'ผลงานรอบเดือน', 'ผลงานสะสม', '%เปลี่ยนแปลง',
  'เดือนรายงาน', 'ผู้รายงาน'
];

const USER_HEADERS = [
  'Username', 'Password', 'Role', 'Agency', 'CreatedDate'
];

// ===== Web App Entry Points =====
function doGet(e) {
  const action = e.parameter.action || '';
  const callback = e.parameter.callback || '';

  let result = {};

  switch (action) {
    case 'getKPI':
      result = getKPIData();
      break;
    case 'getUsers':
      result = getUsersData();
      break;
    case 'ping':
      result = { status: 'ok', message: 'Backend connected!' };
      break;
    default:
      result = { status: 'ok', message: 'Roadmap เกษตรเพิ่มมูลค่า API' };
  }

  // Support JSONP
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const action = e.parameter?.action || '';
    let body = {};

    try {
      body = JSON.parse(e.postData?.contents || '{}');
    } catch (err) {
      return jsonResponse({ status: 'error', message: 'Invalid JSON' });
    }

    switch (action) {
      case 'saveKPI':
        return jsonResponse(saveKPIData(body));
      case 'saveUsers':
        return jsonResponse(saveUsersData(body));
      default:
        return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ===== SAVE KPI DATA =====
function saveKPIData(body) {
  const rows = body.rows || [];
  if (rows.length === 0) return { status: 'error', message: 'No data' };

  // Save to KPIData sheet (overwrite)
  const sheet = getOrCreateSheet(SHEET_NAME_KPI, KPI_HEADERS);

  // Clear old data (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, KPI_HEADERS.length).clearContent();
  }

  // Write new data
  const dataRows = rows.map(r => [
    r.strategy || '',
    r.strategyName || '',
    r.agency || '',
    r.kpiName || '',
    r.target || '',
    r.baseline || '',
    r.monthVal || '',
    r.monthUnit || '',
    r.cumulVal || '',
    r.cumulUnit || '',
    r.pctChange || 0,
    r.reportMonth || '',
    r.reportedBy || '',
    r.timestamp || new Date().toLocaleString('th-TH')
  ]);

  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, KPI_HEADERS.length).setValues(dataRows);
  }

  // Also append to History
  const histSheet = getOrCreateSheet(SHEET_NAME_HISTORY, HISTORY_HEADERS);
  const histRows = rows
    .filter(r => r.cumulVal) // only save rows that have data
    .map(r => [
      new Date().toLocaleString('th-TH'),
      r.strategy || '',
      r.kpiName || '',
      r.baseline || '',
      r.monthVal || '',
      r.cumulVal || '',
      r.pctChange || 0,
      r.reportMonth || '',
      r.reportedBy || ''
    ]);

  if (histRows.length > 0) {
    histSheet.getRange(histSheet.getLastRow() + 1, 1, histRows.length, HISTORY_HEADERS.length)
      .setValues(histRows);
  }

  // Auto-format
  formatSheet(sheet);

  return { status: 'ok', message: 'บันทึก ' + dataRows.length + ' รายการสำเร็จ', count: dataRows.length };
}

// ===== GET KPI DATA =====
function getKPIData() {
  const sheet = getOrCreateSheet(SHEET_NAME_KPI, KPI_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { status: 'ok', data: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, KPI_HEADERS.length).getValues();
  const result = data.map(row => ({
    strategy: row[0],
    strategyName: row[1],
    agency: row[2],
    kpiName: row[3],
    target: row[4],
    baseline: row[5],
    monthVal: row[6],
    monthUnit: row[7],
    cumulVal: row[8],
    cumulUnit: row[9],
    pctChange: row[10],
    reportMonth: row[11],
    reportedBy: row[12],
    timestamp: row[13]
  }));

  return { status: 'ok', data: result };
}

// ===== SAVE USERS =====
function saveUsersData(body) {
  const userList = body.users || [];
  const sheet = getOrCreateSheet(SHEET_NAME_USERS, USER_HEADERS);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).clearContent();
  }

  const dataRows = userList.map(u => [
    u.username || '',
    u.password || '',
    u.role || 'user',
    u.agency || '',
    new Date().toLocaleString('th-TH')
  ]);

  if (dataRows.length > 0) {
    sheet.getRange(2, 1, dataRows.length, USER_HEADERS.length).setValues(dataRows);
  }

  return { status: 'ok', message: 'บันทึกผู้ใช้ ' + dataRows.length + ' คนสำเร็จ' };
}

// ===== GET USERS =====
function getUsersData() {
  const sheet = getOrCreateSheet(SHEET_NAME_USERS, USER_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { status: 'ok', data: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).getValues();
  const result = data.map(row => ({
    username: row[0],
    password: row[1],
    role: row[2],
    agency: row[3]
  }));

  return { status: 'ok', data: result };
}

// ===== FORMAT =====
function formatSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).setVerticalAlignment('middle');
    sheet.autoResizeColumns(1, lastCol);
  }
}

// ===== UTIL =====
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== SETUP (รันครั้งแรก) =====
function setupSheets() {
  getOrCreateSheet(SHEET_NAME_KPI, KPI_HEADERS);
  getOrCreateSheet(SHEET_NAME_HISTORY, HISTORY_HEADERS);
  getOrCreateSheet(SHEET_NAME_USERS, USER_HEADERS);

  // สร้าง Admin เริ่มต้น
  const userSheet = getOrCreateSheet(SHEET_NAME_USERS, USER_HEADERS);
  if (userSheet.getLastRow() <= 1) {
    userSheet.getRange(2, 1, 1, USER_HEADERS.length).setValues([
      ['ARG001', '039001', 'admin', 'ผู้ดูแลระบบ', new Date().toLocaleString('th-TH')]
    ]);
  }

  SpreadsheetApp.getUi().alert('✅ สร้าง Sheet ทั้งหมดเรียบร้อย!\n\nSheet ที่สร้าง:\n• KPIData\n• KPIHistory\n• Users\n\nAdmin เริ่มต้น: ARG001 / 039001');
}
