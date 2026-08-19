/**
 * ระบบตรวจข้อสอบโรงเรียนบ้านสันดาบ (Ban Sandab School OMR System)
 * Google Apps Script Backend — จัดการคลังเฉลยใน Google Sheets
 * รองรับทั้ง Google Apps Script Web App และ Standalone Web / GitHub Pages API
 */

function doGet(e) {
  // หากเรียกผ่าน API (เช่น จาก GitHub Pages)
  if (e && e.parameter && e.parameter.action) {
    let result = {};
    if (e.parameter.action === 'getExamKeys') {
      result = getExamKeys();
    } else if (e.parameter.action === 'checkConnection') {
      result = checkConnection();
    } else {
      result = { status: 'error', message: 'Unknown action: ' + e.parameter.action };
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // หากเปิดตรงๆ ผ่านเบราว์เซอร์
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบตรวจข้อสอบโรงเรียนบ้านสันดาบ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let result = { status: 'error', message: 'No payload' };
  try {
    let contents = {};
    if (e && e.postData && e.postData.contents) {
      contents = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      contents = e.parameter;
    }

    if (contents.action === 'syncAllExamKeys') {
      let exams = contents.data;
      if (typeof exams === 'string') exams = JSON.parse(exams);
      result = syncAllExamKeys(exams);
    } else if (contents.action === 'saveExamKey') {
      let exam = contents.data;
      if (typeof exam === 'string') exam = JSON.parse(exam);
      result = saveExamKey(exam);
    } else if (contents.action === 'getExamKeys') {
      result = getExamKeys();
    }
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// ฟังก์ชันเชื่อมต่อฐานข้อมูล Google Sheets
// ==========================================
function getDatabase() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    // Fallback ID (ปรับเปลี่ยนเป็น Google Sheet ID ของคุณได้)
    const SHEET_ID = '1OOb05Fb2VnaFo8UYVwrCSOn6vXWofvXBGdv9yScEQek';
    ss = SpreadsheetApp.openById(SHEET_ID);
  }
  return ss;
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    const sheets = ss.getSheets();
    for (let s of sheets) {
      if (s.getName().trim().toLowerCase().includes(name.toLowerCase())) {
        sheet = s;
        break;
      }
    }
  }
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0 && headers) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#991b1b")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ==========================================
// 1. ดึงคลังเฉลยทั้งหมดจาก Google Sheets
// ==========================================
function getExamKeys() {
  try {
    const ss = getDatabase();
    const headers = ['รายวิชา', 'ระดับชั้น', 'จำนวนข้อ', 'คะแนนเต็ม', 'เฉลย (JSON)', 'วันที่สร้าง'];
    const sheet = getOrCreateSheet(ss, 'คลังเฉลย', headers);

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { status: 'success', data: [] };
    }

    const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const exams = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (!row[0] || !row[2]) continue; // ข้ามแถวว่าง
      let answerKey = {};
      try { 
        answerKey = typeof row[4] === 'string' ? JSON.parse(row[4] || '{}') : (row[4] || {}); 
      } catch(e) {}
      exams.push({
        subject: String(row[0] || ''),
        level: String(row[1] || ''),
        questionCount: Number(row[2]) || 0,
        maxScore: Number(row[3]) || 0,
        answerKey: answerKey,
        createdAt: String(row[5] || '')
      });
    }

    return { status: 'success', data: exams };
  } catch (error) {
    return { status: 'error', message: error.toString(), data: [] };
  }
}

// ==========================================
// 2. บันทึกเฉลยใหม่ลง Google Sheets
// ==========================================
function saveExamKey(examData) {
  try {
    const ss = getDatabase();
    const headers = ['รายวิชา', 'ระดับชั้น', 'จำนวนข้อ', 'คะแนนเต็ม', 'เฉลย (JSON)', 'วันที่สร้าง'];
    const sheet = getOrCreateSheet(ss, 'คลังเฉลย', headers);

    const answerKeyJson = JSON.stringify(examData.answerKey || {});
    const newRow = [
      examData.subject || '',
      examData.level || '',
      examData.questionCount || 0,
      examData.maxScore || 0,
      answerKeyJson,
      examData.createdAt || new Date().toLocaleString('th-TH')
    ];

    sheet.appendRow(newRow);
    return { status: 'success', message: 'บันทึกเฉลย "' + examData.subject + '" (' + examData.level + ') สำเร็จ' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

// ==========================================
// 3. ลบเฉลยจาก Google Sheets
// ==========================================
function deleteExamKey(index) {
  try {
    const ss = getDatabase();
    const headers = ['รายวิชา', 'ระดับชั้น', 'จำนวนข้อ', 'คะแนนเต็ม', 'เฉลย (JSON)', 'วันที่สร้าง'];
    const sheet = getOrCreateSheet(ss, 'คลังเฉลย', headers);

    const rowToDelete = Number(index) + 2;
    const lastRow = sheet.getLastRow();

    if (rowToDelete >= 2 && rowToDelete <= lastRow) {
      sheet.deleteRow(rowToDelete);
      return { status: 'success', message: 'ลบเฉลยเรียบร้อย' };
    } else {
      return { status: 'error', message: 'ไม่พบแถวที่ต้องการลบ' };
    }
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

// ==========================================
// 4. แทนที่เฉลยทั้งหมด (Full Sync จาก Client)
// ==========================================
function syncAllExamKeys(examsArray) {
  try {
    const ss = getDatabase();
    const headers = ['รายวิชา', 'ระดับชั้น', 'จำนวนข้อ', 'คะแนนเต็ม', 'เฉลย (JSON)', 'วันที่สร้าง'];
    const sheet = getOrCreateSheet(ss, 'คลังเฉลย', headers);

    // ลบข้อมูลเดิมทั้งหมด (เก็บ header ไว้)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }

    // เขียนข้อมูลใหม่ทั้งหมด
    if (examsArray && examsArray.length > 0) {
      const dataToWrite = examsArray.map(function(ex) {
        return [
          ex.subject || '',
          ex.level || '',
          ex.questionCount || 0,
          ex.maxScore || 0,
          JSON.stringify(ex.answerKey || {}),
          ex.createdAt || new Date().toLocaleString('th-TH')
        ];
      });
      sheet.getRange(2, 1, dataToWrite.length, 6).setValues(dataToWrite);
    }

    return { status: 'success', message: 'ซิงค์เฉลย ' + (examsArray ? examsArray.length : 0) + ' ชุดสำเร็จ' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

// ==========================================
// 5. ตรวจสอบสถานะเชื่อมต่อ
// ==========================================
function checkConnection() {
  try {
    const ss = getDatabase();
    return { status: 'success', message: 'เชื่อมต่อ Google Sheets สำเร็จ', sheetName: ss.getName() };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}
