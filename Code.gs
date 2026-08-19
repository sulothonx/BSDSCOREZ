/**
 * ระบบตรวจข้อสอบโรงเรียนบ้านสันดาบ (Ban Sandab School OMR System)
 * Google Apps Script Backend — จัดการคลังเฉลยใน Google Sheets
 * ใช้คู่กับ Index.html (Client-side OMR)
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบตรวจข้อสอบโรงเรียนบ้านสันดาบ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// ฟังก์ชันเชื่อมต่อฐานข้อมูล
// ==========================================
function getDatabase() {
  // ใช้ Active Spreadsheet (Container-bound) ก่อน
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    // Fallback: ใช้ ID โดยตรง (Standalone)
    const SHEET_ID = '1OOb05Fb2VnaFo8UYVwrCSOn6vXWofvXBGdv9yScEQek';
    ss = SpreadsheetApp.openById(SHEET_ID);
  }
  return ss;
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    // ลองค้นหาชื่อคล้ายกัน
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
      try { answerKey = JSON.parse(row[4] || '{}'); } catch(e) {}
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
// 3. ลบเฉลยจาก Google Sheets (ตาม index)
// ==========================================
function deleteExamKey(index) {
  try {
    const ss = getDatabase();
    const headers = ['รายวิชา', 'ระดับชั้น', 'จำนวนข้อ', 'คะแนนเต็ม', 'เฉลย (JSON)', 'วันที่สร้าง'];
    const sheet = getOrCreateSheet(ss, 'คลังเฉลย', headers);

    const rowToDelete = Number(index) + 2; // +1 header, +1 for 1-indexed
    const lastRow = sheet.getLastRow();

    if (rowToDelete >= 2 && rowToDelete <= lastRow) {
      sheet.deleteRow(rowToDelete);
      return { status: 'success', message: 'ลบเฉลยเรียบร้อย' };
    } else {
      return { status: 'error', message: 'ไม่พบแถวที่ต้องการลบ (row ' + rowToDelete + ', lastRow ' + lastRow + ')' };
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