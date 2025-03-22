# Code.gs
/**
 * --- Code.gs ---
 *
 * ไฟล์ Google Apps Script นี้มีตรรกะฝั่งเซิร์ฟเวอร์สำหรับ
 * แบบฟอร์มใบขอเบิกเงิน (form.html) จัดการการโหลดข้อมูล,
 * การส่งแบบฟอร์ม, การสร้างเอกสาร และงานแบ็กเอนด์อื่น ๆ
 */

function doGet(e) {
  try {
    let page = e.parameter.page || 'form'; // กำหนดหน้าเริ่มต้นเป็น 'form'

    const html = HtmlService.createTemplateFromFile(page)
      .evaluate()
      .setTitle('ระบบจัดการการเบิกจ่ายเงิน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

    return html;
  } catch (error) {
    Logger.log("Error in doGet: " + error.message);
    return HtmlService.createHtmlOutput("เกิดข้อผิดพลาด: " + error.message);
  }
}

// --- ตัวแปรส่วนกลาง ---
var SHEET_ID = "17kgkDuFf7XfAq1hVdp4Lny-_noCOcDuDfVgjQLYJYFM"; // แทนที่ด้วย ID ของ Google Sheets ของคุณ
var MASTER_DATA_SHEET_NAME = "MasterData"; // ชื่อชีทที่เก็บข้อมูลหลัก
var REQUISITIONS_SHEET_NAME = "Requisitions"; // ชื่อชีทที่เก็บข้อมูลใบขอเบิก
var DOCUMENT_NUMBER_PREFIX = "RQ"; // คำนำหน้าของเลขที่เอกสาร
const MAX_EXPENSE_ITEMS = 10; // จำนวนรายการค่าใช้จ่ายสูงสุด

// --- ฟังก์ชันอรรถประโยชน์ ---

/**
 * ส่งคืนสเปรดชีตที่ใช้งานอยู่
 * @return {Spreadsheet} สเปรดชีตที่ใช้งานอยู่
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

/**
 * ส่งคืนชีทตามชื่อ
 * @param {string} sheetName ชื่อของชีท
 * @return {Sheet} ชีท
 */
function getSheetByName(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = createSheetWithHeaders(ss, sheetName); // สร้างชีทพร้อมหัวตารางถ้ายังไม่มี
  }
  return sheet;
}

/**
 * บันทึกการเข้าสู่ระบบของผู้ใช้
 * @return {string} อีเมลของผู้ใช้ที่เข้าสู่ระบบ
 */
function getUserLogin() {
  return Session.getActiveUser().getEmail();
}

/**
 * สร้างเลขที่เอกสารที่ไม่ซ้ำกัน
 * @return {string} เลขที่เอกสารที่สร้างขึ้น
 */
function generateDocumentNumber() {
  var sheet = getSheetByName(REQUISITIONS_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var lastDocumentNumber = "";

  // หาเลขที่เอกสารล่าสุด ถ้ามี
  if (lastRow > 1) {
    lastDocumentNumber = sheet.getRange(lastRow, 1).getValue(); // assuming doc number is in the first column
  }

  var sequenceNumber = 1;

  // ถ้ามีเลขที่เอกสารล่าสุด, แยกส่วนที่เป็นตัวเลขออกมา
  if (lastDocumentNumber) {
    var lastSequenceNumber = parseInt(lastDocumentNumber.slice(DOCUMENT_NUMBER_PREFIX.length));
    if (!isNaN(lastSequenceNumber)) {
      sequenceNumber = lastSequenceNumber + 1;
    }
  }

  // สร้างเลขที่เอกสารใหม่
  var newDocumentNumber = DOCUMENT_NUMBER_PREFIX + sequenceNumber.toString().padStart(4, '0');
  return newDocumentNumber;
}

/**
 * จัดรูปแบบอ็อบเจ็กต์วันที่ให้อยู่ในรูปแบบ YYYY-MM-DD
 * @param {Date} date อ็อบเจ็กต์วันที่
 * @return {string} สตริงวันที่ที่จัดรูปแบบแล้ว
 */
function formatDate(date) {
  var year = date.getFullYear();
  var month = String(date.getMonth() + 1).padStart(2, '0');
  var day = String(date.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}

// --- ฟังก์ชันสร้างชีท ---

/**
 * สร้างชีทพร้อมส่วนหัวที่ระบุ หากยังไม่มี
 * @param {Spreadsheet} spreadsheet สเปรดชีตที่จะสร้างชีทใน
 * @param {string} sheetName ชื่อของชีทที่จะสร้าง
 * @return {Sheet} ชีทที่สร้างใหม่
 */
function createSheetWithHeaders(spreadsheet, sheetName) {
  var newSheet = spreadsheet.insertSheet(sheetName);
  var headers;

  if (sheetName === MASTER_DATA_SHEET_NAME) {
    headers = ["Company Name", "Project Name", "Expense Type", "Payment Method", "Recipient Name", "Recipient Bank Account"]; // หัวตารางสำหรับ MasterData
  } else if (sheetName === REQUISITIONS_SHEET_NAME) {
    headers = ["เลขที่เอกสาร", "เลขที่อ้างอิง", "บริษัท", "วันที่บันทึก", "วันที่ครบกำหนด", "ประเภทรายจ่าย", "ผู้รับเงิน", "บัญชีธนาคาร", "โครงการ", "รูปแบบการจ่าย", "หมายเหตุ", "จำนวนเงินรวม"];
    // เพิ่มคอลัมน์สำหรับรายการค่าใช้จ่าย 10 รายการ
    for (let i = 1; i <= MAX_EXPENSE_ITEMS; i++) {
      headers.push(`รายละเอียด ${i}`);
      headers.push(`จำนวนเงิน ${i}`);
    }
    headers.push("ผู้บันทึก");
  } else {
    // จัดการชื่อชีทอื่น ๆ หรือ throw ข้อผิดพลาด
    throw new Error("Unknown sheet name: " + sheetName);
  }

  newSheet.getRange(1, 1, 1, headers.length).setValues([headers]); // ตั้งค่าส่วนหัว
  return newSheet;
}

// --- ฟังก์ชันโหลดข้อมูล ---

/**
 * ดึงข้อมูลหลัก (บริษัท โครงการ ฯลฯ) จากสเปรดชีต
 * @return {object} อ็อบเจ็กต์ที่มีข้อมูลหลัก
 */
function getFormMasterData() {
  var sheet = getSheetByName(MASTER_DATA_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var masterData = {
    companies: [],
    projects: [],
    expenseTypes: [],
    paymentMethods: [],
    recipients: []
  };

  // Assuming data is structured in columns like:
  // | Company Name | Project Name | Expense Type | Payment Method | Recipient Name | Recipient Bank Account |
  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    if (row[0]) masterData.companies.push({ name: row[0] });
    if (row[1]) masterData.projects.push({ name: row[1] });
    if (row[2]) masterData.expenseTypes.push({ name: row[2] });
    if (row[3]) masterData.paymentMethods.push({ name: row[3] });
    if (row[4]) masterData.recipients.push({ name: row[4], bankAccount: row[5] || "" }); // Assuming bank account is in the next column
  }

  // Remove duplicates (optional, but good practice)
  masterData.companies = removeDuplicates(masterData.companies, "name");
  masterData.projects = removeDuplicates(masterData.projects, "name");
  masterData.expenseTypes = removeDuplicates(masterData.expenseTypes, "name");
  masterData.paymentMethods = removeDuplicates(masterData.paymentMethods, "name");
  masterData.recipients = removeDuplicates(masterData.recipients, "name");

  return masterData;
}

/**
 * ลบอ็อบเจ็กต์ที่ซ้ำกันออกจากอาร์เรย์โดยพิจารณาจากคุณสมบัติ
 * @param {array} arr อาร์เรย์ของอ็อบเจ็กต์
 * @param {string} prop คุณสมบัติที่จะตรวจสอบความซ้ำ
 * @return {array} อาร์เรย์ที่ไม่มีอ็อบเจ็กต์ที่ซ้ำกัน
 */
function removeDuplicates(arr, prop) {
  return arr.filter((obj, pos, arr) => {
    return arr.map(mapObj => mapObj[prop]).indexOf(obj[prop]) === pos;
  });
}

/**
 * ดึงข้อมูลใบขอเบิกตามเลขที่เอกสาร
 * @param {string} docNumber เลขที่เอกสารที่ต้องการค้นหา
 * @return {object} อ็อบเจ็กต์ที่มีข้อมูลใบขอเบิก หรือข้อความแสดงข้อผิดพลาด
 */
function getPaymentRequisitionByDocNumber(docNumber) {
  var sheet = getSheetByName(REQUISITIONS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = getSheetHeaders(REQUISITIONS_SHEET_NAME);
  var requisitionData = {};

  // Loop through rows to find the matching document number
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] === docNumber) { // Assuming document number is in the first column
      // Map each column to its header
      for (var j = 0; j < headers.length; j++) {
        requisitionData[headers[j]] = row[j];
      }
      return { success: true, data: requisitionData };
    }
  }

  return { success: false, message: "ไม่พบใบตั้งเบิกเลขที่ " + docNumber };
}

// --- ฟังก์ชันบันทึกข้อมูล ---

/**
 * บันทึกใบขอเบิกใหม่ลงในสเปรดชีต
 * @param {object} formData ข้อมูลแบบฟอร์มที่จะบันทึก
 * @return {object} ข้อความแสดงความสำเร็จหรือข้อผิดพลาด
 */
function savePaymentRequisition(formData) {
  try {
    var sheet = getSheetByName(REQUISITIONS_SHEET_NAME);
    var headers = getSheetHeaders(REQUISITIONS_SHEET_NAME);

    // เตรียมแถวข้อมูล
    var dataRow = [];
    dataRow[headers.indexOf("เลขที่เอกสาร")] = formData.documentNumber;
    dataRow[headers.indexOf("เลขที่อ้างอิง")] = formData.referenceNumber || "";
    dataRow[headers.indexOf("บริษัท")] = formData.company || "";
    dataRow[headers.indexOf("วันที่บันทึก")] = formData.recordDate;
    dataRow[headers.indexOf("วันที่ครบกำหนด")] = formData.dueDate;
    dataRow[headers.indexOf("ประเภทรายจ่าย")] = formData.expenseType;
    dataRow[headers.indexOf("ผู้รับเงิน")] = formData.recipient;
    dataRow[headers.indexOf("บัญชีธนาคาร")] = formData.recipientBankAccount || "";
    dataRow[headers.indexOf("โครงการ")] = formData.project;
    dataRow[headers.indexOf("รูปแบบการจ่าย")] = formData.paymentMethod;
    dataRow[headers.indexOf("หมายเหตุ")] = formData.notes || "";
    dataRow[headers.indexOf("จำนวนเงินรวม")] = formData.totalAmount;

    // บันทึกรายการค่าใช้จ่าย 10 รายการ
    for (let i = 0; i < MAX_EXPENSE_ITEMS; i++) {
      if (formData.expenseItems && formData.expenseItems[i]) {
        dataRow[headers.indexOf(`รายละเอียด ${i + 1}`)] = formData.expenseItems[i].description || "";
        dataRow[headers.indexOf(`จำนวนเงิน ${i + 1}`)] = formData.expenseItems[i].amount || 0;
      } else {
        dataRow[headers.indexOf(`รายละเอียด ${i + 1}`)] = "";
        dataRow[headers.indexOf(`จำนวนเงิน ${i + 1}`)] = 0;
      }
    }
    dataRow[headers.indexOf("ผู้บันทึก")] = getUserLogin(); //บันทึกข้อมูลผู้ใช้

    // เพิ่มข้อมูลลงในชีท
    sheet.appendRow(dataRow);

    return { success: true, message: "บันทึกข้อมูลใบตั้งเบิกเลขที่ " + formData.documentNumber + " สำเร็จ" };

  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาดในการบันทึกข้อมูล: " + e.toString() };
  }
}

/**
 * อัปเดตใบขอเบิกที่มีอยู่ในสเปรดชีต
 * @param {object} formData ข้อมูลแบบฟอร์มที่จะอัปเดต
 * @return {object} ข้อความแสดงความสำเร็จหรือข้อผิดพลาด
 */
function updatePaymentRequisition(formData) {
  try {
    var sheet = getSheetByName(REQUISITIONS_SHEET_NAME);
    var headers = getSheetHeaders(REQUISITIONS_SHEET_NAME);
    var data = sheet.getDataRange().getValues();

    // Find the row to update
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === formData.documentNumber) { // Assuming document number is in the first column
        var row = i + 1; // Row number in the sheet

        // อัปเดตข้อมูลในแถว
        sheet.getRange(row, headers.indexOf("เลขที่อ้างอิง") + 1).setValue(formData.referenceNumber || "");
        sheet.getRange(row, headers.indexOf("บริษัท") + 1).setValue(formData.company || "");
        sheet.getRange(row, headers.indexOf("วันที่บันทึก") + 1).setValue(formData.recordDate);
        sheet.getRange(row, headers.indexOf("วันที่ครบกำหนด") + 1).setValue(formData.dueDate);
        sheet.getRange(row, headers.indexOf("ประเภทรายจ่าย") + 1).setValue(formData.expenseType);
        sheet.getRange(row, headers.indexOf("ผู้รับเงิน") + 1).setValue(formData.recipient);
        sheet.getRange(row, headers.indexOf("บัญชีธนาคาร") + 1).setValue(formData.recipientBankAccount || "");
        sheet.getRange(row, headers.indexOf("โครงการ") + 1).setValue(formData.project);
        sheet.getRange(row, headers.indexOf("รูปแบบการจ่าย") + 1).setValue(formData.paymentMethod);
        sheet.getRange(row, headers.indexOf("หมายเหตุ") + 1).setValue(formData.notes || "");
        sheet.getRange(row, headers.indexOf("จำนวนเงินรวม") + 1).setValue(formData.totalAmount);

        // อัปเดตรายการค่าใช้จ่าย 10 รายการ
        for (let j = 0; j < MAX_EXPENSE_ITEMS; j++) {
          if (formData.expenseItems && formData.expenseItems[j]) {
            sheet.getRange(row, headers.indexOf(`รายละเอียด ${j + 1}`) + 1).setValue(formData.expenseItems[j].description || "");
            sheet.getRange(row, headers.indexOf(`จำนวนเงิน ${j + 1}`) + 1).setValue(formData.expenseItems[j].amount || 0);
          } else {
            sheet.getRange(row, headers.indexOf(`รายละเอียด ${j + 1}`) + 1).setValue("");
            sheet.getRange(row, headers.indexOf(`จำนวนเงิน ${j + 1}`) + 1).setValue(0);
          }
        }
        // Not updating recorder since it should be the original creator

        return { success: true, message: "แก้ไขข้อมูลใบตั้งเบิกเลขที่ " + formData.documentNumber + " สำเร็จ" };
      }
    }

    return { success: false, message: "ไม่พบใบตั้งเบิกเลขที่ " + formData.documentNumber + " สำหรับแก้ไข" };

  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาดในการแก้ไขข้อมูล: " + e.toString() };
  }
}

/**
 * รับส่วนหัวจากชีท
 * @param {string} sheetName ชื่อของชีท
 * @return {array} ส่วนหัว
 */
function getSheetHeaders(sheetName) {
  var sheet = getSheetByName(sheetName);
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

// --- ฟังก์ชัน UI ---

/**
 * เปิดแถบด้านข้างพร้อมเนื้อหา HTML ของแบบฟอร์ม
 */
function showForm() {
  var html = HtmlService.createHtmlOutputFromFile('form')
      .setTitle('Payment Requisition Form');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}

/**
 * รับ URL ของเว็บแอป Google Apps Script
 * @return {string} URL ของเว็บแอป
 */
function getScriptURL() {
  return ScriptApp.getService().getUrl();
}
