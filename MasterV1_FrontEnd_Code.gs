// --- CONFIGURATION ---
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const DATA_SHEET_ID = SCRIPT_PROPS.getProperty('DATA_SHEET_ID');
const DATA_SHEET_NAME = SCRIPT_PROPS.getProperty('DATA_SHEET_NAME');
const LOG_SHEET_ID = SCRIPT_PROPS.getProperty('LOG_SHEET_ID');
const LOG_SHEET_NAME = SCRIPT_PROPS.getProperty('LOG_SHEET_NAME');

// --- WEB APP ---
function doGet(e) {
  logActivity("Web App accessed.");
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ฟอร์มบันทึกข้อมูลและแนบเอกสาร')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- FILE UPLOAD FUNCTION ---
// ฟังก์ชันสำหรับอัปโหลดไฟล์ไปยัง Google Drive
function doUpload(obj) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(obj.data), obj.mimeType, obj.fileName);
    // *** แก้ไข FOLDER ID ตรงนี้ ***
    const id = '1Zujyb1nNIEFoKdKvlDSPo0bkGeNc1g_O'; // <--- !! ใส่ Folder ID ของคุณที่นี่
    const folder = DriveApp.getFolderById(id);
    const file = folder.createFile(blob);
    const fileURL = file.getUrl();
    
    logActivity(`File uploaded successfully: ${obj.fileName}, URL: ${fileURL}`);

    const response = {
      'status': 'success',
      'fileName': obj.fileName,
      'fileUrl': fileURL
    };
    return response;
  } catch (e) {
    logActivity(`ERROR uploading file: ${e.toString()}`);
    return {
      'status': 'error',
      'message': `เกิดข้อผิดพลาดในการอัปโหลดไฟล์: ${e.message}`
    };
  }
}

// --- CORE DATA SAVING FUNCTION ---
// ฟังก์ชันหลักในการบันทึกข้อมูลที่รับมาจากฟอร์ม (รวมข้อมูลไฟล์)
function saveData(formData) {
  logActivity(`Received form submission: ${JSON.stringify(formData)}`);
  try {
    const dataSheet = SpreadsheetApp.openById(DATA_SHEET_ID).getSheetByName(DATA_SHEET_NAME);
    
    const nextReceiptNumber = getNextReceiptNumber(dataSheet);
    const updateDateTime = new Date();
    const status = "Pending";

    // **ปรับปรุงแถวข้อมูลใหม่ ให้มีคอลัมน์สำหรับไฟล์**
    const newRow = [
      nextReceiptNumber,
      formData.title,
      formData.notes,
      formData.username,
      formData.fileName || "", // ใส่ชื่อไฟล์ (ถ้ามี)
      formData.fileUrl || "",  // ใส่ URL ของไฟล์ (ถ้ามี)
      updateDateTime,
      status
    ];

    dataSheet.appendRow(newRow);
    logActivity(`Successfully saved data. New Receipt No: ${nextReceiptNumber}`);
    
    return { 
      status: 'success', 
      message: `บันทึกข้อมูลสำเร็จ! เลขที่รับเข้าคือ: ${nextReceiptNumber}` 
    };

  } catch (e) {
    logActivity(`ERROR saving data: ${e.toString()} | Stack: ${e.stack}`);
    return { 
      status: 'error', 
      message: `เกิดข้อผิดพลาดในการบันทึกข้อมูล: ${e.message}` 
    };
  }
}

// --- HELPER FUNCTIONS ---
// ฟังก์ชันสำหรับหาเลขที่รับเข้าล่าสุด + 1
function getNextReceiptNumber(sheet) {
  logActivity("Getting next receipt number...");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    logActivity("No data found. Starting with 1001.");
    return 1001;
  }
  const lastReceiptNumber = sheet.getRange(lastRow, 1).getValue();
  if (typeof lastReceiptNumber === 'number' && !isNaN(lastReceiptNumber)) {
    const nextNumber = lastReceiptNumber + 1;
    logActivity(`Last number was ${lastReceiptNumber}. Next number is ${nextNumber}.`);
    return nextNumber;
  }
  logActivity(`WARNING: Last receipt number is not a number (${lastReceiptNumber}). Re-calculating from all data.`);
  const allReceiptNumbers = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
                               .flat()
                               .filter(n => typeof n === 'number' && !isNaN(n));
  const maxNumber = allReceiptNumbers.length > 0 ? Math.max(...allReceiptNumbers) : 1000;
  return maxNumber + 1;
}

// ฟังก์ชันสำหรับบันทึก Log
function logActivity(message) {
  try {
    const logSheet = SpreadsheetApp.openById(LOG_SHEET_ID).getSheetByName(LOG_SHEET_NAME);
    logSheet.appendRow([new Date(), message]);
  } catch (e) {
    console.error(`Failed to write to Audit Log: ${e.message}`);
    console.log(`Original Log Message: ${message}`);
  }
}

