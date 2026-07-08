// ==========================================
// การตั้งค่าเริ่มต้น (Configuration)
// ==========================================
const SHEET_ID = '1VhamXrDMql0ACyi7vlBX2DC--AizF2zXTxq_XwdZ8TE';
const FOLDER_ID = '1j-OGh4nDhr1_OQwwdHMRpQuHU1Q3VQS1'; // โฟลเดอร์สำหรับเก็บรูปภาพ
const VALID_USERNAME = '@mmwnews';
const VALID_PASSWORD = 'admin@suwatt';

// ==========================================
// 1. ฟังก์ชันขอสิทธิ์ (รันครั้งแรกเพื่อขออนุญาตเข้าถึง Google Drive)
// ==========================================
function testDrivePermission() {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    console.log("✅ เชื่อมต่อ Drive สำเร็จ! ชื่อโฟลเดอร์: " + folder.getName());
    return "OK: Permission Granted";
  } catch (e) {
    console.error("❌ Error: " + e.toString());
    return "Error: " + e.toString();
  }
}

// ==========================================
// 2. HTTP GET Request (ดึงข้อมูล)
// ==========================================
function doGet(e) {
  try {
    const output = ContentService.createTextOutput();
    output.setMimeType(ContentService.MimeType.JSON);
    
    // ตรวจสอบว่า parameter ว่างหรือไม่
    if (!e || !e.parameter) {
       return output.setContent(JSON.stringify({ success: false, error: 'No parameters provided' }));
    }
    
    const action = e.parameter.action;
    
    if (action === 'getActivities') {
      return getActivities(e.parameter);
    } else if (action === 'login') {
      return handleLogin(e.parameter);
    } else if (action === 'test') {
      return testConnection();
    } else if (action === 'getYears') {
      return getAvailableYears();
    }
    
    return output.setContent(JSON.stringify({
      success: false,
      error: 'Invalid action',
      availableActions: ['getActivities', 'login', 'test', 'getYears']
    }));
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. HTTP POST Request (เพิ่ม, แก้ไข, ลบ, อัปโหลดรูป)
// ==========================================
function doPost(e) {
  try {
    let data = null;
    
    // 1. ลองแปลงจาก JSON Body (สำหรับอัปโหลดรูป)
    if (e.postData && e.postData.contents) {
      const content = e.postData.contents;
      if (content.trim().startsWith('{')) {
        try {
          data = JSON.parse(content);
        } catch (jsonErr) {
          console.log('JSON parse failed, falling back to FormData');
        }
      }
    }
    
    // 2. ลองแปลงจาก FormData (พารามิเตอร์ปกติ)
    if (!data && e.parameter && Object.keys(e.parameter).length > 0) {
      data = e.parameter;
    }
    
    // 3. Fallback สำหรับ URL Encoded
    if (!data && e.postData && e.postData.contents) {
      data = {};
      const params = e.postData.contents.split('&');
      for (let param of params) {
        const [key, value] = param.split('=');
        if (key && value !== undefined) {
          data[decodeURIComponent(key)] = decodeURIComponent(value);
        }
      }
    }

    if (!data || !data.action) {
      throw new Error('Invalid POST action or payload is empty');
    }
    
    const action = data.action;
    
    // แยกการทำงานตาม Action
    if (action === 'incrementView') {
      return incrementView(data);
    } else if (action === 'addActivity') {
      return addActivity(data);
    } else if (action === 'updateActivity') {
      return updateActivity(data);
    } else if (action === 'deleteActivity') {
      return deleteActivity(data);
    } else if (action === 'togglePublish') {
      return togglePublish(data);
    } else if (action === 'uploadImage') {
      return uploadImage(data);
    } else if (action === 'login') {
      return handleLogin(data);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Unknown action: ' + action
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString(),
      stack: error.stack
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// ฟังก์ชันอัปโหลดรูปภาพเข้า Google Drive
// ==========================================
function uploadImage(data) {
  try {
    verifyToken(data.token);
    
    if (!data.image || !data.filename) {
      throw new Error('Image data or filename missing');
    }

    // แยก Header ออกจากเนื้อหา Base64
    const parts = data.image.split(',');
    if (parts.length < 2) throw new Error('Invalid image format');
    
    const contentType = parts[0].split(':')[1].split(';')[0];
    const base64Data = parts[1];
    
    // สร้าง Blob สำหรับอัปโหลด
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, contentType, data.filename);
    
    // บันทึกไฟล์ลงโฟลเดอร์
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const file = folder.createFile(blob);
    
    // ปลดล็อกให้ไฟล์ที่อัปโหลดเป็นสาธารณะ (ให้ทุกคนที่มีลิงก์ดูได้)
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      console.log('Permission warning: ไม่สามารถตั้งค่าแชร์อัตโนมัติได้ กรุณาตั้งค่าโฟลเดอร์ให้เป็นสาธารณะด้วยตนเอง');
    }

    // แก้ไข URL ให้แสดงผลในรูปแบบ lh3.googleusercontent.com
    const fileUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      imageUrl: fileUrl,
      fileId: file.getId()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Upload failed: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// ฟังก์ชันนับยอดวิว (View Counter)
// ==========================================
function incrementView(data) {
  // ใช้ LockService เพื่อป้องกันปัญหาหากมีคนกดดูพร้อมกันหลายคน
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); 
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Server Busy' })).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    const dataRange = sheet.getDataRange().getValues();
    
    // หากแถวแรก (Header) ยังไม่มีคอลัมน์ Views ให้สร้างขึ้นมา
    if (dataRange[0] && dataRange[0].length <= 9) {
       sheet.getRange(1, 10).setValue('Views');
    }

    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][0] === data.id) {
        let currentViews = parseInt(dataRange[i][9]);
        if (isNaN(currentViews)) currentViews = 0;
        
        let newViews = currentViews + 1;
        sheet.getRange(i + 1, 10).setValue(newViews); // อัปเดตในชีต
        
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          views: newViews
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Activity Not found' })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// Helper Functions (ฟังก์ชันตัวช่วย)
// ==========================================

function verifyToken(token) {
  if (!token || !token.startsWith('valid_token_')) {
    throw new Error('Unauthorized - กรุณาเข้าสู่ระบบใหม่');
  }
  return true;
}

function handleLogin(params) {
  if (params.username === VALID_USERNAME && params.password === VALID_PASSWORD) {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      token: 'valid_token_' + Date.now()
    })).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' })).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// การดึงข้อมูลกิจกรรม (จัดการเรื่องวันที่และ Timezone)
// ==========================================
function getActivities(params) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const activities = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // ถ้ามี ID ให้ดำเนินการต่อ
    if (row[0] && i > 0) { 
      let formattedDate = '';
      
      // การตรวจสอบและแปลงวันที่
      if (row[5] instanceof Date) {
        // กรณีเป็น Date Object ใน Google Sheets ให้บวกเวลาเพิ่ม 12 ชั่วโมง 
        // เพื่อป้องกันปัญหา Timezone ทำให้วันถอยหลังไป 1 วันเมื่อแปลงเป็นข้อความ
        let safeDate = new Date(row[5].getTime() + (12 * 60 * 60 * 1000));
        formattedDate = Utilities.formatDate(safeDate, 'Asia/Bangkok', 'yyyy-MM-dd');
      } else {
        // กรณีบันทึกเป็น Text รูปแบบ YYYY-MM-DD
        formattedDate = row[5] ? String(row[5]) : '';
      }
      
      const viewCount = parseInt(row[9]) || 0;

      const act = {
        id: row[0],
        title: row[1] || '',
        description: row[2] || '',
        imageUrl: row[3] || '',
        albumLink: row[4] || '',
        date: formattedDate,
        tags: row[6] || '',
        photographer: row[7] || '',
        isPublished: String(row[8]).toLowerCase() === 'true',
        views: viewCount
      };
      
      // ระบบค้นหาและตัวกรอง (Filter)
      let match = true;
      if (params.query) {
        const q = params.query.toLowerCase();
        if (!act.title.toLowerCase().includes(q) && !act.tags.toLowerCase().includes(q)) match = false;
      }
      if (params.month && act.date) {
        // ตัด YYYY-MM-DD เอาเฉพาะเดือนที่ตำแหน่ง index 1
        const m = act.date.split('-')[1]; 
        if (m !== params.month) match = false;
      }
      if (params.year && act.date) {
        // ตัด YYYY-MM-DD เอาเฉพาะปีที่ตำแหน่ง index 0
        const y = act.date.split('-')[0];
        if (y !== params.year) match = false;
      }
      if (params.tag && !act.tags.includes(params.tag)) match = false;

      if (match) activities.push(act);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    data: activities
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// การบันทึกและแก้ไขข้อมูล
// ==========================================
function addActivity(data) {
  verifyToken(data.token);
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  
  // ถ้าตารางว่างเปล่า ให้สร้าง Header 
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', 'Title', 'Description', 'ImageURL', 'AlbumLink', 'Date', 'Tags', 'Photographer', 'IsPublished', 'Views']);
  }
  
  // ใช้ "'" (single quote) นำหน้า data.date เพื่อบังคับให้ Sheets เก็บค่าเป็น Text เสมอ
  // ช่วยแก้ปัญหาที่ Sheets แอบเปลี่ยนจาก 10/11 เป็นเดือน 10 วันที่ 11
  sheet.appendRow([
    data.id, data.title, data.description, data.imageUrl, data.albumLink, "'" + data.date, data.tags, data.photographer, data.isPublished, 0
  ]);
  return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
}

function updateActivity(data) {
  verifyToken(data.token);
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.id) {
      sheet.getRange(i+1, 2).setValue(data.title);
      sheet.getRange(i+1, 3).setValue(data.description);
      sheet.getRange(i+1, 4).setValue(data.imageUrl);
      sheet.getRange(i+1, 5).setValue(data.albumLink);
      
      // บังคับให้วันที่เป็น Text
      sheet.getRange(i+1, 6).setValue("'" + data.date);
      
      sheet.getRange(i+1, 7).setValue(data.tags);
      sheet.getRange(i+1, 8).setValue(data.photographer);
      sheet.getRange(i+1, 9).setValue(data.isPublished);
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  throw new Error('ID not found');
}

function deleteActivity(data) {
  verifyToken(data.token);
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  throw new Error('ID not found');
}

function togglePublish(data) {
  verifyToken(data.token);
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.id) {
      const current = values[i][8] === true || String(values[i][8]).toLowerCase() === 'true';
      sheet.getRange(i + 1, 9).setValue(!current);
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  throw new Error('ID not found');
}

// ==========================================
// อื่นๆ
// ==========================================
function getAvailableYears() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const years = new Set();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][5]) {
      // แม้ข้อมูลจะถูกบันทึกเป็น Text รูปแบบ YYYY-MM-DD การใช้ new Date() ก็สามารถแยกปีออกมาได้
      let d = new Date(data[i][5]);
      if (!isNaN(d.getTime())) years.add(d.getFullYear());
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    years: Array.from(years).sort((a,b) => b-a) // เรียงลำดับจากปีล่าสุดไปเก่าสุด
  })).setMimeType(ContentService.MimeType.JSON);
}

function testConnection() {
  return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'Connected' })).setMimeType(ContentService.MimeType.JSON);
}