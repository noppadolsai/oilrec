function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ระบบบันทึกและแดชบอร์ดเติมน้ำมัน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function saveRecord(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DataLog");
    
    // ตรวจสอบเรื่องสลิป: ถ้าไม่มีการอัปโหลดใหม่ ให้ใช้ลิงก์สลิปเดิมที่ส่งมา
    var slipUrl = data.existingSlipUrl || "";
    if (data.slipFile && data.slipFile.bytes) {
      var folderId = "ใส่_FOLDER_ID_ของกรุณาตรงนี้"; // ใส่ Folder ID ของคุณตรงนี้
      var folder = DriveApp.getFolderById(folderId);
      var contentType = data.slipFile.mimeType;
      var bytes = Utilities.base64Decode(data.slipFile.bytes);
      var blob = Utilities.newBlob(bytes, contentType, data.slipFile.name);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      slipUrl = file.getUrl();
    }
    
    var dateObj = new Date(data.date);
    var liters = parseFloat(data.liters);
    var pricePerUnit = parseFloat(data.pricePerUnit);
    var totalAmount = liters * pricePerUnit;
    
    var year = dateObj.getFullYear();
    var month = dateObj.getMonth() + 1;
    
    var oneJan = new Date(dateObj.getFullYear(), 0, 1);
    var numberOfDays = Math.floor((dateObj - oneJan) / (24 * 60 * 60 * 1000));
    var weekNum = Math.ceil((dateObj.getDay() + 1 + numberOfDays) / 7);
    
    // เตรียมชุดข้อมูลสำหรับหยอดลงแถว
    var rowData = [
      dateObj,          // A: วันที่
      weekNum,          // B: สัปดาห์ที่
      month,            // C: เดือน
      year,             // D: ปี
      data.fuelType,    // E: ประเภทน้ำมัน
      data.station,     // F: ปั๊มน้ำมัน
      liters,           // G: จำนวนลิตร
      pricePerUnit,     // H: ราคาต่อหน่วย
      totalAmount,      // I: จำนวนเงินรวม
      data.paymentMethod,// J: การชำระเงิน
      slipUrl,          // K: ลิงก์สลิป
      data.notes || ""  // L: หมายเหตุ
    ];
    
    // ตรวจสอบว่าเป็นโหมด "แก้ไขแถวเดิม" หรือ "เพิ่มแถวใหม่"
    var isUpdate = data.rowIndex ? true : false;
    
    if (isUpdate) {
      var rIndex = parseInt(data.rowIndex);
      // สั่งเขียนข้อมูลทับลงไปในพิกัดแถวเดิมที่มีการแก้ไข
      sheet.getRange(rIndex, 1, 1, 12).setValues([rowData]);
    } else {
      // โหมดปกติ: บันทึกต่อท้ายแถวใหม่
      sheet.appendRow(rowData);
      
      // ส่ง LINE แจ้งเตือน (เฉพาะตอนเพิ่มรายการใหม่เท่านั้น เพื่อไม่ให้สแปมตอนแก้ไข)
      if (typeof sendLineMessage === 'function') {
        sendLineMessage(data, totalAmount);
      }
    }
    
    return {
      status: "success", 
      message: isUpdate ? "✏️ อัปเดตแก้ไขข้อมูลในระบบสำเร็จแล้ว!" : "💾 บันทึกข้อมูลการเติมน้ำมันสำเร็จแล้ว!"
    };
  } catch(e) {
    return {status: "error", message: e.toString()};
  }
}

// ปรับปรุงการส่งข้อมูล: ให้ส่งค่าแถว (Row Index) และรูปแบบวันที่ ISO ติดไปด้วย
function getDashboardData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DataLog");
    var values = sheet.getDataRange().getValues();
    
    if (values.length <= 1) return [];
    values.shift(); // เอาหัวตารางออก
    
    var tz = Session.getScriptTimeZone();
    
    return values.map(function(row, index) {
      var rowIndex = index + 2; // คำนวณหาเลขแถวจริงใน Google Sheets (Index เริ่มที่ 0 + หัวตาราง 1 แถว = แถวที่ 2)
      var displayDate = "";
      var dateIso = "";
      if (row[0] instanceof Date) {
        displayDate = Utilities.formatDate(row[0], tz, "dd/MM/yyyy");
        dateIso = Utilities.formatDate(row[0], tz, "yyyy-MM-dd"); // ฟอร์แมตสำหรับใช้ป้อนเข้าช่องคีย์วันที่หน้าฟอร์ม
      }
      return {
        rowIndex: rowIndex,
        dateStr: displayDate,
        dateIso: dateIso,
        week: row[1],
        month: parseInt(row[2]) || 0,
        year: parseInt(row[3]) || 0,
        fuelType: row[4],
        station: row[5],
        liters: parseFloat(row[6]) || 0,
        pricePerUnit: parseFloat(row[7]) || 0,
        totalAmount: parseFloat(row[8]) || 0,
        paymentMethod: row[9],
        slipUrl: row[10],
        notes: row[11]
      };
    });
  } catch(e) {
    return [];
  }
}
