function doGet() { 
  return HtmlService.createTemplateFromFile('INDEX').evaluate()
         .setTitle('New Art. FC Report (Dark Mode)')
         .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getProductData(scannedValue) {
  try {
    if (!scannedValue) return { error: 'กรุณาระบุ Article หรือ Barcode' };
    var masterSs = SpreadsheetApp.openById("1z7q3IGnSszYpzDNATa0ppE4jTgfq_5Msy4QM-GOkf78");
    var dbSheet = masterSs.getSheetByName('QC IMP');
    if (!dbSheet) return { error: 'ไม่พบ Sheet QC IMP' };
    
    var data = dbSheet.getDataRange().getValues();
    var searchId = String(scannedValue).trim();
    
    for (var i = 1; i < data.length; i++) {
      var barcode = String(data[i][9] || '').trim();
      var article = String(data[i][5] || '').trim();
      
      if (barcode === searchId || article === searchId) {
        var rawDate = new Date(data[i][0]);
        var formattedDate = Utilities.formatDate(rawDate, "GMT+7", "dd-MM-yyyy"); // แก้เดือนเป็น MM แล้ว
        
        return {
          success: true,
          date: formattedDate,
          vendor: data[i][1] || '',
          poNumber: data[i][2] || '',
          article: data[i][5] || '',
          description: data[i][6] || '',
          poQty: data[i][7] || 0,
          sampling: data[i][8] || 0
        };
      }
    }
    return { error: 'ไม่พบข้อมูลสินค้า Article/Barcode: ' + searchId };
  } catch (e) { return { error: e.toString() }; }
}

function processForm(formData) {
  try {
    var templateSs = SpreadsheetApp.getActiveSpreadsheet(); // นี่คือไฟล์ Master
    var mainFolder = DriveApp.getFolderById("1qNJtUa1y7pF0i7Mj-R4mwPyFJTlTvXs2");
    var folderName = formData.date + " - " + formData.poNumber;
    
    // สร้างหรือค้นหา Folder
    var folders = mainFolder.getFoldersByName(folderName);
    var targetFolder = folders.hasNext() ? folders.next() : mainFolder.createFolder(folderName);
    
    // สร้างหรือค้นหาไฟล์ Report สำหรับ PO นี้
    var fileName = "New Art. FC - PO_" + formData.poNumber;
    var files = targetFolder.getFilesByName(fileName);
    var reportSs;
    
    if (files.hasNext()) {
      reportSs = SpreadsheetApp.open(files.next()); // ถ้ามีไฟล์ PO เดิม ให้เปิดขึ้นมา
    } else {
      var newFile = DriveApp.getFileById(templateSs.getId()).makeCopy(fileName, targetFolder);
      reportSs = SpreadsheetApp.open(newFile); // ถ้ายังไม่มีไฟล์ PO นี้ ให้ Copy ใหม่
    }

    var sheetName = formData.article.toString();
    
    // ลบ Sheet ข้อมูล Article เก่าถ้ามี (กรณีสแกนแก้ข้อมูลเดิม)
    if (reportSs.getSheetByName(sheetName)) {
      reportSs.deleteSheet(reportSs.getSheetByName(sheetName));
    }
    
    // 🔥 วิธีที่ 2: ดึง Template จากไฟล์ Master ต้นฉบับ (templateSs) เท่านั้น! 🔥
    var masterTemplateSheet = templateSs.getSheetByName("Template");
    if (!masterTemplateSheet) {
      return { success: false, message: 'ไม่พบ Template Sheet ในไฟล์ Master' };
    }
    
    // คัดลอก Template จากไฟล์ Master ไปยังไฟล์ Report ปลายทาง
    var targetSheet = masterTemplateSheet.copyTo(reportSs).setName(sheetName);
    targetSheet.showSheet(); 

    // บันทึกข้อมูลข้อความ
    targetSheet.getRange('B2').setValue(formData.date);
    targetSheet.getRange('B3').setValue(formData.poNumber);
    targetSheet.getRange('B4').setValue(formData.vendor);
    targetSheet.getRange('B5').setValue(formData.article);
    targetSheet.getRange('B6').setValue(formData.description);
    targetSheet.getRange('B7').setValue(formData.poQty + " boxes");
    targetSheet.getRange('B8').setValue(formData.sampling + " boxes");
    
    // จัดรูปแบบวันที่ MFD
    var mfdParts = formData.mfgDate.split('-');
    var formattedMFD = mfdParts[2] + "-" + mfdParts[1] + "-" + mfdParts[0];
    targetSheet.getRange('B9').setValue(formattedMFD);
    
    // ลงผลตรวจและไฮไลท์
    targetSheet.getRange('B10').setValue(formData.inspectionResult).setBackground('#FFFF00');
    targetSheet.getRange('B11').setValue(formData.sampleStatus).setBackground('#FFFF00');

    // แทรกรูปภาพให้อยู่กึ่งกลางเซลล์
    if (formData.group1) insertImagesCentered(formData.group1, targetSheet, 14);
    if (formData.group2) insertImagesCentered(formData.group2, targetSheet, 15);
    if (formData.group3) insertImagesCentered(formData.group3, targetSheet, 16);

    // 🔥 ลบหน้า Template ขยะ ในไฟล์ Report ทิ้ง (เพื่อให้ไฟล์คลีนตลอดเวลา) 🔥
    var localTemplate = reportSs.getSheetByName("Template");
    if (localTemplate) {
      reportSs.deleteSheet(localTemplate);
    }
    
    return { 
      success: true, 
      message: 'บันทึกข้อมูลสำเร็จ!\nArticle: ' + formData.article + '\nPO: ' + formData.poNumber,
      fileUrl: reportSs.getUrl()
    };
  } catch (e) { 
    Logger.log('Error in processForm: ' + e.toString());
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + e.toString() }; 
  }
}

function insertImagesCentered(images, sheet, row) {
  var colStart = 3; // Column C
  var cellWidth = 201; // Width 28.00
  var cellHeight = 192; // Height 144.0
  var padding = 10;
  var maxImages = Math.min(images.length, 4);

  for (var i = 0; i < maxImages; i++) {
    try {
      if (images[i] && images[i].data) {
        var blob = Utilities.newBlob(Utilities.base64Decode(images[i].data), images[i].type, images[i].name);
        var image = sheet.insertImage(blob, colStart + i, row);
        
        var rawW = image.getWidth();
        var rawH = image.getHeight();
        var ratio = Math.min((cellWidth - padding * 2) / rawW, (cellHeight - padding * 2) / rawH);
        
        var newW = rawW * ratio;
        var newH = rawH * ratio;
        
        image.setWidth(newW).setHeight(newH);
        
        var offsetX = (cellWidth - newW) / 2;
        var offsetY = (cellHeight - newH) / 2;
        image.setAnchorCellXOffset(offsetX).setAnchorCellYOffset(offsetY);
      }
    } catch(e) { 
      Logger.log('Error inserting image ' + i + ': ' + e.toString());
    }
  }
}
