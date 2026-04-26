function doGet() { 
  // ต้องมีคำว่า return ตรงนี้นะครับ
  return HtmlService.createTemplateFromFile('Index').evaluate()
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
        var formattedDate = Utilities.formatDate(rawDate, "GMT+7", "dd-MM-yyyy"); 
        
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
    var templateSs = SpreadsheetApp.getActiveSpreadsheet(); 
    var mainFolder = DriveApp.getFolderById("1qNJtUa1y7pF0i7Mj-R4mwPyFJTlTvXs2");
    var folderName = formData.date + " - " + formData.poNumber;
    
    var folders = mainFolder.getFoldersByName(folderName);
    var targetFolder = folders.hasNext() ? folders.next() : mainFolder.createFolder(folderName);
    
    var fileName = "New Art. FC - PO_" + formData.poNumber;
    var files = targetFolder.getFilesByName(fileName);
    var reportSs;
    
    if (files.hasNext()) {
      reportSs = SpreadsheetApp.open(files.next()); 
    } else {
      var newFile = DriveApp.getFileById(templateSs.getId()).makeCopy(fileName, targetFolder);
      reportSs = SpreadsheetApp.open(newFile); 
    }

    var sheetName = formData.article.toString();
    
    var masterTemplateSheet = templateSs.getSheetByName("Template");
    if (!masterTemplateSheet) {
      return { success: false, message: 'ไม่พบ Template Sheet ในไฟล์ Master' };
    }
    
    // 🛠️ จุดที่แก้ไข: สลับลำดับการวาง Sheet เพื่อป้องกัน Error ลบ Sheet จนหมด 🛠️
    // 1. คัดลอก Template มาค้ำไว้ในไฟล์ปลายทางก่อน (ระบบจะตั้งชื่อชั่วคราวเช่น "Copy of Template")
    var targetSheet = masterTemplateSheet.copyTo(reportSs);
    
    // 2. เช็คว่ามีข้อมูล Article เดิมอยู่ไหม ถ้ามีให้ลบทิ้ง (ตอนนี้ลบได้แล้วเพราะมี targetSheet หน้าใหม่ค้ำไว้)
    var existingSheet = reportSs.getSheetByName(sheetName);
    if (existingSheet) {
      reportSs.deleteSheet(existingSheet);
    }
    
    // 3. ค่อยเปลี่ยนชื่อ Sheet ใหม่ให้เป็นเลข Article
    targetSheet.setName(sheetName);
    targetSheet.showSheet(); 
    // -------------------------------------------------------------

    targetSheet.getRange('B2').setValue(formData.date);
    targetSheet.getRange('B3').setValue(formData.poNumber);
    targetSheet.getRange('B4').setValue(formData.vendor);
    targetSheet.getRange('B5').setValue(formData.article);
    targetSheet.getRange('B6').setValue(formData.description);
    targetSheet.getRange('B7').setValue(formData.poQty + " boxes");
    targetSheet.getRange('B8').setValue(formData.sampling + " boxes");
    
    var mfdParts = formData.mfgDate.split('-');
    var formattedMFD = mfdParts[2] + "-" + mfdParts[1] + "-" + mfdParts[0];
    targetSheet.getRange('B9').setValue(formattedMFD);
    
    targetSheet.getRange('B10').setValue(formData.inspectionResult).setBackground('#FFFF00');
    targetSheet.getRange('B11').setValue(formData.sampleStatus).setBackground('#FFFF00');

    if (formData.group1) insertImagesCentered(formData.group1, targetSheet, 14);
    if (formData.group2) insertImagesCentered(formData.group2, targetSheet, 15);
    if (formData.group3) insertImagesCentered(formData.group3, targetSheet, 16);

    // ลบ Template หน้าเปล่าที่ติดมาตอนเริ่มสร้างไฟล์ (ถ้ามี)
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
  var colStart = 3; // เริ่มที่คอลัมน์ C
  var cellWidth = 170; // ความกว้างช่อง (อ้างอิงจาก Template ของคุณ)
  var cellHeight = 145; // ความสูงช่อง
  var padding = 3; // 💡 เผื่อขอบไว้ 10px ป้องกัน Excel ดึงภาพเพี้ยน
  
  // จำนวนรูปรวมสูงสุดต่อแถว
  var maxImages = Math.min(images.length, 10);

  for (var i = 0; i < maxImages; i++) {
    try {
      if (images[i] && images[i].data) {
        var blob = Utilities.newBlob(
          Utilities.base64Decode(images[i].data), 
          images[i].type, 
          images[i].name
        );
        var image = sheet.insertImage(blob, colStart + i, row);
        
        // 1. อ่านขนาดจริงของรูปต้นฉบับ
        var rawW = image.getWidth();
        var rawH = image.getHeight();
        
        // 2. คำนวณอัตราส่วนเพื่อย่อรูป (รูปจะไม่เบี้ยวหน้าอ้วน/หน้าตอบ)
        var ratio = Math.min((cellWidth - padding * 2) / rawW, (cellHeight - padding * 2) / rawH);
        
        var newW = rawW * ratio;
        var newH = rawH * ratio;
        
        // 3. กำหนดขนาดใหม่ให้รูปภาพ
        image.setWidth(newW).setHeight(newH);
        
        // 4. คำนวณเพื่อจัดกึ่งกลางเซลล์พอดีเป๊ะ
        var offsetX = (cellWidth - newW) / 2;
        var offsetY = (cellHeight - newH) / 2;
        image.setAnchorCellXOffset(offsetX).setAnchorCellYOffset(offsetY);
      }
    } catch(e) { 
      Logger.log('Error inserting image ' + i + ': ' + e.toString());
    }
  }
}
