function doGet() { 
  return HtmlService.createTemplateFromFile('INDEX').evaluate()
         .setTitle('New Art. FC Report')
         .addMetaTag('viewport', 'width=device-width, initial-scale=1'); 
}

function getProductData(scannedValue) {
  try {
    // Validation
    if (!scannedValue || scannedValue.toString().trim() === '') {
      return { error: 'กรุณาระบุ Article หรือ Barcode' };
    }

    var masterSs = SpreadsheetApp.openById("1z7q3IGnSszYpzDNATa0ppE4jTgfq_5Msy4QM-GOkf78");
    var dbSheet = masterSs.getSheetByName('QC IMP');
    
    if (!dbSheet) {
      return { error: 'ไม่พบ Sheet QC IMP' };
    }
    
    var data = dbSheet.getDataRange().getValues();
    var searchId = String(scannedValue).trim();

    // ค้นหาข้อมูล
    for (var i = 1; i < data.length; i++) {
      var barcode = String(data[i][9] || '').trim();
      var article = String(data[i][5] || '').trim();
      
      if (barcode === searchId || article === searchId) {
        // ตรวจสอบข้อมูลที่จำเป็น
        if (!data[i][5] || !data[i][2]) {
          return { error: 'ข้อมูลไม่ครบถ้วน: ขาด Article หรือ PO Number' };
        }
        
        var rawDate = new Date(data[i][0]);
        var formattedDate = Utilities.formatDate(rawDate, "GMT+7", "yyyy-MM-dd");
        
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
    
  } catch (e) { 
    Logger.log('Error in getProductData: ' + e.toString());
    return { error: 'เกิดข้อผิดพลาด: ' + e.toString() }; 
  }
}

function processForm(formData) {
  try {
    // Validation
    if (!formData.article || !formData.poNumber) {
      return { success: false, message: 'ข้อมูลไม่ครบถ้วน กรุณาตรวจสอบ Article และ PO Number' };
    }
    
    if (!formData.mfgDate || formData.mfgDate.trim() === '') {
      return { success: false, message: 'กรุณาระบุวันผลิต (MFD)' };
    }
    
    // ตรวจสอบรูปภาพอย่างน้อย 1 กลุ่ม
    if (!formData.group1 && !formData.group2 && !formData.group3) {
      return { success: false, message: 'กรุณาแนบรูปภาพอย่างน้อย 1 กลุ่ม' };
    }
    
    var templateSs = SpreadsheetApp.getActiveSpreadsheet();
    var mainFolder = DriveApp.getFolderById("1qNJtUa1y7pF0i7Mj-R4mwPyFJTlTvXs2");
    var folderName = formData.date + " - " + formData.poNumber;
    
    // สร้างหรือค้นหา Folder
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
    
    // ลบ Sheet เก่าถ้ามี
    if (reportSs.getSheetByName(sheetName)) {
      reportSs.deleteSheet(reportSs.getSheetByName(sheetName));
    }
    
    var templateSheet = reportSs.getSheetByName("Template");
    if (!templateSheet) {
      return { success: false, message: 'ไม่พบ Template Sheet' };
    }
    
    var targetSheet = templateSheet.copyTo(reportSs).setName(sheetName);
    targetSheet.showSheet(); 

    // บันทึกข้อมูล
    targetSheet.getRange('B2').setValue(formData.date);
    targetSheet.getRange('B3').setValue(formData.poNumber);
    targetSheet.getRange('B4').setValue(formData.vendor);
    targetSheet.getRange('B5').setValue(formData.article);
    targetSheet.getRange('B6').setValue(formData.description);
    targetSheet.getRange('B7').setValue(formData.poQty);
    targetSheet.getRange('B8').setValue(formData.sampling);
    targetSheet.getRange('B9').setValue(formData.mfgDate);
    targetSheet.getRange('B10').setValue(formData.inspectionResult).setBackground('#FFFF00');
    targetSheet.getRange('B11').setValue(formData.sampleStatus);

    // แทรกรูปภาพ
    if (formData.group1 && formData.group1.length > 0) {
      insertImagesCentered(formData.group1, targetSheet, 14);
    }
    if (formData.group2 && formData.group2.length > 0) {
      insertImagesCentered(formData.group2, targetSheet, 15);
    }
    if (formData.group3 && formData.group3.length > 0) {
      insertImagesCentered(formData.group3, targetSheet, 16);
    }

    // ซ่อน Template
    if (reportSs.getSheetByName("Template")) {
      reportSs.getSheetByName("Template").hideSheet();
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
  var colStart = 3; 
  var imgWidth = 150;
  var imgHeight = 150;
  var offsetX = 7;
  var offsetY = 8;
  
  // จำกัดจำนวนรูปสูงสุด 10 รูปต่อแถว
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
        image.setWidth(imgWidth).setHeight(imgHeight);
        image.setAnchorCellXOffset(offsetX).setAnchorCellYOffset(offsetY);
      }
    } catch(e) { 
      Logger.log('Error inserting image ' + i + ': ' + e.toString());
    }
  }
}
