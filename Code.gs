function doGet() { 
  return HtmlService.createTemplateFromFile('Index').evaluate()
         .setTitle('New Art. FC Report (Dark Mode)')
         .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getProductData(scannedValue) {
  try {
    if (!scannedValue) return { error: 'กรุณาระบุ Article หรือ Barcode' };
    var masterSs = SpreadsheetApp.openById("1z7q3IGnSszYpzDNATa0ppE4jTgfq_5Msy4QM-GOkf78");
    var dbSheet = masterSs.getSheetByName('QC IMP');
    var data = dbSheet.getDataRange().getValues();
    var searchId = String(scannedValue).trim();
    
    for (var i = 1; i < data.length; i++) {
      var barcode = String(data[i][9] || '').trim();
      var article = String(data[i][5] || '').trim();
      
      if (barcode === searchId || article === searchId) {
        var rawDate = new Date(data[i][0]);
        // Format เป็น dd-mm-yyyy สำหรับแสดงผล
        var formattedDate = Utilities.formatDate(rawDate, "GMT+7", "dd-mm-yyyy");
        
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
    return { error: 'ไม่พบข้อมูลสินค้า: ' + searchId };
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
    if (reportSs.getSheetByName(sheetName)) {
      reportSs.deleteSheet(reportSs.getSheetByName(sheetName));
    }
    
    var templateSheet = reportSs.getSheetByName("Template");
    var targetSheet = templateSheet.copyTo(reportSs).setName(sheetName);
    targetSheet.showSheet(); 

    // บันทึกข้อมูลพร้อมใส่หน่วย " boxes"
    targetSheet.getRange('B2').setValue(formData.date);
    targetSheet.getRange('B3').setValue(formData.poNumber);
    targetSheet.getRange('B4').setValue(formData.vendor);
    targetSheet.getRange('B5').setValue(formData.article);
    targetSheet.getRange('B6').setValue(formData.description);
    targetSheet.getRange('B7').setValue(formData.poQty + " boxes");
    targetSheet.getRange('B8').setValue(formData.sampling + " boxes");
    
    // แปลงวันที่ MFD (yyyy-mm-dd) เป็น dd-mm-yyyy
    var mfdParts = formData.mfgDate.split('-');
    var formattedMFD = mfdParts[2] + "-" + mfdParts[1] + "-" + mfdParts[0];
    targetSheet.getRange('B9').setValue(formattedMFD);
    
    targetSheet.getRange('B10').setValue(formData.inspectionResult).setBackground('#FFFF00');
    targetSheet.getRange('B11').setValue(formData.sampleStatus).setBackground('#FFFF00'); // Highlight สีเหลือง

    // แทรกรูปภาพ (Row 14, 15, 16)
    if (formData.group1) insertImagesCentered(formData.group1, targetSheet, 14);
    if (formData.group2) insertImagesCentered(formData.group2, targetSheet, 15);
    if (formData.group3) insertImagesCentered(formData.group3, targetSheet, 16);

    // ลบ Template Sheet ออกก่อนส่งไฟล์
    if (reportSs.getSheetByName("Template")) {
      reportSs.deleteSheet(reportSs.getSheetByName("Template"));
    }
    
    return { success: true, message: 'บันทึกสำเร็จ!', fileUrl: reportSs.getUrl() };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function insertImagesCentered(images, sheet, row) {
  var colStart = 3; // Column C
  var cellWidth = 201; // Width 28.00 (201 pixels)
  var cellHeight = 192; // Height 144.0 (192 pixels)
  var padding = 10;
  
  var maxImages = Math.min(images.length, 4); // C D E F มี 4 คอลัมน์

  for (var i = 0; i < maxImages; i++) {
    try {
      if (images[i] && images[i].data) {
        var blob = Utilities.newBlob(Utilities.base64Decode(images[i].data), images[i].type, images[i].name);
        var image = sheet.insertImage(blob, colStart + i, row);
        
        // คำนวณขนาดรูปให้ใหญ่ที่สุดแต่ไม่เกินขอบเซลล์
        var rawW = image.getWidth();
        var rawH = image.getHeight();
        var ratio = Math.min((cellWidth - padding * 2) / rawW, (cellHeight - padding * 2) / rawH);
        
        var newW = rawW * ratio;
        var newH = rawH * ratio;
        
        image.setWidth(newW).setHeight(newH);
        
        // จัดกึ่งกลางด้วย Offset
        var offsetX = (cellWidth - newW) / 2;
        var offsetY = (cellHeight - newH) / 2;
        image.setAnchorCellXOffset(offsetX).setAnchorCellYOffset(offsetY);
      }
    } catch(e) { Logger.log(e.toString()); }
  }
}
