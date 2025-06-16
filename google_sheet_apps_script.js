function doPost(e) {
  try {
    var sheetId = "1wJxcdocXqXTs-nVoqM06MN0wwb8wI-kHUxoVq_W_1dE"; // เปลี่ยน Spreadsheet ID
    var folderId = "1lGvOpyyWhrzQ2g20yCgIIxSZiPty93Jm"; // เปลี่ยน Google Drive Folder ID
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName("QC_Data");

    if (!sheet) {
      sheet = SpreadsheetApp.openById(sheetId).insertSheet("QC_Data");
    }

    var data = JSON.parse(e.postData.contents);
    var lastRow = sheet.getLastRow() + 1;

    // บันทึกค่าตัวเลข (Fat Percentage) ลงในเซลล์โดยตรง
    var newRow = [
      data.id_manufacturer,
      data.id_qc,
      data.lot_no,
      data.date_value,
      data.target_fat,
      data.avg_fat,
      data.fat_diff,
      data.status,
      parseFloat(data.img_a1_fat) || 0, parseFloat(data.img_a2_fat) || 0, 
      parseFloat(data.img_b1_fat) || 0, parseFloat(data.img_b2_fat) || 0,
      parseFloat(data.img_c1_fat) || 0, parseFloat(data.img_c2_fat) || 0,
      parseFloat(data.img_d1_fat) || 0, parseFloat(data.img_d2_fat) || 0
    ];

    sheet.appendRow(newRow);
    
    var imageKeys = ["img_a1", "img_a2", "img_b1", "img_b2", "img_c1", "img_c2", "img_d1", "img_d2"];
    var imageCols = [9, 10, 11, 12, 13, 14, 15, 16]; // คอลัมน์ที่ต้องแทรกรูป

    var folder = DriveApp.getFolderById(folderId);

    for (var i = 0; i < imageKeys.length; i++) {
      if (data[imageKeys[i]]) {
        var decodedData = Utilities.base64Decode(data[imageKeys[i]]);
        var blob = Utilities.newBlob(decodedData, "image/png", imageKeys[i] + ".png");
        var file = folder.createFile(blob);

        // แทรกรูปภาพในเซลล์ให้ตรงกับคอลัมน์ที่กำหนด
        var img = sheet.insertImage(file.getBlob(), imageCols[i], lastRow);
        img.setWidth(100);
        img.setHeight(100);
      }
    }

    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.message).setMimeType(ContentService.MimeType.TEXT);
  }
}
