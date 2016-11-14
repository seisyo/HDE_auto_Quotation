function main() {
  // Open this spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Application");
  
  // Get new data's values
  var pastNumRange = SpreadsheetApp.openById("1dq_bzZQbfWUc8kOEFiJXqWw_kbXouXiQYhYmThmMkqw").getSheetByName("Num").getRange(1, 1);
  var pastNum = pastNumRange.getValue()
  var currentNum = sheet.getLastRow();
  
  // Checking Is there new one
  var gap = currentNum - pastNum;
  
  for(var i = 1; i <= gap; i++) {
    var quotationID = pastNum + i - 1;
    var companyName = sheet.getRange(quotationID + 1, 1).getValue();
    var seatNum = sheet.getRange(quotationID + 1, 6).getValue();
    newQuotation(quotationID, companyName, seatNum);
    Logger.log(quotationID+ " - Quotation for " + companyName);
  }
  pastNumRange.setValue(currentNum);
}

function newQuotation(qid, name, num) {
  // Copy new Quotation from template
  var qt = DriveApp.getFileById("1_ze9v8gXjmUZyB4bBgb3i3b5fOA899O9U3wtwWjmXhc");
  var qt_new = qt.makeCopy("[TN-AUZHOS-" + qid + "] Quotation of " + name + " for Acer U-Zham HDE Online Storage", DriveApp.getFolderById("0B3ZXQ_--j_DYRm9YMkxYM24ycFk"));
  
  // Open this new Quotation
  var qt_new_ss = SpreadsheetApp.openById(qt_new.getId());
  var qt_new_sheet = qt_new_ss.getSheetByName("Quotation");
  
  // Write the data to the new Quotation
  date = new Date();
  qt_new_sheet.getRange("I4:J5").setValue(": TN-AUZHOS - " + qid);
  qt_new_sheet.getRange("F46:F47").setValue(num);
  qt_new_sheet.getRange("I30:J31").setValue(name);
  qt_new_sheet.getRange("I7:J8").setValue(Utilities.formatDate(date, "GMT+8", ": yyyy年MM月dd日"));
  qt_new_sheet.getRange("C24:E25").setValue(Utilities.formatDate(new Date(date.getTime() + 30 * (24*3600*1000)), "GMT+8", "yyyy年MM月dd日"));
  qt_new_sheet.getRange("B64:J65").setValue("");
  
  // Create PDF file
  var pdf = qt_new_ss.getAs('application/pdf');
  var pdfFile = DriveApp.getFolderById('0B3ZXQ_--j_DYRm9YMkxYM24ycFk').createFile(pdf);
  var pdfFileId = pdfFile.getId();
  
  //Send Alert Mail
  sentAlertMail(name, pdfFileId);
}

function sentAlertMail(companyName, pdfId) {
  var attachment = DriveApp.getFileById(pdfId);
  var recipient = "tingyu.chen@g.hde.co.jp";
  var subject = "Quotation Created Alert Mail";
  var body = "Quotation for " + companyName + " created, please check it ! ";
  var options = {
    attachments: [attachment],
    cc: "seisho.jo@g.hde.co.jp"
  };
  MailApp.sendEmail(recipient, subject, body, options);
  Logger.log(companyName + "mail sent.");
}


