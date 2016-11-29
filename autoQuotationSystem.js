function main() {
    // Folder and Files ID setting
    var numFormID = "1dq_bzZQbfWUc8kOEFiJXqWw_kbXouXiQYhYmThmMkqw";
    var quptationTemplateID = "1_ze9v8gXjmUZyB4bBgb3i3b5fOA899O9U3wtwWjmXhc";
    var quotationFolderID = "0B3ZXQ_--j_DYRm9YMkxYM24ycFk";
    
    // Open this application form
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("Application");
    } catch(error) {
        logWriter("Open [Acer U-Zham HOS Application Form] spreadsheet failed(" + error + ")");
        Logger.log("Open this this spreadsheet failed(" + error + ")");
    }
    // Get new data's values
    try {
        var pastNumRange = SpreadsheetApp.openById(numFormID).getSheetByName("Num").getRange(1, 1);
        var pastNum = pastNumRange.getValue()
        var currentNum = sheet.getLastRow();
    } catch(error) {
        logWriter("Open [For Count Order Number] failed(" + error + ")");
        Logger.log("Open [For Count Order Number] failed(" + error + ")");
    }
    // Checking Is there new one
    var gap = currentNum - pastNum;
    if (gap > 0) {
        // Generate quotations
        for(var i = 1; i <= gap; i++) {
            var quotationID = pastNum + i - 1;
            var companyName = sheet.getRange(quotationID + 1, 1).getValue();
            var seatNum = sheet.getRange(quotationID + 1, 6).getValue();

            newQuotation(quotationID, companyName, seatNum, quptationTemplateID, quotationFolderID);
            logWriter("Quotation for " + companyName + " complete");
            Logger.log("Quotation for " + companyName + " complete");
        }
        pastNumRange.setValue(currentNum);
    } else {
        logWriter("System checked there is no new quotation");
    }
}

function newQuotation(qid, cname, seat, qtid, qfid) {
    // Copy & Create
    try {
        // Copy new Quotation from template
        var qt = DriveApp.getFileById(qtid);
        var qt_new = qt.makeCopy("[TN-AUZHOS-" + qid + "] Quotation of " + cname + " for Acer U-Zham HDE Online Storage", DriveApp.getFolderById(qfid));
        // Open this new Quotation
        var qt_new_ss = SpreadsheetApp.openById(qt_new.getId());
        var qt_new_sheet = qt_new_ss.getSheetByName("Quotation");
    } catch(error) {
        logWriter("Copy and create new Quotation failed(" + error + ")");
        Logger.log("Copy and create new Quotation failed(" + error + ")");
    }
    logWriter("New quotation for " + cname + " copied & created");
    
    // Write data
    try {
      // Write the data to the new Quotation
      date = new Date();
      qt_new_sheet.getRange("I4:J5").setValue(": TN-AUZHOS - " + qid);
      qt_new_sheet.getRange("F46:F47").setValue(seat);
      qt_new_sheet.getRange("I30:J31").setValue(cname);
      qt_new_sheet.getRange("I7:J8").setValue(Utilities.formatDate(date, "GMT+8", ": yyyy年MM月dd日"));
      qt_new_sheet.getRange("C24:E25").setValue(Utilities.formatDate(new Date(date.getTime() + 30 * (24*3600*1000)), "GMT+8", "yyyy年MM月dd日"));
      // For saving
      qt_new_sheet.getRange("B64:J65").setValue("");
    } catch(error) {
        logWriter("Update new Quotation's data failed(" + error + ")");
        Logger.log("Update new Quotation's data failed(" + error + ")");
    }
    logWriter("New quotation for " + cname + "'s data updated");
    
    // Create PDF file
    try {
      var pdf = qt_new_ss.getAs('application/pdf');
      var pdfFile = DriveApp.getFolderById(qfid).createFile(pdf);
      var pdfFileId = pdfFile.getId();
    } catch(error) {
        logWriter("Create new Quotation's pdf file failed(" + error + ")");
        Logger.log("Create new Quotation's pdf file failed(" + error + ")");
    }
    logWriter("Quotation for " + cname + "'s data updated");
    
    //Send Alert Mail
    sentAlertMail(cname, pdfFileId);
}

function sentAlertMail(cname, pdfId) {
    var attachment = DriveApp.getFileById(pdfId);
    var recipient = "mark.lee@acer.com";
    var cc = "tw-event@hde.co.jp";
    var subject = "Quotation for " + cname + " Created Alert Mail";
    var body = "Quotation for " + cname + " created, please check it ! ";
    var options = {
        attachments: [attachment],
        cc: cc
    };
    try {
        MailApp.sendEmail(recipient, subject, body, options);
    } catch(error) {
        logWriter("Sending mail occured some problem(" + error + ")");
        Logger.log("Sending mail occured some problem(" + error + ")");
    }
    logWriter(cname + "'s quotation mail sent.");
    Logger.log(cname + "'s quotation mail sent.");
}

function logWriter(message) {
    var logSSId = "1TesmX85M8rgPdAgqXsX14WNKH0LU70wySd3zIwgHAyU";
     // Open log file
    try {
        var logSS = SpreadsheetApp.openById(logSSId).getSheetByName("Log");
        lastRow = logSS.getLastRow();
        logSS.getRange(lastRow + 1, 1).setValue(Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss"));
        logSS.getRange(lastRow + 1, 2).setValue(message);
    } catch(error) {
        Logger.log("Open this System Log spreadsheet failed(" + error + ")");
    }
}


