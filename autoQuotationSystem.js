function main() {
    // Folder and Files setting
    var numFormID = "1dq_bzZQbfWUc8kOEFiJXqWw_kbXouXiQYhYmThmMkqw";
    var quptationTemplateID = "1_ze9v8gXjmUZyB4bBgb3i3b5fOA899O9U3wtwWjmXhc";
    var quotationFolderID = "0B3ZXQ_--j_DYRm9YMkxYM24ycFk";
    // Open this application form
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName("Application");
    } catch {
        Logger.log("Open this this spreadsheet failed");
    }
    // Get new data's values
    try {
        var pastNumRange = SpreadsheetApp.openById(numFormID).getSheetByName("Num").getRange(1, 1);
        var pastNum = pastNumRange.getValue()
        var currentNum = sheet.getLastRow();
    } catch {
        Logger.log("Open this numForm failed");
    }
    // Checking Is there new one
    var gap = currentNum - pastNum;
    
    // Generate quotations
    for(var i = 1; i <= gap; i++) {
        var quotationID = pastNum + i - 1;
        var companyName = sheet.getRange(quotationID + 1, 1).getValue();
        var seatNum = sheet.getRange(quotationID + 1, 6).getValue();
        try {
            newQuotation(quotationID, companyName, seatNum, quptationTemplateID, quotationFolderID);
        } catch {
            Logger.log("Generate new Quotation Failed");
        }
        Logger.log(quotationID+ " - Quotation for " + companyName);
    }
    pastNumRange.setValue(currentNum);
}

function newQuotation(qid, cname, seat, qtid, qfid) {
    // Copy new Quotation from template
    var qt = DriveApp.getFileById(qtid);
    var qt_new = qt.makeCopy("[TN-AUZHOS-" + qid + "] Quotation of " + cname + " for Acer U-Zham HDE Online Storage", DriveApp.getFolderById(qfid));
    
    // Open this new Quotation
    var qt_new_ss = SpreadsheetApp.openById(qt_new.getId());
    var qt_new_sheet = qt_new_ss.getSheetByName("Quotation");
    
    // Write the data to the new Quotation
    date = new Date();
    qt_new_sheet.getRange("I4:J5").setValue(": TN-AUZHOS - " + qid);
    qt_new_sheet.getRange("F46:F47").setValue(seat);
    qt_new_sheet.getRange("I30:J31").setValue(cname);
    qt_new_sheet.getRange("I7:J8").setValue(Utilities.formatDate(date, "GMT+8", ": yyyy年MM月dd日"));
    qt_new_sheet.getRange("C24:E25").setValue(Utilities.formatDate(new Date(date.getTime() + 30 * (24*3600*1000)), "GMT+8", "yyyy年MM月dd日"));
    // For saving
    qt_new_sheet.getRange("B64:J65").setValue("");
    
    // Create PDF file
    var pdf = qt_new_ss.getAs('application/pdf');
    var pdfFile = DriveApp.getFolderById(qfid).createFile(pdf);
    var pdfFileId = pdfFile.getId();
    
    //Send Alert Mail
    sentAlertMail(cname, pdfFileId);
}

function sentAlertMail(cname, pdfId) {
    var attachment = DriveApp.getFileById(pdfId);
    var recipient = "seisho.jo@g.hde.co.jp";
    var cc = "seisyo1234@gmail.com";
    var subject = "Quotation Created Alert Mail";
    var body = "Quotation for " + cname + " created, please check it ! ";
    var options = {
        attachments: [attachment],
        cc: cc
    };
    try {
        MailApp.sendEmail(recipient, subject, body, options);
    } catch {
        Logger.log("Sending mail occured some problem.");
    }
    Logger.log(cname + "mail sent.");
}


