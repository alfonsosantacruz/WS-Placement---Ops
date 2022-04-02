var fieldPositions = {
  studentEmail: 2,
  studentFullName: 3,
  citizenship: 4,
  wsEligible: 6,
  onRotation: 7,
  company: 8,
  role: 9,
  manager: 10,
  salaryRate: 11,
  projectCode: 12,
  ctdSignOff: 13,
  paycomID: 14,
  paycomDeptCode: 17,
  managerEmail: 18,
  avgHours: 21,
  positionUpdateBool: 22
};

var actionsByColumn = {
  2: "StudentEmail",
  3: "StudentFullName",
  4: "Citizenship",
  6: "Eligibility",
  7: "Rotation",
  8: "Company",
  9: "Role",
  10: "Manager",
  11: "Salary",
  12: "Project Code",
  13: "CTD Sign Off",
  14: "PaycomID",
  17: "Dept Code",
  18: "Manager Email",
  21: "Average Hours",
  22: "Position Update Bool",
  101: "Doc_Unsigned_Offer_Link",
  102: "PDF_Unsigned_Offer_Link",
  103: "HelloSign_Offer_UUID",
  104: "Signed_Offer_Link",
  105: "Onboarding_Email_Sent"
};

var currentDashboardSheetName = "Spring Dashboard";
var previousDashboardSheetName = "Dashboard";
var updatesSheetName = "Updates";
var projectCodesSheetName = "PaycomProjectCodes";
var checklistSheetName = "Offers Checklist";
var managersSheetName = "PaycomManagersByEmail";
var paycomAnalysisSheetName = "PaycomYTDAnalysis";
var formConfigSheetName = "Form Config";

var listOfCCedAdmins = [
].join();

// Google Doc ID for file containing email template
var positionUpdateNotificationFileId = "1RmlUzPa-wDnDr-D5E3M9jbkaBmpxKixDL8mfbFd1UYE";
// Google Form for Hours Submission/Approval System - See Contractors Approval System directory for more information
var submissionFormURL = "https://docs.google.com/forms/d/1XWNo07nvTxrdzxnmSpkQeINrd0g1hypLMHI1WC4SeYw/edit?edit_requested=true";

function processUpdateEmails() {
  var form = FormApp.openByUrl(submissionFormURL);
  
  // If hours submission for this PP is closed
  if (form.isAcceptingResponses() == false) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var currentDashboardSheet = sheet.getSheetByName(currentDashboardSheetName);
    var previousDashboardSheet = sheet.getSheetByName(previousDashboardSheetName);
    var startRow = 2;
    var currentDashboardDataRange = currentDashboardSheet.getRange(startRow, 1, currentDashboardSheet.getLastRow(), currentDashboardSheet.getLastColumn());
    var previousDashboardDataRange = previousDashboardSheet.getRange(startRow, 1, previousDashboardSheet.getLastRow(), previousDashboardSheet.getLastColumn());
    // Fetch values for each row in the Range.
    var currentDashboardData = currentDashboardDataRange.getValues();
    var previousDashboardData = previousDashboardDataRange.getValues();

    for (var i = 0; i < currentDashboardData.length; i++) {
      var positionUpdate = currentDashboardData[i][fieldPositions.positionUpdateBool - 1];
      if (positionUpdate) {
        var studentEmail = currentDashboardData[i][fieldPositions.studentEmail - 1],
            studentName = currentDashboardData[i][fieldPositions.studentFullName - 1],
            currentManagerName = currentDashboardData[i][fieldPositions.manager - 1],
            currentManagerEmail = currentDashboardData[i][fieldPositions.managerEmail - 1],
            currentPositionName = currentDashboardData[i][fieldPositions.role - 1],
            previousManagerName = previousDashboardData[i][fieldPositions.manager - 1],
            previousManagerEmail = previousDashboardData[i][fieldPositions.managerEmail - 1],
            previousPositionName = previousDashboardData[i][fieldPositions.role - 1];
        
        var message = getGoogleDocumentAsHTML(positionUpdateNotificationFileId);
        var subject = DocumentApp.openById(positionUpdateNotificationFileId).getName();

        message = message.replace("###StudentName###", studentName);
        message = message.replace("###NewRole###", currentPositionName);
        message = message.replace("###NewManager###", currentManagerName);
        message = message.replace("###PreviousRole###", previousPositionName);
        message = message.replace("###PreviousManager###", previousManagerName);

        var recipients = [studentEmail]

        if (currentManagerEmail) {
          recipients.push(currentManagerEmail);
        }

        if (previousManagerEmail && currentManagerEmail != previousManagerEmail) {
          recipients.push(previousManagerEmail);
        }

        var recipients = recipients.join();

        if (currentManagerEmail || previousManagerEmail) {
          sendIndividualEmail(recipients, listOfCCedAdmins, subject, message);
          updateTracker(currentDashboardSheet, startRow + i, fieldPositions.ctdSignOff, "Official");
          copyAndPasteHardValues(currentDashboardSheetName, previousDashboardSheetName, i + startRow);
        }
      }
    }
  }
}

function getHardCopy(copyFromSheetName, rowNum) {
  var og_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copyFromSheetName);
  var og_values = og_sheet.getRange(rowNum, 1, og_sheet.getLastRow(), og_sheet.getLastColumn()).getValues();

  return og_values
}

function putHardCopy(values, copyToSheetName, rowNum) {
  var copy_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copyToSheetName);
  copy_sheet.getRange(rowNum, 1, values.length, values[0].length).setValues(values);
}

function copyAndPasteHardValues(copyFromSheetName, copyFromSheetName, rowNum) {
  var values = getHardCopy(copyFromSheetName, rowNum);
  putHardCopy(values, copyToSheetName, rowNum);
}

function getGoogleDocumentAsHTML(fileId){
  var forDriveScope = DriveApp.getStorageUsed(); //needed to get Drive Scope requested
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+fileId+"&exportFormat=html";
  var param = {
    method      : "get",
    headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions:true,
  };
  var html = UrlFetchApp.fetch(url,param).getContentText();
  
  return html
}

function sendIndividualEmail(recipients, cc, subject, message) {
  MailApp.sendEmail(recipients, subject, message, {
    cc: cc,
    htmlBody: message,
    name: "Student Payroll",
    noReply: true,
    replyTo: "studentpayroll@minerva.edu",
  });
}

function updateTracker(sheet, rowNum, colNum, val) {
  sheet.getRange(rowNum, colNum).setValue(val);
}

