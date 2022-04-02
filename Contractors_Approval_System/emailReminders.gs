// Google Docs File IDs for email templates for each scenario. See templates directory for more.
var managerApprovalReminderFileId = "************************************";
var studentSubmissionReminderFileId = "************************************";

var submissionFormURL = "https://docs.google.com/forms/d/1XWNo07nvTxrdzxnmSpkQeINrd0g1hypLMHI1WC4SeYw/edit?edit_requested=true";

var listOfCCedAdmins = [
  "asantacruz@uni.minerva.edu"
].join();

function onlyUnique(value, index, self) {
  return self.indexOf(value) === index;
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

function sendBulkEmail(recipients, subject, message) {
  MailApp.sendEmail(recipients, subject, message, {
    cc: listOfCCedAdmins,
    htmlBody: message,
    name: "Student Payroll",
    noReply: true,
    replyTo: "studentpayroll@minerva.edu",
  });
}

function managerApprovalReminder() {
  var form = FormApp.openByUrl(submissionFormURL);
  
  if (form.isAcceptingResponses() == false) {

    var source = SpreadsheetApp.getActiveSpreadsheet();
    var activeSheet = source.getSheets()[0];
    var managersEmails = activeSheet.getRange(dataStartingRow, columnToNumberMap.managerEmail.colNum, activeSheet.getLastRow()).getValues().reduce((acc, val) => acc.concat(val), []);
    var uniqueManagersEmails = managersEmails.filter(onlyUnique);

    var activeSheet = source.getSheets()[0];
    activePP = activeSheet.getName();
      
    var message = getGoogleDocumentAsHTML(managerApprovalReminderFileId);
    var subject = DocumentApp.openById(managerApprovalReminderFileId).getName();
    subject = subject.replace("###PP###", activePP);

    sendBulkEmail(uniqueManagersEmails.join(), subject, message);
  }
}

function contractorsSubmissionReminder() {
  var form = FormApp.openByUrl(submissionFormURL);
  
  if (form.isAcceptingResponses() == true) {
    var source = SpreadsheetApp.getActiveSpreadsheet();
    var activeSheet = source.getSheets()[0];
    var studentsEmails = activeSheet.getRange(dataStartingRow, columnToNumberMap.studentEmail.colNum, activeSheet.getLastRow()).getValues().reduce((acc, val) => acc.concat(val), []);
    var uniqueStudentsEmails = studentsEmails.filter(onlyUnique);

    var activeSheet = source.getSheets()[0];
    activePP = activeSheet.getName();
      
    var message = getGoogleDocumentAsHTML(studentSubmissionReminderFileId);
    var subject = DocumentApp.openById(studentSubmissionReminderFileId).getName();
    subject = subject.replace("###PP###", activePP);

    sendBulkEmail(uniqueStudentsEmails.join(), subject, message);
  }
}
