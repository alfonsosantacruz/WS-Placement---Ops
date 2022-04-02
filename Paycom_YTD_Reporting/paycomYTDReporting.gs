var reportDataStartRow = 2; // First row of data to process
var reportDataStartColumn = 1; // First column of data to process

var listOfCCedAdmins = [
  // "asantacruz@uni.minerva.edu",
  // "hello@minervaworkstudy.com",
].join();

var paycomManagersSheetName = "PaycomManagersByEmailFromTracker",
    paycomAnalysisSheetName = "PaycomYTDAnalysis",
    configSheetName = "Config";

// Google Docs Templates. See templates directory
var paycomYTDEveryoneManagerReportFileId = "************************************";
var paycomYTDAboveAverageManagerReportFileId = "************************************";
var paycomYTDStudentReportFileId = "************************************";
var paycomYTDEndOfSemesterManagerReportFileId = "************************************";

function getSheetDataBySheetName(sheetName) {
  // Imports data from managers sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var numRows = sheet.getLastRow(); // Number of rows to process
  var numColumns = sheet.getLastColumn(); // Numbers of columns to process
  var dataRange = sheet.getRange(reportDataStartRow, reportDataStartColumn, numRows, numColumns);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();

  return data
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

function run_sendReportToManagers() {
  sendReportToManagers(paycomManagersSheetName, paycomAnalysisSheetName);
}

function sendReportToManagers(managersSheetName, analysisSheetName) {
  // Imports data from managers sheet
  var managers_data = getSheetDataBySheetName(managersSheetName);

  // Imports the data from paycom analysis sheet
  var interns_data = getSheetDataBySheetName(analysisSheetName);

  //Import the data from the config sheet
  var config_data = getSheetDataBySheetName(configSheetName);

  var payPeriodNumber = config_data[0][1],
      payPeriodEndDate = config_data[1][1],
      expectedHours = config_data[2][1],
      reportTo = config_data[3][1];
  
  Logger.log(reportTo);

  if (reportTo == "Everyone") {
    var reportFileId = paycomYTDEveryoneManagerReportFileId;
  } else if (reportTo == "Above Average") {
    var reportFileId = paycomYTDAboveAverageManagerReportFileId;
  } else if (reportTo == "End of Semester") {
    var reportFileId = paycomYTDEndOfSemesterManagerReportFileId;
  }

  for (var i = 0; i < managers_data.length; i++) {
    var is_manager_active = managers_data[i][9];

    if (is_manager_active == true) {
      var message = getGoogleDocumentAsHTML(reportFileId);
      var [ tableFirstIndex, tableLastIndex, table ] = getHTMLTableFromGoogleDocHTMLText(message)

      var manager_email = managers_data[i][0],
          manager_name = managers_data[i][5],
          html_interns_table_content = "";

      var [ rowFirstIndex, rowLastIndex, html_interns_table_row ] = getHTMLContentRowFromGoogleDocHTMLTable(table);

      for (var j = 0; j < interns_data.length; j++) {
        var interns_manager_email = interns_data[j][6],
            interns_manager_name = interns_data[j][4],
            intern_average_status = interns_data[j][13];

        var average = interns_data[j][10];

        var reportingCondition = checkConditionalInternAverageStatus(reportTo, intern_average_status);

        if (manager_name == interns_manager_name && manager_email == interns_manager_email && reportingCondition && typeof average == "number") {

          // Add row to table from interns data
          var studentName = interns_data[j][2],
              studentEmail = interns_data[j][1],
              totalHours = interns_data[j][8];

          var html_interns_table_row_copy = html_interns_table_row;

          html_interns_table_row_copy = html_interns_table_row_copy.replace("###StudentName###", studentName);
          html_interns_table_row_copy = html_interns_table_row_copy.replace("###StudentEmail###", studentEmail);
          html_interns_table_row_copy = html_interns_table_row_copy.replace("###Average###", average);
          html_interns_table_row_copy = html_interns_table_row_copy.replace("###TotalHours###", totalHours);
          html_interns_table_row_copy = html_interns_table_row_copy.replace("###HoursAboveLimit###", totalHours - expectedHours);

          html_interns_table_content += html_interns_table_row_copy
        }
      }

      if (html_interns_table_content != "") {
        var new_html_table_with_content = insertStringInRange(table, html_interns_table_content, rowFirstIndex, rowLastIndex);
        var newMessage = insertStringInRange(message, new_html_table_with_content, tableFirstIndex, tableLastIndex);

        var subject = DocumentApp.openById(reportFileId).getName();

        newMessage = newMessage.replace("###ManagerName###", manager_name);
        newMessage = newMessage.replace("###PPNum###", payPeriodNumber);
        newMessage = newMessage.replace("###PPEndDate###", payPeriodEndDate);
        newMessage = newMessage.replace("###ExpectedHours###", expectedHours);
        
        subject = subject.replace("###PP###", payPeriodNumber);
        
        // Send the email for that specific manager
        sendIndividualEmail(manager_email, listOfCCedAdmins, subject, newMessage);
      }
    }
  }
}

function checkConditionalInternAverageStatus(reportTo, intern_average_status) {
  if (reportTo == "Everyone" || reportTo == "End of Semester") {
    return true
  } else if (reportTo == "Above Average") {
    return intern_average_status == true
  }
}

function getHTMLTableFromGoogleDocHTMLText(text) {
  var firstIndex = text.indexOf("<table");
  var lastIndex = text.indexOf("</table>") + "</table>".length

  var table = text.slice(firstIndex, lastIndex)

  return [ firstIndex, lastIndex, table ]
}

function getHTMLContentRowFromGoogleDocHTMLTable(table) {
  var firstIndex = table.indexOf("</tr><tr") + "</tr>".length;
  var lastIndex = table.indexOf("</td></tr></tbody></table>") + "</td></tr>".length;

  var row = table.slice(firstIndex, lastIndex);

  return [ firstIndex, lastIndex, row ]
}

function insertStringInRange(existingString, insertingString, firstIndex, lastIndex) {
  newStr = existingString.slice(0, firstIndex) + insertingString + existingString.slice(lastIndex);
  return newStr
}

function run_sendReportToInterns() {
  sendReportToInterns(paycomAnalysisSheetName)
}

function sendReportToInterns(internsSheetName) {
  // Imports the data from paycom analysis sheet
  var interns_data = getSheetDataBySheetName(internsSheetName);

  //Import the data from the config sheet
  var config_data = getSheetDataBySheetName(configSheetName);

  var payPeriodNumber = config_data[0][1],
      payPeriodEndDate = config_data[1][1],
      expectedHours = config_data[2][1];

  for (var j = 0; j < interns_data.length; j++) {

    var studentActiveStatus = interns_data[j][7],
        isStudentInReport = interns_data[j][8];

    if (studentActiveStatus == "Active" && isStudentInReport != "Not in Report") {
      var studentName = interns_data[j][2],
          studentEmail = interns_data[j][1],
          average = interns_data[j][10],
          totalHours = interns_data[j][8];
      
      var subject = DocumentApp.openById(paycomYTDStudentReportFileId).getName();
      subject = subject.replace("###PP###", payPeriodNumber);

      var message = getGoogleDocumentAsHTML(paycomYTDStudentReportFileId);
      message = message.replace("###StudentName###", studentName);
      message = message.replace("###Average###", average);
      message = message.replace("###TotalHours###", totalHours);
      message = message.replace("###PPNum###", payPeriodNumber);
      message = message.replace("###PPEndDate###", payPeriodEndDate);
      message = message.replace("###ExpectedHours###", expectedHours);

      sendIndividualEmail(studentEmail, "", subject, message);
    }
  }
}
