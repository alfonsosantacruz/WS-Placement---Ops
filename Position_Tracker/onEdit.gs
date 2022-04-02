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

function onEdit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  if (sheet.getActiveSheet().getName() == currentDashboardSheetName) {
    // Gets row number of the cell that was modified
    var updatedRow = e.range.getRow(),
        oldValue = e.oldValue;
  
    // Run function to check for the entire sheet row
    var newValues = e.range.getValues()[0],
        firstUpdatedColumn = e.range.getColumn();

    for (var i = 0; i < newValues.length; i++) {
      var modifiedColumn = i + firstUpdatedColumn,
          newValue = newValues[i];
      
      processUpdate(sheet, updatedRow, modifiedColumn, newValue, oldValue);
    };
  };
};

function processUpdate(sheet, updatedRow, modifiedColumn, newValue, oldValue) {
  // Checks whether the modified cell corresponds to a relevant field.
  // Returns the position as an integer to check whether the modified field matters from our fieldPositions object
  var modifiedField = Object.values(fieldPositions).find(position => {
    return position == modifiedColumn;
  });
  // If a relevant field was modified, proceeds to log the update
  if (!!modifiedField) {   

    handleProjectCodesOnCompanyChange(sheet, updatedRow, modifiedColumn, newValue, oldValue);

    // If there was a change in the value, posts the update in the logs.
    postUpdatedStudentInfoAsUpdate(sheet, updatedRow, modifiedField, newValue);
  }
}

function handleProjectCodesOnCompanyChange(sheet, updatedRow, modifiedColumn, newValue, oldValue) {
  if (modifiedColumn == fieldPositions.company) {
    var dashboardSheet = sheet.getSheetByName(currentDashboardSheetName);
    var impactedProjectCodeCell = dashboardSheet.getRange(updatedRow, fieldPositions.projectCode);

    if (newValue == 'MU') {
      impactedProjectCodeCell.setDataValidation(null);
      impactedProjectCodeCell.setValue("Unassigned");
      postUpdatedStudentInfoAsUpdate(sheet, updatedRow, fieldPositions.projectCode, "Unassigned");
    } else if (newValue == 'MP') {
      var projectCodesSheet = sheet.getSheetByName(projectCodesSheetName);
      var projectCodesList = projectCodesSheet.getRange(2, 2, projectCodesSheet.getLastRow());
      var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(projectCodesList).build();
      
      impactedProjectCodeCell.setValue("");
      impactedProjectCodeCell.setDataValidation(validationRule);
    } else if (!newValue && !!oldValue) {
      impactedProjectCodeCell.setDataValidation(null);
      impactedProjectCodeCell.setValue("");
    }
  };
};

function postUpdatedStudentInfoAsUpdate(sheet, updatedRow, modifiedField, newValue) {
  if (!!newValue) {
    var updatesSheet = sheet.getSheetByName(updatesSheetName);
    var { studentEmail, studentFullName, paycomID } = getStudentInfoForUpdate(sheet, updatedRow);
    var newRowNum = updatesSheet.getLastRow() + 1
    var actionNum = Object.keys(actionsByColumn).find(code => {
      return modifiedField == code
    });
    var actionCode = actionsByColumn[actionNum];
    
    updatesSheet.getRange(newRowNum, 1).setValue(new Date());
    updatesSheet.getRange(newRowNum, 2).setValue(studentEmail);
    updatesSheet.getRange(newRowNum, 3).setValue(studentFullName);
    updatesSheet.getRange(newRowNum, 4).setValue(paycomID);
    updatesSheet.getRange(newRowNum, 5).setValue(actionCode);
    updatesSheet.getRange(newRowNum, 6).setValue(newValue);
  };
};

function getStudentInfoForUpdate(sheet, updatedRow) {
  var dashboardSheet = sheet.getSheetByName(currentDashboardSheetName);
  // Gets first occurrence since the data is returned as a list of lists with only one elemnt (one sheet row)
  var studentInfo = dashboardSheet.getRange(updatedRow, 1, 1, dashboardSheet.getLastColumn()).getValues()[0];
  var studentEmail = studentInfo[1],
      studentFullName = studentInfo[2],
      paycomID = studentInfo[fieldPositions.paycomID - 1];

  return { studentEmail, studentFullName, paycomID }
};
