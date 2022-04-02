var columnToNumberMap = {
  studentEmail: {
    colNum: 2,
    submissionsColNum: null
  },
  week1: {
    colNum: 3,
    submissionsColNum: 4
  },
  week2: {
    colNum: 4,
    submissionsColNum: 5
  },
  pending: {
    colNum: 5,
    submissionsColNum: 6
  },
  sum: {
    colNum: 6,
    submissionsColNum: null
  },
  tasks: {
    colNum: 7,
    submissionsColNum: 7
  },
  managerName: {
    colNum: 8,
    submissionsColNum: null
  },
  managerApproval: {
    colNum: 9,
    submissionsColNum: null
  },
  managerComments: {
    colNum: 10,
    submissionsColNum: null
  },
  managerEmail: {
    colNum:11,
    submissionsColNum: null
  }
}

var dataStartingRow = 4,
    contentRangeInA1 = "A:W",
    sheetNameCellInA1 = "$W$1",
    configSheetName = "Config",
    currentPPCellInA1 = "$B$4"

// Folks that have edit access to closed pay periods sheets.
var absoluteEditors = [
  'asantacruz@uni.minerva.edu'
]

// fileURL for the position tracker
var positionTrackerFileURL = "https://docs.google.com/spreadsheets/d/****************************/edit?usp=sharing",
    checklistSheetNameInPositionTracker = "Offers Checklist",
    checklistSheetNameInApprovalSheet = "Checklist";
    

function getSheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function updateChecklistData() {
  values = getHardCopy(positionTrackerFileURL, checklistSheetNameInPositionTracker);
  putHardCopy(values, checklistSheetNameInApprovalSheet);

  sortByManagerName();
}

function getHardCopy(fileURL, copyFromSheetName) {
  var ss = SpreadsheetApp.openByUrl(fileURL);
  var og_sheet = ss.getSheetByName(copyFromSheetName);
  var og_values = og_sheet.getRange(1, 1, og_sheet.getLastRow(), og_sheet.getLastColumn()).getValues();

  return og_values
}

function putHardCopy(values, copyToSheetName) {
  var copy_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copyToSheetName);
  copy_sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}

function sortByManagerName() {
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = source.getSheets()[0];
  var activeFilter = activeSheet.getFilter();
  activeFilter.sort(columnToNumberMap.managerName.colNum, true);
}

function copySheet() {
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = source.getSheets()[0];
  var configSheet = source.getSheetByName(configSheetName);

  currentPP = configSheet.getRange(currentPPCellInA1).getValue();
  activePP = activeSheet.getName();
  
  if(currentPP != activePP) {

    // Creates a copy of the activeSheet and appends it at the end of the sheets list
    activeSheet.copyTo(source);

    var allSheets = source.getSheets();
    var copySheetIndex = allSheets.length - 1
    var copiedSheet = source.getSheets()[copySheetIndex];
    
    var requests = [{
      "clearBasicFilter": {
        "sheetId": copiedSheet.getSheetId()
      }
    }];
    Sheets.Spreadsheets.batchUpdate({'requests': requests}, source.getId());
  
    // Adds the current values in the activeSheet to the copiedSheet valuesOnly
    activeSheet.getRange(contentRangeInA1).copyTo(copiedSheet.getRange(contentRangeInA1), {contentsOnly:true});
    
    // Updates the name of the activeSheet
    activeSheet.setName(currentPP);
    // Changes the name of the copiedSheet to the previous name of the activeSheet
    copiedSheet.setName(activePP);
    
    // Locks the copiedSheet
    var protectionDescription = `${activePP} - Closed`
    lockPPSheet(copiedSheet, protectionDescription)
    
    // Clear manager approval section cells
    activeSheet.getRange(dataStartingRow, columnToNumberMap.managerApproval.colNum, activeSheet.getLastRow(), 2).clearContent();
    
    // Makes sure formulas are in place after making the copy, since managers could have made updates.
    refillFormulas(activeSheet);

    // Restates the new name as a constant
    activeSheet.getRange(sheetNameCellInA1).setFormula("=getSheetName()");
  }
}

function refillFormulas(activeSheet) {
  var submissionsRange = activeSheet.getRange(dataStartingRow, columnToNumberMap.week1.colNum, activeSheet.getLastRow()).getValues();
  for (var rowNum = dataStartingRow; rowNum <= submissionsRange.length; rowNum++) {
    for (var colNum = columnToNumberMap.week1.colNum; colNum <= columnToNumberMap.tasks.colNum; colNum++) {
      if (colNum == columnToNumberMap.sum.colNum) {
        var formula = getSumFormula(rowNum);
      } else if (colNum == columnToNumberMap.tasks.colNum) {
        var formula = getTasksFormula(rowNum);
      } else {
        var formula = getHoursFormula(rowNum, colNum);
      }
      activeSheet.getRange(rowNum, colNum).setFormula(formula);
    }
  }
}

function lockPPSheet(sheet, protectionDescription) {
  var protection = sheet.protect().setDescription(protectionDescription);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
  for (var i = 0; i < absoluteEditors.length; i++) {
    protection.addEditor(absoluteEditors[i])
  }
}

function getSumFormula(rowNum) {
  return `=if(or($C${rowNum} <> "", $D${rowNum} <> ""), sum($C${rowNum}:$E${rowNum}), "")`
}

function getTasksFormula(rowNum) {
  return `=iferror(vlookup(concatenate($B${rowNum}, " - ",${sheetNameCellInA1}),Submissions!$A$2:$G,${columnToNumberMap.tasks.submissionsColNum},False), "")`
}

function getHoursFormula(rowNum, colNum) {
  return `=if($B${rowNum} <> "",iferror(vlookup(concatenate($B${rowNum}, " - ",${sheetNameCellInA1}),Submissions!$A$2:$G,${colNum + 1},False), 0),"")`
}

function justificationValuesOnly() {
  var form = FormApp.openByUrl(submissionFormURL);
  
  if (form.isAcceptingResponses() == false) {

    var source = SpreadsheetApp.getActiveSpreadsheet();
    var activeSheet = source.getSheets()[0];
    
    var justificationsRange = activeSheet.getRange(dataStartingRow, columnToNumberMap.tasks.colNum, activeSheet.getLastRow(), 1);
    var justificationsValues = justificationsRange.getValues(); 
    justificationsRange.setValues(justificationsValues) , {contentsOnly:true};
  }
}


