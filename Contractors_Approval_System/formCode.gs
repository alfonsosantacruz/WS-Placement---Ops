// File ID to Position Tracker Sheet. See Position Tracker Directory for more.
var trackerSheetDBId = "******************************";
var sheetConfigName = "Config";

function updateForm() {
  var ss = SpreadsheetApp.openById(trackerSheetDBId);
  var sheet = ss.getSheetByName(sheetConfigName);
  var data = sheet.getDataRange().getValues(); // Data for pre-fill
  var formUrl = ss.getFormUrl();
  var form = FormApp.openByUrl(formUrl);
  var items = form.getItems();
  
  var formItem_0 = items[0].asSectionHeaderItem();
  var formItem_1 = items[1].asTextItem(); //Item for Week 1
  var formItem_2 = items[2].asTextItem(); //Item for Week 2
  
  formItem_0.setTitle(data[5][1]);
  formItem_1.setTitle(data[6][1]);
  formItem_2.setTitle(data[7][1]);
  
}

function restrict(){
  var form = FormApp.openByUrl("https://docs.google.com/forms/d/1XWNo07nvTxrdzxnmSpkQeINrd0g1hypLMHI1WC4SeYw/edit?edit_requested=true")
  if (form.isAcceptingResponses() == true){
    form.setAcceptingResponses(false)
    form.setCustomClosedFormMessage("Dear Intern, We are sorry to let you know that the Pay Period is currently closed. The form will re-open next Saturday and you can fill the missing hours section on the next Pay Period.");    
} else {
    form.setAcceptingResponses(true) }
}
