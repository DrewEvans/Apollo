const destinationDocument = SpreadsheetApp.openByUrl(
  "https://docs.google.com/spreadsheets/d/1NrOXIXIzZjzVTkW5nZERC8B6r0HnqmMSo_LbxwfR_Tk/edit#gid=0"
);

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Applications")
    .addItem("Capture Data", "dataCapture")
    .addToUi();
}

function dataCapture() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dealSelection = ss.getSheetByName("SELECTED DEALS");
  var database = ss.getSheetByName("deal_log");

  var current_rows = dealSelection.getLastRow();
  var database_rows = database.getLastRow() + 1;
  var database_rows_new = current_rows + database_rows - 5;
  var rows_new = dealSelection.getRange("A5:V" + current_rows).getValues();

  database
    .getRange("A" + database_rows + ":V" + database_rows_new)
    .setValues(rows_new);
}

function createNewTab() {
  var activeSheet = activeDocument.getSheetByName("Apollo");
  var sendDate = activeSheet.getRange("B2").getValue();
  var sendNumber = activeSheet.getRange("B3").getValue();
  var mobileListNumber = activeSheet.getRange("B4").getValue();

  destinationDocument.insertSheet(
    sendDate + " Send " + sendNumber + " ML " + mobileListNumber
  );
}

function pushData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getSheetByName("Apollo");
  var sendDate = activeSheet.getRange("B1").getValue();
  var sendNumber = activeSheet.getRange("B2").getValue();
  var mobileListNumber = activeSheet.getRange("B3").getValue();
  var sourceData = activeSheet.getRange("A8:K42").getValues();
  var destinationTab = destinationDocument.getSheetByName(
    sendDate + " Send " + sendNumber + " ML " + mobileListNumber
  );

  if (destinationTab) {
    destinationTab.getRange("A1:K35").setValues(sourceData);
  }
}
