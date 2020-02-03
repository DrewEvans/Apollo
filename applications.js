const destinationDocument = SpreadsheetApp.openByUrl(
    'https://docs.google.com/spreadsheets/d/1JVk5QK_UM7H9NPBQ2GdZ2e3ma45Jnn2Efax_79DoZIg/edit#gid=0'
);
const activeDocument = SpreadsheetApp.getActiveSpreadsheet();

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Applications')
        .addItem('Capture Data', 'dataCapture')
        .addItem('Create New Tab', 'createNewTab')
        .addItem('Insert DA-File Data','pushData')
        .addToUi();
}

function dataCapture() {

    // get current sheet and tabs after send data has been selected
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dealSelection = ss.getSheetByName('SELECTED DEALS');
    var database = ss.getSheetByName('deal_log');

    // count rows of data to capture
    var current_rows = dealSelection.getLastRow();
    var database_rows = database.getLastRow() + 1;
    var database_rows_new = current_rows + database_rows - 5;
    var rows_new = dealSelection.getRange("A5:V" + current_rows).getValues();

    // Database range that is being captured
    database.getRange("A" + database_rows + ":V" + database_rows_new).setValues(rows_new);
}

function createNewTab() {
    var activeSheet = activeDocument.getActiveSheet();
    var countryCode = activeSheet.getRange('B1').getValue();
    var sendDate = activeSheet.getRange('B2').getValue();
    var sendNumber = activeSheet.getRange('B3').getValue();
    var mobileListNumber = activeSheet.getRange('B4').getValue();

    destinationDocument.insertSheet(countryCode + '_' + sendDate + '_S' + sendNumber + '_ML' + mobileListNumber);
}

function pushData() {
    var activeSheet = activeDocument.getActiveSheet();
    var countryCode = activeSheet.getRange('B1').getValue();
    var sendDate = activeSheet.getRange('B2').getValue();
    var sendNumber = activeSheet.getRange('B3').getValue();
    var mobileListNumber = activeSheet.getRange('B4').getValue();
    var sourceData = activeSheet.getRange("A6:Y40").getValues();
    var destinationTab = destinationDocument.getSheetByName(
        countryCode + '_' + sendDate + '_S' + sendNumber + '_ML' + mobileListNumber);

    Logger.log(sourceData);

    if (destinationTab) {
        destinationTab.getRange("A1:Y35").setValues(sourceData);
    }
}
