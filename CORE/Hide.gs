function nighthide() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('240:748').activate();
  spreadsheet.getActiveSheet().hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
};

function dayhide() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('784:1454').activate();
  spreadsheet.getActiveSheet().hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
};

function eveninghide() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('1490:1896').activate();
  spreadsheet.getActiveSheet().hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
};
