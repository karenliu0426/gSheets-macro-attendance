/** @OnlyCurrentDoc */

function Test() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H1').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['No'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(8, criteria);
  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setFormula('=concat("Organizations attended YTD: ",countunique(Q:Q)-1)');
  spreadsheet.getRange('E1').activate();
  spreadsheet.getCurrentCell().setFormula('=CONCAT("Unique attendees YTD: ",countunique(L:L))');
  spreadsheet.getRange('E2').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Spotlight'), true);
  spreadsheet.getRange('H1').activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['', 'N/A', 'No'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(8, criteria);
  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getCurrentCell().setFormula('=CONCAT("Organizations attended YTD: ",countunique(Q:Q)-1)');
  spreadsheet.getRange('E1').activate();
  spreadsheet.getCurrentCell().setFormula('=CONCAT("Unique attendees YTD: ",countunique(L:L))');
  spreadsheet.getRange('E2').activate();
};
