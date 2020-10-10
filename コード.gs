function 日付変換() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("出席");

  menu.addItem("日付変換","日付変換");
  menu.addItem("出席集計","出席集計");
  menu.addToUi();
  
  var ss = SpreadsheetApp.getActive();
  var sht = ss.getSheetByName('シート1');
  var lastRow = sht.getLastRow();
  var data = sht.getRange(2, 1, lastRow - 1, 4).getValues();
  for(var i=0;i<lastRow - 1;i++){
    var depDate = data[i][0];
    var depDateStrings = Utilities.formatDate(new Date(depDate), "Asia/Tokyo" , "YYYY/MM/dd");
    sht.getRange(i + 2, 4).setValue(depDateStrings);
  }
}

function 出席集計() {
  var ss = SpreadsheetApp.getActive();
  var sht = ss.getSheetByName('シート1');
  var lastRow = ss.getLastRow();
  ss.getRange('A1').activate();
  var data = sht.getRange(1, 1, lastRow, 4);

  ss.insertSheet(ss.getActiveSheet().getIndex() + 1).activate();
  ss.getActiveSheet().setHiddenGridlines(true);
  var pivotTable = ss.getRange('A1').createPivotTable(data);
  pivotTable = ss.getRange('A1').createPivotTable(data);
  var pivotGroup = pivotTable.addRowGroup(2);
  pivotTable = ss.getRange('A1').createPivotTable(data);
  pivotGroup = pivotTable.addRowGroup(2);
  pivotGroup = pivotTable.addColumnGroup(4);
  pivotTable = ss.getRange('A1').createPivotTable(data);
  var pivotValue = pivotTable.addPivotValue(4, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotGroup = pivotTable.addRowGroup(2);
  pivotGroup = pivotTable.addColumnGroup(4);
};



