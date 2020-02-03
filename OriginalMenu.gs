function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Simplit');
  menu.addItem('Simplitインポート', 'UpdateData');
  //menu.addItem('契約状況をまとめたファイルを作成する', 'makeLatestData');
  menu.addToUi();  
}