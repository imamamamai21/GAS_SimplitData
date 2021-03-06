/// このシートの情報をもとに、別ファイルを作成or別ファイルへアップデートするコード ///

/**
 * Latestスプシの書き出し
 * Storeから、最新のデータのみを抜いたデータ
 * 作られるフォルダ▶https://drive.google.com/drive/u/1/folders/1Vq_KVfKhDEUzkEsRMoSBSy_OE6FnKjaC
 
function makeLatestData() {
  // 新しいシートを対象フォルダに作る
  var timeStamp = Utilities.formatDate(new Date(), 'JST', 'yy-MM/dd_hh:mm');
  var newFile = SpreadsheetApp.create('LatestSimplit_' + timeStamp);
  var folderId = '1Vq_KVfKhDEUzkEsRMoSBSy_OE6FnKjaC';
  var folder = DriveApp.getFolderById(folderId);
  folder.addFile(DriveApp.getFileById(newFile.getId()));
  
  // Storedシートの内容をLatestシートにコピー
  var storeData = simplitStoredSheet.values;
  var newSheet = newFile.getSheets()[0];
  newSheet.getRange(1, 1, storeData.length, storeData[0].length).setValues(storeData);
  newSheet.setName('Latest');
  
  var index = simplitStoredSheet.getIndex();
  newSheet // Storedシートの内容を契約番号とYRL管理番号で重複削除
    .getRange(2, 1, newSheet.getLastRow() - 1, newSheet.getLastColumn())
    .removeDuplicates([index.agreementNo + 1, index.yrlNo]);
    
  // popupでお知らせ
  var htmlOutput = HtmlService
      .createHtmlOutput('<p><a href="' + SHEET_URL + newFile.getId() +  '" target="blank">Latestシート▶</a></p><p><a href="https://drive.google.com/drive/u/1/folders/' + folderId + '" target="blank">保存ディレクトリ▶</a></p>')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(150)
      .setHeight(150);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, '作成しました');
}*/
