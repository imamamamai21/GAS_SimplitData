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

/**
 * 別スプシへデータを書き出す
 * 内定者インターン・レンタル: https://docs.google.com/spreadsheets/d/1ZgQoZFQUHmglUv14QRplSI-iCpBmU6zun-X7AJMK5oY/edit#gid=78777832
 * ライブラリ RentalPc： https://script.google.com/d/M734HhKXBXrZHXEzrZjZBtv2vJ94xa-wl/edit?mid=ACjPJvF31u2fW0cZ_va5Gdxc13jvynMxx9KBDkdLq54UBm9QfXFJTuRFDjmvwlIdhKQjhO5w2ID9a9C7DwUybXB5mZgP0W3MCmfV2E3g7rXjRNJdTKA9E3dnrPmuymWSEwvbvhEsAYQHuuU&uiv=2 
 */
function updateInternlistSinc() {
  var targetSheet = new RentalPc.SimplitSheet();
  targetSheet.sheet.getRange(2, 2, targetSheet.values.length - 1, targetSheet.values[0].length -1).clearContent(); // 対象スプレッドシートの対象シートのデータを削除
  
  var csvData = simplitCSVSheet.values
    .slice(simplitCSVSheet.titleRow + 1)
    .map(function(value) { value.splice(simplitCSVSheet.getIndex().timeStamp, 1); return value; }); // タイトル以下 タイムスタンプを抜いたデータ
 
  targetSheet.sheet
    .getRange(2, 2, Number(simplitCSVSheet.getLastRow()) - 1, Number(simplitCSVSheet.getLastColumn()) - 1)
    .setValues(csvData);//CSVデータシートから取得した内容を対象シートにセット
}

