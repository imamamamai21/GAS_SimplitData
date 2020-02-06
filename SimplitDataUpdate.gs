var RESULT_TEXT__SUCCESS = '成功';
var RESULT_TEXT__ERROR = '成功';

/**
 * CSVファイルをインポートしてデータを更新する
 */
function UpdateData() {
  var today = new Date();
  var week = today.getDay();
  // 土日を避ける
  if (week === 0 && week === 6) return;
  
  try {
    var Simplit_Data_Upfolder = DriveApp.getFolderById('1JOnTxUsT8rxRYd8ougVkSEwpWt_OB_9J');//SimplitDataのCSVファイルアップロードフォルダオブジェクトの取得
    var Simplit_Data_Upfiles = Simplit_Data_Upfolder.getFilesByType(MimeType.CSV);//CSVファイルオブジェクトをイテレータで取得
    var Simplit_Data_UpfileObject = Simplit_Data_Upfiles.next()//CSVファイルオブジェクトを取得
    var fileName = Simplit_Data_UpfileObject.getName();//CSVファイル名を取得
    var timeStamp = fileName.slice('レンタル契約台帳_'.length, fileName.length - ('.csv'.length));
    
    // 適当な位置のタイムスタンプを確認
    var oltStamp = simplitCSVSheet.sheet.getRange('A2').getValue();
    if (timeStamp === oltStamp) return RESULT_TEXT__ERROR;
    
    //CSVファイルをShiftJISで開く
    var Simplit_Data_Upfile = Simplit_Data_UpfileObject.getBlob().getDataAsString('Shift_JIS');
    
    //CSVファイルの内容をオブジェクトとして取得
    var Simplit_Data = Utilities.parseCsv(Simplit_Data_Upfile);
  
    //ダウンロードしたCSVファイルの項目数が不足する場合は、ポップアップ・メール・エラーを出して終了
    if (Simplit_Data[0].length < 45){
      Browser.msgBox('Simplit Managerからは全項目ダウンロードしてください');
      MailApp.sendEmail('infosys_sup@cyberagent.co.jp', 'Simplitデータが不正です', '資産管理チーム担当者各位\n\nお疲れ様です！\n\nアップロードされたSimplitデータの項目数が不足しています。\nSimplit Managerからは全項目をダウンロードしてください。');
    }
    createStoreData(Simplit_Data);
    updateCSVData(Simplit_Data, timeStamp);
    RentalReturnTaskSheet.createNewEndPc();
    postEndData(today.getDate());
    return RESULT_TEXT__SUCCESS;
  } catch (e) {
    return RESULT_TEXT__ERROR;
  }
}

/**
 * フォルダに入ったcsvデータを使って、CSVシートを作成する
 * @param simplitData { object }
 * @param timeStamp { string } タイムスタンプ
 */
function updateCSVData(simplitData, timeStamp) {
  //スプレッドシートにCSVファイルの内容を書き込む
  simplitCSVSheet.sheet.getRange(1, 1, simplitCSVSheet.getLastRow(), simplitCSVSheet.getLastColumn()).clearContent();
  simplitCSVSheet.sheet.getRange(1, 2, simplitData.length, simplitData[0].length).setValues(simplitData);
  simplitCSVSheet.sheet.getRange(2, 1, simplitData.length -1, 1).setValue(timeStamp);
  simplitCSVSheet.sheet.getRange('A1').setValue(TITLE__TIME_STAMP);
}

/**
 * csvシートのデータを使って、Storedシートを作成する
 * 別のファイルにデータを送って書き込み命令を出す
 *
 * @param newData {object} csvのデータ
 * @param timeStamp {object} アップデート前のタイムスタンプ
 */
function createStoreData(newData) {
  var index = simplitCSVSheet.getIndex();
  
  // データの差分を見て、なくなったデータを保存
  var disappearedData = simplitCSVSheet.values.slice(1).filter(function(value, count) {
    for (var i = 1; i < newData.length; i++) {
      if (Number(newData[i][index.yrlNo - 1]) === Number(value[index.yrlNo]) && 
        Number(newData[i][index.agreementNo - 1]) === Number(value[index.agreementNo])) {
        return false;
      }
    }
    return true;
  });
  // Stored用のシートに書き込み命令を送る
  var len = disappearedData.length;
  if (disappearedData.length > 0) StoredSheet.editDeletedRentalData(disappearedData);
  
  /*TODO以下を消す
  var csvData = simplitCSVSheet.values.slice(simplitCSVSheet.titleRow + 1); // タイトル以下のデータ
  var storedData = simplitCSVSheet.values.slice(simplitCSVSheet.titleRow + 1); // タイトル以下のデータ
  var index = simplitCSVSheet.getIndex();
  var filterdData = [].concat(csvData);
  
  storedData.forEach(function(value) {
    for(var i = 0; i < csvData.length; i++ ) {
      var data = csvData[i];
      if( // データが同じものは抜く
        data[index.agreementNo] === value[index.agreementNo] && 
        data[index.agreementIndex] === value[index.agreementIndex] && 
        data[index.endDate] === value[index.endDate] && 
        data[index.code] === value[index.code] && 
        data[index.serialNo] === value[index.serialNo]
      ) {
        filterdData.splice(i, 1);
        return;
      }
    }
  });
  if (filterdData.length === 0) return;
  
  simplitStoredSheet.sheet.getRange(simplitStoredSheet.getLastRow() + 1, 1, filterdData.length, filterdData[0].length).setValues(filterdData);
  
  // 契約番号、行番号、タイムスタンプでソート
  simplitStoredSheet.sheet
    .getRange(index.agreementNo + 1, 1, simplitStoredSheet.getLastRow() - 1, simplitStoredSheet.getLastColumn())
    .sort([{column: index.agreementNo + 1, ascending: true},{column: index.agreementIndex + 1, ascending: false}, {column: index.timeStamp + 1, ascending: false}]);
  */
}
