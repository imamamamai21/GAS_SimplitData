var csvData = simplitCSVSheet.values.splice(1);

/**
 * 全シスデータをpickupして台帳に登録・更新する
 */
function updateZenshisuData() {
  var startTime = new Date();
  var time = Utilities.formatDate(startTime, 'JST', 'yyMMddHHmm');
  
  // タイムアウトの壁を超えるための対応もする
  var properties = PropertiesService.getScriptProperties();  // 途中経過保存用
  var startRowKey = 'startRow-' + time;  // 何行目まで処理したかを保存するときに使用するkey
  var triggerKey = 'trigger-' + time;    // トリガーIDを保存するときに使用するkey
  
  //途中から実行した場合、ここに何行目まで実行したかが入る
  var startRow = parseInt(properties.getProperty(startRowKey));
  if (!startRow) startRow = 1;
  
  var index = simplitCSVSheet.getIndex();
  
  var dataTitles = KintonePCData.pcDataSheet.getTitles();
  var kintoneData = KintonePCData.pcDataSheet.getRentalData();
  var putData = [];
  
  // 台帳データを取得しまとめているシートから参照する
  kintoneData.forEach(function(pcValue, i) {
    var diff = parseInt((new Date() - startTime) / (1000 * 60));
    if (diff >= 5) { // 5分経過していたら処理を中断
      properties.setProperty(startRowKey, i);  // 何行まで処理したかを保存
      setTrigger(triggerKey, 'updateData-' + time);  // トリガーを発行
      Logger.log('@set Trigger');
      return;
    }
    var data = [];
    var hitIndex = null;
    var rentalId = pcValue[dataTitles.rentalid.index];
    if (rentalId && rentalId.length === 8) rentalId = rentalId.replace('-', '');
    csvData.some(function(target, j) {
      if (target[index.yrlNo] == rentalId) {
        data = target;
        hitIndex = j;
        Logger.log('HIT! :' + j + ' : csvData.length = ' + csvData.length);
        return true;
      }
    });
    
    if (data.length === 0) return;
    
    csvData.splice(hitIndex, 1); // 同レンタルIDがあったらcsvDataから除外
    var rendalDate = Utilities.formatDate(data[index.endDate], 'JST', 'yyyy-MM-dd');
    var editedDate = pcValue[dataTitles.rental_end.index] === '' ? '' : Utilities.formatDate(pcValue[dataTitles.rental_end.index], 'JST', 'yyyy-MM-dd');
    if (editedDate === rendalDate && Number(pcValue[dataTitles.rental_fee.index]) === Number(data[index.money])) return;
    
    Logger.log('putKintone: レコードID = ' + pcValue[dataTitles[KintoneApi.KEY_ID].index]);
    
    putData.push({
      id: pcValue[dataTitles[KintoneApi.KEY_ID].index],
      record: {
        rental_end       : { value: rendalDate },
        rental_contractid: { value: data[index.agreementNo] },
        rental_fee       : { value: data[index.money] }
      }
    });
  });
  Logger.log('@putData length = ' + putData.length);
  if (putData.length === 0) return;
  Logger.log('@put records :' + putData.length);
  
  // 契約番号と、レンタル終了日を更新
  KintoneApi.caApi.api.putRecords(putData);
  
  // 全て実行終えたらトリガーと何行目まで実行したかを削除する
  deleteTrigger(triggerKey);
  properties.deleteProperty(startRowKey);
}

/**
 * トリガーを発行
 */
function setTrigger(triggerKey, funcName) {
  deleteTrigger(triggerKey); // 保存しているトリガーがあったら削除
  var date = new Date();
  date.setMinutes(dt.getMinutes() + 1);  // １分後に再実行
  var triggerId = ScriptApp.newTrigger(funcName).timeBased().at(date).create().getUniqueId();
  // あとでトリガーを削除するためにトリガーIDを保存しておく
  PropertiesService.getScriptProperties().setProperty(triggerKey, triggerId);
}

/**
 * 指定したkeyに保存されているトリガーIDを使って、トリガーを削除する
 */
function deleteTrigger(triggerKey) {
  Logger.log('@delete trigger');
  var triggerId = PropertiesService.getScriptProperties().getProperty(triggerKey);
  
  if(!triggerId) return;
  
  ScriptApp.getProjectTriggers().filter(function(trigger){
    return trigger.getUniqueId() == triggerId;
  }).forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  PropertiesService.getScriptProperties().deleteProperty(triggerKey);
}
