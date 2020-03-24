/**
 * 台帳データを更新する
 * レンタル番号に基づく台帳にレンタル終了日を入れます
 * @param targetData { [[string]] } シートのデータ
 */
function updateKintone(targetData) {
  var index = simplitCSVSheet.getIndex();
  var dataTitles = KintonePCData.pcDataSheet.getTitles();
  var putData = [];
  
  KintonePCData.pcDataSheet.getRentalData().forEach(function(pcValue) {
    // 台帳データを取得しまとめているシートから参照する
    var data = null;
    
    targetData.some(function(target, i) {
      if (target[index.yrlNo] === pcValue[dataTitles.rentalid.index].replace('-', '')) {
        data = target;
        targetData.splice(i, 1); // 同レンタルIDがあったらtargetDataから除外
        Logger.log('HIT! :' + target + ' : ' + targetData.length);
        return true;
      }
    });
    if (!data) return;
    var rendalDate = Utilities.formatDate(data[index.endDate], 'JST', 'yyyy-MM-dd');
    var editedDate = pcValue[dataTitles.rental_end.index] === '' ? '' : Utilities.formatDate(pcValue[dataTitles.rental_end.index], 'JST', 'yyyy-MM-dd');
    if ( editedDate === rendalDate &&
         Number(pcValue[dataTitles.rental_fee.index]) === Number(data[index.money])
    ) return;
    
    Logger.log('putKintone: レコードID = ' + data[dataTitles[KintoneApi.KEY_ID].index]);
    
    putData.push({
      id: targetValue[dataTitles[KintoneApi.KEY_ID].index],
      record: {
        rental_end       : { value: rendalDate },
        rental_contractid: { value: data[index.agreementNo] },
        rental_fee       : { value: data[index.money] }
      }
    });
  });
  if (putData.length === 0) return;
  Logger.log('put records :' + putData.length);
  
  // 契約番号と、レンタル終了日を更新
  KintoneApi.caApi.api.putRecords(putData);
}


function testK() {
  updateKintone(simplitCSVSheet.getZenshisuData());
  return;
  
  //putKintone(simplitCSVSheet.getZenshisuData()); // 台帳の自動更新を走らせる
  //var data = simplitCSVSheet.values.slice(simplitCSVSheet.titleRow + 1); // タイトル以下のデータ
  //var data = [simplitCSVSheet.values[236]];
  //putKintone(data);
}
