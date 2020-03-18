/**
 * 台帳データを更新する
 * レンタル番号に基づく台帳にレンタル終了日を入れます
 * @param tagetData { [[string]] } シートのデータ
 */
function updateKintone(tagetData) {
  var index = simplitCSVSheet.getIndex();
  var dataTitles = KintonePCData.pcDataSheet.getTitles();
  var pcData = KintonePCData.pcDataSheet.values.slice(2).filter(function(pcValue) {
     return (pcValue[dataTitles.rentalid.index] != '' &&
            pcValue[dataTitles.rental_status.index] != '終了' &&
            pcValue[dataTitles.pc_status.index] != '廃止' &&
            pcValue[dataTitles.pc_status.index] != '紛失' &&
            pcValue[dataTitles.pc_status.index] != '譲渡済' &&
            pcValue[dataTitles.pc_status.index] != '使用不可');
  });
  var putData = [];
  
  tagetData.forEach(function(value) {
    var myId = simplitCSVSheet.getRentalNo(value[index.yrlNo].toString());
    var rendalDate = Utilities.formatDate(value[index.endDate], 'JST', 'yyyy-MM-dd');
    
    // 台帳データを取得しまとめているシートから参照する
    var data = pcData.filter(function(pcValue) {
      var editedDate = pcValue[dataTitles.rental_end.index] != '' ? Utilities.formatDate(pcValue[dataTitles.rental_end.index], 'JST', 'yyyy-MM-dd') : '';
      return (
        pcValue[dataTitles.rentalid.index] === myId &&
        (editedDate != rendalDate || Number(pcValue[dataTitles.rental_fee.index]) != Number(value[index.money]))
      );
    });
    if (data.length === 0) return;
    Logger.log('putKintone: レコードID = ' + data[0][dataTitles[KintoneApi.KEY_ID].index])
    
    putData.push({
      id: data[0][dataTitles[KintoneApi.KEY_ID].index],
      record: {
        rental_end       : { value: rendalDate },
        rental_contractid: { value: value[index.agreementNo] },
        rental_fee       : { value: value[index.money] }
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
  return
  
  //putKintone(simplitCSVSheet.getZenshisuData()); // 台帳の自動更新を走らせる
  //var data = simplitCSVSheet.values.slice(simplitCSVSheet.titleRow + 1); // タイトル以下のデータ
  //var data = [simplitCSVSheet.values[236]];
  //putKintone(data);
}
