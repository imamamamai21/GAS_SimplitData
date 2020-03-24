/**
 * ２週間以内に終了するレンタルPCリスト
 * @return [Object] [{rentalNo, endDate, agreementNo}]
 */
function getEndList() {
  var index = simplitCSVSheet.getIndex();
  var today = new Date();
  return simplitCSVSheet.getZenshisuData()
    .filter(function(value) {
      var dif = (value[index.endDate] - today) / 1000 / 60 / 60 / 24; // 日で割る
      return dif < 14;
    })
    .map(function(value){
      return {
        rentalNo: simplitCSVSheet.getRentalNo(value[index.yrlNo].toString()),
        endDate: value[index.endDate],
        agreementNo: value[index.agreementNo],
        demandCompany: value[index.demandCompany]
      };
    });
}

/**
 * 終了日が早い順にソートして返す
 * @return [Object] [{rentalNo, endDate, agreementNo}]
 */
function getSortedEndList() {
  return getEndList().sort(function(a, b) {
    if (a.endDate < b.endDate) return -1;
    if (a.endDate > b.endDate) return 1;
    return 0;
  })
}

/**
 * レンタル終了間近のものをworkplaceにBOTから投稿する
 * 前回投稿したものは投稿しない
 */
function postEndData() {
  const TITLE = '# ▼レンタルの終了日が近づいているPCがあります\n';
  const MY_SHEET_ROW__BOT_MEMO = 'E';
  var text = TITLE;
  var noOnlyText = ''; // レンタル番号のみを表示するテキスト
  var memoSheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('はじめに');
  var lastMemo = memoSheet.getRange(MY_SHEET_ROW__BOT_MEMO + ':' + MY_SHEET_ROW__BOT_MEMO).getValues().filter(String).length;
  var beforeDate;
  var editValues = [];
  var day = (new Date()).getDate();
  
  // 毎月1日は投稿履歴をリセット(全てのものを投稿する)
  var isTopDay = (day === 1);
  if (isTopDay) {
    for (var i = 2; i <= lastMemo; i++) {
      memoSheet.getRange(MY_SHEET_ROW__BOT_MEMO + i).setValue('');
    }
  }
  getSortedEndList().forEach(function (value, index) {
    if (!isTopDay) {
      // すでにBOTで吐かれた契約番号か否か
      var isPosted = memoSheet.getDataRange().getValues().filter(function(memo) {
        return memo[4] === value.rentalNo;
      });
      if (isPosted.length > 0) return;
    }
    editValues.push([value.rentalNo]);
    var date = Utilities.formatDate(value.endDate, "JST", "yyyy年MM月dd日");
    
    if (beforeDate !== date) {
      text += '## レンタル終了予定日: ' + date + '\n';
      beforeDate = date;
    }
    text += '```\nNO: ' + value.rentalNo + '  契約番号: ' + value.agreementNo + '　需要先: ' + value.demandCompany + '\n```\n';
    noOnlyText += value.rentalNo + '\n';
  });
  if (text === TITLE || editValues.length === 0) return;
  
  // メモに契約番号を書き込む
  memoSheet.getRange(lastMemo + 1, 5, editValues.length, 1).setValues(editValues);
    
  text += '\n\n以上が期間が過ぎているor２週間以内に過ぎるリストです。' +
  '\nチェックは平日13:00~14:00に走ります。\n毎月1日には全ての終了間近PCを表示しますが、それ以降は新たにリストに上がったPCがあった時のみ表示します。\n\n' +
  '\n▼関連シート\n* [Simplitデータ](' + SHEET_URL + MY_SHEET_ID + ')\n* [人事用返却シート](' + SHEET_URI__RETURN_JINJI + ')\n* [返却タスクシート](' + SHEET_URI__RETURN_TASK + ')\n' +
  '\n▼検索用\n```\n' + noOnlyText + '```\n';
  
  //WorkplaceApi.postBotForTest(text);
  WorkplaceApi.postBotForArms(text);
}
