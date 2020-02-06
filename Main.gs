/**
 * v1.0.0
 * 外部に公開 ID: 1pJfY7otnFZBrCcScFVBf-3x1VYHeNGJ6rGTEVeUQRmL6Er7Z1Iz13zbh

 * ▼読み込みライブラリ▼
 * WorkplaceApi : https://script.google.com/a/cyberagent.co.jp/d/1FTr1sWuuRXA98QtjWlhvjxiZEpnGqaCXvFLpESuS3vI2IUEKkAM4tIXJ/edit?usp=drive_web&folder=1EpZaef-5PS-9BY2B1peeXd6Uox5bRERn&splash=yes
 * KintoneApi : https://script.google.com/a/cyberagent.co.jp/d/1JVXoRoeGsMWfVC_kPdnc1DuTxUjXh4dyG3x15kk6fn-9Upg7KMi3HEuz/edit?usp=drive_web&folder=1XUmNNRDId1HwJbTAWetjdZrmQyDbN8pD&splash=yes
 * RentalPc： https://script.google.com/d/M734HhKXBXrZHXEzrZjZBtv2vJ94xa-wl/edit?mid=ACjPJvF31u2fW0cZ_va5Gdxc13jvynMxx9KBDkdLq54UBm9QfXFJTuRFDjmvwlIdhKQjhO5w2ID9a9C7DwUybXB5mZgP0W3MCmfV2E3g7rXjRNJdTKA9E3dnrPmuymWSEwvbvhEsAYQHuuU&uiv=2
 * RentalReturnTaskSheet: https://script.google.com/d/McF98qFj4HgFkKtrQe7WoaP2vJ94xa-wl/edit?mid=ACjPJvHCJD-apER8BFJalsYDzwcJV3yqj_YC8c7vWShGjz7PCaFj903yYYZPv0GTvBcswGa9PjzV3ro34JgW7OOqMXh5WpeFOb2ph5Vsibd_dlSmCxsrH3qhMc0wDIU09seCu2tHyrlkja0&uiv=2
 */

/**
 * エラーpopup 指定行が見つからないときに使う
 */
function showTitleError(key) {
  Browser.msgBox('データが見つかりません', '表のタイトル名を変えていませんか？ : ' + key, Browser.Buttons.OK);
}

/**
 * 深夜に行う更新
 * トリガー登録：池田 毎日1:00~2:00
 */
function midnightUpdate() {
  putKintone(simplitCSVSheet.getZenshisuData()); // 台帳の自動更新を走らせる
}
/**
 * 午後に行う更新
 * トリガー登録：池田 毎日14:00~15:00
 */
function afternoonUpdate() {
  // レンタルPC(インターン・内定者用)シートのデータを更新する
  updateInternlistSinc();
}


// TODO: これ使っているか聞く　＞＜

function TriggeredUpdate() {
  var DateObject = new Date();
  var now = Utilities.formatDate(DateObject, 'JST', 'yyyy/MM/dd hh:mm')
  var Result = UpdateData();
  if (Result != RESULT_TEXT__ERROR){
    MailApp.sendEmail('okada_tomoko@cyberagent.co.jp', 'CSVファイル日次処理成功', now)
  }
}

function UploadReauest() {
  var body = '資産管理チーム各位\n\nお疲れ様です！\n本日のSimplitデータのアップロードをお願いします！\n\nSimplitManagerからダウンロードしたレンタル台帳データをファイル名を変更せずに下記にアップロードしてください。\nhttps://drive.google.com/drive/folders/1JOnTxUsT8rxRYd8ougVkSEwpWt_OB_9J\n\n12時までにアップロードすると、自動的に下記のスプレッドシートに反映されます。\nhttps://docs.google.com/spreadsheets/d/1j4KatfSqVIiUnVckXfu7KCvh_ZTzemGUbct5UeWDgHM\n\n手動で反映する場合は、このスプレッドシートのメニューにある「Simplit」から「Simplitインポート」を実行してください。'
  MailApp.sendEmail('infosys_sup@cyberagent.co.jp', 'Simplitデータの更新をお願いします！', body);
}