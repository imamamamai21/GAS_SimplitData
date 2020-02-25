/**
 * Simplitシート どのシートもほとんど一緒なのでnew時にシート名指定して共通で使う
 * @param sheetName {string} シート名
 */
var SimplitSheet = function(sheetName) {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName(sheetName);
  this.values = this.sheet.getDataRange().getValues();
  this.titleRow = 0;
  this.index = {};
  
  this.createIndex = function() {
    const TEXT_NO = 'YRL管理番号';
    var me = this;
    var filterData = (function() {
      for(var i = 0; i < me.values.length; i++) {
        if (me.values[i].indexOf(TEXT_NO) > -1) {
          me.titleRow = i + 1;
          return me.values[i];
        }
      }
    }());
    if(!filterData || filterData.length === 0) {
      showTitleError();
      return;
    }
    
    this.index = {
      timeStamp         : filterData.indexOf(TITLE__TIME_STAMP),
      yrlNo             : filterData.indexOf(TEXT_NO),
      agreementNo       : filterData.indexOf('契約番号'),
      agreementIndex    : filterData.indexOf('契約行番号'),
      code              : filterData.indexOf('商品コード'),
      serialNo          : filterData.indexOf('シリアル番号'),
      startDate         : filterData.indexOf('レンタル開始日'),
      startExtensionDate: filterData.indexOf('レンタル延長開始日'),
      endDate           : filterData.indexOf('レンタル終了予定日'),
      contractedParty   : filterData.indexOf('契約先事業所名称'),
      demandCompany     : filterData.indexOf('需要先会社名称'),
      money             : filterData.indexOf('レンタル月額(契約先)')
    };
    return this.index;
  }
}
  
SimplitSheet.prototype = {
  getRowKey: function(target) {
    var targetIndex = this.getIndex()[target];
    var returnKey = (targetIndex > -1) ? SHEET_ROWS[targetIndex] : '';
    if (!returnKey || returnKey === '') showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  },
  /**
   * 対象のレンタルPCのデータを渡す
   * @param レンタルPC番号 {string} '00-00000'
   * @return array 
   */
  getTargetData: function(rentalNo) {
    var index = this.getIndex().yrlNo;
    return this.values.filter(function(value) {
      return value[index].toString() === rentalNo.replace('-', '');
    })[0];
  },
  /**
   * 全シスのレンタルPCのデータを渡す
   */
  getZenshisuData: function() {
    var index = this.getIndex().contractedParty ;
    return this.values.filter(function(value) {
      return value[index] === '全社システム本部';
    });
  },
  /**
   * yrlNoをレンタルNOになおして返す
   * yrlNo {string}
   */
  getRentalNo: function(yrlNo) {
    return yrlNo.substr(0, 2) + '-' + yrlNo.substr(2, 5);
  },
  getLastColumn: function() {
    return this.sheet.getLastColumn();
  },
  getLastRow: function() {
    return this.sheet.getLastRow();
  }
};

var simplitCSVSheet = new SimplitSheet('CSVデータ');
