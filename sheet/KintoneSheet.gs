/**
 * 固定資産の台帳情報をまとめるシート
 */
var KintoneSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('台帳データ');
  this.values = this.sheet.getDataRange().getValues();
  this.titleRow = 0;
  this.index = {};
  
  this.createIndex = function() {
    const KEY_TEXT = 'レコード番号';
    var me = this;
    var filterData = (function() {
      for(var i = 0; i < me.values.length; i++) {
        if (me.values[i].indexOf(KEY_TEXT) > -1) {
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
      recodeNo       : filterData.indexOf(KEY_TEXT),
      code           : filterData.indexOf('固定資産コード'),
      ringy          : filterData.indexOf('稟議番号'),
      status         : filterData.indexOf('ステータス'),
      getDate        : filterData.indexOf('取得日付'),
      money          : filterData.indexOf('購入金額'),
      assetName      : filterData.indexOf('資産名'),
      num            : filterData.indexOf('数量'),
      subjectName    : filterData.indexOf('資産勘定科目名'),
      monthPrice     : filterData.indexOf('月末帳簿価額'),
      monthPriceDate : filterData.indexOf('月末帳簿価額 更新日'),
      appendix       : filterData.indexOf('備考'),
      retailerName   : filterData.indexOf('購入先名'),
      departmentCode : filterData.indexOf('部門コード'),
      departmentName : filterData.indexOf('部門名'),
      place          : filterData.indexOf('設置場所名'),
      expenseDivision: filterData.indexOf('費目区分名'),
      row            : filterData.indexOf('行数')
    };
    return this.index;
  };
}
  
KintoneSheet.prototype = {
  getRowKey: function(target) {
    var targetIndex = this.getIndex()[target];
    var returnKey = (targetIndex > -1) ? SHEET_ROWS[targetIndex] : '';
    if (!returnKey || returnKey === '') showTitleError(target);
    return returnKey;
  },
  getIndex: function() {
    return Object.keys(this.index).length ? this.index : this.createIndex();
  },
  updateData: function() {
    this.getIndex();
    var lastRow = this.sheet.getRange('A:A').getValues().filter(String).length + 1;
    var t = this.titleRow
    var f = this.values[this.titleRow]
    this.sheet.getRange(this.titleRow + 1, 1, lastRow, this.values[this.titleRow - 1].length).clearContent();
    var data = KintoneApi.fixedAssetApi.getAllData().map(function(record) {
      return [
        record[KintoneApi.KEY_ID].value,
        record.fixedasset.value,
        record.ringy.value,
        record.status.value,
        record.get_date.value,
        record.purchase_amount.value,
        record.fixed_price.value,
        record.fixed_price_date.value,
        record.asset_name.value,
        record.assets_num.value,
        record.subject_name.value,
        record.appendix.value,
        record.retailer_name.value,
        record.department_name.value,
        record.department_code.value,
        record.place.value,
        record.expense_division.value
      ];
    });
    this.sheet.getRange(this.titleRow + 1, 1, data.length, data[0].length).setValues(data);
    this.sheet.getRange('E1').setValue(Utilities.formatDate(new Date(), 'JST', 'yyyy-MM/dd(E) HH:mm'));
  }
};
var kintoneSheet = new KintoneSheet();
