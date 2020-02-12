/**
 * 経理からもらう固定資産のシート
 */
var FixedAssetSheet = function() {
  this.sheet = SpreadsheetApp.openById(MY_SHEET_ID).getSheetByName('最新');
  this.values = this.sheet.getDataRange().getValues();
  this.titleRow = 0;
  this.index = {};
  
  this.createIndex = function() {
    const KEY_TEXT = '資産勘定科目コード';
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
      subjectCode    : filterData.indexOf(KEY_TEXT),
      subjectName    : filterData.indexOf('資産勘定科目名'),
      code           : filterData.indexOf('資産コード'),
      assetName      : filterData.indexOf('資産名'),
      getDate        : filterData.indexOf('取得日付'),
      num            : filterData.indexOf('数量'),
      price          : filterData.indexOf('取得価額'),
      monthPrice     : filterData.indexOf('月末帳簿価額（会計）'),
      ringy          : filterData.indexOf('稟議番号'),
      appendix       : filterData.indexOf('備考'),
      retailerName   : filterData.indexOf('購入先名'),
      departmentCode : filterData.indexOf('部門コード'),
      departmentName : filterData.indexOf('部門名'),
      place          : filterData.indexOf('設置場所名'),
      expenseDivision: filterData.indexOf('費目区分名')
    };
    return this.index;
  };
}
  
FixedAssetSheet.prototype = {
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
   * 工具器具備品&稟議番号があるデータのみを返す
   */
  getMachineryData: function() {
    var index = fixedAssetSheet.getIndex();
    return fixedAssetSheet.values.filter(function(value) {
      return value[index.subjectName] === '工具器具備品' && value[index.ringy] != '' && value[index.code] != '';
    });
  },
  /**
   * 指定の行のデータを新規登録する
   * @param rows [number] 登録したい行数を配列に入れて指定
   */
  postRecords: function(values) {
    var index = fixedAssetSheet.getIndex();
    var today = new Date();
    // postする内容 [{key1: {value: 'hoge'}, key2: {value: 'fuga'}}]
    var data = values.map(function(value) {
      var d = value[index.getDate]
      var f =Utilities.formatDate(value[index.getDate], 'JST', 'yyyy-MM-dd');
      return {
        fixedasset      : { value: value[index.code]},
        ringy           : { value: ('0000000000' + value[index.ringy]).slice(-10)},
        status          : { value: '有効'},
        get_date        : { value: Utilities.formatDate(value[index.getDate], 'JST', 'yyyy-MM-dd')},
        purchase_amount : { value: value[index.price]},
        fixed_price     : { value: value[index.monthPrice]},
        fixed_price_date: { value: Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd')},
        asset_name      : { value: value[index.assetName]},
        assets_num      : { value: value[index.num]},
        subject_name    : { value: value[index.subjectName]},
        appendix        : { value: value[index.appendix]},
        retailer_name   : { value: value[index.retailerName]},
        department_code : { value: value[index.departmentCode]},
        department_name : { value: value[index.departmentName]},
        place           : { value: value[index.place]},
        expense_division: { value: value[index.expenseDivision]}
      };
    });
    KintoneApi.fixedAssetApi.api.postRecords(data);
  },
  /**
   * 指定のデータを更新する
   * @param values [object] 更新したい行数のデータを配列に入れて指定
   */
  putRecords: function(values) {
    var index = this.getIndex();
    var today = new Date();
    // 変更する内容 [{id: レコードID, record: {value: 'hoge'}]
    var data = values.map(function(value) {
      var idd = value[index.recodeNo]
      return {
        id: value[index.recodeNo],
        record: {
          ringy           : { value: ('0000000000' + value[index.ringy]).slice(-10)},
          fixed_price     : { value: value[index.monthPrice]},
          fixed_price_date: { value: Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd')},
        }
      };
    });
    KintoneApi.fixedAssetApi.api.putRecords(data);
  }
};

var fixedAssetSheet = new FixedAssetSheet();
