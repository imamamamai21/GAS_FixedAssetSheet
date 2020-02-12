/* =============================================================
　　固定資産台帳のデータをもとに、
　　kintoneの固定資産データアプリに登録するためのコードです。
============================================================= */

/**
 * メニューを設定する
 */
function createMenu() {
  var ui = SpreadsheetApp.getUi();         // Uiクラスを取得する
  var menu = ui.createMenu('▼スクリプト');  // Uiクラスからメニューを作成する
  // メニューにアイテムを追加する
  menu.addItem('固定資産アップデート', 'onClickUpdateData');
  menu.addToUi(); // メニューをUiクラスに追加する
}

function onClickUpdateData() {
  var conf = Browser.msgBox('固定資産をアップデートします。', '[最新]シートの内容は最新の固定資産台帳データになっていますか？', Browser.Buttons.YES_NO);
  if (conf === 'yes') addKintone();
}

/**
 * Kintone台帳に登録(or更新)
 */
function addKintone() {
  var index = fixedAssetSheet.getIndex();
  var kintoneIndex = kintoneSheet.getIndex();
  
  var updateValues = [];
  var kintoneValues = kintoneSheet.values.slice(kintoneSheet.titleRow);
  
  var today = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
  
  var unregisteredData = fixedAssetSheet.getMachineryData().filter(function(value) {
    var i = 0;
    var needPost = true;
    
    while (needPost && i < kintoneValues.length) {
      var data = kintoneValues[i];
      if (data[kintoneIndex.code] === value[index.code]) { // すでに台帳に登録があるか
        if (data[kintoneIndex.monthPrice] != value[index.monthPrice] || data[kintoneIndex.ringy] != value[index.ringy]) { // 更新する内容があるかチェック
          Logger.log('UPDATE!!');
          var putValue = { // 台帳データを更新するためのObject
            id: data[kintoneIndex.recodeNo],
            record: {
              ringy           : { value: ('0000000000' + value[index.ringy]).slice(-10)},
              fixed_price     : { value: value[index.monthPrice]},
              fixed_price_date: { value: today}
            }
          };
          updateValues.push(putValue);
        }
        kintoneValues.splice(i, 1); // kintoneValuesから削除
        needPost = false;
      }
      i++;
    }
    return needPost;
  });
  
  // 無効(削除されたデータ)の内容をpush
  kintoneValues.forEach(function(value) {
    updateValues.push({ // 台帳データを更新するためのObject
      id: value[kintoneIndex.recodeNo],
      record: { status: { value: '無効'} }
    });
  });
  
  // 新規で登録する
  if (unregisteredData.length > 0) fixedAssetSheet.postRecords(unregisteredData);
  
  // データを更新する
  if (updateValues.length > 0) KintoneApi.fixedAssetApi.api.putRecords(updateValues);
  
  // 最新のデータを書き込む
  kintoneSheet.updateData();
}
