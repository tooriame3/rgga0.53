function add_menuVUP() {
  "use strict";
  // 現在アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var spreadsheetObj = SpreadsheetApp.getActiveSpreadsheet();
  var menuList       = [];
  menuList.push({
    name : "１：まだありません",
    functionName : "vup0"
  });
  /*/
  menuList.push({
    name : "２：myTest",
    functionName : "vup1"
  });
  /*/
  ss.addMenu("バージョンアップ", menuList);
}
function vup0(){ 
}

function vup1() {
  var verNo = '1.01';
  var kousin = '1.01　テスト。 ライブラリー：aggr 1.03'
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('使用方法');
  if (sh.getRange('A1').getValue()!= verNo){
  sh.insertRowBefore(3);
  sh.getRange('B3').setValue(kousin);
  sh.getRange('A1').setValue(verNo);
  }
}

function setNameLookup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('対応表');
  ss.setNamedRange('戦略名', sh.getRange('F2:F23'));
  ss.setNamedRange('番号から戦略名', sh.getRange('D2:F23'));
  ss.setNamedRange('メモから戦略名', sh.getRange('H2:K43'));
};
function sortLookup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('対応表');
  sh.getRange('D2:F23').sort([{column: 4, ascending: true}]);
  sh.getRange('H2:K43').sort([{column: 8, ascending: true}]);
};
