function addClosingPriceData() {
  var folderName = "okasan";
  var myfolder=DriveApp.getRootFolder().getFoldersByName(folderName).next();
  var fileName = "最新終値データ.csv";
  var folders = DriveApp.getFoldersByName(folderName);
  
  //指定フォルダを検索
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() == folderName) {
      myfolder = folder;
      var files = DriveApp.getFilesByName(fileName);
      //指定したCSVファイルを検索
      while (files.hasNext()) {
        var file = files.next();
        if (file.getName() == fileName) {
          //var data = file.getBlob().getDataAsString("Shift_JIS"); 
          //var csv = Utilities.parseCsv(data);            
          uploadClosingPriceData(file);
          folder.removeFile(file);
          return;
        }
      }
    }
  }
}


function uploadClosingPriceData(file){
  // 現在アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheetName ='終値データの補正'; 
  var sh = ss.getSheetByName(sheetName);
  sh.activate;
  // 指定したセル範囲にフォーマットを適用
  var numFormats1 = 'yyyy/MM/dd';
  var numFormats2 = 'hh:mm:ss';
  //
  　var range = sh.getRange("A2:C");
    range.clearContent();
    var csvData = Utilities.parseCsv(file.getBlob().getDataAsString("Shift_JIS"));
    var data = csvData;
    appendData(sheetName, data);
    // 指定したセル範囲にフォーマットを適用
      sh.getRange("M:M").setNumberFormat(numFormats1);
      sh.getRange("A:A").setNumberFormat(numFormats1);  
      sh.getRange("B:B").setNumberFormat(numFormats2);
      sh.getRange("N:N").setNumberFormat(numFormats2);
    //補正したものをdata2に
    var data2 = [[]];
    var last_rowM = getLastRowNumber_ColumnNo(sheetName,13);
    data2 = sh.getRange(2, 13, last_rowK-1, 3).getValues();//１行目なし
  
　//data2を終値データに張り付け  
  var sheetName ='終値データ'; 
  var sh = ss.getSheetByName(sheetName);
  sh.activate;
  appendData(sheetName, data2);
  // 指定したセル範囲にフォーマットを適用
  sh.getRange("K:K").setNumberFormat(numFormats1);
  sh.getRange("A:A").setNumberFormat(numFormats1);  
  sh.getRange("B:B").setNumberFormat(numFormats2);
  sh.getRange("L:L").setNumberFormat(numFormats2);
  //重複を削除したものdata3に換えておく
  var last_rowK = getLastRowNumber_ColumnA(sheetName,11);
  var data3 = [[]];
  data3 = sh.getRange(2, 11, last_rowK-1, 3).getValues();//１行目なし
    var range = sh.getRange("A2:C");
    range.clearContent();
  appendData(sheetName, data3);
  // 指定したセル範囲にフォーマットを適用
  sh.getRange("K:K").setNumberFormat(numFormats1);
  sh.getRange("A:A").setNumberFormat(numFormats1);  
  sh.getRange("B:B").setNumberFormat(numFormats2);
  sh.getRange("L:L").setNumberFormat(numFormats2);  
  //sort
  last_row = getLastRowNumber_ColumnA(sheetName);
  if (last_row>2){
    rng = sh.getRange(2, 1, last_row-1, 3);  // <--対象範囲  
    //そーと
    sh.getRange(2, 1, last_row-1, 3).sort([{column: 1, ascending: true},{column: 2, ascending: true}]);
  }
  //タイトル行が最終行にきたら削除
  if (sh.getRange(last_row,1).getValue()=="日付"){
    sh.deleteRow(last_row);
  };
  
  // 指定したセル範囲にフォーマットを適用
  var numFormats1 = 'yyyy/MM/dd';
  var numFormats2 = 'hh:mm:ss';
  sh.getRange("K:K").setNumberFormat(numFormats1);
  sh.getRange("A:A").setNumberFormat(numFormats1);  
  sh.getRange("B:B").setNumberFormat(numFormats2);
  sh.getRange("L:L").setNumberFormat(numFormats2);
}

//最終行を返す
function getLastRowNumber_ColumnNo(sheetName,col){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var  sh = ss.getSheetByName(sheetName);
  var last_row = sh.getLastRow();
  var data = [[]];
  data = sh.getRange(1, col, last_row, col).getValues();//
  i = data.filter(String).length;
  return i;
}

