function getrequest() {
//スプレッドシートを開いて取得
var ss = SpreadsheetApp.getActiveSpreadsheet();
//データsampleシートを指定
var sht = ss.getSheetByName("GETリクエストsample");
//ｇｅtvalueでgetrangeの値を取得

// セル範囲の値を２次元配列で取得する
var value = sht.getRange("A4").getValues();
// セル範囲の値を１次元配列に変換する
var valuesFlat = value.flat()

console.log(valuesFlat)
  while (valuesFlat.length){
      var Elem = valuesFlat.shift(); 
      Logger.log(Elem)
      var response = UrlFetchApp.fetch(Elem);
  console.log(response.getContentText())
  console.log(response.getResponseCode())
  }
}


