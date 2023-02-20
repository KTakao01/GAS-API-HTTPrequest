function sendRequesttesttest(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sht = ss.getSheetByName('リクエストsample');
Logger.log(sht.getLastRow()+100);
var methodsArray = sht.getRange(2,2,sht.getLastRow()-1).getValues();　//複数セルの値を取得する - getValues -
var idsArray = sht.getRange(2,1,sht.getLastRow()-1).getValues();
Logger.log(methodsArray);//[[POST], [POST], [GET]]
Logger.log(idsArray);//[[1.0], [2.0], [3.0]] //rowNumber= ids+１(最初の行は見出しなのでカウントしない)
//methodsArrayとidsArrayを１次配列に変換する
var methodArray = methodsArray.flat();
var idArray = idsArray.flat();
//Logger.log(methodArray);//	[POST, POST, GET, ]
Logger.log(idArray);//[1.0, 2.0, 3.0, ] //rowNumber= id+１(最初の行は見出しなのでカウントしない)//rownumber = [2,3,4]だと好都合
var id = 0;
var　method = "";
for(var k=0;k<idArray.length;k++){
var id = idArray[k];
var method = methodArray[k];
//Logger.log(method);//	[POST, POST, GET]
//Logger.log(id);
var rowNumber = id + 1;
//Logger.log(rowNumber+10);
if(method == "GET") {
  sendGetRequest()
} 
else if(method == "POST"){
  sendPostRequest()
}
}


function sendGetRequest() {
//スプレッドシートを開いて取得
//function sendRequestで指定しているので不要
//var ss = SpreadsheetApp.getActiveSpreadsheet();
//データsampleシートを指定
//var sht = ss.getSheetByName("リクエストsample");
//ｇｅtvalueでgetrangeの値を取得
// セル範囲の値を２次元配列で取得する
var value = sht.getRange(rowNumber,3).getValues();
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

//入力のある行のパラメータを取得
function sendPostRequest() {
  //パラメータの一覧keyを取得
  //function sendRequestで指定しているので不要
 // let mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リクエストsample');
  
    const range = sht.getRange(1, 4, 1 ,sht.getLastColumn()-3)
    const keys2array = range.getValues() 
    var keyFlat = keys2array.flat()
 
  //パラメータの内容valueを取得
  
    const rangeParam = sht.getRange(rowNumber, 4, 1, sht.getLastColumn()-3)
    const values2array = rangeParam.getValues() 
    var valueFlat = values2array.flat()

  //key-valueを対応させる
  //key-valueの二次元配列を作成する
  var obj2array = [keyFlat,valueFlat];
  
  //console.log(obj2array); 
  

   
  //key-valueの二次元配列から連想配列への変換 
　  const keys = obj2array[0];
    const values = obj2array[1];
    var obj = {};
  // 空valueの削除
    for(let j=0; j<=keys.length; j++){
      if (values[j] !== ""){
      obj[keys[j]] = values[j];}
      else {
      }
     //URLを配列で取得
    const val = sht.getRange(rowNumber,3).getValues();  // B列の全ての行を取得
  //  const numberOfValues = val.filter(String).length; // 空以外の配列の数を数える
  //  const urls = mySheet.getRange(2,2,numberOfValues-1).getValues();   
   // var urlFlat = urls.flat()
    var urlFlat = val.flat()
    //console.log(urlFlat);
    //URL配列からURLを抽出
    
    var url = urlFlat[0]; 
    
    }

  console.log(obj);
  console.log(url);
  //配列の文字列化
  // var string = Object.values(obj).join(",");
  // console.log(string);
  //上記：keyがなくなるので意味ない

  //配列のJSON化
  var string =　JSON.stringify(obj)
  console.log(string);

  //POSTリクエスト　parameter
    var options = {
    'method' : 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload' : string
    };
 

  
     //POSTリクエスト　url
     var response = UrlFetchApp.fetch(url, options);
  

  console.log(response.getContentText())
  console.log(response.getResponseCode())
Logger.log(idArray);
 }
 Logger.log(idArray);
}