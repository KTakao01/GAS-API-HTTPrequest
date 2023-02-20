

//入力のある行のパラメータを取得
function sendPostRequestnew() {
  //パラメータの一覧keyを取得

    let mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リクエストsample');
  for(let i=2;i<=mySheet.getLastRow();i++){
    const range = mySheet.getRange(1, 3, 1 ,mySheet.getLastColumn()-2)
    const keys2array = range.getValues() 
    var keyFlat = keys2array.flat()
 
  //パラメータの内容valueを取得
  
    const rangeParam = mySheet.getRange(i, 3, 1, mySheet.getLastColumn()-2)
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
    const val = mySheet.getRange("B:B").getValues();  // B列の全ての行を取得
    const numberOfValues = val.filter(String).length; // 空以外の配列の数を数える
    const urls = mySheet.getRange(2,2,numberOfValues-1).getValues();   
    var urlFlat = urls.flat()
    //console.log(urlFlat);
    //URL配列からURLを抽出
    
    var url = urlFlat[i-2]; 
    
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
  }


 }
