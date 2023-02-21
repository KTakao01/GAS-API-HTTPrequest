function main() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht = ss.getSheetByName('リクエストsample 修正1');
  //test//Logger.log(sht.getLastRow()+100);　//100+4//行を取得できているかの確認

  //複数セルの値を二次元配列として取得する - getValues -
  var methodsArray = sht.getRange(2, 2, sht.getLastRow() - 1).getValues();
  var idsArray = sht.getRange(2, 1, sht.getLastRow() - 1).getValues();
  //Logger.log(methodsArray);//[[POST], [POST], [GET]]
  //Logger.log(idsArray);//[[1.0], [2.0], [3.0]] //rowNumber= ids+１(最初の行は見出しなのでカウントしない)

  //取得した二次元配列：methodsArrayとidsArrayを１次配列に変換する
  var methodArray = methodsArray.flat();
  var idArray = idsArray.flat();
  //test　//id取得できているかの確認、後に行数に変換する.rownumberとして定義する。
  //Logger.log(idArray);//[1.0, 2.0, 3.0] //rowNumber= id+１(最初の行は見出しなのでカウントしない)//rownumber = [2,3,4]だと好都合
  var id = 0;
  var method = "";
  for (var k = 0; k < idArray.length; k++) {
    var id = idArray[k];
    var method = methodArray[k];

    //検証//Logger.log(method);//	[POST, POST, GET]
    //検証//Logger.log(id);
    var rowNumber = id + 1;
    //検証//Logger.log(rowNumber+10);

    if (method == "GET") {
      sendGetRequest(sht,rowNumber)
    }
    else if (method == "POST") {
      sendPostRequest(sht,rowNumber)
    }
  }
}


function sendGetRequest(sht,rowNumber) {
  //getvalueでgetrangeの値を取得
  // セル範囲の値（ここではURL）を２次元配列で取得する
  var value = sht.getRange(rowNumber, 7).getValues();
  // セル範囲の値（ここではURL）を１次元配列に変換する
  var valuesFlat = value.flat()
  //検証//console.log(valuesFlat)
  while (valuesFlat.length) {
    var Elem = valuesFlat.shift();
    //一次元配列からURLテキスト抽出
    Logger.log(Elem)

    var options = {
      'method': 'get',
      "muteHttpExceptions": true,
    };


    var response = UrlFetchApp.fetch(Elem, options);
    var getMessage= response.getContentText();
    var getStatusCode = response.getResponseCode();
    //console.log(response.getContentText())
    //console.log(response.getResponseCode())
    console.log(getMessage);
    console.log(getStatusCode);  
    return {getMessage,getStatusCode};
  }
}

//入力のある行のパラメータを取得
function sendPostRequest(sht,rowNumber) {
  //パラメータの一覧keyを取得
  const range = sht.getRange(1, 8, 1, sht.getLastColumn() - 7)
  const keys2array = range.getValues()
  var keyFlat = keys2array.flat()

  //パラメータの内容valueを取得  
  const rangeParam = sht.getRange(rowNumber, 8, 1, sht.getLastColumn() - 7)
  const values2array = rangeParam.getValues()
  var valueFlat = values2array.flat()

  //key-valueを対応させる
  //key-valueの二次元配列を作成する
  var obj2array = [keyFlat, valueFlat];

  //console.log(obj2array); 


  //key-valueの二次元配列から連想配列への変換 
  const keys = obj2array[0];
  const values = obj2array[1];
  var obj = {};
  // 空valueの削除
  for (let j = 0; j <= keys.length; j++) {
    if (values[j] !== "") {
      obj[keys[j]] = values[j];
    }
    else {
    }
    //URLを二次元配列で取得
    const val = sht.getRange(rowNumber, 7).getValues();  // G列:URLカラムの全ての行を取得

    //URLの二次元配列を一次元配列に変換
    var urlFlat = val.flat()
    //console.log(urlFlat);

    //URL配列からURLを抽出
    var url = urlFlat[0];

  }

  //検証//console.log(obj);
  console.log(url);

  //配列のJSON化
  var string = JSON.stringify(obj)
  console.log(string);

  //POSTリクエスト　parameter
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': string,
    "muteHttpExceptions": true,
  };

  //POSTリクエスト　url
  var postResponse = UrlFetchApp.fetch(url, options);
  var postMessage = postResponse.getContentText();
  var postStatusCode = postResponse.getResponseCode();
  //console.log(response.getContentText())
  //console.log(response.getResponseCode())
  //Logger.log(idArray);
  console.log(postMessage);
  console.log(postStatusCode);
  return {postMessage,postStatusCode};
}
  //テスト//Logger.log(idArray);
