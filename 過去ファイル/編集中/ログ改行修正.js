function updateTrigger() {
  //Deletes all triggers in the current project.//過去のトリガーが残って処理が遅くなるときに利用する
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  //有効なトリガーを取得する
  Logger.log("トリガーを更新します。");
  Logger.log('Current project has ' + ScriptApp.getProjectTriggers().length + ' triggers.');

  //スプシ起動時にステータスコード書き込みを自動実行するトリガーを作成
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var autoRun = ScriptApp.newTrigger("main").forSpreadsheet(ssId).onOpen().create();
}


//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function main() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht = ss.getSheetByName('オートメーションテスト');
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
      //sendGetRequest()の返却値：レスポンスメッセージとステータスコードをメイン関数で再利用する。
      //書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
      const { getMessageObj, getStatusCode } = sendGetRequest(sht, rowNumber , method);
      sht.getRange(rowNumber, 5).setValue(getStatusCode);


      //検証//console.log(getStatusCode); 
    }
    else if (method == "POST") {
      const { postMessageObj, postStatusCode } = sendPostRequest(sht, rowNumber , method);
      sht.getRange(rowNumber, 5).setValue(postStatusCode);

    }

    else if (method == "PUT") {
      const { putMessageObj, putStatusCode } = sendPutRequest(sht, rowNumber , method);
      sht.getRange(rowNumber, 5).setValue(putStatusCode);

    }

    else if (method == "DELETE") {
      const { deleteMessageObj, deleteStatusCode } = sendDeleteRequest(sht, rowNumber ,method);
      sht.getRange(rowNumber, 5).setValue(deleteStatusCode);

    }


  }
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendGetRequest(sht, rowNumber, method) {
  //getvalueでgetrangeの値を取得
  // セル範囲の値（ここではURL）を２次元配列で取得する
  var urlValue = sht.getRange(rowNumber, 7).getValues();
  // セル範囲の値（ここではURL）を１次元配列に変換する
  var urlValuesFlat = urlValue.flat()
  //検証//console.log(valuesFlat)

  //ログ排出用にapi,組み合わせカラムのバリューを取得する。
  var getApi = sht.getRange(rowNumber, 3).getValue();
  var getCombi = sht.getRange(rowNumber, 4).getValue();
  //console.log(getApi);
  //console.log(getCombi);

  //一次元配列からURLテキスト抽出
  while (urlValuesFlat.length) {
    var getUrl = urlValuesFlat.shift();

    var options = {
      'method': 'get',
      "muteHttpExceptions": true,
    };

    //レスポンス
    var response = UrlFetchApp.fetch(getUrl, options);
    var getStatusCode = String(response.getResponseCode());
    var getMessage = response.getContentText();
    var getMessageObj = JSON.parse(getMessage);
    console.log(method + " " + getApi + "についてaccess_tokenとuser_idの組み合わせが" + getCombi + "の時、リクエストとレスポンスは以下の通りです。\n※URL、レスポンスメッセージ、ステータスコードの順に記載")
    console.log(getUrl);
    console.log(getMessageObj);
    console.log(getStatusCode);

    return { getMessage, getStatusCode };

  }
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendPostRequest(sht, rowNumber, method) {
  //パラメータの一覧keyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(1, 8, 1, sht.getLastColumn() - 7)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  var keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(rowNumber, 8, 1, sht.getLastColumn() - 7)
  const values2array = rangeParam.getValues()
  var valueFlat = values2array.flat()

  //key-valueの各要素を対応させる
  //key-valueの二次元配列を作成する
  var obj2array = [keyFlat, valueFlat];

  //console.log(obj2array); 
  //key-valueの二次元配列から連想配列への変換 
  const keys = obj2array[0];
  const values = obj2array[1];
  var obj = {};

  //入力のあるパラメータを取得.valueが空の時はkeyが削除される
  for (let j = 0; j <= keys.length; j++) {
    if (values[j] !== "") {
      obj[keys[j]] = values[j];
    }
    else {
    }
    //URLを二次元配列で取得
    const urlPost2Array = sht.getRange(rowNumber, 7).getValues();  // G列:URLカラムの全ての行を取得

    //URLの二次元配列を一次元配列に変換
    var urlPostFlat = urlPost2Array.flat()
    //console.log(urlPostFlat);

    //URL配列からURLを抽出
    var urlPost = urlPostFlat[0];
  }


  //key-value配列のJSON化
  var string = JSON.stringify(obj,null , "\t")
  

  //POSTリクエスト　parameter
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': string,
    "muteHttpExceptions": true,
  };

  //POSTリクエスト　urlと合わせて送信
  var postResponse = UrlFetchApp.fetch(urlPost, options);

  //ログ排出用にapi,組み合わせカラムのバリューを取得する。
  var postApi = sht.getRange(rowNumber, 3).getValue();
  var postCombi = sht.getRange(rowNumber, 4).getValue();

  //レスポンス
  var postMessage = postResponse.getContentText();
  var postMessageObj = JSON.parse(postMessage);
  var postStatusCode = String(postResponse.getResponseCode());
  //console.log(response.getContentText())
  //console.log(response.getResponseCode())
  //テスト//Logger.log(idArray);

  //リクエストとレスポンスのログ排出
  console.log(method + " " +  postApi + "についてaccess_tokenとuser_idの組み合わせが" + postCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
  console.log(urlPost);
  console.log(string);
  console.log(postMessageObj);
  console.log(postStatusCode);
  return { postMessage, postStatusCode };
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendPutRequest(sht, rowNumber, method) {
  //パラメータの一覧keyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(1, 8, 1, sht.getLastColumn() - 7)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  var keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(rowNumber, 8, 1, sht.getLastColumn() - 7)
  const values2array = rangeParam.getValues()
  var valueFlat = values2array.flat()

  //key-valueの各要素を対応させる
  //key-valueの二次元配列を作成する
  var obj2array = [keyFlat, valueFlat];

  //console.log(obj2array); 
  //key-valueの二次元配列から連想配列への変換 
  const keys = obj2array[0];
  const values = obj2array[1];
  var obj = {};

  //入力のあるパラメータを取得.valueが空の時はkeyが削除される
  for (let j = 0; j <= keys.length; j++) {
    if (values[j] !== "") {
      obj[keys[j]] = values[j];
    }
    else {
    }
    //URLを二次元配列で取得
    const urlPut2Array = sht.getRange(rowNumber, 7).getValues();  // G列:URLカラムの全ての行を取得

    //URLの二次元配列を一次元配列に変換
    var urlPutFlat = urlPut2Array.flat()
    //console.log(urlPostFlat);

    //URL配列からURLを抽出
    var urlPut = urlPutFlat[0];
  }


  //key-value配列のJSON化
  var string = JSON.stringify(obj,null , "\t")

  //PUTリクエスト　parameter
  var options = {
    'method': 'put',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': string,
    "muteHttpExceptions": true,
  };

  //PUTリクエスト　urlと合わせて送信
  var putResponse = UrlFetchApp.fetch(urlPut, options);

  //ログ排出用にapi,組み合わせカラムのバリューを取得する。
  var putApi = sht.getRange(rowNumber, 3).getValue();
  var putCombi = sht.getRange(rowNumber, 4).getValue();

  //レスポンス
  var putMessage = putResponse.getContentText();
  var putMessageObj = JSON.parse(putMessage);
  var putStatusCode = String(putResponse.getResponseCode());
  //console.log(response.getContentText())
  //console.log(response.getResponseCode())
  //テスト//Logger.log(idArray);

  //リクエストとレスポンスのログ排出
  console.log(method + " " + putApi + "についてaccess_tokenとuser_idの組み合わせが" + putCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
  console.log(urlPut);
  console.log(string);
  console.log(putMessageObj);
  console.log(putStatusCode);
  return { putMessage, putStatusCode };
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendDeleteRequest(sht, rowNumber, method) {
  //パラメータの一覧keyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(1, 8, 1, sht.getLastColumn() - 7)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  var keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(rowNumber, 8, 1, sht.getLastColumn() - 7)
  const values2array = rangeParam.getValues()
  var valueFlat = values2array.flat()

  //key-valueの各要素を対応させる
  //key-valueの二次元配列を作成する
  var obj2array = [keyFlat, valueFlat];

  //console.log(obj2array); 
  //key-valueの二次元配列から連想配列への変換 
  const keys = obj2array[0];
  const values = obj2array[1];
  var obj = {};

  //入力のあるパラメータを取得.valueが空の時はkeyが削除される
  for (let j = 0; j <= keys.length; j++) {
    if (values[j] !== "") {
      obj[keys[j]] = values[j];
    }
    else {
    }
    //URLを二次元配列で取得
    const urlDelete2Array = sht.getRange(rowNumber, 7).getValues();  // G列:URLカラムの全ての行を取得

    //URLの二次元配列を一次元配列に変換
    var urlDeleteFlat = urlDelete2Array.flat()
    //console.log(urlDeleteFlat);

    //URL配列からURLを抽出
    var urlDelete = urlDeleteFlat[0];
  }


  //key-value配列のJSON化
  var string = JSON.stringify(obj,null , "\t")

  //DELETEリクエスト　parameter
  var options = {
    'method': 'delete',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': string,
    "muteHttpExceptions": true,
  };

  //DELETEリクエスト　urlと合わせて送信
  var deleteResponse = UrlFetchApp.fetch(urlDelete, options);

  //ログ排出用にapi,組み合わせカラムのバリューを取得する。
  var deleteApi = sht.getRange(rowNumber, 3).getValue();
  var deleteCombi = sht.getRange(rowNumber, 4).getValue();

  //レスポンス
  var deleteMessage = deleteResponse.getContentText();
  var deleteMessageObj = JSON.parse(deleteMessage);
  var deleteStatusCode = String(deleteResponse.getResponseCode());
  //console.log(response.getContentText())
  //console.log(response.getResponseCode())
  //テスト//Logger.log(idArray);

  //リクエストとレスポンスのログ排出
  console.log(method + " " + deleteApi + "についてaccess_tokenとuser_idの組み合わせが" + deleteCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
  console.log(urlDelete);
  console.log(string);
  console.log(deleteMessageObj);
  console.log(deleteStatusCode);
  return { deleteMessage, deleteStatusCode };
}


  //テスト//Logger.log(idArray);
