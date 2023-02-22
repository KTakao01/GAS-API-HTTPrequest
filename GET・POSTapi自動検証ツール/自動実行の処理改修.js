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

    //書き込み先セルを文字列で取得
    //var writtenCell = "E" + String(rowNumber)
    //rowNumberあるので不要

    if (method == "GET") {
      //sendGetRequest(sht,rowNumber);//重複するので削除

      //sendGetRequest()の返却値：レスポンスメッセージとステータスコードをメイン関数で再利用する。
      //書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
      const { getMessage, getStatusCode } = sendGetRequest(sht, rowNumber);
      sht.getRange(rowNumber, 5).setValue(getStatusCode);


      //検証//console.log(getStatusCode); 
    }
    else if (method == "POST") {
      //sendPostRequest(sht,rowNumber);//重複するので削除
      const { postMessage, postStatusCode } = sendPostRequest(sht, rowNumber);
      sht.getRange(rowNumber, 5).setValue(postStatusCode);

    }
  }
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendGetRequest(sht, rowNumber) {
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
    var getMessage = response.getContentText();
    var getStatusCode = String(response.getResponseCode());
    //console.log(response.getContentText())
    //console.log(response.getResponseCode())

    console.log(getApi + "についてaccess_tokenとuser_idの組み合わせが" + getCombi + "の時、リクエストとレスポンスは以下の通りです。※URL、レスポンスメッセージ、ステータスコードの順に記載")
    console.log(getUrl);
    console.log(getMessage);
    console.log(getStatusCode);

    return { getMessage, getStatusCode };

  }
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendPostRequest(sht, rowNumber) {
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
  var string = JSON.stringify(obj)

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
  var postStatusCode = String(postResponse.getResponseCode());
  //console.log(response.getContentText())
  //console.log(response.getResponseCode())
  //テスト//Logger.log(idArray);

  //リクエストとレスポンスのログ排出
  console.log(postApi + "についてaccess_tokenとuser_idの組み合わせが" + postCombi + "の時、リクエストとレスポンスは以下の通りです。※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
  console.log(urlPost);
  console.log(string);
  console.log(postMessage);
  console.log(postStatusCode);
  return { postMessage, postStatusCode };
}
  //テスト//Logger.log(idArray);
