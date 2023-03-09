function updateTrigger() {
  //Deletes all triggers in the current project.//過去のトリガーが残っているときに利用
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


  //トリガー同時実行→POST先に処理、PUT・DELETE・GET後処理となったが、POSTしたデータをPUTしたりDELETEしたりはできなかった。
  //サーバーの通信・同期時間を考慮できていない？クライアント側では処理にPOSTとPUTの序列ついているが、サーバーではまとめて処理される。
  //var autoRun = ScriptApp.newTrigger("mainnoPost").forSpreadsheet(ssId).onOpen().create();

  //var autoRun = ScriptApp.newTrigger("mainNoPost").forSpreadsheet(ssId).doPost();
}


function main() {
  mainPost();
  mainNoPost();
}

//POSTしたデータをPUT、DELETE、GETするので、POST処理の優先順位を上げる。
//リクエスト処理の実行、書き込む処理はmainPost()で行う。sendXX()はリクエストとログ書き出し。
function mainPost() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht = ss.getSheetByName('改修オートメーション');
  //test//Logger.log(sht.getLastRow()+100);　//100+4//行を取得できているかの確認


  //for (var i = 2; i < sht.getLastRow()+1 ;i = i +2) {}

  //認証情報を参照して認証必要なAPIのuser_idとaccess_tokenの列に認証結果を書き出す。
  //ここから記述する
  // var methodsArray = sht.getRange(2, 2, sht.getLastRow() - 1).getValues();
  //var idsArray = sht.getRange(2, 1, sht.getLastRow() - 1).getValues();


  //認証不要のときは、リクエスト処理を行う。
  //ここから記述する



  //    
  //複数セルの値を二次元配列として取得する - getValues -
  var methodsArray = sht.getRange(1, 2, sht.getLastRow()).getValues();
  var idsArray = sht.getRange(1, 1, sht.getLastRow()).getValues();
  var authsArray = sht.getRange(1, 4, sht.getLastRow()).getValues();
  //Logger.log(methodsArray);//[[method],[POST],[method],[POST],[method], [POST],[method], [GET]]
  //Logger.log(idsArray);//[[No],[認証],[No], [1.0],[No], [2.0],[No], [3.0]] 


  //取得した二次元配列：methodsArrayとidsArrayを１次配列に変換する
  var methodArray = methodsArray.flat();
  var idArray = idsArray.flat();
  var authArray = authsArray.flat();

  // Logger.log(methodArray);

  //test　//id取得できているかの確認、後に行数に変換する.valueRowNumberとして定義する。
  //Logger.log(idArray);//[No,認証,No,1.0,No, 2.0,No, 3.0] //[0][2][4][6]・・・・がkeyでmainでは不要（sendXXで必要）//[1][3][5]・・・・がvalueで必要
  var id = 0;
  var method = "";
  var auth = "";

  for (var k = 0; k < idArray.length; k = k + 2) {
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    var arrayNo = k + 1
    var id = idArray[arrayNo];
    var method = methodArray[arrayNo];
    var auth = authArray[arrayNo];
    //検証//Logger.log(method);//	[POST, POST, POST, GET]
    //検証//Logger.log(id);//[認証,1,2,3]

    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    var keyRowNumber = k + 1 //k=0のとき認証APIを指す
    var valueRowNumber = k + 2;//k=0のとき認証APIを指す
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);


    //認証、idの切り分け 
    //認証と非認証APIの切り分け
    // if (id == "認証"){
    //認証はリクエストしてレスポンスのaccess_tokenとuser_idを各種APIの該当項目に書き出し

    //リクエスト//認証APIはPOSTの想定
    //   if (method == "POST") {
    // const { getUrlReference,postMessageObj, postStatusCode } = sendPostRequest(sht, valueRowNumber , method);
    // sht.getRange(valueRowNumber, 6).setValue(getUrlReference);
    // sht.getRange(valueRowNumber, 7).setValue(postMessageObj);
    // sht.getRange(valueRowNumber, 8).setValue(postStatusCode);

    //}

    //該当レスポンスの取得
    //access_tokenは255文字なので、substring(indexOf(""access_token":""),indexOf(""access_token":"") + 255)
    //user_id substring(indexOf(""user_id":"),indexOf(","last_login_datetime"""))

    //書き出し（このブロックでは無理。値を保持しておく）
    //

    //非認証APIである各種APIはさらに認証要不要で分類される。
    //認証要の非認証APIは認証APIの情報書き出し。
    //if (認証要不要 == "必要"){

    //}

    //}//ーーーーα

    //認証不要の非認証APIはリクエスト
    //非認証APIについて
    if (id != "認証" && method == "POST") {
      if (auth == "不要" || auth == "不明") {
        outputPostToWritten(method, sht, keyRowNumber, valueRowNumber);
      }
      else { }




      //α参照
      //if (auth == "必要") {
      //書き出してからリクエスト
      //書き出すセルの場所を検索する必要がある。
      //      sht.getRange(valueRowNumber, ?).setValue(token);
      //    sht.getRange(valueRowNumber, ?).setValue(userid);
      //}

    }
    else {

    }
  }
}

//POSTしたデータをPUT、DELETE、GETするので、POST以外のメソッドの処理の優先順位を下げる。
//PUT,DELETE,GETのリクエスト実行、書き込み処理
function mainNoPost() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht = ss.getSheetByName('改修オートメーション');
  //test//Logger.log(sht.getLastRow()+100);　//100+4//行を取得できているかの確認


  //for (var i = 2; i < sht.getLastRow()+1 ;i = i +2) {}

  //認証情報を参照して認証必要なAPIのuser_idとaccess_tokenの列に認証結果を書き出す。
  //ここから記述する
  // var methodsArray = sht.getRange(2, 2, sht.getLastRow() - 1).getValues();
  //var idsArray = sht.getRange(2, 1, sht.getLastRow() - 1).getValues();


  //認証不要のときは、リクエスト処理を行う。
  //ここから記述する



  //    
  //複数セルの値を二次元配列として取得する - getValues -
  var methodsArray = sht.getRange(1, 2, sht.getLastRow()).getValues();
  var idsArray = sht.getRange(1, 1, sht.getLastRow()).getValues();
  var authsArray = sht.getRange(1, 4, sht.getLastRow()).getValues();
  //Logger.log(methodsArray);//[[method],[POST],[method],[POST],[method], [POST],[method], [GET]]
  //Logger.log(idsArray);//[[No],[認証],[No], [1.0],[No], [2.0],[No], [3.0]] 


  //取得した二次元配列：methodsArrayとidsArrayを１次配列に変換する
  var methodArray = methodsArray.flat();
  var idArray = idsArray.flat();
  var authArray = authsArray.flat();

  // Logger.log(methodArray);

  //test　//id取得できているかの確認、後に行数に変換する.valueRowNumberとして定義する。
  //Logger.log(idArray);//[No,認証,No,1.0,No, 2.0,No, 3.0] //[0][2][4][6]・・・・がkeyでmainでは不要（sendXXで必要）//[1][3][5]・・・・がvalueで必要
  var id = 0;
  var method = "";
  var auth = "";

  for (var k = 0; k < idArray.length; k = k + 2) {
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    var arrayNo = k + 1
    var id = idArray[arrayNo];
    var method = methodArray[arrayNo];
    var auth = authArray[arrayNo];
    //検証//Logger.log(method);//	[POST, POST, POST, GET]
    //検証//Logger.log(id);//[認証,1,2,3]

    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    var keyRowNumber = k + 1 //k=0のとき認証APIを指す
    var valueRowNumber = k + 2;//k=0のとき認証APIを指す
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);


    //認証、idの切り分け 
    //認証と非認証APIの切り分け
    // if (id == "認証"){
    //認証はリクエストしてレスポンスのaccess_tokenとuser_idを各種APIの該当項目に書き出し

    //リクエスト//認証APIはPOSTの想定
    //   if (method == "POST") {
    // const { getUrlReference,postMessageObj, postStatusCode } = sendPostRequest(sht, valueRowNumber , method);
    // sht.getRange(valueRowNumber, 6).setValue(getUrlReference);
    // sht.getRange(valueRowNumber, 7).setValue(postMessageObj);
    // sht.getRange(valueRowNumber, 8).setValue(postStatusCode);

    //}

    //該当レスポンスの取得
    //access_tokenは255文字なので、substring(indexOf(""access_token":""),indexOf(""access_token":"") + 255)
    //user_id substring(indexOf(""user_id":"),indexOf(","last_login_datetime"""))

    //書き出し（このブロックでは無理。値を保持しておく）
    //

    //非認証APIである各種APIはさらに認証要不要で分類される。
    //認証要の非認証APIは認証APIの情報書き出し。
    //if (認証要不要 == "必要"){

    //}

    //}//ーーーーα

    //認証不要の非認証APIはリクエスト
    //非認証APIについて
    if (id != "認証" && method != "POST") {
      if (auth == "不要" || auth == "不明") {
        outputNoPostToWritten(method, sht, keyRowNumber, valueRowNumber);
      }

      else {
    
      }




      //α参照
      //if (auth == "必要") {
      //書き出してからリクエスト
      //書き出すセルの場所を検索する必要がある。
      //      sht.getRange(valueRowNumber, ?).setValue(token);
      //    sht.getRange(valueRowNumber, ?).setValue(userid);
      //}

    }
    else {


    }
  }
}


function outputPostToWritten(method, sht, keyRowNumber, valueRowNumber) {

  if (method == "POST") {
    const { postString, postMessage, postStatusCode } = sendPostRequest(sht, keyRowNumber, valueRowNumber, method);
    sht.getRange(valueRowNumber, 6).setValue(postString);
    sht.getRange(valueRowNumber, 7).setValue(postMessage);
    sht.getRange(valueRowNumber, 8).setValue(postStatusCode);

  }

  else {

  }
}


function outputNoPostToWritten(method, sht, keyRowNumber, valueRowNumber) {

  if (method == "GET") {
    //sendGetRequest()の返却値：レスポンスメッセージとステータスコードをメイン関数で再利用する。
    //書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
    const { getUrlReference, getMessage, getStatusCode } = sendGetRequest(sht, keyRowNumber, valueRowNumber, method);
    sht.getRange(valueRowNumber, 6).setValue(getUrlReference);
    sht.getRange(valueRowNumber, 7).setValue(getMessage);
    sht.getRange(valueRowNumber, 8).setValue(getStatusCode);


    //検証//console.log(getStatusCode); 
  }

  else if (method == "PUT") {
    const { putString, putMessage, putStatusCode } = sendPutRequest(sht, keyRowNumber, valueRowNumber, method);
    sht.getRange(valueRowNumber, 6).setValue(putString);
    sht.getRange(valueRowNumber, 7).setValue(putMessage);
    sht.getRange(valueRowNumber, 8).setValue(putStatusCode);

  }

  else if (method == "DELETE") {
    const { deleteString, deleteMessage, deleteStatusCode } = sendDeleteRequest(sht, keyRowNumber, valueRowNumber, method);
    sht.getRange(valueRowNumber, 6).setValue(deleteString);
    sht.getRange(valueRowNumber, 7).setValue(deleteMessage);
    sht.getRange(valueRowNumber, 8).setValue(deleteStatusCode);

  }

  else {

  }
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendGetRequest(sht, keyRowNumber, valueRowNumber, method) {
  //getvalueでgetrangeの値を取得
  // セル範囲の値（ここではURL）を２次元配列で取得する
  var urlValue = sht.getRange(valueRowNumber, 10).getValues();
  // セル範囲の値（ここではURL）を１次元配列に変換する
  var urlValuesFlat = urlValue.flat()
  //検証//console.log(valuesFlat)

  //ログ排出用にapi,組み合わせカラムのバリューを取得する。
  var getApi = sht.getRange(valueRowNumber, 3).getValue();
  var getCombi = sht.getRange(valueRowNumber, 5).getValue();
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
    var getMessagePre = response.getContentText();

    var getMessageJson = JSON.stringify(JSON.parse(getMessagePre));
    var searchKey0 = /[,]/g;
    var getMessage = getMessageJson.replace(searchKey0, ",\n");

    //リクエスト出力のため改行して整形
    var searchKey1 = /[?]/g;
    var arrangedUrl1 = getUrl.replace(searchKey1, "?\n");
    var searchKey2 = /[&]/g;
    var arrangedUrl2 = getUrl.replace(searchKey2, "&\n");
    var getUrlReference = arrangedUrl2;

    console.log(method + " " + getApi + "についてaccess_tokenとuser_idの組み合わせが" + getCombi + "の時、リクエストとレスポンスは以下の通りです。\n※URL、レスポンスメッセージ、ステータスコードの順に記載")
    //console.log(getUrl);
    console.log(getUrlReference);
    console.log(getMessage);
    console.log(getStatusCode);

    return { getUrlReference, getMessage, getStatusCode };

  }
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendPostRequest(sht, keyRowNumber, valueRowNumber, method) {
  //パラメータのkeyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(keyRowNumber, 11, 1, sht.getLastColumn() - 10)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  var keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(valueRowNumber, 11, 1, sht.getLastColumn() - 10)
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
    const urlPost2Array = sht.getRange(valueRowNumber, 10).getValues();  // G列:URLカラムの全ての行を取得

    //URLの二次元配列を一次元配列に変換
    var urlPostFlat = urlPost2Array.flat()
    //console.log(urlPostFlat);

    //URL配列からURLを抽出
    var urlPost = urlPostFlat[0];
  }


  //key-value配列のJSON化
  var postString = JSON.stringify(obj, null, "\t")


  //POSTリクエスト　parameter
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': postString,
    "muteHttpExceptions": true,
  };

  //POSTリクエスト　urlと合わせて送信
  var postResponse = UrlFetchApp.fetch(urlPost, options);

  //ログ排出用にapi,組み合わせカラムのバリューを取得する。
  var postApi = sht.getRange(valueRowNumber, 3).getValue();
  var postCombi = sht.getRange(valueRowNumber, 5).getValue();

  //レスポンス
  var postMessagePre = postResponse.getContentText();
  var postMessageJson = JSON.stringify(JSON.parse(postMessagePre));
  var searchKey0 = /[,]/g;
  var postMessage = postMessageJson.replace(searchKey0, ",\n");


  var postStatusCode = String(postResponse.getResponseCode());
  //console.log(response.getContentText())
  //console.log(response.getResponseCode())
  //テスト//Logger.log(idArray);

  //リクエストとレスポンスのログ排出
  console.log(method + " " + postApi + "についてaccess_tokenとuser_idの組み合わせが" + postCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
  console.log(urlPost);
  console.log(postString);
  console.log(postMessage);
  console.log(postStatusCode);
  return { postString, postMessage, postStatusCode };
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendPutRequest(sht, keyRowNumber, valueRowNumber, method) {
  //パラメータの一覧keyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(keyRowNumber, 11, 1, sht.getLastColumn() - 10)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  var keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(valueRowNumber, 11, 1, sht.getLastColumn() - 10)
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
    const urlPut2Array = sht.getRange(valueRowNumber, 10).getValues();  // G列:URLカラムの全ての行を取得

    //URLの二次元配列を一次元配列に変換
    var urlPutFlat = urlPut2Array.flat()
    //console.log(urlPostFlat);

    //URL配列からURLを抽出
    var urlPut = urlPutFlat[0];
  }


  //key-value配列のJSON化
  var putString = JSON.stringify(obj, null, "\t")

  //PUTリクエスト　parameter
  var options = {
    'method': 'put',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': putString,
    "muteHttpExceptions": true,
  };

  //PUTリクエスト　urlと合わせて送信
  var putResponse = UrlFetchApp.fetch(urlPut, options);

  //ログ排出用にapi,組み合わせカラムのバリューを取得する。
  var putApi = sht.getRange(valueRowNumber, 3).getValue();
  var putCombi = sht.getRange(valueRowNumber, 5).getValue();

  //レスポンス
  var putMessagePre = putResponse.getContentText();
  var putMessageJson = JSON.stringify(JSON.parse(putMessagePre));

  var searchKey0 = /[,]/g;
  var putMessage = putMessageJson.replace(searchKey0, ",\n");

  var putStatusCode = String(putResponse.getResponseCode());
  //console.log(response.getContentText())
  //console.log(response.getResponseCode())
  //テスト//Logger.log(idArray);

  //リクエストとレスポンスのログ排出
  console.log(method + " " + putApi + "についてaccess_tokenとuser_idの組み合わせが" + putCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
  console.log(urlPut);
  console.log(putString);
  console.log(putMessage);
  console.log(putStatusCode);
  return { putString, putMessage, putStatusCode };
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendDeleteRequest(sht, keyRowNumber, valueRowNumber, method) {
  //パラメータの一覧keyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(keyRowNumber, 11, 1, sht.getLastColumn() - 10)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  var keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(valueRowNumber, 11, 1, sht.getLastColumn() - 10)
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
    const urlDelete2Array = sht.getRange(valueRowNumber, 10).getValues();  // G列:URLカラムの全ての行を取得

    //URLの二次元配列を一次元配列に変換
    var urlDeleteFlat = urlDelete2Array.flat()
    //console.log(urlDeleteFlat);

    //URL配列からURLを抽出
    var urlDelete = urlDeleteFlat[0];
  }


  //key-value配列のJSON化
  var deleteString = JSON.stringify(obj, null, "\t")

  //DELETEリクエスト　parameter
  var options = {
    'method': 'delete',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': deleteString,
    "muteHttpExceptions": true,
  };

  //DELETEリクエスト　urlと合わせて送信
  var deleteResponse = UrlFetchApp.fetch(urlDelete, options);

  //ログ排出用にapi,組み合わせカラムのバリューを取得する。
  var deleteApi = sht.getRange(valueRowNumber, 3).getValue();
  var deleteCombi = sht.getRange(valueRowNumber, 5).getValue();

  //レスポンス
  var deleteMessagePre = deleteResponse.getContentText();
  var deleteMessageJson = JSON.stringify(JSON.parse(deleteMessagePre));
  var deleteStatusCode = String(deleteResponse.getResponseCode());
  var searchKey0 = /[,]/g;
  var deleteMessage = deleteMessageJson.replace(searchKey0, ",\n");

  //console.log(response.getContentText())
  //console.log(response.getResponseCode())
  //テスト//Logger.log(idArray);

  //リクエストとレスポンスのログ排出
  console.log(method + " " + deleteApi + "についてaccess_tokenとuser_idの組み合わせが" + deleteCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
  console.log(urlDelete);
  console.log(deleteString);
  console.log(deleteMessage);
  console.log(deleteStatusCode);
  return { deleteString, deleteMessage, deleteStatusCode };
}


  //テスト//Logger.log(idArray);
