function updateTrigger() {
  //Deletes all triggers in the current project.//過去のトリガーが残っているときに利用する
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


function main() {
  mainAuthPost();
  mainPost();
  mainNoPost();
  //var { authId, authMethod, authAuth, sht,authKeyRowNumber, authValueRowNumber } = mainPost();
  //var { authId, authMethod, authAuth, sht,authKeyRowNumber, authValueRowNumber } = mainNoPost();
}


//if(auth == "必要")　//のときkey取得。
//配列から要素抽出
//if(key == "access_token" ) //のときvalueに認証APIの値を書き出し
//認証APIを参照して参照が必要なAPIにはPOSTするので認証APIの処理の優先順位を上げる。
//リクエスト処理の実行、書き込む処理はmainAuthPost()で行う。sendXX()はリクエストとログ書き出し。
function mainAuthPost() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht = ss.getSheetByName('改修オートメーション認証あり');
  //test//Logger.log(sht.getLastRow()+100);　//100+4//行を取得できているかの確認


  //    
  //複数セルの値を二次元配列として取得する - getValues -
  var methodsArray = sht.getRange(1, 2, sht.getLastRow()).getValues();
  var idsArray = sht.getRange(1, 1, sht.getLastRow()).getValues();
  var authsArray = sht.getRange(1, 4, sht.getLastRow()).getValues();
  //Logger.log(methodsArray);//[[method],[POST],[method],[POST],[method], [POST],[method], [GET]]
  //Logger.log(idsArray);//[[No],[認証],[No], [1.0],[No], [2.0],[No], [3.0]] 


  //取得した二次元配列：methodsArrayとidsArrayを１次配列に変換する
  var authMethodArray = methodsArray.flat();
  var authIdArray = idsArray.flat();
  var authArray = authsArray.flat();

  // Logger.log(authMethodArray);

  //test　//id取得できているかの確認、後に行数に変換する.valueRowNumberとして定義する。
  //Logger.log(authIdArray);//[No,認証,No,1.0,No, 2.0,No, 3.0] //[0][2][4][6]・・・・がkeyでmainでは不要（sendXXで必要）//[1][3][5]・・・・がvalueで必要


  var authId = 0;
  var authMethod = "";
  var authAuth = "";

  //k

  for (var k = 0; k < authIdArray.length; k = k + 2) {
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    var authArrayNo = k + 1
    var authId = authIdArray[authArrayNo];
    var authMethod = authMethodArray[authArrayNo];
    var authAuth = authArray[authArrayNo];
    //検証//Logger.log(method);//	[POST, POST, POST, GET]
    //検証//Logger.log(id);//[認証,1,2,3]

    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    var authKeyRowNumber = k + 1 //k=0のとき認証APIを指す
    var authValueRowNumber = k + 2;//k=0のとき認証APIを指す
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);


    //認証、idの切り分け 
    //認証と非認証APIの切り分け
    if (authId.indexOf("認証") != -1 && authMethod == "POST") {
      outputAuthPostToWritten(authId, authMethod, sht, authKeyRowNumber, authValueRowNumber);

    }
    else {

    }


    //return { authId, authMethod, authAuth, sht, authKeyRowNumber, authValueRowNumber };
  }
}




//POSTしたデータをPUT、DELETE、GETするので、POST処理の優先順位を上げる。
//リクエスト処理の実行、書き込む処理はmainPost()で行う。sendXX()はリクエストとログ書き出し。
function mainPost(authId, authMethod, authAuth, sht, authKeyRowNumber, authValueRowNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht = ss.getSheetByName('改修オートメーション認証あり');
  //test//Logger.log(sht.getLastRow()+100);　//100+4//行を取得できているかの確認



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

  //k

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

    //認証と非認証の切り分け
    //認証不要APIについて
    if (id.indexOf("認証") == -1 && method == "POST") {

      //認証不要の場合、リクエスト処理
      if (auth == "不要" || auth == "不明") {
        outputPostToWritten(id, method, sht, keyRowNumber, valueRowNumber);
      }
      else { }


      //認証必要APIについて
      if (auth.lastIndexOf("認証") != -1) {
        //referenceAuth()とmainPost()は依存関係にない
        referenceAuth();
        outputPostToWritten(id, method, sht, keyRowNumber, valueRowNumber);



      }
      else {

      }
    }
  }
}


//POSTしたデータをPUT、DELETE、GETするので、POST以外のメソッドの処理の優先順位を下げる。
//PUT,DELETE,GETのリクエスト実行、書き込み処理
function mainNoPost(authId, authMethod, authAuth, sht, authKeyRowNumber, authValueRowNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht = ss.getSheetByName('改修オートメーション認証あり');
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

    //認証APIはPOSTのみの想定。GET,PUT,DELETEの認証APIが必要な場合は別途記述する必要あり。

    //認証不要の非認証APIはリクエスト
    //非認証APIについて
    if (id.indexOf("認証") == -1 && method != "POST") {
      if (auth == "不要" || auth == "不明") {
        outputNoPostToWritten(id, method, sht, keyRowNumber, valueRowNumber);
      }

      else {

      }

      //認証必要APIについて
      if (auth.lastIndexOf("認証") != -1) {
        //referenceAuth()とmainNoPost()は依存関係にない
        referenceAuth();
        outputNoPostToWritten(id, method, sht, keyRowNumber, valueRowNumber);



      }
      else {

      }


    }
    else {


    }
  }
}


//認証APIのレスポンスを認証APIのvalueに書き出した結果を参照APIのvalueに書き出し
//他のoutputxxx()と違いmainXXX()と依存関係にない点に注意。引数渡すと思っている挙動はできない。
function referenceAuth() {

  //参照するAPIを選択する（改修予定）//"必要"部分を変数化する必要がある参照

  //認証APIをループで参照しつつ、参照APIを参照しながら書き換えていくので二重ループ
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht = ss.getSheetByName('改修オートメーション認証あり');

  //複数セルの値を二次元配列として取得する - getValues -
  var referMethodsArray = sht.getRange(1, 2, sht.getLastRow()).getValues();
  var referIdsArray = sht.getRange(1, 1, sht.getLastRow()).getValues();
  var referAuthsArray = sht.getRange(1, 4, sht.getLastRow()).getValues();
  //Logger.log(referMethodsArray);//[[method],[POST],[method],[POST],[method], [POST],[method], [GET]]
  //Logger.log(referIdsArray);//[[No],[認証],[No], [1.0],[No], [2.0],[No], [3.0]] 


  //取得した二次元配列：methodsArrayとidsArrayを１次配列に変換する
  var referMethodArray = referMethodsArray.flat();
  var referIdArray = referIdsArray.flat();
  var referAuthArray = referAuthsArray.flat();

  // Logger.log(methodArray);

  //test　//id取得できているかの確認、後に行数に変換する.valueRowNumberとして定義する。
  //Logger.log(idArray);//[No,認証,No,1.0,No, 2.0,No, 3.0] //[0][2][4][6]・・・・がkeyでmainでは不要（sendXXで必要）//[1][3][5]・・・・がvalueで必要


  var referId = 0;
  var referMethod = "";
  var referAuth = "";

  //認証APIを参照する
  for (var m = 0; m < referIdArray.length; m = m + 2) {
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    var referArrayNo = m + 1
    var referId = referIdArray[referArrayNo];
    var referMethod = referMethodArray[referArrayNo];
    var referAuth = referAuthArray[referArrayNo];
    //検証//Logger.log(referMethod);//	[POST, POST, POST, GET]
    //検証//Logger.log(referId);//[認証,1,2,3]

    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    var referKeyRowNumber = m + 1 //k=0のとき認証APIを指す
    var referValueRowNumber = m + 2;//k=0のとき認証APIを指
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);

    console.log(referMethod);
    console.log(referAuth);
    console.log(referId);

    //認証APIのkey-valueを抽出。keyはあとで条件一致の時に使う。”返却された”という文字以外を抽出した
    if (referId.indexOf("認証") != -1) {

      var referValueAccessToken = sht.getRange(referValueRowNumber, 11).getValue();
      var referValueUserId = sht.getRange(referValueRowNumber, 12).getValue();

      var referKeyAccessTokenPre = sht.getRange(referKeyRowNumber, 11).getValue();
      var referKeyUserIdPre = sht.getRange(referKeyRowNumber, 12).getValue();

      var referKeyAccessToken = referKeyAccessTokenPre.substring(referKeyAccessTokenPre.indexOf("access_token"), referKeyAccessTokenPre.length);
      var referKeyUserId = referKeyUserIdPre.substring(referKeyUserIdPre.indexOf("user_id"), referKeyUserIdPre.length);
      console.log(referValueUserId);
      console.log(referValueAccessToken);
      console.log(referKeyUserId);
      console.log(referKeyAccessToken);
    }

    else {
    }

    //参照APIを参照//認証APIでvalueに値がある時に、参照APIが選択した認証APIのvalueを参照APIのkey-valueに書き込み
    for (var n = 0; n < referIdArray.length; n = n + 2) {
      var referredAuths2Array = sht.getRange(1, 4, sht.getLastRow()).getValues();
      var referredArrayNo = n + 1
      var referredAuth1Array = referredAuths2Array.flat();
      var referredAuth = referredAuth1Array[referredArrayNo];


      console.log(referredAuth + "aaaaa");
      if (referValueAccessToken != "" || referValueUserId != "") {

        console.log(referId);
        if (referredAuth == referId) {
          var referredKeyRowNumber = n + 1 //k=0のとき認証APIを指すが、条件ではじかれる
          var referredValueRowNumber = n + 2;//k=0のとき認証APIを指すが条件ではじかれる

          //参照APIのパラメータkey取得
          var rangeParam = sht.getRange(referredKeyRowNumber, 11, 1, sht.getLastColumn() - 10)
          const referredKeys2array = rangeParam.getValues()
          var referredKey1array = referredKeys2array.flat()
          console.log(referredKeyRowNumber);
          console.log(referredKey1array);



          for (count = 0; count < referredKey1array.length; count++) {
            console.log(referredKey1array[count]);
            if (referValueUserId != "" && referredKey1array[count] == referKeyUserId) {
              console.log(referredKey1array[count] + "bbbb");
              sht.getRange(referredValueRowNumber, 11 + count).setValue(referValueUserId);
            }
            else if (referValueAccessToken != "" && referredKey1array[count] == referKeyAccessToken) {
              sht.getRange(referredValueRowNumber, 11 + count).setValue(referValueAccessToken);
            }
            else {
            }
          }
        }
        else {
        }
      }
      else {
      }

    }

  }
}




//認証APIのレスポンスを認証APIのvalueに書き出し
function outputAuthPostToWritten(authId, authMethod, sht, authKeyRowNumber, authValueRowNumber) {

  if (authMethod == "POST" && authId.indexOf("認証") != -1) {
    //認証APIのレスポンスを認証APIのvalueに書き出し
    const { authPostString, authPostMessage, authPostStatusCode, authAccessToken, authUserId } = sendAuthPostRequest(sht, authKeyRowNumber, authValueRowNumber, authMethod);
    sht.getRange(authValueRowNumber, 6).setValue(authPostString);
    sht.getRange(authValueRowNumber, 7).setValue(authPostMessage);
    sht.getRange(authValueRowNumber, 8).setValue(authPostStatusCode);

    sht.getRange(authValueRowNumber, 11).setValue(authAccessToken);
    sht.getRange(authValueRowNumber, 12).setValue(authUserId);


  }
  //認証APIはPOSTの想定。他メソッドでの認証は想定していない。
  else {
  }
}



function outputPostToWritten(id, method, sht, keyRowNumber, valueRowNumber) {

  if (method == "POST" && id.indexOf("認証") == -1) {
    const { postString, postMessage, postStatusCode } = sendPostRequest(sht, keyRowNumber, valueRowNumber, method);
    sht.getRange(valueRowNumber, 6).setValue(postString);
    sht.getRange(valueRowNumber, 7).setValue(postMessage);
    sht.getRange(valueRowNumber, 8).setValue(postStatusCode);

  }

  else {

  }
}


function outputNoPostToWritten(id, method, sht, keyRowNumber, valueRowNumber) {

  if (method == "GET" && id.indexOf("認証") == -1) {
    //sendGetRequest()の返却値：レスポンスメッセージとステータスコードをメイン関数で再利用する。
    //書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
    const { getUrlReference, getMessage, getStatusCode } = sendGetRequest(sht, keyRowNumber, valueRowNumber, method);
    sht.getRange(valueRowNumber, 6).setValue(getUrlReference);
    sht.getRange(valueRowNumber, 7).setValue(getMessage);
    sht.getRange(valueRowNumber, 8).setValue(getStatusCode);


    //検証//console.log(getStatusCode); 
  }

  else if (method == "PUT" && id.indexOf("認証") == -1) {
    const { putString, putMessage, putStatusCode } = sendPutRequest(sht, keyRowNumber, valueRowNumber, method);
    sht.getRange(valueRowNumber, 6).setValue(putString);
    sht.getRange(valueRowNumber, 7).setValue(putMessage);
    sht.getRange(valueRowNumber, 8).setValue(putStatusCode);

  }

  else if (method == "DELETE" && id.indexOf("認証") == -1) {
    const { deleteString, deleteMessage, deleteStatusCode } = sendDeleteRequest(sht, keyRowNumber, valueRowNumber, method);
    sht.getRange(valueRowNumber, 6).setValue(deleteString);
    sht.getRange(valueRowNumber, 7).setValue(deleteMessage);
    sht.getRange(valueRowNumber, 8).setValue(deleteStatusCode);

  }

  else {

  }
}



//認証APIのPOSTの書き込む処理は別関数で行う。sendXX()はリクエストとログ書き出し。
function sendAuthPostRequest(sht, authKeyRowNumber, authValueRowNumber, authMethod) {
  //パラメータのkeyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(authKeyRowNumber, 13, 1, sht.getLastColumn() - 12)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  var keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(authValueRowNumber, 13, 1, sht.getLastColumn() - 12)
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
    const urlPost2Array = sht.getRange(authValueRowNumber, 10).getValues();  // G列:URLカラムの全ての行を取得

    //URLの二次元配列を一次元配列に変換
    var urlPostFlat = urlPost2Array.flat()
    //console.log(urlPostFlat);

    //URL配列からURLを抽出
    var urlPost = urlPostFlat[0];
  }


  //key-value配列のJSON化
  var authPostString = JSON.stringify(obj, null, "\t")


  //POSTリクエスト　parameter
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': authPostString,
    "muteHttpExceptions": true,
  };

  //POSTリクエスト　urlと合わせて送信
  var authPostResponse = UrlFetchApp.fetch(urlPost, options);

  //ログ排出用にapi,組み合わせカラムのバリューを取得する。
  var authPostApi = sht.getRange(authValueRowNumber, 3).getValue();
  var authPostCombi = sht.getRange(authValueRowNumber, 5).getValue();

  //レスポンス
  var authPostMessagePre = authPostResponse.getContentText();
  var authPostMessageObj = JSON.parse(authPostMessagePre);
  var authPostMessageJson = JSON.stringify(authPostMessageObj);
  var searchKey0 = /[,]/g;
  var authPostMessage = authPostMessageJson.replace(searchKey0, ",\n");

  var authPostStatusCode = String(authPostResponse.getResponseCode());
  //console.log(response.getContentText())
  //console.log(response.getResponseCode())
  //テスト//Logger.log(idArray);

  //認証情報を参照して認証必要なAPIのuser_idとaccess_tokenの列に認証結果を書き出す。
  // レスポンスメッセージをJavaScriptオブジェクトの状態でkeyでフィルタリング可能
  //objectを文字列化して書き出す
  var authAccessToken = String(authPostMessageObj.access_token);
  var authUserId = String(authPostMessageObj.user_id);

  //リクエストとレスポンスのログ排出
  console.log(authMethod + " " + authPostApi + "についてaccess_tokenとuser_idの組み合わせが" + authPostCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
  console.log(urlPost);
  console.log(authPostString);
  console.log(authPostMessage);

  console.log(authPostStatusCode);
  console.log(authAccessToken);
  console.log(authUserId);
  return { authPostString, authPostMessage, authPostStatusCode, authAccessToken, authUserId };

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