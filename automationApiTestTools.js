//スクリプトを編集するたびにupdateTrigger()実行を推奨
function updateTrigger() {
  //Deletes all triggers in the current project.//過去のトリガーが残っているときに利用する
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {

    ScriptApp.deleteTrigger(triggers[i]);
  }

  //有効なトリガーを取得する
  Logger.log("トリガーを更新します。");
  Logger.log('Current project has ' + ScriptApp.getProjectTriggers().length + ' triggers.');

  //スプシ起動時にステータスコード書き込みを自動実行するトリガーを作成
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const autoRun = ScriptApp.newTrigger("main").forSpreadsheet(ssId).onOpen().create();

  Logger.log('スプレッドシート起動時に実行するトリガーを作成しました。');
}


function main() {
  mainAuthPost();
  mainPost();
  mainNoPost();
}

//sendAuthPost()関数：シートを取得して認証APIのリクエストおよびレスポンスの書き出しのための準備。および書き出す関数の呼び出しを行う
//認証APIを参照して参照が必要なAPIにはPOSTするので認証APIの処理の優先順位を上げている。
//outputAuthPostWritten():sendAuthPostRequest()の大元。
//outputXX():リクエストとレスポンスの書き込み、sendXX():リクエスト
function mainAuthPost() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sht = ss.getSheetByName('sample');
  //test//Logger.log(sht.getLastRow()+100);　//100+4//行を取得できているかの確認

  //シートの値を二次元配列として取得する。
  //APIのNo=ID、メソッド、認証APIの要不要について
  const methodsArray = sht.getRange(1, 2, sht.getLastRow()).getValues();
  const idsArray = sht.getRange(1, 1, sht.getLastRow()).getValues();
  const authsArray = sht.getRange(1, 4, sht.getLastRow()).getValues();
  //Logger.log(methodsArray);//[[method],[POST],[method],[POST],[method], [POST],[method], [GET]]
  //Logger.log(idsArray);//[[No],[認証],[No], [1.0],[No], [2.0],[No], [3.0]] 

  //取得した2次元配列を１次元配列に変換する
  const authMethodArray = methodsArray.flat();
  const authIdArray = idsArray.flat();
  const authArray = authsArray.flat();

  // Logger.log(authMethodArray);

  //test　//id取得できているかの確認、後に行数に変換する.valueRowNumberとして定義する。
  //Logger.log(authIdArray);
  //[No,認証,No,1.0,No, 2.0,No, 3.0] //[0][2][4][6]・・・・がkeyでmainでは不要（sendXXで必要）//[1][3][5]・・・・がvalueで必要

  for (let k = 0; k < authIdArray.length; k = k + 2) {

    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    const authKeyRowNumber = k + 1 //k=0のとき認証APIを指す
    const authValueRowNumber = k + 2;//k=0のとき認証APIを指す
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);

    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    const authId = authIdArray[authKeyRowNumber];
    const authMethod = authMethodArray[authKeyRowNumber];
    const authAuth = authArray[authKeyRowNumber];
    //検証//Logger.log(method);//	[POST, POST, POST, GET]
    //検証//Logger.log(id);//[認証,1,2,3]


    //認証と非認証APIの切り分け
    //認証APIはPOSTメソッドの想定
    if (authId.indexOf("認証") != -1 && authMethod == "POST") {

      //認証APIならリクエストとレスポンスを認証APIの該当シートに書き出し
      //※レスポンスの書き出しは、access_tokenとuser_idにのみ対応
      outputAuthPostToWritten(authId, authMethod, sht, authKeyRowNumber, authValueRowNumber);

    }
    else {
    }
  }
}

//mainPost():シートの値を取得してPOSTメソッドのリクエストと書き出しの準備。referenceAuth();とoutputPostToWritten()を呼び出す。
//POSTにより登録したデータをPUT、DELETE、GETできるので、POST処理の優先順位を上げている。
//outputXXX()はsendXX()を呼び出す。sendXX()はリクエストを行い、outputXX()でリクエスト、レスポンスを書き出し
//referenceXX()は認証APIを参照する関数。認証APIのaccess_tokenとuser_idを対象の参照APIに書き出してからリクエスト処理を行えるようにする
function mainPost() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sht = ss.getSheetByName('sample');
  //test//Logger.log(sht.getLastRow()+100);　//100+4//行を取得できているかの確認

  //シートの値を二次元配列として取得する。
  //APIのNo=ID、メソッド、認証APIの要不要について
  const methodsArray = sht.getRange(1, 2, sht.getLastRow()).getValues();
  const idsArray = sht.getRange(1, 1, sht.getLastRow()).getValues();
  const authsArray = sht.getRange(1, 4, sht.getLastRow()).getValues();
  //Logger.log(methodsArray);//[[method],[POST],[method],[POST],[method], [POST],[method], [GET]]
  //Logger.log(idsArray);//[[No],[認証],[No], [1.0],[No], [2.0],[No], [3.0]] 

  //取得した2次元配列を１次元配列に変換する
  const methodArray = methodsArray.flat();
  const idArray = idsArray.flat();
  const authArray = authsArray.flat();

  // Logger.log(methodArray);

  //test　//id取得できているかの確認、後に行数に変換する.valueRowNumberとして定義する。
  //Logger.log(idArray);//[No,認証,No,1.0,No, 2.0,No, 3.0] //[0][2][4][6]・・・・がkeyでmainでは不要（sendXXで必要）//[1][3][5]・・・・がvalueで必要

  for (let k = 0; k < idArray.length; k = k + 2) {
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    const keyRowNumber = k + 1 //k=0のとき認証APIを指す
    const valueRowNumber = k + 2;//k=0のとき認証APIを指す
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);

    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    const id = idArray[keyRowNumber];
    const method = methodArray[keyRowNumber];
    const auth = authArray[keyRowNumber];
    //検証//Logger.log(method);//	[POST, POST, POST, GET]
    //検証//Logger.log(id);//[認証,1,2,3]


    //認証と参照の切り分け.参照APIのみリクエストして、リクエストとレスポンスを書き出し
    if (id.indexOf("認証") == -1 && method == "POST") {

      //認証API不要の場合、そのままリクエスト処理
      if (auth == "不要" || auth == "不明") {
        outputPostToWritten(id, method, sht, keyRowNumber, valueRowNumber);
      }

      //認証API必要の場合、referenceAuth()にて認証APIのレスポンスを対象の参照APIのキーパラメータに上書きする
      else if (auth.indexOf("認証") != -1) {
        //referenceAuth()とmainPost()は依存関係にない
        referenceAuth();
        outputPostToWritten(id, method, sht, keyRowNumber, valueRowNumber);
      }
      else {
      }
    }
  }
}


//mainNoPost():シートの値を取得してGET,DELETE,PUTメソッドのリクエストと書き出しの準備。referenceAuth();とoutputNoPostToWritten()を呼び出す。
//POSTしたデータをPUT、DELETE、GETするので、POST以外のメソッドの処理の優先順位を下げている。
//outputXXX()はsendXX()を呼び出す。sendXX()はリクエストを行い、outputXX()でリクエスト、レスポンスを書き出し
//referenceXX()は認証APIを参照する関数。認証APIのaccess_tokenとuser_idを対象の参照APIに書き出してからリクエスト処理を行えるようにする
function mainNoPost() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sht = ss.getSheetByName('sample');
  //test//Logger.log(sht.getLastRow()+100);　//100+4//行を取得できているかの確認

  //シートの値を二次元配列として取得する。
  //APIのNo=ID、メソッド、認証APIの要不要について
  const methodsArray = sht.getRange(1, 2, sht.getLastRow()).getValues();
  const idsArray = sht.getRange(1, 1, sht.getLastRow()).getValues();
  const authsArray = sht.getRange(1, 4, sht.getLastRow()).getValues();
  //Logger.log(methodsArray);//[[method],[POST],[method],[POST],[method], [POST],[method], [GET]]
  //Logger.log(idsArray);//[[No],[認証],[No], [1.0],[No], [2.0],[No], [3.0]] 


  //取得した2次元配列を１次元配列に変換する
  const methodArray = methodsArray.flat();
  const idArray = idsArray.flat();
  const authArray = authsArray.flat();

  // Logger.log(methodArray);

  //test　//id取得できているかの確認、後に行数に変換する.valueRowNumberとして定義する。
  //Logger.log(idArray);
  //[No,認証,No,1.0,No, 2.0,No, 3.0] //[0][2][4][6]・・・・がkeyでmainでは不要（sendXXで必要）//[1][3][5]・・・・がvalueで必要

  for (let k = 0; k < idArray.length; k = k + 2) {

    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    const keyRowNumber = k + 1 //k=0のとき認証APIを指す
    const valueRowNumber = k + 2;//k=0のとき認証APIを指す
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);


    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    const id = idArray[keyRowNumber];
    const method = methodArray[keyRowNumber];
    const auth = authArray[keyRowNumber];
    //検証//Logger.log(method);//	[POST, POST, POST, GET]
    //検証//Logger.log(id);//[認証,1,2,3]

    //認証APIはPOSTのみの想定。GET,PUT,DELETEの認証APIが必要な場合は別途記述する必要あり。
    //認証と参照の切り分け.参照APIのみリクエストして、リクエストとレスポンスを書き出し
    if (id.indexOf("認証") == -1 && method != "POST") {

      //認証API不要の場合、そのままリクエスト処理
      if (auth == "不要" || auth == "不明") {
        outputNoPostToWritten(id, method, sht, keyRowNumber, valueRowNumber);
      }

      else {

      }

      //認証API必要の場合、referenceAuth()にて認証APIのレスポンスを対象の参照APIのキーパラメータに上書きする
      if (auth.indexOf("認証") != -1) {
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sht = ss.getSheetByName('sample');

  //シートの値を参照用に取得
  const maxRow = sht.getLastRow()
  const maxCol = sht.getLastColumn()
  // シート全範囲の値
  const arraySheet = sht.getRange(1, 1, maxRow, maxCol).getValues()

  const referAuthArray = arraySheet.map(item => item[3]);
  const referIdArray = arraySheet.map(item => item[0]);
  const referMethodArray = arraySheet.map(item => item[1]);
  //Logger.log(referMethodsArray);//[[method],[POST],[method],[POST],[method], [POST],[method], [GET]]
  //Logger.log(referIdsArray);//[[No],[認証],[No], [1.0],[No], [2.0],[No], [3.0]] 

  const referredAuth1Array = arraySheet.map(item => item[3]);
  const referredId1Array = arraySheet.map(item => item[0]);

  //認証APIの行に関する情報を取得する
  for (let m = 0; m < referIdArray.length; m = m + 2) {

    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    const referKeyRowNumber = m + 1 //k=0のとき認証APIを指す
    const referValueRowNumber = m + 2;//k=0のとき認証APIを指
    //検証//Logger.log(valueRowNumber+10);

    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    const referId = referIdArray[referKeyRowNumber];
    const referMethod = referMethodArray[referKeyRowNumber];
    const referAuth = referAuthArray[referKeyRowNumber];
    //検証//Logger.log(referMethod);//	[POST, POST, POST, GET]
    //検証//Logger.log(referId);//[認証,1,2,3]


    //検証//console.log(referMethod);
    //検証//console.log(referAuth);
    //検証//console.log(referId);
    console.log("認証API(No." + referId + ")に対応する参照APIの有無を調べます。")

    //認証APIのkey-valueを抽出。keyはあとで条件一致の時に使う。”返却された”という文字以外を抽出した
    if (referId.indexOf("認証") != -1) {
      console.log("認証API(No." + referId + ")のaccess_tokenおよびuser_idを参照APIに渡すために取得します。\n取得結果は以下の通りです。")
      var referValueAccessToken = arraySheet[referValueRowNumber - 1][10];
      var referValueUserId = arraySheet[referValueRowNumber - 1][11];

      var referKeyAccessTokenPre = arraySheet[referKeyRowNumber - 1][10]
      var referKeyUserIdPre = arraySheet[referKeyRowNumber - 1][11]

      var referKeyAccessToken = referKeyAccessTokenPre.substring(referKeyAccessTokenPre.indexOf("access_token"), referKeyAccessTokenPre.length);
      var referKeyUserId = referKeyUserIdPre.substring(referKeyUserIdPre.indexOf("user_id"), referKeyUserIdPre.length);
      //console.log(referValueUserId);
      //console.log(referValueUserId);
      //console.log(referValueAccessToken);
      //console.log(referKeyUserId);
      //console.log(referKeyAccessToken);
    }

    else {
      console.log("参照API(NO." + referId + ")のaccess_tokenおよびuser_idは参照APIに渡すためには取得しません。")
    }

    //参照APIの行に関する情報を取得
    //認証APIでvalueに値がある時に、参照APIが選択した認証APIのvalueを参照APIのkey-valueに書き込むための前処理
    for (let n = 0; n < referIdArray.length; n = n + 2) {
      const referredKeyRowNumber = n + 1 //k=0のとき認証APIを指すが、前のifではじかれる
      const referredValueRowNumber = n + 2;//k=0のとき認証APIを指すが前のifではじかれる

      const referredAuth = referredAuth1Array[referredKeyRowNumber];
      const referredId = referredId1Array[referredKeyRowNumber];

      //認証APIreferと参照APIreferredの切り分けを行い、認証APIのvalueを参照APIのvalueに書き込み
      if (referredId.indexOf("認証") == -1 && referId.indexOf("認証") != -1) {
        //認証APIにvalueがなくても参照する。空valueでリクエスト投げて結果を返さないと上書きしたときエラーの原因がわからない。
        console.log("認証API(No." + referId + ")を参照するAPI(No." + referredId + ")のパラメータを取得しています。");

        //指定した認証APIを参照する(No.カラムと認証APIカラムが一致するときに参照APIに認証APIのレスポンスを書き込む
        if (referredAuth == referId) {
          //参照APIのパラメータkey取得
          const referredKey1array = arraySheet.slice(referredKeyRowNumber - 1, referredKeyRowNumber).map(row => row.slice(10, maxCol)).flat()
          //console.log(referredKey1array)
          //console.log("aaasdkaf;aslj")
          //console.log(referredKeyRowNumber);
          console.log("参照しているAPI(No." + referId + ")が指定した認証API(No." + referredId + ")と一致しました。\n参照API(No." + referId + ")パラメーターのkeyは以下の通りです。")
          console.log(referredKey1array);

          //取得した認証APIパラムのkey要素を順次みていって、user_idとaccess_tokenの列を取り出し、認証APIの該当valueを書き込み
          for (let count = 0; count < referredKey1array.length; count++) {
            console.log("参照APIパラメーターのkeyのうち" + referredKey1array[count] + "について、以下の通り処理を行います。")
            //console.log(count)
            //console.log(referredKey1array[count])
            //console.log(referKeyUserId)
            //console.log(referKeyAccessToken)
            //console.log("aaaklmsladmflk")

            if (referValueUserId != "" && referredKey1array[count] == referKeyUserId) {
              console.log("user_idのkeyです。書き込み処理を行います。");
              sht.getRange(referredValueRowNumber, 11 + count).setValue(referValueUserId);
            }
            else if (referValueAccessToken != "" && referredKey1array[count] == referKeyAccessToken) {
              console.log("access_tokenのkeyです。書き込み処理を行います。");
              sht.getRange(referredValueRowNumber, 11 + count).setValue(referValueAccessToken);
            }
            else {
              console.log("user_idおよびaccess_tokenのkeyではありません。書き込み処理は行いません。")
            }
          }
        }
        else {
          console.log("参照API(No." + referredId + ")が指定した認証API(No." + referredAuth + ")と一致しませんでした。処理いたしません。")
        }
      }
      else {
        console.log("対象API(No." + referId + ")は認証APIではないか、対象API(No." + referredId + ")は参照APIでないため、処理いたしません。")
      }
    }
  }
}





//認証APIのリクエスト、レスポンスを認証APIのvalueに書き出し
function outputAuthPostToWritten(authId, authMethod, sht, authKeyRowNumber, authValueRowNumber) {

  if (authMethod == "POST" && authId.indexOf("認証") != -1) {
    //console.log("呼び出し確認開始");
    //認証APIのレスポンスを認証APIのvalueに書き出し
    const [authPostString, authPostMessage, authPostStatusCode, authAccessToken, authUserId, error] = sendAuthPostRequest(sht, authKeyRowNumber, authValueRowNumber, authMethod);
    //console.log("呼び出し確認完了");

    if (error == null) {
      //console.log("正常系の処理を行います。\nシートにリクエストとレスポンスを書き出します。");
      //リクエスト,レスポンスの書き込み
      sht.getRange(authValueRowNumber, 6).setValue(authPostString);
      sht.getRange(authValueRowNumber, 7).setValue(authPostMessage);
      sht.getRange(authValueRowNumber, 8).setValue(authPostStatusCode);
      //認証APIのaccess_token,user_idを書き込み
      sht.getRange(authValueRowNumber, 11).setValue(authAccessToken);
      sht.getRange(authValueRowNumber, 12).setValue(authUserId);
    }

    else if (error != null) {
      //console.log("異常系の処理を行います。\nシートにリクエストとレスポンスを書き出します。");
      //リクエスト,レスポンスについて
      sht.getRange(authValueRowNumber, 6).setValue(authPostString);
      sht.getRange(authValueRowNumber, 7).setValue(error);
      sht.getRange(authValueRowNumber, 8).setValue(authPostStatusCode);
      //認証APIのaccess_token,user_idを書き込み
      sht.getRange(authValueRowNumber, 11).setValue(authAccessToken);
      sht.getRange(authValueRowNumber, 12).setValue(authUserId);
    }
  }

  //認証APIはPOSTの想定。他メソッドでの認証は想定していない。
  else {
    console.log("想定外（：認証APIが非POSTメソッド）の処理です。");
  }

}


//参照APIのPOSTメソッドのリクエスト、レスポンスを認証APIのvalueに書き出し
function outputPostToWritten(id, method, sht, keyRowNumber, valueRowNumber) {
  if (method == "POST" && id.indexOf("認証") == -1) {
    const [postString, postMessage, postStatusCode, error] = sendPostRequest(sht, keyRowNumber, valueRowNumber, method);

    if (error == null) {
      console.log("正常系の処理を行います。\nシートにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(postString);
      sht.getRange(valueRowNumber, 7).setValue(postMessage);
      sht.getRange(valueRowNumber, 8).setValue(postStatusCode);
    }

    else if (error != null) {
      console.log("異常系の処理を行います。\nシートにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(postString);
      sht.getRange(valueRowNumber, 7).setValue(error);
      sht.getRange(valueRowNumber, 8).setValue(postStatusCode);
    }
  }
  else {
    console.log("想定外（：参照APIが非POSTメソッド）の処理です。");
  }
}


//参照APIのGET,DELETE,PUTメソッドのリクエスト、レスポンスを認証APIのvalueに書き出し
function outputNoPostToWritten(id, method, sht, keyRowNumber, valueRowNumber) {

  //GETメソッド、参照APIについて
  if (method == "GET" && id.indexOf("認証") == -1) {
    const [getUrlReference, getMessage, getStatusCode, error] = sendGetRequest(sht, keyRowNumber, valueRowNumber, method);

    if (error == null) {

      console.log("正常系の処理を行います。\nシートにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(getUrlReference);
      sht.getRange(valueRowNumber, 8).setValue(getStatusCode);
      console.log("実行確認")
      //console.log(typeof(getMessage))

      //セルへの書き出しは50000文字の制限がある。
      if (getMessage.length < 50000) {
        console.log("レスポンスは50000文字以内です。書き出しいたします。")
        sht.getRange(valueRowNumber, 7).setValue(getMessage);
        //検証//console.log(getStatusCode); 
      }
      else {
        //入力内容が 1 つのセルに最大 50000 文字の制限を超えている場合
        //try~catchでは実装不可だった。スプレッドシート側の問題なのでGASではエラー処理されないっぽい
        console.log("異常を検知しました。レスポンスは50000文字を超えています。\nセルにリクエストを書き出します。レスポンスは書き出せません。")
        sht.getRange(valueRowNumber, 7).setValue("レスポンスが50000文字を超えているためセルに書き出せません。\nログをご確認ください。");
      }
    }

    else if (error != null) {
      console.log("異常系の処理を行います。\nシートにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(getUrlReference);
      sht.getRange(valueRowNumber, 7).setValue(error);
      sht.getRange(valueRowNumber, 8).setValue(getStatusCode);
    }
  }

  //PUTメソッド、参照APIについて
  else if (method == "PUT" && id.indexOf("認証") == -1) {
    const [putString, putMessage, putStatusCode, error] = sendPutRequest(sht, keyRowNumber, valueRowNumber, method);
    if (error == null) {
      console.log("正常系の処理を行います。\nシートにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(putString);
      sht.getRange(valueRowNumber, 7).setValue(putMessage);
      sht.getRange(valueRowNumber, 8).setValue(putStatusCode);
    }

    else if (error != null) {
      console.log("異常系の処理を行います。\nシートにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(putString);
      sht.getRange(valueRowNumber, 7).setValue(error);
      sht.getRange(valueRowNumber, 8).setValue(putStatusCode);
    }
  }

  //DELETEメソッド、参照APIについて
  else if (method == "DELETE" && id.indexOf("認証") == -1) {
    const [deleteString, deleteMessage, deleteStatusCode, error] = sendDeleteRequest(sht, keyRowNumber, valueRowNumber, method);
    if (error == null) {
      console.log("正常系の処理を行います。\nシートにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(deleteString);
      sht.getRange(valueRowNumber, 7).setValue(deleteMessage);
      sht.getRange(valueRowNumber, 8).setValue(deleteStatusCode);
    }

    else if (error != null) {
      console.log("異常系の処理を行います。\nシートにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(deleteString);
      sht.getRange(valueRowNumber, 7).setValue(error);
      sht.getRange(valueRowNumber, 8).setValue(deleteStatusCode);
    }

  }

  else {
    console.log("想定外（：参照APIがPOSTメソッド）の処理です。");
  }
}


//認証APIのPOSTメソッドをリクエストする
function sendAuthPostRequest(sht, authKeyRowNumber, authValueRowNumber, authMethod) {
  //パラメータのkeyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(authKeyRowNumber, 13, 1, sht.getLastColumn() - 12)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  const keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(authValueRowNumber, 13, 1, sht.getLastColumn() - 12)
  const values2array = rangeParam.getValues()
  const valueFlat = values2array.flat()

  //key-valueの各要素を対応させる//key-valueの二次元配列を作成する
  const obj2array = [keyFlat, valueFlat];
  //console.log(obj2array); 
  //key-valueの二次元配列から連想配列への変換

  const keys = obj2array[0];
  const values = obj2array[1];
  const obj = {};

  //入力のあるパラメータを取得.valueが空の時はkeyが削除される
  for (let j = 0; j <= keys.length; j++) {
    if (values[j] !== "") {
      obj[keys[j]] = values[j];
    }
    else {
    }
  }
  //URLを二次元配列で取得
  const authUrlPost2Array = sht.getRange(authValueRowNumber, 10).getValues();  // G列:URLカラムの全ての行を取得

  //URLの二次元配列を一次元配列に変換
  const authUrlPostFlat = authUrlPost2Array.flat()
  //console.log(authUrlPostFlat);

  //URL配列からURLを抽出


  const authUrlPost = authUrlPostFlat[0];

  //console.log(authUrlPost)

  //key-value配列のJSON化
  const authPostString = JSON.stringify(obj, null, "\t")


  //POSTリクエスト　parameter
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': authPostString,
    "muteHttpExceptions": true,
  }
  try {
    //POSTリクエスト　urlと合わせて送信
    var authPostResponse = UrlFetchApp.fetch(authUrlPost, options);

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
    //console.log(authUserId)//返却されないときはundefined
    //console.log(1000000000000000000)

    if (authUserId == "undefined") {
      var authUserId = String(obj.user_id);
      console.log(authUserId);
      console.log(999);
    }

    else if (authAccessToken == "undefined") {
      var authAccessToken = String(obj.access_token);
    }


    let error = null
    //リクエストとレスポンスのログ排出
    console.log(authMethod + " " + authPostApi + "についてaccess_tokenとuser_idの組み合わせが" + authPostCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
    console.log(authUrlPost);
    console.log(authPostString);
    console.log(authPostMessage);

    console.log(authPostStatusCode);
    console.log(authAccessToken);
    console.log(authUserId);
    return [authPostString, authPostMessage, authPostStatusCode, authAccessToken, authUserId, error];
  } catch (e) {
    console.log(authMethod + " " + authPostApi + "についてaccess_tokenとuser_idの組み合わせが" + authPostCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
    //例外エラー処理
    let error = JSON.stringify(e.message)
    console.log('Error:' + error)
    return [authPostString, authPostMessage, authPostStatusCode, authAccessToken, authUserId, error];
  }
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendGetRequest(sht, keyRowNumber, valueRowNumber, method) {
  //getvalueでgetrangeの値を取得
  // セル範囲の値（ここではURL）を２次元配列で取得する
  const urlValue = sht.getRange(valueRowNumber, 10).getValues();
  // セル範囲の値（ここではURL）を１次元配列に変換する
  const urlValuesFlat = urlValue.flat()
  //検証//console.log(valuesFlat)

  //ログ排出用にapi,組み合わせカラムのバリューを取得する。
  const getApi = sht.getRange(valueRowNumber, 3).getValue();
  const getCombi = sht.getRange(valueRowNumber, 5).getValue();
  //console.log(getApi);
  //console.log(getCombi);

  //一次元配列からURLテキスト抽出
  while (urlValuesFlat.length) {
    const getUrl = urlValuesFlat.shift();

    const options = {
      'method': 'get',
      "muteHttpExceptions": true,
    }

    try {
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
      var error = null;
      console.log(method + " " + getApi + "についてaccess_tokenとuser_idの組み合わせが" + getCombi + "の時、リクエストとレスポンスは以下の通りです。\n※URL、レスポンスメッセージ、ステータスコードの順に記載")
      //console.log(getUrl);
      console.log(getUrlReference);
      console.log(getMessage);
      console.log(getStatusCode);

      return [getUrlReference, getMessage, getStatusCode, error];
    } catch (e) {
      console.log(method + " " + getApi + "についてaccess_tokenとuser_idの組み合わせが" + getCombi + "の時、リクエストとレスポンスは以下の通りです。\n※URL、レスポンスメッセージ、ステータスコードの順に記載")
      //例外エラー処理
      var error = JSON.stringify(e.message)
      console.log('Error:' + error)
      return [getUrlReference, getMessage, getStatusCode, error];
    }
  }
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendPostRequest(sht, keyRowNumber, valueRowNumber, method) {
  //パラメータのkeyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(keyRowNumber, 11, 1, sht.getLastColumn() - 10)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  const keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(valueRowNumber, 11, 1, sht.getLastColumn() - 10)
  const values2array = rangeParam.getValues()
  const valueFlat = values2array.flat()

  //key-valueの各要素を対応させる
  //key-valueの二次元配列を作成する
  const obj2array = [keyFlat, valueFlat];

  //console.log(obj2array); 
  //key-valueの二次元配列から連想配列への変換 
  const keys = obj2array[0];
  const values = obj2array[1];
  const obj = {};

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
    const urlPostFlat = urlPost2Array.flat()
    //console.log(urlPostFlat);

    //URL配列からURLを抽出
    var urlPost = urlPostFlat[0];
  }


  //key-value配列のJSON化
  const postString = JSON.stringify(obj, null, "\t")


  //POSTリクエスト　parameter
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': postString,
    "muteHttpExceptions": true,
  }

  try {
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
    var error = null
    //console.log(response.getContentText())
    //console.log(response.getResponseCode())
    //テスト//Logger.log(idArray);

    //リクエストとレスポンスのログ排出
    console.log(method + " " + postApi + "についてaccess_tokenとuser_idの組み合わせが" + postCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
    console.log(urlPost);
    console.log(postString);
    console.log(postMessage);
    console.log(postStatusCode);
    return [postString, postMessage, postStatusCode, error]
  } catch (e) {
    console.log(method + " " + postApi + "についてaccess_tokenとuser_idの組み合わせが" + postCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
    var error = JSON.stringify(e.message)
    console.log('Error:' + error)

    return [postString, postMessage, postStatusCode, error]
  }

}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendPutRequest(sht, keyRowNumber, valueRowNumber, method) {
  //パラメータの一覧keyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(keyRowNumber, 11, 1, sht.getLastColumn() - 10)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  const keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(valueRowNumber, 11, 1, sht.getLastColumn() - 10)
  const values2array = rangeParam.getValues()
  const valueFlat = values2array.flat()

  //key-valueの各要素を対応させる
  //key-valueの二次元配列を作成する
  const obj2array = [keyFlat, valueFlat];

  //console.log(obj2array); 
  //key-valueの二次元配列から連想配列への変換 
  const keys = obj2array[0];
  const values = obj2array[1];
  const obj = {};

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
    const urlPutFlat = urlPut2Array.flat()
    //console.log(urlPostFlat);

    //URL配列からURLを抽出
    var urlPut = urlPutFlat[0];
  }


  //key-value配列のJSON化
  const putString = JSON.stringify(obj, null, "\t")

  //PUTリクエスト　parameter
  const options = {
    'method': 'put',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': putString,
    "muteHttpExceptions": true,
  };

  try {
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
    var error = null;
    //console.log(response.getContentText())
    //console.log(response.getResponseCode())
    //テスト//Logger.log(idArray);

    //リクエストとレスポンスのログ排出
    console.log(method + " " + putApi + "についてaccess_tokenとuser_idの組み合わせが" + putCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
    console.log(urlPut);
    console.log(putString);
    console.log(putMessage);
    console.log(putStatusCode);
    return [putString, putMessage, putStatusCode, error];
  } catch (e) {
    console.log(method + " " + putApi + "についてaccess_tokenとuser_idの組み合わせが" + putCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
    var error = JSON.stringify(e.message)
    console.log('Error:' + error)
    return [putString, putMessage, putStatusCode, error];

  }
}

//書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
function sendDeleteRequest(sht, keyRowNumber, valueRowNumber, method) {
  //パラメータの一覧keyを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const range = sht.getRange(keyRowNumber, 11, 1, sht.getLastColumn() - 10)
  const keys2array = range.getValues()
  //Logger.log(keys2array);
  const keyFlat = keys2array.flat()
  //console.log(keyFlat);

  //パラメータの内容valueを2次元配列として取得。1次元配列へ変換して配列要素抽出
  const rangeParam = sht.getRange(valueRowNumber, 11, 1, sht.getLastColumn() - 10)
  const values2array = rangeParam.getValues()
  const valueFlat = values2array.flat()

  //key-valueの各要素を対応させる
  //key-valueの二次元配列を作成する
  const obj2array = [keyFlat, valueFlat];

  //console.log(obj2array); 
  //key-valueの二次元配列から連想配列への変換 
  const keys = obj2array[0];
  const values = obj2array[1];
  const obj = {};

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
    const urlDeleteFlat = urlDelete2Array.flat()
    //console.log(urlDeleteFlat);

    //URL配列からURLを抽出
    var urlDelete = urlDeleteFlat[0];
  }


  //key-value配列のJSON化
  const deleteString = JSON.stringify(obj, null, "\t")

  //DELETEリクエスト　parameter
  const options = {
    'method': 'delete',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': deleteString,
    "muteHttpExceptions": true,
  };

  try {
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
    var error = null

    //console.log(response.getContentText())
    //console.log(response.getResponseCode())
    //テスト//Logger.log(idArray);

    //リクエストとレスポンスのログ排出
    console.log(method + " " + deleteApi + "についてaccess_tokenとuser_idの組み合わせが" + deleteCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
    console.log(urlDelete);
    console.log(deleteString);
    console.log(deleteMessage);
    console.log(deleteStatusCode);
    return [deleteString, deleteMessage, deleteStatusCode, error];
  } catch (e) {
    console.log(method + " " + deleteApi + "についてaccess_tokenとuser_idの組み合わせが" + deleteCombi + "の時、リクエストとレスポンスは以下の通りです。\n※エンドポイント、JSONペイロード、レスポンスメッセージ、ステータスコードの順に記載")
    var error = JSON.stringify(e.message)
    console.log('Error:' + error)
    return [deleteString, deleteMessage, deleteStatusCode, error];

  }
}


  //テスト//Logger.log(idArray);