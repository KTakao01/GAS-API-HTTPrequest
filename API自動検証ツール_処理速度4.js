
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
  var sht = ss.getSheetByName('testのコピー');
  //test//Logger.log(sht.getLastRow()+100);　//100+4//行を取得できているかの確認
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
  for (var k = 0; k < authIdArray.length; k = k + 2) {
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    var authKeyRowNumber = k + 1 //k=0のとき認証APIを指す
    var authValueRowNumber = k + 2;//k=0のとき認証APIを指す
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    var authId = authIdArray[authKeyRowNumber];
    var authMethod = authMethodArray[authKeyRowNumber];
    var authAuth = authArray[authKeyRowNumber];
    //検証//Logger.log(method);// [POST, POST, POST, GET]
    //検証//Logger.log(id);//[認証,1,2,3]
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
  var sht = ss.getSheetByName('testのコピー');
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
  for (var k = 0; k < idArray.length; k = k + 2) {
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    var keyRowNumber = k + 1 //k=0のとき認証APIを指す
    var valueRowNumber = k + 2;//k=0のとき認証APIを指す
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    var id = idArray[keyRowNumber];
    var method = methodArray[keyRowNumber];
    var auth = authArray[keyRowNumber];
    //検証//Logger.log(method);// [POST, POST, POST, GET]
    //検証//Logger.log(id);//[認証,1,2,3]
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
  var sht = ss.getSheetByName('testのコピー');
  //test//Logger.log(sht.getLastRow()+100);　//100+4//行を取得できているかの確認
  //for (var i = 2; i < sht.getLastRow()+1 ;i = i +2) {}
  //認証情報を参照して認証必要なAPIのuser_idとaccess_tokenの列に認証結果を書き出す。
  //認証不要のときは、リクエスト処理を行う。
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
  for (var k = 0; k < idArray.length; k = k + 2) {
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    var keyRowNumber = k + 1 //k=0のとき認証APIを指す
    var valueRowNumber = k + 2;//k=0のとき認証APIを指す
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    var id = idArray[keyRowNumber];
    var method = methodArray[keyRowNumber];
    var auth = authArray[keyRowNumber];
    //検証//Logger.log(method);// [POST, POST, POST, GET]
    //検証//Logger.log(id);//[認証,1,2,3]
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
function referenceAuthtest6() {
  //参照するAPIを選択する（改修予定）//"必要"部分を変数化する必要がある参照
  //認証APIをループで参照しつつ、参照APIを参照しながら書き換えていくので二重ループ
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sht = ss.getSheetByName('testのコピー');
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
  var startH = 0

  // const startJ = 0
  var numHI = referIdArray.length
  //const numJ = referredKey1array.length
  //認証APIを参照する

  //全ての行の取得と、参照APIの各種カラム取得
  var lineArray = [];

  for (var m = startH; m < numHI; m = m + 2) {
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5→keyの行に対応
    //k=0のときk+2=2、k=2のときk+2=4,k=4のときk+2=6→valueの行に対応
    var referKeyRowNumber = m + 1 //k=0のとき認証APIを指す
    var referValueRowNumber = m + 2;//k=0のとき認証APIを指
    //admin/login,user/loginのように認証apiが2つ以上ある場合
    //検証//Logger.log(valueRowNumber+10);
    //k=0のときk+1=1、k=2のときk+1=3,k=4のときk+1=5
    var referId = referIdArray[referKeyRowNumber];
    var referMethod = referMethodArray[referKeyRowNumber];
    var referAuth = referAuthArray[referKeyRowNumber];
    //検証//Logger.log(referMethod);//  [POST, POST, POST, GET]
    //検証//Logger.log(referId);//[認証,1,2,3]
    //検証//console.log(referMethod);
    //検証//console.log(referAuth);
    //検証//console.log(referId);
    console.log("認証API(No." + referId + ")に対応する参照APIの有無を調べます。")
    //認証APIのkey-valueを抽出。keyはあとで条件一致の時に使う。”返却された”という文字以外を抽出した
    if (referId.indexOf("認証") != -1) {
      console.log("認証API(No." + referId + ")のaccess_tokenおよびuser_idを参照API に渡すために取得します。\n取得結果は以下の通りです。")
      var referValueAccessToken = sht.getRange(referValueRowNumber, 11).getValue();
      var referValueUserId = sht.getRange(referValueRowNumber, 12).getValue();
      var referKeyAccessTokenPre = sht.getRange(referKeyRowNumber, 11).getValue();
      var referKeyUserIdPre = sht.getRange(referKeyRowNumber, 12).getValue();
      var referKeyAccessToken = referKeyAccessTokenPre.substring(referKeyAccessTokenPre.indexOf("access_token"), referKeyAccessTokenPre.length);
      var referKeyUserId = referKeyUserIdPre.substring(referKeyUserIdPre.indexOf("user_id"), referKeyUserIdPre.length);
      //console.log(referValueUserId);
      //console.log(referValueAccessToken);
      //console.log(referKeyUserId);
      //console.log(referKeyAccessToken);
    }
    else {
      console.log("参照API(NO." + referId + ")のaccess_tokenおよびuser_idは参照APIに渡すためには取得しません。")
    }
    //参照APIを参照。keyを取得
    //認証APIでvalueに値がある時に、参照APIが選択した認証APIのvalueを参照APIのkey-valueに書き込むための前処理
    var startI = 0
    //console.log(numHI)//18
    //console.log(dataId)//[]
    ////for (var n = startI; n < numHI; n = n + 2) {

    //for (var line = startI; line < numHI; line++) {

    var startJ = 0
    //全ての列を取得→ジャグ配列を取得//独立していていい
    var numJ = sht.getLastColumn();
    //console.log(numJ);//26の想定→OK//ジャグ配列の最長要素を返してくれる。
    //console.log(777)
    //シートを1行ずつ取得

    //console.log(lineArray)
    //console.log(9999)


    for (var n = 0; n < referIdArray.length; n += 2) {
      var referredKeyRowNumber = n + 1 //k=0のとき認証APIを指すが、前のifではじかれる
      var referredValueRowNumber = n + 2;//k=0のとき認証APIを指すが前のifではじかれる

      lineArray[referredKeyRowNumber - 1] = [];
      lineArray[referredValueRowNumber - 1] = [];


      //console.log("最終行の範囲内です。")
      var referredAuths2Array = sht.getRange(1, 4, sht.getLastRow()).getValues();
      var referredAuth1Array = referredAuths2Array.flat();
      var referredAuth = referredAuth1Array[referredKeyRowNumber];
      var referredIds2Array = sht.getRange(1, 1, sht.getLastRow()).getValues();
      var referredId1Array = referredIds2Array.flat();
      var referredId = referredId1Array[referredKeyRowNumber];


      //全ての列を取得する
      //取得した行の配列に空白があれば、nullを入れる
      for (col = startJ; col < numJ; col++) {
        var colLog = col + 1
        lineArray[referredKeyRowNumber - 1][col] = sht.getRange(referredKeyRowNumber, col + 1).getValue();
        lineArray[referredValueRowNumber - 1][col] = sht.getRange(referredValueRowNumber, col + 1).getValue();
        if (lineArray[referredKeyRowNumber - 1][col] == "") {
          console.log("ジャグ配列の最大数に揃うように、空白セル"+ referredKeyRowNumber + "行" + colLog + "列にnullを代入しています。")
          lineArray[referredKeyRowNumber - 1][col] = null
        }
        else if (lineArray[referredValueRowNumber - 1][col] == "") {
          lineArray[referredValueRowNumber - 1][col] = null

          console.log("ジャグ配列の最大数に揃うように、空白セル" + referredKeyRowNumber + "行" + colLog + "列にnullを代入しています。")
        }
        else {
          console.log(referredKeyRowNumber + "行" + colLog + "列の既存の値を取得しています。")
        }


        //}
        //console.log(lineArray)
        //console.log(10000)



        //認証APIreferと参照APIreferredの切り分けを行い、認証APIのvalueを参照APIのvalueに書き込み
        if (referredId.indexOf("認証") == -1 && referId.indexOf("認証") != -1) {
          //認証APIにvalueがある場合のみ参照APIは参照する？
          //→認証APIにvalueがなくても参照するようにする。空valueでリクエスト投げて結果を返さないと上書きしたときエラーの原因がわからない。
          //if (referValueAccessToken != "" || referValueUserId != "") {
          console.log("認証API(No." + referId + ")を参照するAPI(No." + referredId + ")のパラメータを取得しています。");
          //指定した認証APIを参照する
          if (referredAuth == referId) {
            //参照APIのパラメータkey取得
            var rangeParam = sht.getRange(referredKeyRowNumber, 11, 1, sht.getLastColumn() - 10)
            const referredKeys2array = rangeParam.getValues()
            //console.log(referredKeys2array)
            var referredKey1array = referredKeys2array.flat()
            //console.log(referredKeyRowNumber);
            console.log("参照しているAPI(No." + referId + ")が指定した認証API(No." + referredId + ")と一致しました。\n参照API(No." + referId + ")パラメーターのkeyは以下の通りです。")

            //指定した認証APIと一致した参照APIのキーパラメータを配列取得
            console.log(referredKey1array);


            //console.log(numJ*10)//180

            //console.log(referredValueRowNumber-1 + 100000);

            //取得した認証APIパラムのkey要素を順次みていって、user_idとaccess_tokenの列を取り出し、認証APIの該当valueを書き込み
            //for (count = startJ; count < referredKey1array.length; count++) {
            //console.log(dataToken)
            //console.log(dataId)//18

            var count = col;
            console.log(count);
            //for (count = 0; count < referredKey1array.length; count++) {
            console.log("参照APIパラメーターのkeyのうち" + referredKey1array[count] + "について、以下の通り上書き処理の準備を行います。")
            var countLog = count + 1

            //認証APIのuser_idバリューが存在していて、参照APIのキー配列に認証APIのuser_idキーが一致するところで、認証APIのuser_idバリューを渡す
            if (referValueUserId != "" && referredKey1array[count] == referKeyUserId) {

              console.log("user_idのkeyです。" + referredValueRowNumber + "行" + countLog + "列のセルの書き込み処理を行う予定です。");
              lineArray[referredValueRowNumber - 1][count] = referValueUserId
              console.log(lineArray[referredValueRowNumber - 1][count]);
              //console.log(referValueRowNumber)
              //console.log(10 + count)
              console.log("差異trdhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhhh")
              //console.log(referValueUserId)
              //console.log(lineArray)
            }
            //認証APIのaccess_tokenバリューが存在していて、参照APIのキー配列に認証APIのaccess_tokenキーが一致するところで、認証APIのaccess_tokenバリューを渡す
            else if (referValueAccessToken != "" && referredKey1array[count] == referKeyAccessToken) {
              console.log("access_tokenのkeyです。" + referredValueRowNumber + "行" + countLog + "列のセルの書き込み処理を行う予定です。");
              //console.log(count)
              //console.log(1)
              //console.log(referredKey1array[count])
              //console.log(2)
              //console.log(referKeyAccessToken)
              //console.log(3)
              //console.log(referredValueRowNumber - 1)
              //console.log(4)
              //console.log(lineArray)
              //console.log(5)

              lineArray[referredValueRowNumber - 1][count] = referValueAccessToken
              console.log(lineArray[referredValueRowNumber - 1][count])
              //console.log(referValueRowNumber )
              //console.log(10 + count)
              //console.log(lineArray)
            }
            //認証APIのuser_idバリュー,access_tokenバリューが存在しない、または、参照APIのキー配列に認証APIのuser_id,access_tokenキーが一致しない。
            else {
              console.log("user_idおよびaccess_tokenのkeyではありません。該当の配列要素を書き換えません。\n書き込み処理は行いません。")

              //console.log(referValueRowNumber - 1)
            }
            //}
            //}
          }
          else {
            console.log("参照API(No." + referredId + ")が指定した認証API(No." + referredAuth + ")と一致しませんでした。処理いたしません。")
          }

        }


        else {
          console.log("対象API(No." + referId + ")は認証APIではないか、対象API(No." + referredId + ")は参照APIでないため、処理いたしません。")
        }

        //}//for colは独立しててよい
      }
    }
  }


  //}
  //console.log(lineArray)

  console.log("認証を必要とする参照APIパラメーターのkeyのうち、user_idとaccess_tokenについて認証APIのvalueの値を書き込みます")
  //console.log("user_idのの書き込み処理を行います。");
  //console.log("access_tokenの書き込み処理を行います。");
  sht.getRange(startI + 1, startJ + 1, numHI, numJ).setValues(lineArray);
  //sht.getRange(referredValueRowNumber, 11 + count).setValue(referValueAccessToken);


}
//認証APIのレスポンスを認証APIのvalueに書き出し
function outputAuthPostToWritten(authId, authMethod, sht, authKeyRowNumber, authValueRowNumber) {
  if (authMethod == "POST" && authId.indexOf("認証") != -1) {
    //console.log("呼び出し確認開始");
    //認証APIのレスポンスを認証APIのvalueに書き出し
    const [authPostString, authPostMessage, authPostStatusCode, authAccessToken, authUserId, error] = sendAuthPostRequest(sht, authKeyRowNumber, authValueRowNumber, authMethod);
    //console.log("呼び出し確認完了");
    if (error == null) {
      //console.log("正常系の処理を行います。\nセルにリクエストとレスポンスを書き出します。");
      sht.getRange(authValueRowNumber, 6).setValue(authPostString);
      sht.getRange(authValueRowNumber, 7).setValue(authPostMessage);
      sht.getRange(authValueRowNumber, 8).setValue(authPostStatusCode);
      sht.getRange(authValueRowNumber, 11).setValue(authAccessToken);
      sht.getRange(authValueRowNumber, 12).setValue(authUserId);
    }
    else if (error != null) {
      //console.log("異常系の処理を行います。\nセルにリクエストとレスポンスを書き出します。");
      sht.getRange(authValueRowNumber, 6).setValue(authPostString);
      sht.getRange(authValueRowNumber, 7).setValue(error);
      sht.getRange(authValueRowNumber, 8).setValue(authPostStatusCode);
      sht.getRange(authValueRowNumber, 11).setValue(authAccessToken);
      sht.getRange(authValueRowNumber, 12).setValue(authUserId);
    }
  }
  //認証APIはPOSTの想定。他メソッドでの認証は想定していない。
  else {
    console.log("想定外（：認証APIが非POSTメソッド）の処理です。");
  }
}
function outputPostToWritten(id, method, sht, keyRowNumber, valueRowNumber, e) {
  if (method == "POST" && id.indexOf("認証") == -1) {
    const [postString, postMessage, postStatusCode, error] = sendPostRequest(sht, keyRowNumber, valueRowNumber, method);
    if (error == null) {
      console.log("正常系の処理を行います。\nセルにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(postString);
      sht.getRange(valueRowNumber, 7).setValue(postMessage);
      sht.getRange(valueRowNumber, 8).setValue(postStatusCode);
    }
    else if (error != null) {
      console.log("異常系の処理を行います。\nセルにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(postString);
      sht.getRange(valueRowNumber, 7).setValue(error);
      sht.getRange(valueRowNumber, 8).setValue(postStatusCode);
    }
  }
  else {
    console.log("想定外（：参照APIが非POSTメソッド）の処理です。");
  }
}
function outputNoPostToWritten(id, method, sht, keyRowNumber, valueRowNumber) {
  if (method == "GET" && id.indexOf("認証") == -1) {
    const [getUrlReference, getMessage, getStatusCode, error] = sendGetRequest(sht, keyRowNumber, valueRowNumber, method);
    if (error == null) {
      console.log("正常系の処理を行います。\nセルにリクエストとレスポンスを書き出します。")
      //sendGetRequest()の返却値：レスポンスメッセージとステータスコードをメイン関数で再利用する。
      //書き込む処理はmain()で行う。sendXX()はリクエストとログ書き出し。
      sht.getRange(valueRowNumber, 6).setValue(getUrlReference);
      sht.getRange(valueRowNumber, 8).setValue(getStatusCode);
      console.log("実行確認")
      try {
        console.log("実行確認2")
        sht.getRange(valueRowNumber, 7).setValue(getMessage);
        //検証//console.log(getStatusCode); 
      } catch (e) {
        //入力内容が 1 つのセルに最大 50000 文字の制限を超えている場合
        console.log("異常を検知しました。\nセルにリクエストとレスポンスを書き出します。")
        sht.getRange(valueRowNumber, 7).setValue(e.message);
      }
    }
    else if (error != null) {
      console.log("異常系の処理を行います。\nセルにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(getUrlReference);
      sht.getRange(valueRowNumber, 7).setValue(error);
      sht.getRange(valueRowNumber, 8).setValue(getStatusCode);
    }
  }
  else if (method == "PUT" && id.indexOf("認証") == -1) {
    const [putString, putMessage, putStatusCode, error] = sendPutRequest(sht, keyRowNumber, valueRowNumber, method);
    if (error == null) {
      console.log("正常系の処理を行います。\nセルにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(putString);
      sht.getRange(valueRowNumber, 7).setValue(putMessage);
      sht.getRange(valueRowNumber, 8).setValue(putStatusCode);
    }
    else if (error != null) {
      console.log("異常系の処理を行います。\nセルにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(putString);
      sht.getRange(valueRowNumber, 7).setValue(error);
      sht.getRange(valueRowNumber, 8).setValue(putStatusCode);
    }
  }
  else if (method == "DELETE" && id.indexOf("認証") == -1) {
    const [deleteString, deleteMessage, deleteStatusCode, error] = sendDeleteRequest(sht, keyRowNumber, valueRowNumber, method);
    if (error == null) {
      console.log("正常系の処理を行います。\nセルにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(deleteString);
      sht.getRange(valueRowNumber, 7).setValue(deleteMessage);
      sht.getRange(valueRowNumber, 8).setValue(deleteStatusCode);
    }
    else if (error != null) {
      console.log("異常系の処理を行います。\nセルにリクエストとレスポンスを書き出します。")
      sht.getRange(valueRowNumber, 6).setValue(deleteString);
      sht.getRange(valueRowNumber, 7).setValue(error);
      sht.getRange(valueRowNumber, 8).setValue(deleteStatusCode);
    }
  }
  else {
    console.log("想定外（：参照APIがPOSTメソッド）の処理です。");
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
    const authUrlPost2Array = sht.getRange(authValueRowNumber, 10).getValues();  // G列:URLカラムの全ての行を取得
    //URLの二次元配列を一次元配列に変換
    var authUrlPostFlat = authUrlPost2Array.flat()
    //console.log(authUrlPostFlat);
    //URL配列からURLを抽出
    var authUrlPost = authUrlPostFlat[0];
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
      //console.log(999);
    }

    else if (authAccessToken == "undefined") {
      var authAccessToken = String(obj.access_token);
    }



    var error = null
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
    var error = JSON.stringify(e.message)
    console.log('Error:' + error)
    return [authPostString, authPostMessage, authPostStatusCode, authAccessToken, authUserId, error];
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
