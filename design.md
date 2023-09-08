・想定される使用場面
1.アプリ側：新規開発時、複数のAPIを考慮する場合があるとき、アプリでAPI連携した時、テスト結果がNG。原因がアプリ側かサーバー側かの切り分け
2.サーバー側：リグレッションテスト。新規プロジェクトでのAPIの改修/新規追加

・既存API検証ツールの痒い所（デメリット）
Postman:無料プランは閲覧権限３人まで、パラメーターを一気にかえるにはエクスポートしてjson見る必要があり、一手間かかるand若干編集しづらい(https://qiita.com/panpanpanyasan/items/d1eadb7d69a45e6ef866#シナリオテストもできるよ)、https://dev.classmethod.jp/articles/postman_newman_githubactions/

REST Client(VSCode):複数のAPIを検証するためにボタンをクリックする必要がある。一斉に検証するのが手間。単発検証・
curl:記述,編集しにくい。リクエストが見にくい。単発検証。

・既存API検証ツールのいいところ（メリット）
Postman;直感的で分かりやすい。パラメータのkey-valueが見やすい。シナリオテスト（一斉検証）可能、CI/CDに組み込める、テストはjsonで共有可能、

Rest Client:検証ファイルを共有しやすい。わかりやすい。
curl：リクエストとレスポンスをテキストで共有しやすい。

・・GASによるマクロツールのメリット/デメリット
メリット：スクリプトと結果を共有しやすい。複数のAPIを一斉に検証できる。パラメータを検索/置換でまとめて変えやすい。自動実行が楽。設定ほぼ不要。CICDに組み込める(?)(https://zenn.dev/furnqse/articles/a138962560db560)
デメリット：作るのに時間かかる。

・対象のAPI
https://www.chatwork.com/#!rid280703476-1681895945547063296

・・GASの公式ドキュメント
実現可能なパラメータ


https://developers.google.com/apps-script/reference/url-fetch/http-response?hl=ja

・他参考仕様
https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app?hl=ja

検討項目

・メソッド別にシート分ける（DELETEはPOSTしないと400がかえるので、PUTも同様、GETも同様？基本的にDBに入っているが実装したばかりだとデータセットない場合あるのでPOST通してデータあることが前提になる。）
→POSTの優先順位上げてその他は下げることで分けずに済む

・POSTはログイン認証必要なAPIある。（task系：api/user/login/admin系:api/admin/;login）access_tokenだけいる？refresh_tokenとは？不要？
https://stg.mobile-backend.com/admin/api/155
https://stg.mobile-backend.com/admin/api/28
→access_tokenを更新するためにあって、不要

→ログイン必要なAPIだけ検討するか、それともログイン必要ないapiもまとめて一律にログインしてしまって通すか。
→前者、180個あるapiを仕分けするのは結構大変そう。問題なければ後者で実装したい。
→NG。access_token不要のAPIについてデフォルトで渡すのはいかがなものか。

見通し
1.認証を先に処理してaccess_token（とuser_id)書き出し。
2.書き出した内容を各種APIのaccesstokenのセルに書き込み
3.key-valueのオブジェクトを取得、JSON化
4.リクエストしてレスポンスをセルに書き出し




・リクエストデータとレスポンスを書き込むようにする

・他必要な項目→MTG結果の章に記載
https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app?hl=ja

フォーマット

叩き台(すぐできそう） 
・できれば現状のシートを応用する形にしたい。



・パラメータを横に並べていて入力しにくい、参照しにくい→API別にフォーマットをかえる。処理が遅くならないか。実装はちょっと複雑になる（ちょっと考える時間いる）




認証して書き出し
見通し
1.特定の列{parameter,value記載しているセル行の番号＝開始行番号＋id＋(id-1),parameter,行番号}を取得
2.1.の配列から１つおきに抽出＝valueのみ抽出
for (var i = 0; i < getLastRow()/2; i = i + 2)

3.user_idやaccess_token書き出し

リクエストーレスポンス
1.特定の列{parameter,value記載しているセル行の番号＝開始行番号＋id＋(id-1),parameter,行番号}を取得
2.1.の配列から１つおきに抽出＝valueのみ抽出
for (var i = 0; i < getLastRow()/2; i = i + 2)

3.value記載しているセル行の番号-1=parameter記載しているセル行の番号

2.と3.組み合わせてJSONにする
ツール使用者へのヒアリングMTG結果

1.書き出す項目の追加：一旦完了0228
リクエスト内容
レスポンス結果（ステータスコードも含む）

2.API別に分けて指定したAPIパラメータのkey-valueを取得する：一旦完了0228

3.認証APIを参照する:
一旦完了0302→refresh_token不要。

→access_token必要ないapiについて通してもOK
→access_token必要ないapiにaccess_tokenが付与されてしまうので認証必要なAPIとそうでないAPIはシート・スクリプトをそれぞれ分ける必要がある

→カラム追加して処理を分割して対応

仕分け：ルーティング参照
→カラム追加して対応。入力者が認証apiを要するかどうか設定できる

4.レスポンス、リクエストの可読性向上：一旦完了0227
・エラーメッセージ　変換する必要あるなら、時短にならないのでは。デコード処理はほしい。

→JSONレスポンス長文　ベタ書きだとリクエスのエラー箇所特定できない

→改行
key-value
key-value
key-value

・,の後ろに”¥n" 改行コード入れる

その他追記

5.POSTリクエスト処理の優先順位あげる。GET,DELETE,PUTはPOSTより下で同等。：一旦完了0301
リクエスト処理をPOSTとそれ以外で関数それぞれ作成する。

6.リファクタリング。800行-900行くらいになりそう。共通部分の関数はファイルを分けて管理する。管理できないならクラス化→実装しない。

検証ツールの注意点

・認証APIはPOSTで投げる想定で、それ以外のメソッドを認証APIとすることはできない
・現時点で、認証APIを複数用意して APIに応じて認証するAPIを選択する機能はない。
→対応済み
・つまり複数の認証APIを参照することはできない。
→複数の認証apiに対応済み
7.シートにカラム:参照APIを追加して、対応するAPIを選択できるようにする→たぶん完了　0302 要テスト

時間内にできるかどうか微妙。フォーマットのイメージ。とりあえず7は無視して実装する。あとで改修。



8.使い方のドキュメント整備：完了0309
・シートのフォーマット(カラム変更していい場所、変更してはいけない場所)

・シートの項目の意味の説明
・トリガーの設定の仕方
・スクリプトの編集後のトリガーの設定方法
・関数の説明(ざっくりとフローを示す
・スクリプトシート名をスクリプトに入力する必要ある

9.テスト：0309一旦完了
・認証apiが複数ある時対応できるかどうか
・POST以外のapiで認証apiを参照できるかどうか
・同一APIでPOSTした後PUT、DELETE、GETできるかどうか

10.
レスポンスが返ってこない時のエラー処理をtry~catchで実装：0307完了

dnsサーバーエラーやアクセスできないurlをエンドポイントに指定した場合などhttps://www.monotalk.xyz/blog/google-app-script-の-urlfetchapp-の-例外ハンドリングについて/

11.GETで表示する文字の量が多い時、実行できなくなる。：完了0308
Exception: 入力内容が 1 つのセルに最大 50000 文字の制限を超えています。

12.処理速度の改善(認証APIあると実行時間70秒かかる)：検討中
①キャッシュの利用：NG
→更新したとき、キャッシュとかちあう。
→キャッシュを行別に取得は困難
→認証APIありのシートと認証APIなしのシートにわける。スクリプトも別々にする。

独自キャッシュ機構をつくる。
https://befool.co.jp/blog/8823-scholar//gas-use-cache/

GASのキャッシュは時間制約、容量制約などあり。
https://developers.google.com/apps-script/reference/cache?hl=ja

②3重ループのコード改善：進行中：ロジックの大幅な書き換え必要で終わるかどうか不明
書き出し処理をループ外へ

13:個別apiのリクエスト：検討終了、実装しない
ボタンを実装
→新規APIの追加に対応できない
→シートをプロジェクト別に分ける

14.認証APIでuser_idを返さない場合にundefinedを書き出してしまう不具合：一旦完了0309

      
   
   
      
         