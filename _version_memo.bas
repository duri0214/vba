Attribute VB_Name = "●バージョンメモ"
'ただのメモ書きです

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'ver20101013の変更点
'LogWritterにタイムスタンプ機能がついた
'ストリームをあけると スタートラップを取り、ストリームを閉じる瞬間にエンドラップを取る
'その後、エンドラップとスタートラップの差をログに書き込む

'ver20101021の変更点
'LogWritterとstopWatchを修正
'分、秒表示しかできなかったが、時に対応した xx時間xx分xx秒(xxxx秒)

'ver20101022の変更点
'リモコンフォームを修正
'デフォルトコントロールである日付のテキストボックスがNULLのとき、実行時エラーを起こしていたが、
'Nz関数を使用して、日付のテキストボックスがNULLだった場合は、長さ 0 の文字列を返すようにした。Nz超便利！

'ver20101118の変更点
'LogWritterの変更に伴い、御茶義理が余計な一行を追加するようになった
'ZEUS登録用に320000行を出力すると、320001行あるのでダメですと返される
'なので、御茶義理がLogWritterをまるごと含むようにした。そしてこれを期に _CrrConn のみにした

'ver20101202の変更点
'UtilYMクラスにGetYesterdayの実装
'昨日営業日を取り出す関数です YYYY/MM/DD で返ります

'ver20110117の変更点【割と致命的】
'Genociderクラスの皆殺しにおいて、アクションクエリは消すが選択クエリは消してくれなかった
'cat.Proceduresと絞っていたのが原因で、cat.ViewsのForLoop文をコピーしてくっつけた
'KillerForQueryの部分も修正

'ver20110118の変更点
'Genociderクラスにおいて、自MDBのみへの処理に絞った

'ver20110119の変更点
'MDBManipulatorにおいて、CreateTableの際に、作成テーブル名と同じテーブルがあったら削除するようにした

'ver20110203の変更点
'SpLogic_Stacker > SpLogic_Stack (クラス名の変更。内容は変えてない）
'SpLogic_Que の作成

'ver20110302の変更点
'Linkerクラスにおいて、コードの整理を行った。MDBManipulatorに一部リンク系のコマンドがあったがすべてこちらに移送

'ver20110317の変更点
'UtilYMクラスにおいて、GetFormalDate2の新設。これはYYYYMMDDのYYYY/MM/DDへの変換である

'ver20110405の変更点
'FSOSuiteクラスにおいて、IsFolderExistsAndMakeFolderの新設。引数フォルダが無かった場合にフォルダを作成

'ver20110427の変更点
'MDBManipuratorクラスにおいて、CreateQueryObjectの無駄を排除　mdbクラスが変数として宣言されていた（無駄な宣言）

'ver20120608の変更点
'MDBManipuratorクラスにおいて、ExportForAcTable、ExportForAcTableAndNameEdit　を追加（外のMDBへテーブルをバシルーラ）
