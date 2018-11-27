# ExcelManipulator
セル範囲を文字結合する
```vb
Private Sub btnConcat_Click()
    Dim u As New ExcelManipulator
    MsgBox u.ConcatenateRangeValue(ActiveSheet.Range("B23:C24"), ",")
End Sub
```

vba本流のCopyFileは、自分自身が開いている場合の別名コピーができない
```vb
Private Sub btnFSOCopy_Click()
    Dim fso As New FSOSuite
    fso.CopyFile "C:\Users\yoshi\Desktop\パソコン処分.pdf", "C:\Users\yoshi\Desktop\パソコン処分_cp.pdf"
End Sub
```

そのシート名があるかを bool で返す
```vb
Private Sub checkTheSheet_Click()

    Dim u As New ExcelManipulator
    MsgBox u.IsSheetExists(ThisWorkbook, "A"), vbInformation, "A"
    MsgBox u.IsSheetExists(ThisWorkbook, "Q"), vbInformation, "Q"

End Sub
```

左上隅をA1に移動する
```vb
Private Sub btnMoveA1_Click()

    Dim u As New ExcelManipulator
    
    u.MoveA1 ActiveSheet

End Sub
```

そのフォルダのなかで一番新しくできたファイルパスを返す
```vb
Private Sub btnNewestFile_Click()

    Dim fso As New FSOSuite
    
    MsgBox fso.GetNewestFileInTheFolder("E:\OneDrive\Desktop")
    
End Sub
```

pathが誰かによって開かれている場合、ファイル名の末尾に年月日時分秒をつけて返す
```vb
Private Sub btnNewname_Click()

    Dim fso As New FSOSuite
    
    MsgBox fso.GetNewNameIfTheFileIsUsing("E:\OneDrive\Desktop\キャッシュフロー計算書.xlsx")
    
End Sub
```

ブックの名前の定義を削除
```vb
Private Sub btnNameKill_Click()

    Dim u As New ExcelManipulator
    
    u.GenocideNameInTheBook ThisWorkbook
    
End Sub
```

外部ファイルへのリンクを解除
```vb
Private Sub btnLinkKill_Click()

    Dim u As New ExcelManipulator
    
    'エクセルファイルにはびこる「外部ファイルへのリンク」を解除します
    u.GenocideExternalBookLink ThisWorkbook
    
End Sub
```

各種自パス取得
```vb
Private Sub btnOwnPath_Click()
    
    Dim u As New ExcelManipulator
    
    MsgBox "u.GetOwnFileName()：" & u.GetOwnFileName()
    MsgBox "u.GetOwnFolderPath()：" & u.GetOwnFolderPath()
    MsgBox "u.GetOwnFullPath()：" & u.GetOwnFullPath()
    
End Sub
```

グラフの範囲変更とグラフ種変更のサンプル
```vb
Private Sub btnGraph_Click()
    
    Dim u As New ExcelManipulator
    Dim g As New Graphman
    Dim hanrei As Range
    Dim header As Range
    
    Set hanrei = u.GetRegion(ActiveSheet.Range("A2"), xlDown)
    Set header = u.GetRegion(ActiveSheet.Range("B1"), xlToRight)
    
    '「グラフ 1」の範囲を変更する
    g.Init ActiveSheet, "グラフ 1"
    g.SetGraphRange header, hanrei
    
    g.SetGraphType ActiveSheet, "グラフ 1", "合計", xlLine, xlSecondary, True, , 49407
    g.SetGraphLabel 1, 9, rgb(191, 191, 191)
    g.SetGraphLabel_adjustXY 1, 9, -20, -20
    
End Sub
```
オートフィルタ後の範囲をRange取得するサンプル
```vb
Private Sub btnFilter_Click()

    Dim u As New ExcelManipulator
    Dim r As Range
    
    '重ね掛けができます
    ActiveSheet.AutoFilterMode = False
    u.GetFilteredRange ActiveSheet.Range("A1:C1"), "MK:1000,9999"
    Set r = u.GetFilteredRange(ActiveSheet.Range("A1:C1"), "ステータス:B,C")
    r.Copy ThisWorkbook.Worksheets("Autofiltered").Range("A1")
    Application.CutCopyMode = False
    
End Sub
```
ピボットテーブルを作成するサンプル
```vb
Private Sub btnPivot_Click()
    
    Dim u As New ExcelManipulator
    Dim newSheet As Worksheet

    '作成するシートの決定とシート名の決定
    Set newSheet = Sheets.Add
    newSheet.Name = "pvt"

    Dim pvt_name As String
    Dim pvt_group As String
    Dim pvt_agg As String
    Dim pvt_data As Range
    Dim pvt_destination As Range
    
    Set pvt_data = ThisWorkbook.Worksheets("db").Range("J27:L63")
    
    'A）Pivotテーブルを仕込む
    pvt_name = "pivot"
    pvt_group = "MK,ステータス"
    pvt_agg = "dammy"
    Set pvt_destination = newSheet.Range("A1")
    u.CreatePivotTable pvt_name, pvt_group, pvt_agg, pvt_data, pvt_destination
    
    'A）Pivotテーブルにフィルターを仕込む
    u.SetFilterOnPivotTable ActiveSheet.PivotTables(1), "ステータス", "A,C", False
    u.SetFilterOnPivotTable ActiveSheet.PivotTables(1), "ステータス", "A,C", True
    
End Sub
```

# Array2d_v
variant2次元配列を操作します
```vb
Private Sub btnStart_Click()

    Dim a As New Array2d_v

    Dim u As New ExcelManipulator
    Dim r As Range
    Dim data As Variant
    Dim oneline As Variant
    
    '基準セルB10のend.rightとend.downで矩形を範囲取り、variantに代入
    Set r = u.GetRegion(ActiveSheet.Range("B10"), xlToRight)
    Set r = u.GetRegion(r, xlDown)
    data = r
    
    '次の行に入れたいonelineを作成する
    oneline = r.Resize(1).Offset(-1)
    
    '追加するonelineと追加されるdataのヨコサイズは一致していないといけない
    data = a.AddNewRecordAndConcat(data, oneline)
    
    '追加された配列をセルに貼り付けるために範囲を把握
    Set r = r(1, 1).Resize(UBound(data), UBound(data, 2))
    
    '貼り付け
    r = data

End Sub
```

# SQLServer
SQLServerからデータを取ってセルにペーストするまでのサンプル
```vb
Private Sub btnSQLServer_Click()
    
    Dim r As Range
    Dim mssql As New SQLServer
    Dim rs As ADODB.Recordset
    Dim editSql As String
    
    Set r = ActiveSheet.Range("A10")
    
    mssql.Connect "localhost\SQLExpress", "db", "python", "python"
    
    editSql = "SELECT * FROM [db].[dbo].[顧客]"
    Set rs = mssql.GetCursor(editSql)
    
    If Not rs.EOF Then
        r.CopyFromRecordset rs
    End If
    
    mssql.DisConnect
    
End Sub
```

# ExcelDb
シートのB2に以下のサンプルの表を貼り付けてから実行

| 図鑑番号 | ポケモン名 | タイプ１ | タイプ２ | 通常特性１ | 通常特性２ | 夢特性 | HP | こうげき | ぼうぎょ | とくこう | とくぼう | すばやさ | 合計 |
|----------|:------------:|---------:|----------|------------------|-----------------|----------------|-----|----------|----------|----------|----------|----------|------|
| 686 | マーイーカ | あく | エスパー | あまのじゃく |  きゅうばん | すりぬけ | 53 | 54 | 53 | 37 | 46 | 45 | 288 |
| 687 | カラマネロ | あく | エスパー | あまのじゃく |  きゅうばん | すりぬけ | 86 | 92 | 88 | 68 | 75 | 73 | 482 |
| 559 | ズルッグ | あく | かくとう | だっぴ |  じしんかじょう | いかく | 50 | 75 | 70 | 35 | 70 | 48 | 348 |
| 560 | ズルズキン | あく | かくとう | だっぴ |  じしんかじょう | いかく | 65 | 90 | 115 | 45 | 115 | 58 | 488 |
| 302 | ヤミラミ | あく | ゴースト | するどいめ |  あとだし | いたずらごころ | 50 | 75 | 75 | 65 | 65 | 50 | 380 |
| 302-1 | メガヤミラミ | あく | ゴースト | マジックミラー | 　 | 　 | 50 | 85 | 125 | 85 | 115 | 20 | 480 |
| 215 | ニューラ | あく | こおり | せいしんりょく |  するどいめ | わるいてぐせ | 55 | 95 | 55 | 35 | 75 | 115 | 430 |
| 461 | マニューラ | あく | こおり | プレッシャー | 　 | わるいてぐせ | 70 | 120 | 65 | 45 | 85 | 125 | 510 |
| 633 | モノズ | あく | ドラゴン | はりきり | 　 | 　 | 52 | 65 | 50 | 45 | 50 | 38 | 300 |
| 634 | ジヘッド | あく | ドラゴン | はりきり | 　 | 　 | 72 | 85 | 70 | 65 | 70 | 58 | 420 |
| 635 | サザンドラ | あく | ドラゴン | ふゆう | 　 | 　 | 92 | 105 | 90 | 125 | 90 | 98 | 600 |
| 799 | アクジキング | あく | ドラゴン | ビーストブースト | 　 | 　 | 223 | 101 | 53 | 97 | 53 | 43 | 570 |

```vb
'Variant返し
Private Sub btnSearch_Click()

    'ExcelDb のテスト
    
    Dim db As New ExcelDb
    Dim hdr As Range
    Dim data As Range
    
    Dim search() As String
    Dim ret As Variant
    Dim r As Range
    
    Set hdr = ActiveSheet.Range("B2:O2")
    Set data = ActiveSheet.Range("B3:O14")
    
    'Search #Init
    db.SetInit hdr, data
    
    'Search #1（1発目の検索）
    search = Split(InputBox("項目名,検索語句", , "タイプ２,ゴースト"), ",")
    ret = db.GetCurser(search(0), search(1))
    '検索結果を書き込む
    Set r = ActiveSheet.Range("B18").Resize(UBound(ret), UBound(ret, 2))
    r = ret
    
    'Search #2（1発目の検索結果を使って2発目の検索）
    search = Split(InputBox("項目名,検索語句", , "通常特性１,するどいめ"), ",")
    ret = db.GetCurser(search(0), search(1))
    '検索結果を書き込む
    Set r = ActiveSheet.Range("B22").Resize(UBound(ret), UBound(ret, 2))
    r = ret

End Sub
```
```vb
'Range返し
Private Sub btnRetRange_Click()
    
    Dim db As New ExcelDb
    Dim hdr As Range
    Dim data As Range
    Dim search() As String
    Dim r As Range
    
    Set hdr = ActiveSheet.Range("B2:O2")
    Set data = ActiveSheet.Range("B3:O14")
    
    'Search #Init
    db.SetInit hdr, data
    
    'Search
    search = Split(InputBox("項目名,検索語句", , "タイプ２,ゴースト"), ",")
    Set r = db.GetCurser_r(search(0), search(1))

    If Not r Is Nothing Then
        r.Select
    Else
        MsgBox "検索結果0件"
    End If
    
End Sub
```

# EditString.cls
```vb
GetStringLengthB(str As String) As Integer
Byteで文字カウント
StringCutterForFixed(koteichou As String, separate() As Integer) As String()
固定長文字カッター
StringCutterForCSV(commaMixString As String) As String()
Fit(text As String, retLengthB As Integer) As String
```
# FSOSuite.cls
```vb
DeleteFile(fileFullPath As String) As String
GetFilesCount(folderPath As String) As Integer
GetFoldersCount(folderPath As String) As Integer
GetFilesCountForCSV(folderPath As String) As Integer
GetFilesCountForMDB(folderPath As String) As Integer
IsFolderExists(folderPath As String) As Boolean
IsFolderExistsAndMakeFolder(folderPath As String)
IsFileExists(filePath As String) As Boolean
IsValidFileForMdb(mdbPath As String) As Boolean
IsValidFileForCSV(csvPath As String) As Boolean
GetFileName(filePath As String) As String
GetParentFolderPath(filePath As String) As String
GetParentFolderName(filePath As String) As String
ReadHeaderLineByTextFile(textFileFullPath As String) As String
ReadAllByTextFile(textFileFullPath As String) As String
InvokeTextArrayByTextFile(textFileFullPath As String) As String()
```
# Genocider.cls
```vb
LinkGenocider() As String()
LinkGenociderByTheKeywordBeginning(keyword As String) As String()
TableGenocider() As String()
TableGenociderByTheKeywordBeginning(keyword As String) As String()
QueryGenociderByTheKeywordBeginning(keyword As String) As String()
FileGenociderForCsvFile(folderPath As String) As String()
FileGenociderForExcelFile(folderPath As String) As String()
```
# Importer.cls
```vb
ImportForExcelFiles(folderPath As String, topRecordIsFieldName As Boolean) As String()
ImportForCSVFiles(folderPath As String, ImportTeigiMei As String, topRecordIsFieldName As Boolean) As String()
ImportForExcelFile(xlFileFullPath As String, topRecordIsFieldName As Boolean) As String
ImportForCSVFile(csvFileFullPath As String, ImportTeigiMei As String, topRecordIsFieldName As Boolean) As String
ImportForAcTable(ToImportMDBFullPath As String, targetTableName As String) As String
ImportForAcTableAndNameEdit(ToImportMDBFullPath As String, targetTableName As String, newName As String) As String
ImportForCSVFiles_UsingIni(csvFolder As String) As String()
ImportForCSVFile_UsingIni(csvFolder As String, csvFileName As String) As String
```
# Linker.cls
```vb
LinkForExcelFiles(folderPath As String, topRecordIsFieldName As Boolean) As String()
LinkForCSVFiles(folderPath As String, linkTeigiMei As String, topRecordIsFieldName As Boolean) As String()
LinkForExcelFile(xlFileFullPath As String, topRecordIsFieldName As Boolean) As String
LinkForCSVFile(csvFileFullPath As String, linkTeigiMei As String, topRecordIsFieldName As Boolean) As String
LinkForAcTable(ToLinkMDBFullPath As String, targetTableName As String) As String
LinkForAcTableAndNameEdit(ToLinkMDBFullPath As String, targetTableName As String, newName As String) As String
LinkForCSVFiles_UsingIni(csvFolder As String) As String()
LinkForCSVFile_UsingIni(csvFolder As String, csvFileName As String) As String
```
# LogWritter.cls
```vb
OpenTextStream(Optional logFileFullPath As String)
CloseTextStream()
WriteLineLog(txt As String)
WriteLine(txt As String)
```

# MDBManipulator.cls
MDB操作に必要な処理はだいたいこれに詰め込んだ
```vb
GetOwnFolderPath() As String
GetOwnFileName() As String
GetOwnFullPath() As String
GetQueryNamesByTheKeywordBeginning(keyword As String) As String()
ChangeIndexAtThisArray(oldArray() As String, changePlan As String) As String()
GetTableSchema(tableName As String, colLikeKeyword As String) As String()
CreateMDBFile(newMDBFullPath As String) As String
DeleteMDBFile(MDBFullPath As String) As String
CreateTable(newTableName As String)
AddColumns(tableName As String, columnSchema() As String, typeModel As Integer, Optional columnSize As Integer)
AddColumns2(tableName As String, columnSchema() As String)
RenameColumns(targetTableName As String, before As String, after As String)
RenameTable(targetTableName As String, newTableName As String)
DeleteColumn(targetTableName As String, delColumnName As String)
DeleteTable(delTableName As String)
DeleteQueryObject(delQueryName As String)
ExecuteSQL(sql As String)
ExecuteQuery(queryName As String)
Export2Excel(outputTableOrQuery As String, outputFolder As String) As String
Export2ExcelAndNameEdit(outputTableOrQuery As String, newFileFullPath As String) As String
CountOfTableRecs(targetTable As String) As Long
SetTableDescription(targetTable As String, setumei As String)
SetPrimaryKey(targetTable As String, keyString As String)
DeletePrimaryKey(targetTable As String)
CreateQueryObject(queryContext As String, newQueryName As String)
ExportForAcTable(ToExportMDBFullPath As String, targetTableName As String) As String
ExportForAcTableAndNameEdit(ToExportMDBFullPath As String, targetTableName As String, newName As String) As String
```
# UtilYM.cls
日付処理に必要な処理はだいたいこれに詰め込んだ
```vb
Private Sub btnUtilYM_Click()
    
    Dim ym As New UtilYM
    
    Dim yyyyab As String
    Dim yyyymm As String
    Dim ym_shift As Integer
    Dim yyyymmdd As String
    Dim kaishiYM As String
    Dim shuryoYM As String
    Dim yyyy As String
    
    '1年には上期(4月-9月)と下期(10月-3月)があり、「上期」「下期」の単語は、ソートすると順番的には逆になってしまう。
    'そこでそれぞれ「A」「B」になるという経緯があった
    yyyyab = "2018A"
    yyyymm = "201806"
    ym_shift = 3
    yyyymmdd = "20180714"
    kaishiYM = "201807"
    shuryoYM = "201812"
    yyyy = "2018"
    
    '「期」の情報を入力し、その「期」の期初月を返す（例：2018A→201804）
    MsgBox ym.GetFirstYM_InThisPeriod(yyyyab)

    '「期」の情報を入力し、その「期」の期末月を返す（例：2018A→201809）
    MsgBox ym.GetLastYM_InThisPeriod(yyyyab)

    '西暦年月を年月の末日に変換する（例：201806 → 20180630）
    MsgBox ym.GetTheEndOfTheMonth(yyyymm)

    '和暦年月を年月の末日に変換する（例：201806 → 300630）
    MsgBox ym.GetTheEndOfTheMonth_RetEM(yyyymm)

    'systemYMを返す
    MsgBox ym.GetNowYM()

    'systemYMに値を加えます（例：201810を現在として、3と引数指定すると201901）
    MsgBox ym.GetAddYM(ym_shift)

    '基準YMに値を加えます（例：201806を基準として、3と引数指定すると201809）
    MsgBox ym.GetAddYM2(yyyymm, ym_shift)

    'kaishiYM引数に201807, shuryoYM引数に201812と指定すると、5が返る。
    MsgBox ym.GetYMInterval(kaishiYM, shuryoYM)

    '指定年月の会計年度を返します（例：201801～201803 → 2017）
    MsgBox ym.GetKaikeiNendo(yyyymm)

    '引数にYMを入力すると、YYYYAB形式で返る（例：201806 → 2018A）
    MsgBox ym.GetPeriodYAB(yyyymm)

    '引数にYMを入力すると、YYYYAB形式で返る（例：201806 → 30A）
    MsgBox ym.GetPeriodEAB(yyyymm)

    '引数値YYYYMMを、YYYY/MM/01のフォーマットに変換
    MsgBox ym.GetFormalDate(yyyymm)

    '引数値YYYYMMDDフォーマットの年月を、YYYY/MM/DDのフォーマットに変換
    MsgBox ym.GetFormalDate2(yyyymmdd)

    '昨日の「営業日」を返す。昨日が週末(土, 日)なら再起的に -1 の引き算して平日を返す
    MsgBox ym.GetYesterday()

    'YMをEMに変換（例：201804 → 3004）
    MsgBox ym.ConvertYM2EM(yyyymm)

    'YをEに変換（例：2018 → 30）
    MsgBox ym.ConvertY2E(yyyy)
    
    'マニアックなやつ
    '当月（systemYM - 1）から、期をステップバックした西暦期（YYYYAB）を算出
    MsgBox ym.GetPeriod_Prev_RetYAB(ym_shift)
    
    '当月（systemYM - 1）から、期をステップバックした和暦期（EEAB）を算出
    MsgBox ym.GetPeriod_Prev_RetEAB(ym_shift)
    
    '当月（systemYM - 1）から、期をステップバックした西暦期（YYYYAB）の期初月次を算出
    MsgBox ym.GetPeriod_Prev_RetFirstYM(ym_shift)
    
    '当月（systemYM - 1）から、期をステップバックした和暦期（EEAB）の期初月次を算出
    MsgBox ym.GetPeriod_Prev_RetFirstEM(ym_shift)
    
    '当月（systemYM - 1）から、期をステップバックした西暦期（YYYYAB）の期末月次を算出
    MsgBox ym.GetPeriod_Prev_RetLastYM(ym_shift)
    
    '当月（systemYM - 1）から、期をステップバックした和暦期（EEAB）の期末月次を算出
    MsgBox ym.GetPeriod_Prev_RetLastEM(ym_shift)
    
    '201807～201812までを指定すると「2018A,2018B」という配列を返す
    MsgBox Join(ym.GetPeriodListYAB(kaishiYM, shuryoYM), ",")
    
    '201807～201812までを指定すると、「30A,30B」という配列を返す
    MsgBox Join(ym.GetPeriodListEAB(kaishiYM, shuryoYM), ",")

End Sub
```
