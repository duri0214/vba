# ExcelManipulator
## グラフの範囲変更とグラフ種変更のサンプル
```
Private Sub btnGraph_Click()
    
    Dim u As New ExcelManipulator
    Dim hanrei As Range
    Dim header As Range
    
    Set hanrei = u.GetRegion(ActiveSheet.Range("A2"), xlDown)
    Set header = u.GetRegion(ActiveSheet.Range("B1"), xlToRight)
    
    '「グラフ 1」の範囲を変更する
    u.ChangingTheGraphRange "グラフ 1", hanrei, header, xlColumnStacked
    
    u.ChangingTheGraphType ActiveSheet, "グラフ 1", "合計", xlLine, xlSecondary, True, , 49407
    
End Sub
```
## オートフィルタ後の範囲をRange取得するサンプル
```
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
## ピボットテーブルを作成するサンプル
```
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

# SQLServer
SQLServerからデータを取ってセルにペーストするまでのサンプル
```
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

```
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
    Set r = ActiveSheet.Range("B18").Resize(UBound(ret), UBound(ret, 2))
    r = ret
    ActiveSheet.Cells(db.origin_y, db.origin_x).Resize(UBound(ret), UBound(ret, 2)).Select
    MsgBox "db.origin_y:" & db.origin_y & ", db.origin_x:" & db.origin_x
    
    'Search #2（1発目の検索結果を使って2発目の検索）
    search = Split(InputBox("項目名,検索語句", , "通常特性１,するどいめ"), ",")
    ret = db.GetCurser(search(0), search(1))
    Set r = ActiveSheet.Range("B22").Resize(UBound(ret), UBound(ret, 2))
    r = ret
    ActiveSheet.Cells(db.origin_y, db.origin_x).Resize(UBound(ret), UBound(ret, 2)).Select
    MsgBox "db.origin_y:" & db.origin_y & ", db.origin_x:" & db.origin_x

End Sub
```

# EditString.cls
## GetStringLengthB(str As String) As Integer
Byteで文字カウント
## StringCutterForFixed(koteichou As String, separate() As Integer) As String()
固定長文字カッター
## StringCutterForCSV(commaMixString As String) As String()
Fit(text As String, retLengthB As Integer) As String

# FSOSuite.cls
## DeleteFile(fileFullPath As String) As String
## GetFilesCount(folderPath As String) As Integer
## GetFoldersCount(folderPath As String) As Integer
## GetFilesCountForCSV(folderPath As String) As Integer
## GetFilesCountForMDB(folderPath As String) As Integer
## IsFolderExists(folderPath As String) As Boolean
## IsFolderExistsAndMakeFolder(folderPath As String)
## IsFileExists(filePath As String) As Boolean
## IsValidFileForMdb(mdbPath As String) As Boolean
## IsValidFileForCSV(csvPath As String) As Boolean
## GetFileName(filePath As String) As String
## GetParentFolderPath(filePath As String) As String
## GetParentFolderName(filePath As String) As String
## ReadHeaderLineByTextFile(textFileFullPath As String) As String
## ReadAllByTextFile(textFileFullPath As String) As String
## InvokeTextArrayByTextFile(textFileFullPath As String) As String()

# Genocider.cls
## LinkGenocider() As String()
## LinkGenociderByTheKeywordBeginning(keyword As String) As String()
## TableGenocider() As String()
## TableGenociderByTheKeywordBeginning(keyword As String) As String()
## QueryGenociderByTheKeywordBeginning(keyword As String) As String()
## FileGenociderForCsvFile(folderPath As String) As String()
## FileGenociderForExcelFile(folderPath As String) As String()

# Importer.cls
## ImportForExcelFiles(folderPath As String, topRecordIsFieldName As Boolean) As String()
## ImportForCSVFiles(folderPath As String, ImportTeigiMei As String, topRecordIsFieldName As Boolean) As String()
## ImportForExcelFile(xlFileFullPath As String, topRecordIsFieldName As Boolean) As String
## ImportForCSVFile(csvFileFullPath As String, ImportTeigiMei As String, topRecordIsFieldName As Boolean) As String
## ImportForAcTable(ToImportMDBFullPath As String, targetTableName As String) As String
## ImportForAcTableAndNameEdit(ToImportMDBFullPath As String, targetTableName As String, newName As String) As String
## ImportForCSVFiles_UsingIni(csvFolder As String) As String()
## ImportForCSVFile_UsingIni(csvFolder As String, csvFileName As String) As String

# Linker.cls
## LinkForExcelFiles(folderPath As String, topRecordIsFieldName As Boolean) As String()
## LinkForCSVFiles(folderPath As String, linkTeigiMei As String, topRecordIsFieldName As Boolean) As String()
## LinkForExcelFile(xlFileFullPath As String, topRecordIsFieldName As Boolean) As String
## LinkForCSVFile(csvFileFullPath As String, linkTeigiMei As String, topRecordIsFieldName As Boolean) As String
## LinkForAcTable(ToLinkMDBFullPath As String, targetTableName As String) As String
## LinkForAcTableAndNameEdit(ToLinkMDBFullPath As String, targetTableName As String, newName As String) As String
## LinkForCSVFiles_UsingIni(csvFolder As String) As String()
## LinkForCSVFile_UsingIni(csvFolder As String, csvFileName As String) As String

# LogWritter.cls
## OpenTextStream(Optional logFileFullPath As String)
## CloseTextStream()
## WriteLineLog(txt As String)
## WriteLine(txt As String)

# MDBManipulator.cls
MDB操作に必要な処理はだいたいこれに詰め込んだ
## GetOwnFolderPath() As String
## GetOwnFileName() As String
## GetOwnFullPath() As String
## GetQueryNamesByTheKeywordBeginning(keyword As String) As String()
## ChangeIndexAtThisArray(oldArray() As String, changePlan As String) As String()
## GetTableSchema(tableName As String, colLikeKeyword As String) As String()
## CreateMDBFile(newMDBFullPath As String) As String
## DeleteMDBFile(MDBFullPath As String) As String
## CreateTable(newTableName As String)
## AddColumns(tableName As String, columnSchema() As String, typeModel As Integer, Optional columnSize As Integer)
## AddColumns2(tableName As String, columnSchema() As String)
## RenameColumns(targetTableName As String, before As String, after As String)
## RenameTable(targetTableName As String, newTableName As String)
## DeleteColumn(targetTableName As String, delColumnName As String)
## DeleteTable(delTableName As String)
## DeleteQueryObject(delQueryName As String)
## ExecuteSQL(sql As String)
## ExecuteQuery(queryName As String)
## Export2Excel(outputTableOrQuery As String, outputFolder As String) As String
## Export2ExcelAndNameEdit(outputTableOrQuery As String, newFileFullPath As String) As String
## CountOfTableRecs(targetTable As String) As Long
## SetTableDescription(targetTable As String, setumei As String)
## SetPrimaryKey(targetTable As String, keyString As String)
## DeletePrimaryKey(targetTable As String)
## CreateQueryObject(queryContext As String, newQueryName As String)
## ExportForAcTable(ToExportMDBFullPath As String, targetTableName As String) As String
## ExportForAcTableAndNameEdit(ToExportMDBFullPath As String, targetTableName As String, newName As String) As String

# UtilYM.cls
日付処理に必要な処理はだいたいこれに詰め込んだ
## GetFirstYM_InThisPeriod(YYYYAB As String) As String
## GetLastYM_InThisPeriod(YYYYAB As String) As String
## GetTheEndOfTheMonth(YYYYMM As String) As String
## GetTheEndOfTheMonth_RetEM(YYYYMM As String) As String
## GetNowYM() As String
## GetAddYM(i As Integer) As String
## GetAddYM2(YYYYMM As String, i As Integer) As String
## GetYMInterval(kaishiYM As String, shuryoYM As String) As Integer
## GetKaikeiNendo(YYYYMM As String) As String
## GetPeriodYAB(YYYYMM As String) As String
## GetPeriodEAB(YYYYMM As String) As String
## GetPeriod_Prev_RetYAB(i As Integer) As String
## GetPeriod_Prev_RetEAB(i As Integer) As String
## GetPeriod_Prev_RetFirstYM(i As Integer) As String
## GetPeriod_Prev_RetFirstEM(i As Integer) As String
## GetPeriod_Prev_RetLastYM(i As Integer) As String
## GetPeriod_Prev_RetLastEM(i As Integer) As String
## GetFormalDate(YYYYMM As String) As String
## GetFormalDate2(YYYYMMDD As String) As String
## GetPeriodListYAB(kaishiYM As String, shuryoYM As String) As String()
## GetPeriodListEAB(kaishiYM As String, shuryoYM As String) As String()
## GetYesterday() As String
## ConvertYM2EM(YYYYMM As String) As String
## ConvertY2E(yyyy As String) As String
