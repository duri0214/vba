# VBA

## ExcelManipulator
グラフの範囲変更とグラフ種変更のサンプル
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

## SQLServer
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

## ExcelDb
| 図鑑番号 | ポケモン名   | タイプ１ | タイプ２ | 通常特性１       | 通常特性２     | 夢特性         | HP  | こうげき | ぼうぎょ | とくこう | とくぼう | すばやさ | 合計 |
|----------|--------------|----------|----------|------------------|----------------|----------------|-----|----------|----------|----------|----------|----------|------|
| 686      | マーイーカ   | あく     | エスパー | あまのじゃく     | きゅうばん     | すりぬけ       | 53  | 54       | 53       | 37       | 46       | 45       | 288  |
| 687      | カラマネロ   | あく     | エスパー | あまのじゃく     | きゅうばん     | すりぬけ       | 86  | 92       | 88       | 68       | 75       | 73       | 482  |
| 559      | ズルッグ     | あく     | かくとう | だっぴ           | じしんかじょう | いかく         | 50  | 75       | 70       | 35       | 70       | 48       | 348  |
| 560      | ズルズキン   | あく     | かくとう | だっぴ           | じしんかじょう | いかく         | 65  | 90       | 115      | 45       | 115      | 58       | 488  |
| 302      | ヤミラミ     | あく     | ゴースト | するどいめ       | あとだし       | いたずらごころ | 50  | 75       | 75       | 65       | 65       | 50       | 380  |
| 302-1    | メガヤミラミ | あく     | ゴースト | マジックミラー   |                |                | 50  | 85       | 125      | 85       | 115      | 20       | 480  |
| 215      | ニューラ     | あく     | こおり   | せいしんりょく   | するどいめ     | わるいてぐせ   | 55  | 95       | 55       | 35       | 75       | 115      | 430  |
| 461      | マニューラ   | あく     | こおり   | プレッシャー     |                | わるいてぐせ   | 70  | 120      | 65       | 45       | 85       | 125      | 510  |
| 633      | モノズ       | あく     | ドラゴン | はりきり         |                |                | 52  | 65       | 50       | 45       | 50       | 38       | 300  |
| 634      | ジヘッド     | あく     | ドラゴン | はりきり         |                |                | 72  | 85       | 70       | 65       | 70       | 58       | 420  |
| 635      | サザンドラ   | あく     | ドラゴン | ふゆう           |                |                | 92  | 105      | 90       | 125      | 90       | 98       | 600  |
| 799      | アクジキング | あく     | ドラゴン | ビーストブースト |                |                | 223 | 101      | 53       | 97       | 53       | 43       | 570  |
| 019-1    | コラッタ:A   | あく     | ノーマル | くいしんぼう     | はりきり       | あついしぼう   | 30  | 56       | 35       | 25       | 35       | 72       | 253  |
| 020-1    | ラッタ:A     | あく     | ノーマル | くいしんぼう     | はりきり       | あついしぼう   | 75  | 71       | 70       | 40       | 80       | 77       | 413  |
| 624      | コマタナ     | あく     | はがね   | まけんき         | せいしんりょく | プレッシャー   | 45  | 85       | 70       | 40       | 40       | 60       | 340  |
| 625      | キリキザン   | あく     | はがね   | まけんき         | せいしんりょく | プレッシャー   | 65  | 125      | 100      | 60       | 70       | 70       | 490  |
| 198      | ヤミカラス   | あく     | ひこう   | ふみん           | きょううん     | いたずらごころ | 60  | 85       | 42       | 85       | 42       | 91       | 405  |
| 430      | ドンカラス   | あく     | ひこう   | ふみん           | きょううん     | じしんかじょう | 100 | 125      | 52       | 105      | 52       | 71       | 505  |
| 629      | バルチャイ   | あく     | ひこう   | はとむね         | ぼうじん       | くだけるよろい | 70  | 55       | 75       | 45       | 65       | 60       | 370  |
| 630      | バルジーナ   | あく     | ひこう   | はとむね         | ぼうじん       | くだけるよろい | 110 | 65       | 105      | 55       | 95       | 80       | 510  |
| 717      | イベルタル   | あく     | ひこう   | ダークオーラ     |                |                | 126 | 131      | 95       | 131      | 98       | 99       | 680  |
| 228      | デルビル     | あく     | ほのお   | はやおき         | もらいび       | きんちょうかん | 45  | 60       | 30       | 80       | 50       | 65       | 330  |
| 229      | ヘルガー     | あく     | ほのお   | はやおき         | もらいび       | きんちょうかん | 75  | 90       | 50       | 110      | 80       | 95       | 500  |
| 229-1    | メガヘルガー | あく     | ほのお   | サンパワー       |                |                | 75  | 90       | 90       | 140      | 90       | 115      | 600  |
| 052-1    | ニャース:A   | あく     |          | ものひろい       | テクニシャン   | びびり         | 40  | 35       | 35       | 50       | 40       | 90       | 290  |
| 053-1    | ペルシアン:A | あく     |          | ファーコート     | テクニシャン   | びびり         | 65  | 60       | 60       | 75       | 65       | 115      | 440  |
| 197      | ブラッキー   | あく     |          | シンクロ         |                | せいしんりょく | 95  | 65       | 110      | 60       | 130      | 65       | 525  |
| 261      | ポチエナ     | あく     |          | にげあし         | はやあし       | びびり         | 35  | 55       | 35       | 30       | 30       | 35       | 220  |
| 262      | グラエナ     | あく     |          | いかく           | はやあし       | じしんかじょう | 70  | 90       | 70       | 60       | 60       | 70       | 420  |
| 359      | アブソル     | あく     |          | プレッシャー     | きょううん     | せいぎのこころ | 65  | 130      | 60       | 75       | 60       | 75       | 465  |
| 359-1    | メガアブソル | あく     |          | マジックミラー   |                |                | 65  | 150      | 60       | 115      | 60       | 115      | 565  |
| 491      | ダークライ   | あく     |          | ナイトメア       |                |                | 70  | 90       | 90       | 135      | 90       | 125      | 600  |
| 509      | チョロネコ   | あく     |          | じゅうなん       | かるわざ       | いたずらごころ | 41  | 50       | 37       | 50       | 37       | 66       | 281  |
| 510      | レパルダス   | あく     |          | じゅうなん       | かるわざ       | いたずらごころ | 64  | 88       | 50       | 88       | 50       | 106      | 446  |
| 570      | ゾロア       | あく     |          | イリュージョン   |                |                | 40  | 65       | 40       | 80       | 40       | 65       | 330  |

```vb:Sample
    Dim db As New ExcelDb
    Dim hdr As Range
    Dim r As Range
    
    Set hdr = Target.Parent.Range("B2:O2")
    db.SetInit hdr:=hdr, key:="タイプ２2"
    Set r = db.GetCurser("ゴースト")
    If Not r Is Nothing Then
        '検索がヒットしたとき
        db.RecentResult.Select
        Stop
    Else
        '検索が失敗したとき
        Stop
    End If
    
    db.SetInit hdr:=hdr, key:="通常特性１", data:=r
    Set r = db.GetCurser("するどいめ")
    MsgBox db.RecentResult.Address
```

## EditString.cls
### GetStringLengthB(str As String) As Integer
Byteで文字カウント
### StringCutterForFixed(koteichou As String, separate() As Integer) As String()
固定長文字カッター
### StringCutterForCSV(commaMixString As String) As String()
Fit(text As String, retLengthB As Integer) As String

## FSOSuite.cls
### DeleteFile(fileFullPath As String) As String
### GetFilesCount(folderPath As String) As Integer
### GetFoldersCount(folderPath As String) As Integer
### GetFilesCountForCSV(folderPath As String) As Integer
### GetFilesCountForMDB(folderPath As String) As Integer
### IsFolderExists(folderPath As String) As Boolean
### IsFolderExistsAndMakeFolder(folderPath As String)
### IsFileExists(filePath As String) As Boolean
### IsValidFileForMdb(mdbPath As String) As Boolean
### IsValidFileForCSV(csvPath As String) As Boolean
### GetFileName(filePath As String) As String
### GetParentFolderPath(filePath As String) As String
### GetParentFolderName(filePath As String) As String
### ReadHeaderLineByTextFile(textFileFullPath As String) As String
### ReadAllByTextFile(textFileFullPath As String) As String
### InvokeTextArrayByTextFile(textFileFullPath As String) As String()

## Genocider.cls
### LinkGenocider() As String()
### LinkGenociderByTheKeywordBeginning(keyword As String) As String()
### TableGenocider() As String()
### TableGenociderByTheKeywordBeginning(keyword As String) As String()
### QueryGenociderByTheKeywordBeginning(keyword As String) As String()
### FileGenociderForCsvFile(folderPath As String) As String()
### FileGenociderForExcelFile(folderPath As String) As String()

## Importer.cls
### ImportForExcelFiles(folderPath As String, topRecordIsFieldName As Boolean) As String()
### ImportForCSVFiles(folderPath As String, ImportTeigiMei As String, topRecordIsFieldName As Boolean) As String()
### ImportForExcelFile(xlFileFullPath As String, topRecordIsFieldName As Boolean) As String
### ImportForCSVFile(csvFileFullPath As String, ImportTeigiMei As String, topRecordIsFieldName As Boolean) As String
### ImportForAcTable(ToImportMDBFullPath As String, targetTableName As String) As String
### ImportForAcTableAndNameEdit(ToImportMDBFullPath As String, targetTableName As String, newName As String) As String
### ImportForCSVFiles_UsingIni(csvFolder As String) As String()
### ImportForCSVFile_UsingIni(csvFolder As String, csvFileName As String) As String

## Linker.cls
### LinkForExcelFiles(folderPath As String, topRecordIsFieldName As Boolean) As String()
### LinkForCSVFiles(folderPath As String, linkTeigiMei As String, topRecordIsFieldName As Boolean) As String()
### LinkForExcelFile(xlFileFullPath As String, topRecordIsFieldName As Boolean) As String
### LinkForCSVFile(csvFileFullPath As String, linkTeigiMei As String, topRecordIsFieldName As Boolean) As String
### LinkForAcTable(ToLinkMDBFullPath As String, targetTableName As String) As String
### LinkForAcTableAndNameEdit(ToLinkMDBFullPath As String, targetTableName As String, newName As String) As String
### LinkForCSVFiles_UsingIni(csvFolder As String) As String()
### LinkForCSVFile_UsingIni(csvFolder As String, csvFileName As String) As String

## LogWritter.cls
### OpenTextStream(Optional logFileFullPath As String)
### CloseTextStream()
### WriteLineLog(txt As String)
### WriteLine(txt As String)

## MDBManipulator.cls
MDB操作に必要な処理はだいたいこれに詰め込んだ
### GetOwnFolderPath() As String
### GetOwnFileName() As String
### GetOwnFullPath() As String
### GetQueryNamesByTheKeywordBeginning(keyword As String) As String()
### ChangeIndexAtThisArray(oldArray() As String, changePlan As String) As String()
### GetTableSchema(tableName As String, colLikeKeyword As String) As String()
### CreateMDBFile(newMDBFullPath As String) As String
### DeleteMDBFile(MDBFullPath As String) As String
### CreateTable(newTableName As String)
### AddColumns(tableName As String, columnSchema() As String, typeModel As Integer, Optional columnSize As Integer)
### AddColumns2(tableName As String, columnSchema() As String)
### RenameColumns(targetTableName As String, before As String, after As String)
### RenameTable(targetTableName As String, newTableName As String)
### DeleteColumn(targetTableName As String, delColumnName As String)
### DeleteTable(delTableName As String)
### DeleteQueryObject(delQueryName As String)
### ExecuteSQL(sql As String)
### ExecuteQuery(queryName As String)
### Export2Excel(outputTableOrQuery As String, outputFolder As String) As String
### Export2ExcelAndNameEdit(outputTableOrQuery As String, newFileFullPath As String) As String
### CountOfTableRecs(targetTable As String) As Long
### SetTableDescription(targetTable As String, setumei As String)
### SetPrimaryKey(targetTable As String, keyString As String)
### DeletePrimaryKey(targetTable As String)
### CreateQueryObject(queryContext As String, newQueryName As String)
### ExportForAcTable(ToExportMDBFullPath As String, targetTableName As String) As String
### ExportForAcTableAndNameEdit(ToExportMDBFullPath As String, targetTableName As String, newName As String) As String

## UtilYM.cls
日付処理に必要な処理はだいたいこれに詰め込んだ
### GetFirstYM_InThisPeriod(YYYYAB As String) As String
### GetLastYM_InThisPeriod(YYYYAB As String) As String
### GetTheEndOfTheMonth(YYYYMM As String) As String
### GetTheEndOfTheMonth_RetEM(YYYYMM As String) As String
### GetNowYM() As String
### GetAddYM(i As Integer) As String
### GetAddYM2(YYYYMM As String, i As Integer) As String
### GetYMInterval(kaishiYM As String, shuryoYM As String) As Integer
### GetKaikeiNendo(YYYYMM As String) As String
### GetPeriodYAB(YYYYMM As String) As String
### GetPeriodEAB(YYYYMM As String) As String
### GetPeriod_Prev_RetYAB(i As Integer) As String
### GetPeriod_Prev_RetEAB(i As Integer) As String
### GetPeriod_Prev_RetFirstYM(i As Integer) As String
### GetPeriod_Prev_RetFirstEM(i As Integer) As String
### GetPeriod_Prev_RetLastYM(i As Integer) As String
### GetPeriod_Prev_RetLastEM(i As Integer) As String
### GetFormalDate(YYYYMM As String) As String
### GetFormalDate2(YYYYMMDD As String) As String
### GetPeriodListYAB(kaishiYM As String, shuryoYM As String) As String()
### GetPeriodListEAB(kaishiYM As String, shuryoYM As String) As String()
### GetYesterday() As String
### ConvertYM2EM(YYYYMM As String) As String
### ConvertY2E(yyyy As String) As String
