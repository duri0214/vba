# VBA

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
```Visual Basic:Sample
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
