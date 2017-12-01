VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Importer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Function ImportForExcelFiles(folderPath As String, topRecordIsFieldName As Boolean) As String()

    'acSpreadsheetTypeExcel9:Excel2000
    'topRecordIsFieldNameをTrueにすると、１行目が項目行扱いになる。
    
    Dim xlFilePath As String
    Dim buf As String
    Dim ImportageTableNewName As String
    
    Dim retValue As String
        
    buf = Dir(folderPath & "\*.xls")                    'フォルダ内からエクセルファイルを探す
    
    Do While buf <> ""
        
        xlFilePath = folderPath & "\" & buf             'リンクするエクセルファイルへのアドレス取得
        ImportageTableNewName = Replace(buf, ".xls", "")  '「.xls」の文字を消したものがリンクテーブル名
        
        'エクセルリンク。バージョンはexcel2000, リンクテーブル名, エクセルファイル保存場所, トップレコードを項目名とするかどうか
        DoCmd.TransferSpreadsheet _
            acImport, _
            acSpreadsheetTypeExcel9, _
            ImportageTableNewName, _
            xlFilePath, _
            topRecordIsFieldName
        
        retValue = retValue & ImportageTableNewName & ","

        buf = Dir()                                     '次のエクセルファイルを探す
    Loop
    
    ImportForExcelFiles = Split(Left(retValue, Len(retValue) - 1), ",")

End Function

Function ImportForCSVFiles(folderPath As String, ImportTeigiMei As String, topRecordIsFieldName As Boolean) As String()

    'acSpreadsheetTypeExcel9:Excel2000
    'ImportTeigiMeiはリンク定義を文字列で指定する
    'topRecordIsFieldNameをTrueにすると、１行目が項目行扱いになる。
    
    Dim csvFilePath As String
    Dim buf As String
    Dim ImportageTableNewName As String
    
    Dim retValue As String
    
    buf = Dir(folderPath & "\*.csv")                    'フォルダ内からCSVファイルを探す
    
    Do While buf <> ""
        
        csvFilePath = folderPath & "\" & buf             'リンクするCSVファイルへのアドレス取得
        ImportageTableNewName = Replace(buf, ".csv", "")  '「.csv」の文字を消したものがリンクテーブル名
        
        'CSVリンクモード,リンク定義名,リンクテーブル名, CSVファイル保存場所, トップレコードを項目名とするかどうか
        DoCmd.TransferText _
            acImportDelim, _
            ImportTeigiMei, _
            ImportageTableNewName, _
            csvFilePath, _
            topRecordIsFieldName
            
        retValue = retValue & ImportageTableNewName & ","
        
        buf = Dir()                                     '次のCSVファイルを探す
    Loop
    
    ImportForCSVFiles = Split(Left(retValue, Len(retValue) - 1), ",")

End Function

Function ImportForExcelFile(xlFileFullPath As String, topRecordIsFieldName As Boolean) As String

    'acSpreadsheetTypeExcel9:Excel2000
    'topRecordIsFieldNameをTrueにすると、１行目が項目行扱いになる。
    
    Dim buf As String
    Dim ImportageTableNewName As String
    
    Dim retValue As String
    
    buf = Dir(xlFileFullPath)                           'フォルダ内からエクセルファイルを探す
    
    If buf <> "" Then
        
        ImportageTableNewName = Replace(buf, ".xls", "")  '「.xls」の文字を消したものがリンクテーブル名
        
        'エクセルリンク。バージョンはexcel2000, リンクテーブル名, エクセルファイル保存場所, トップレコードを項目名とするかどうか
        DoCmd.TransferSpreadsheet _
            acImport, _
            acSpreadsheetTypeExcel9, _
            ImportageTableNewName, _
            xlFileFullPath, _
            topRecordIsFieldName
            
        retValue = retValue & ImportageTableNewName
    
    End If
    
    ImportForExcelFile = retValue
    
End Function

Function ImportForCSVFile(csvFileFullPath As String, ImportTeigiMei As String, topRecordIsFieldName As Boolean) As String

    'acSpreadsheetTypeExcel9:Excel2000
    'ImportTeigiMeiはリンク定義を文字列で指定する
    'topRecordIsFieldNameをTrueにすると、１行目が項目行扱いになる。
    
    Dim buf As String
    Dim ImportageTableNewName As String
    
    Dim retValue As String
    
    buf = Dir(csvFileFullPath)                          'フォルダ内からCSVファイルを探す
    
    If buf <> "" Then
        
        ImportageTableNewName = Replace(buf, ".csv", "")  '「.csv」の文字を消したものがリンクテーブル名
        
        'CSVリンクモード,リンク定義名,リンクテーブル名, CSVファイル保存場所, トップレコードを項目名とするかどうか
        DoCmd.TransferText _
            acImportDelim, _
            ImportTeigiMei, _
            ImportageTableNewName, _
            csvFileFullPath, _
            topRecordIsFieldName
        
        retValue = retValue & ImportageTableNewName
            
    End If
    
    ImportForCSVFile = retValue

End Function

Function ImportForAcTable(ToImportMDBFullPath As String, targetTableName As String) As String
    
    '▲notice
    'ToImportMDBFullPath  :原リンク側（リンクしたいテーブル）

    Dim retValue As String

    'accessリンク。リンクモード,データベースタイプはACCESS, リンクするMDB, テーブルをリンクする, リンクしたいテーブル名, リンクした後の新しいテーブル名
    DoCmd.TransferDatabase _
        acImport, _
        "Microsoft Access", _
        ToImportMDBFullPath, _
        acTable, _
        targetTableName, _
        targetTableName

    retValue = targetTableName

    ImportForAcTable = retValue

    Application.RefreshDatabaseWindow
    
End Function

Function ImportForAcTableAndNameEdit(ToImportMDBFullPath As String, targetTableName As String, newName As String) As String
    
    '▲notice
    'ToImportMDBFullPath  :原リンク側（リンクしたいテーブル）

    Dim retValue As String

    'accessリンク。リンクモード,データベースタイプはACCESS, リンクするMDB, テーブルをリンクする, リンクしたいテーブル名, リンクした後の新しいテーブル名
    DoCmd.TransferDatabase _
        acImport, _
        "Microsoft Access", _
        ToImportMDBFullPath, _
        acTable, _
        targetTableName, _
        newName

    retValue = targetTableName

    ImportForAcTableAndNameEdit = retValue

    Application.RefreshDatabaseWindow
    
End Function

Function ImportForCSVFiles_UsingIni(csvFolder As String) As String()

    Dim csvFilePath As String
    Dim buf As String
    Dim ImportageTableNewName As String                   'リンク完了後のリンク名
    Dim retValue As String
    
    buf = Dir(csvFolder & "\*.csv")                     'フォルダ内からエクセルファイルを探す
    
    Do While buf <> ""
        
        csvFilePath = csvFolder & "\" & buf             'リンクするエクセルファイルへのアドレス取得
        ImportageTableNewName = ImportForCSVFile_UsingIni(csvFolder, buf)
        retValue = retValue & ImportageTableNewName & ","
        
        buf = Dir()                                     '次のエクセルファイルを探す
    Loop
    
    ImportForCSVFiles_UsingIni = Split(Left(retValue, Len(retValue) - 1), ",")
    
    Application.RefreshDatabaseWindow
    
End Function

Function ImportForCSVFile_UsingIni(csvFolder As String, csvFileName As String) As String

    'csvFolder      : 「C\...\○○○○」    ＣＳＶファイルのあるフォルダ（最後尾に「\」はいらない）
    'csvFileName    : 「○○○○.csv」      パスはいらない。ファイル名と拡張子があればよい
    
    '▲notice
    'スキーマファイル（schema.ini）はあってもなくても同じ挙動をするので注意すること。
    'つまり、スキーマファイルがない場合はコンピュータの勝手な判断で型定義される。
    
    Dim td As dao.TableDef
    Const context As String = "text;DATABASE={CSV_FOLDER};TABLE={CSV_FILENAME};"
    Dim editContext As String
    Dim ImportTableName As String                         'リンク完了したあとのリンクテーブル名
    ImportTableName = Replace(csvFileName, ".csv", "")    'リンクテーブル名はcsvファイル名から拡張子を取った名前
    
    Set td = CurrentDb.CreateTableDef(ImportTableName)    'リンクテーブルの名前
    
    editContext = Replace(context, "{CSV_FOLDER}", csvFolder)
    editContext = Replace(editContext, "{CSV_FILENAME}", csvFileName)
    td.Connect = editContext
    
    td.SourceTableName = csvFileName
    
    CurrentDb.TableDefs.Append td
    
    ImportForCSVFile_UsingIni = ImportTableName
    
    Application.RefreshDatabaseWindow
    
End Function

Sub IsExistsTableName(MDBFullPath As String, targetTableName As String)

'まだ編集中（未着手）
'    Dim cat As New ADOX.Catalog                                     'カタログ
'    Dim tbl As ADOX.Table                                           'テーブル
'    Dim col As ADOX.Column                                          '列
'
'    cat.ActiveConnection = PROVIDER & MDBFullPath
'
'    If targetTableName = vbNullString Then
'        Exit Sub                                                    '実行要件を満たさなかった...
'    End If
'
'    For Each tbl In cat.Tables                                      'すべてのテーブルオブジェクトをめぐって
'        If tbl.Name = targetTableName Then                          'ユーザの指定したテーブル名の、
'            If tbl.Type = "TABLE" Then                              'テーブルがあったときに
'
'                For Each col In tbl.Columns                         'すべての項目をめぐって
'                    col.Name = Replace(col.Name, before, after)     '置換していく
'                Next
'
'            End If
'        End If
'    Next
'
'    Set cat = Nothing
'
'    Application.RefreshDatabaseWindow

End Sub



