VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MDBManipulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const PROVIDER As String = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = "

Function GetOwnFolderPath() As String
    
    '「C:\hoge\hage.mdb」上で、このメソッドを使用した場合、
    '「C:\hoge」が返る
    
    GetOwnFolderPath = Application.CurrentProject.path
    
End Function

Function GetOwnFileName() As String
    
    '「C:\hoge\hage.mdb」上で、このメソッドを使用した場合、
    '「hage.mdb」が返る
    
    GetOwnFileName = Application.CurrentProject.Name
    
End Function

Function GetOwnFullPath() As String
    
    '「C:\hoge\hage.mdb」上で、このメソッドを使用した場合、
    '「C:\hoge\hage.mdb」が返る
    
    GetOwnFullPath = Application.CurrentProject.FullName
    
End Function

Function GetQueryNamesByTheKeywordBeginning(keyword As String) As String()
    
    Dim cat As New ADOX.Catalog                         'カタログ
    Dim obj As Variant                                  'テーブル
    Dim logMessage As String                            'ログに出力するメッセージ
    
    cat.ActiveConnection = CurrentProject.Connection    '処理対象は、我MDB
    For Each obj In cat.Procedures                      'すべての 【アクションクエリ】 オブジェクトをめぐって
        If obj.Name Like (keyword & "*") Then           'なおかつ先頭文字が keyword で始まるとき
            logMessage = logMessage & obj.Name & ","
        End If
    Next
    For Each obj In cat.Views                           'すべての 【選択クエリ】 オブジェクトをめぐって
        If obj.Name Like (keyword & "*") Then           'なおかつ先頭文字が keyword で始まるとき
            logMessage = logMessage & obj.Name & ","
        End If
    Next
    Set cat = Nothing
    
    If logMessage = vbNullString Then logMessage = "キーワードで始まるクエリはありませんでした。"

    GetQueryNamesByTheKeywordBeginning = Split(Left(logMessage, Len(logMessage) - 1), ",")
    Application.RefreshDatabaseWindow
    
End Function

Function ChangeIndexAtThisArray(oldArray() As String, changePlan As String) As String()
    
    Dim idxList() As String
    Dim retArray() As String
    
    Dim i As Integer
    
    idxList = Split(changePlan, ",")
    retArray = oldArray
    
    For i = 0 To UBound(idxList)
        retArray(idxList(i)) = oldArray(i)
    Next
    
    ChangeIndexAtThisArray = retArray
    
End Function

Function GetTableSchema(tableName As String, colLikeKeyword As String) As String()
    
    'colLikeKeyword     :   「"*_" & ym.GetAddYM(-1)」など、収集したい項目名を入力（ﾜｲﾙﾄﾞｶｰﾄﾞok）
    
    '▲notice
    '配列に返る項目名は、項目名の昇順に勝手に並び替えられる。項目が並んでいる順番ではない。
    
    Dim cat                 As New ADOX.Catalog                     'カタログ
    Dim tbl                 As ADOX.TABLE                           'テーブル
    Dim col                 As ADOX.Column                          '列
        
    Dim colCount            As Integer
    Dim colCollection()     As String
    
    Dim i                   As Integer
    
    cat.ActiveConnection = CurrentProject.Connection
    For Each tbl In cat.Tables
        If tbl.Name = tableName Then
            If tbl.Type = "LINK" Or tbl.Type = "TABLE" Then
                
                '必要な最下層項目の数をget
                colCount = 0
                For Each col In tbl.Columns
                    If col.Name Like colLikeKeyword Then
                        colCount = colCount + 1
                    End If
                Next
                ReDim colCollection(colCount - 1)
                
                '最下層項目名をget
                colCount = 0
                For Each col In tbl.Columns
                    For i = LBound(colCollection) To UBound(colCollection)
                        If col.Name Like colLikeKeyword Then
                            colCollection(colCount) = col.Name      '収納
                            colCount = colCount + 1                 '収納idx + 1
                            Exit For
                        End If
                    Next
                Next
                
            End If
        End If
    Next
    
    Set cat = Nothing
    Application.RefreshDatabaseWindow
    
    GetTableSchema = colCollection
    
End Function

Function CreateMDBFile(newMDBFullPath As String) As String
    
    '参照設定　Microsoft ADO Ext 2.x DLL AND Securityを有効にします
    
    Dim cat As New ADOX.Catalog
    Dim retValue As String
    
    cat.CREATE (PROVIDER & newMDBFullPath)
    
    retValue = Dir(newMDBFullPath) & "を生成しました。"
    
    Set cat = Nothing
    
    CreateMDBFile = retValue
    
End Function

Function DeleteMDBFile(MDBFullPath As String) As String
        
    Dim buf As String
    Dim objFSO As Object
    Dim retValue As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    buf = Dir(MDBFullPath)
    
    If buf <> "" Then
        objFSO.DeleteFile MDBFullPath
        retValue = buf & "を削除しました。"
    Else
        retValue = "削除対象ファイルはありませんでした。"
    End If
    
    Set objFSO = Nothing
    
    DeleteMDBFile = retValue

End Function

Sub CreateTable(newTableName As String)
    
    DeleteTable (newTableName)                      'これから作るテーブル名があったら削除しとく
    
    Dim cat As New ADOX.Catalog
    Dim tbl As New ADOX.TABLE
    
    cat.ActiveConnection = CurrentProject.Connection
    tbl.Name = newTableName
    
    cat.Tables.Append tbl
    
    Set cat = Nothing
    
    Application.RefreshDatabaseWindow
    
End Sub

Sub AddColumns(tableName As String, columnSchema() As String, typeModel As Integer, Optional columnSize As Integer)
    
    Dim cat As New ADOX.Catalog
    Dim col As New ADOX.Column
    Dim columnName As Variant
    
    cat.ActiveConnection = CurrentProject.Connection
    
    For Each columnName In columnSchema
    
        col.Name = columnName                       '列名
        col.Type = typeModel                        '型
        
        col.DefinedSize = columnSize                'フィールドサイズ（columnSizeの初期値は0）
        
        col.Attributes = adColNullable              'Null値を許容します
        
        cat.Tables(tableName).Columns.Append col    'テーブルに列追加
        
        Set col = New ADOX.Column                   'objectReset

    Next columnName
    
    Set cat = Nothing
    
    Application.RefreshDatabaseWindow
    
End Sub

Sub AddColumns2(tableName As String, columnSchema() As String)

    '各列に個別にフィールドサイズをつけることができる
    
    '▲notice
    '必ずテキスト型になる
    
    'columnSchemaの指定の仕方　「hoge列,10」　と指定すると
    '「hoge列」という名前の列が作成され
    '「テキスト型」で
    '「フィールドサイズが１０」になる
    
    Dim cat As New ADOX.Catalog
    Dim col As New ADOX.Column
    Dim i As Integer
    Dim x() As String
    
    cat.ActiveConnection = CurrentProject.Connection
    
    For i = 0 To UBound(columnSchema)
    
        x = Split(columnSchema(i), ",")
        
        col.Name = x(0)
        col.Type = adVarWChar                       '現状、必ずテキスト型です（フレキシブル対応したいけど）
        col.DefinedSize = x(1)
        col.Attributes = adColNullable              'Null値を許容します
        cat.Tables(tableName).Columns.Append col    'テーブルに列追加
        
        Set col = New ADOX.Column                   'objectReset
        
    Next
        
    Set cat = Nothing
    
    Application.RefreshDatabaseWindow
    
End Sub

Sub RenameColumns(targetTableName As String, before As String, after As String)
    
    Dim cat As New ADOX.Catalog                                     'カタログ
    Dim tbl As ADOX.TABLE                                           'テーブル
    Dim col As ADOX.Column                                          '列
    
    cat.ActiveConnection = CurrentProject.Connection                'カレントデータベースに接続
    
    If targetTableName = vbNullString Then
        Exit Sub                                                    '実行要件を満たさなかった...
    End If
    
    For Each tbl In cat.Tables                                      'すべてのテーブルオブジェクトをめぐって
        If tbl.Name = targetTableName Then                          'ユーザの指定したテーブル名の、
            If tbl.Type = "TABLE" Then                              'テーブルがあったときに
            
                For Each col In tbl.Columns                         'すべての項目をめぐって
                    col.Name = Replace(col.Name, before, after)     '置換していく
                Next
                
            End If
        End If
    Next
    
    Set cat = Nothing
    
    Application.RefreshDatabaseWindow

End Sub

Sub RenameTable(targetTableName As String, newTableName As String)

    '参照設定　Microsoft ADO Ext 2.x DLL AND Securityを有効にします
    Dim cat As New ADOX.Catalog     'カタログ
        
    cat.ActiveConnection = CurrentProject.Connection    'カレントデータベースに接続
        
    cat.Tables(targetTableName).Name = newTableName
        
    Set cat = Nothing
    
    Application.RefreshDatabaseWindow

End Sub

Sub DeleteColumn(targetTableName As String, delColumnName As String)
    
    Dim cmd As New ADODB.Command
    Dim sql As String
    
    sql = "Alter table [" & targetTableName & "] drop " & delColumnName & ";"
    
    On Error Resume Next
    cmd.ActiveConnection = CurrentProject.Connection
    cmd.CommandText = sql
    cmd.Execute
    On Error GoTo 0
    
    Application.RefreshDatabaseWindow
    
End Sub

Sub DeleteTable(delTableName As String)
    
    Dim cat As New ADOX.Catalog
    
    On Error Resume Next
    cat.ActiveConnection = CurrentProject.Connection
    cat.Tables.DELETE delTableName
    On Error GoTo 0
    
    Set cat = Nothing
    
    Application.RefreshDatabaseWindow
    
End Sub

Sub DeleteQueryObject(delQueryName As String)
    
    Dim cat As New ADOX.Catalog

    cat.ActiveConnection = CurrentProject.Connection
    
    On Error Resume Next
    cat.Views.DELETE delQueryName
    cat.Procedures.DELETE delQueryName
    On Error GoTo 0

    Set cat = Nothing

    Application.RefreshDatabaseWindow
    
End Sub

Sub ExecuteSQL(sql As String)
    
    '▲notice
    'SQLテキストかアクションクエリー名を代入します。
    '選択クエリは実行できない。何もせずに終わる
    
    Dim cmd As New ADODB.Command
    
    cmd.ActiveConnection = CurrentProject.Connection
    cmd.CommandText = sql
    
    cmd.Execute
    Application.RefreshDatabaseWindow
    
End Sub

Sub ExecuteQuery(queryName As String)

    '▲notice
    'アクションクエリ名でも選択クエリでも実行できる。
    '出来合いのクエリを使うだけならExecuteSQLよりこっちの方が優秀
    
    DoCmd.SetWarnings False                     'システム メッセージの表示をオフ
    DoCmd.OpenQuery queryName, acViewNormal
    DoCmd.SetWarnings True                      'システム メッセージの表示をオン
    
    Application.RefreshDatabaseWindow
    
End Sub

Function Export2Excel(outputTableOrQuery As String, outputFolder As String) As String

    '★上書き用のファイル削除ロジック開始
    Dim buf As String
    Dim objFSO As Object                                            'FileSysObj
    Dim retValue As String                                          'ログに出力するメッセージ
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    buf = Dir(outputFolder & "\" & outputTableOrQuery & ".xls")
    If buf <> "" Then
        objFSO.DeleteFile outputFolder & "\" & buf                  'エクスポートする同名ファイルを削除
        retValue = buf & "を削除しました。"
    Else
        retValue = "削除対象ファイルはありませんでした。"
    End If
    Set objFSO = Nothing
    '★上書き用のファイル削除ロジック終了

    DoCmd.SetWarnings False                     'ワーニングオフ
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, outputTableOrQuery, outputFolder & "\" & outputTableOrQuery & ".xls", True
    DoCmd.SetWarnings True                      'ワーニングオン

    Application.RefreshDatabaseWindow           'Refresh
    
    Export2Excel = retValue
    
End Function

Function Export2ExcelAndNameEdit(outputTableOrQuery As String, newFileFullPath As String) As String

    '★上書き用のファイル削除ロジック開始
    Dim buf As String
    Dim objFSO As Object                                            'FileSysObj
    Dim retValue As String                                          'ログに出力するメッセージ
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    buf = Dir(newFileFullPath)                                      'newFileFullPathのファイルを探す
    If buf <> "" Then
        objFSO.DeleteFile newFileFullPath                           'エクスポートする同名ファイルを削除
        retValue = buf & "を削除しました。"
    Else
        retValue = "削除対象ファイルはありませんでした。"
    End If
    Set objFSO = Nothing
    '★上書き用のファイル削除ロジック終了

    DoCmd.SetWarnings False                     'ワーニングオフ
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, outputTableOrQuery, newFileFullPath, True
    DoCmd.SetWarnings True                      'ワーニングオン

    Application.RefreshDatabaseWindow           'Refresh
    Export2ExcelAndNameEdit = retValue
    
End Function

Function CountOfTableRecs(targetTable As String) As Long

    Dim retValue As Long
    
    retValue = DCount("*", targetTable)
    
    CountOfTableRecs = retValue
    
End Function

Sub SetTableDescription(targetTable As String, setumei As String)
    
    '▲notice
    'テーブルに「説明:Description」を書き付けます
    
    Dim setProps As dao.Property
    
    Set setProps = CurrentDb.TableDefs(targetTable).CreateProperty      'Nullプロパティ生成
    
    setProps.Name = "Description"                                       '【必須】説明 プロパティ を
    setProps.Type = dbText                                              '【必須】dbText テキスト型 (Text) で
    setProps.value = setumei                                            '書き付けます
    
    'プロパチーが無いのに消そうとするとエラーがでる
    'しかし、プロパチーがあるのに追加するとエラーがでる
    'したがって、「消す」をエラー無視にした
    On Error Resume Next
    CurrentDb.TableDefs(targetTable).Properties.DELETE ("Description")
    On Error GoTo 0
    
    CurrentDb.TableDefs(targetTable).Properties.Append setProps     'プロパティ結合：よっこらしょっと
    Application.RefreshDatabaseWindow                               'Refresh
    
End Sub

Sub SetPrimaryKey(targetTable As String, keyString As String)
    
    '▲notice
    'SQLの構文上、setPrimaryKeyという文字は任意の "作成するインデックスの名前" である。
    
    Dim cmd As New ADODB.Command
    Dim sql As String
    
    sql = "create index setPrimaryKey on [" & targetTable & "](" & keyString & ") with primary;"
    
    cmd.ActiveConnection = CurrentProject.Connection
    cmd.CommandText = sql
    
    cmd.Execute
    
End Sub

Sub DeletePrimaryKey(targetTable As String)

    '▲notice
    'したがって、あくまで、SetPrimaryKeyメソッドによって作成された、プライマリキーのみ、
    'Killすることになっている。（これを名称指定できるようにしてもまわりくどいだけ）
    'あくまで、setPrimaryKeyというインデックスを、削除します。
    'つまり、SetPrimaryKeyメソッドを介さないプライマリキー（たとえば、手動でセットしたとか）は、
    '削除できない

    Dim cmd As New ADODB.Command
    Dim sql As String
    
    sql = "drop index setPrimaryKey on " & targetTable & ";"
    
    cmd.ActiveConnection = CurrentProject.Connection
    cmd.CommandText = sql
    
    cmd.Execute

End Sub

Sub CreateQueryObject(queryContext As String, newQueryName As String)

    Dim cat As New ADOX.Catalog
    Dim cmd As New ADODB.Command
    
    cat.ActiveConnection = CurrentProject.Connection

    Set cmd = New ADODB.Command

    DeleteQueryObject newQueryName                  'クエリを削除します

    cmd.CommandText = queryContext                          'SQLを引き渡して・・・
    cat.Views.Append newQueryName, cmd                      'クエリを作成します
    
End Sub

Function ExportForAcTable(ToExportMDBFullPath As String, targetTableName As String) As String
    
    '▲notice
    'ToExportMDBFullPath    :飛ばしたい先のMDBパス
    'targetTableName        :飛ばしたいテーブル名（自MDB）

    Dim retValue As String

    'accessリンク。リンクモード,データベースタイプはACCESS, リンクするMDB, テーブルをリンクする, リンクしたいテーブル名, リンクした後の新しいテーブル名
    DoCmd.TransferDatabase _
        acExport, _
        "Microsoft Access", _
        ToExportMDBFullPath, _
        acTable, _
        targetTableName, _
        targetTableName

    retValue = targetTableName
    ExportForAcTable = retValue
    Application.RefreshDatabaseWindow
    
End Function

Function ExportForAcTableAndNameEdit(ToExportMDBFullPath As String, targetTableName As String, newName As String) As String
    
    '▲notice
    'ToExportMDBFullPath    :飛ばしたい先のMDBパス
    'targetTableName        :飛ばしたいテーブル名（自MDB）
    'newName                :テーブル名変更後の名前

    Dim retValue As String

    'accessリンク。リンクモード,データベースタイプはACCESS, リンクするMDB, テーブルをリンクする, リンクしたいテーブル名, リンクした後の新しいテーブル名
    DoCmd.TransferDatabase _
        acExport, _
        "Microsoft Access", _
        ToExportMDBFullPath, _
        acTable, _
        targetTableName, _
        newName

    retValue = targetTableName
    ExportForAcTableAndNameEdit = retValue
    Application.RefreshDatabaseWindow
    
End Function
