VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Genocider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const PROVIDER As String = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = "

Function LinkGenocider() As String()

    Dim retValue() As String
    
    retValue = GenociderCore("LINK", "")
    LinkGenocider = retValue
    
End Function

Function LinkGenociderByTheKeywordBeginning(keyword As String) As String()

    Dim retValue() As String
    
    retValue = GenociderCore("LINK", keyword)
    LinkGenociderByTheKeywordBeginning = retValue
    
End Function

Function TableGenocider() As String()

    Dim retValue() As String
    
    retValue = GenociderCore("TABLE", "")
    TableGenocider = retValue

End Function

Function TableGenociderByTheKeywordBeginning(keyword As String) As String()

    Dim retValue() As String
    
    retValue = GenociderCore("TABLE", keyword)
    TableGenociderByTheKeywordBeginning = retValue

End Function

Function QueryGenociderByTheKeywordBeginning(keyword As String) As String()

    Dim retValue() As String
    
    retValue = GenociderCore("QUERY", keyword)
    QueryGenociderByTheKeywordBeginning = retValue

End Function

Function FileGenociderForCsvFile(folderPath As String) As String()

    'LogWriterを組み合わせるとベリグッド
    'ルーチンに入る前に「フォルダ内のファイルが消されてしまうがよろしいか」と聞いて、いいえなら処理中断するように。
    
    Dim buf As String
    
    Dim objFSO As Object                                                    'FileSysObj
    Dim logMessage As String                'ログに出力するメッセージ
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    buf = Dir(folderPath & "\*.csv")
    
    If buf <> "" Then
    
        Do While buf <> ""
            
            objFSO.DeleteFile folderPath & "\" & buf
            logMessage = logMessage & buf & "を削除しました。" & ","

            buf = Dir()
        Loop
    Else
        logMessage = "削除対象ファイルはありませんでした。"
    End If
    
    Set objFSO = Nothing
    
    FileGenociderForCsvFile = Split(Left(logMessage, Len(logMessage) - 1), ",")
    
    Application.RefreshDatabaseWindow

End Function

Function FileGenociderForExcelFile(folderPath As String) As String()

    'LogWriterを組み合わせるとベリグッド
    'ルーチンに入る前に「フォルダ内のファイルが消されてしまうがよろしいか」と聞いて、いいえなら処理中断するように。
    
    Dim buf As String
    
    Dim objFSO As Object                                                    'FileSysObj
    Dim logMessage As String                'ログに出力するメッセージ
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    buf = Dir(folderPath & "\*.xls")
    
    If buf <> "" Then
    
        Do While buf <> ""
            
            objFSO.DeleteFile folderPath & "\" & buf
            logMessage = logMessage & buf & "を削除しました。" & ","

            buf = Dir()
        Loop
    Else
        logMessage = "削除対象ファイルはありませんでした。"
    End If
    
    Set objFSO = Nothing
    FileGenociderForExcelFile = Split(Left(logMessage, Len(logMessage) - 1), ",")
    
    Application.RefreshDatabaseWindow
    
End Function

Private Function KillerForTable(cat As ADOX.Catalog, killingName As String) As String
    
    Dim retMessage As String

    cat.Tables.DELETE (killingName)
    retMessage = killingName & "を削除しました。"

    KillerForTable = retMessage

End Function

Private Function KillerForQuery(cat As ADOX.Catalog, killingName As String) As String
    
    Dim retMessage As String

    On Error Resume Next
    cat.Views.DELETE killingName
    cat.Procedures.DELETE killingName
    On Error GoTo 0
    
    retMessage = killingName & "を削除しました。"

    KillerForQuery = retMessage

End Function

Private Function GenociderCore(mode As String, keyword As String, Optional MDBFullPath As String) As String()

    Dim cat As New ADOX.Catalog             'カタログ
    Dim obj As Variant                      'テーブル
    Dim findOut As Boolean
    Dim logMessage As String                'ログに出力するメッセージ
    
    Dim firstProc As Boolean
    
    If MDBFullPath <> vbNullString Then
        cat.ActiveConnection = PROVIDER & MDBFullPath       '処理対象は、オプショナル引数に入ってきた外部MDBパス
    Else
        cat.ActiveConnection = CurrentProject.Connection    '処理対象は、我MDB
    End If
    findOut = True
    firstProc = True
    
    '■notice
    'なぜかデリートのコマンド（cat.Tables.Delete）は、１つ対象を消すと１つ飛びに次の対象のデリートが走るので、
    '消し終わったらもう一回ローラーをかけて、「消す対象がまったくない」ところまで、無限ループで
    '消し続ける必要がある。
    '例：５個の消す対象があったら、１、３、５のオブジェクトが消されて、２、４が残る
    'ログを見れば一目瞭然だ
    
    '種類           意味
    'ACCESS TABLE   Accessシステムテーブル
    'LINK           ODBCデータソース以外からリンクされたテーブル
    'SYSTEM TABLE   Microsoft Jetシステムテーブル
    'TABLE          Accessデータベースの標準テーブル
    'VIEW           Accessデータベースのパラメータを含まない選択クエリ
    'QUERY          作った。ActionQueryをジェノサイドしたいときに分岐させるキーワード
    
    Do Until findOut = False
        findOut = False
        
        If mode = "QUERY" Then                          '▲notice MAGIC STRING!!
        
            For Each obj In cat.Procedures              'すべての 【アクションクエリ】 オブジェクトをめぐって
                If obj.Name Like (keyword & "*") Then   'なおかつ先頭文字が keyword で始まるとき
                    logMessage = logMessage & KillerForQuery(cat, obj.Name) & ","
                    findOut = True
                    firstProc = False
                End If
            Next
            For Each obj In cat.Views                   'すべての 【選択クエリ】 オブジェクトをめぐって
                If obj.Name Like (keyword & "*") Then   'なおかつ先頭文字が keyword で始まるとき
                    logMessage = logMessage & KillerForQuery(cat, obj.Name) & ","
                    findOut = True
                    firstProc = False
                End If
            Next
        Else
            For Each obj In cat.Tables                  'すべての 【テーブル】 オブジェクトをめぐって
                If obj.Type = mode Then                 '（引数）種だったときに
                If obj.Name Like (keyword & "*") Then   'なおかつ先頭文字が keyword で始まるとき
                    logMessage = logMessage & KillerForTable(cat, obj.Name) & ","
                    findOut = True
                    firstProc = False
                End If
                End If
            Next
        End If
    Loop
    Set cat = Nothing
    
    If firstProc = True Then logMessage = "削除対象ファイルはありませんでした。"

    GenociderCore = Split(Left(logMessage, Len(logMessage) - 1), ",")
    Application.RefreshDatabaseWindow
    
End Function
