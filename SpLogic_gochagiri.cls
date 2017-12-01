VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpLogic_御茶義理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Const PROVIDER As String = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = "


Private objFSO As Object            'LogWritterクラス のコピペで使用
Private txtStream As Object         'LogWritterクラス のコピペで使用

Sub Gochagiri(targetTable As String, csvOutPutFolder As String, cutSegment As Long)

    '==================================================
    'グループ登録用CSV吐き出しロジック
    '==================================================
    'targetTable    :
    'csvOutPutName  :csvアウトプットネーム。（拡張子抜き）
    'cutSegment     :切り刻む件数。ZEUS登録用ならまず5000でほぼ固定かな。
    
    Dim fso As New FSOSuite
    
    Dim rs As New ADODB.Recordset       'レコードセット
'    Dim csvFile As New LogWritter       'csvオブジェクト
    Dim txt As String                   'CSVに書き込む文字
    Dim recCount As Long                '累積レコード数
    Dim csvNumber As Long               'CSVファイルナンバー
    Dim col As Variant                  '単項目
    Dim now_StreamIsOpen As Boolean     'LogWritterクラスは現在1個しか開けないので、フラグが必要
    
    Dim outputFileName As String
    
    If fso.IsFolderExists(csvOutPutFolder) = False Then GoTo illegalRoot                    '【検問】
    
    rs.ActiveConnection = CurrentProject.Connection
    rs.Open (targetTable)
    
    csvNumber = 1                                                                       'CSVファイルナンバー初期値１
    recCount = 1                                                                        '累積レコード数初期値１
    
    outputFileName = csvOutPutFolder & "\" & targetTable & Format(csvNumber, "00") & ".csv"
    OpenTextStream (outputFileName)                                                     '[内部ロジック使用↓]CSVファイルオープン
    now_StreamIsOpen = True                                                             'flg on
    
    Do Until rs.EOF
        
        txt = ""                                                                        '【書込】データ取得
        For Each col In rs.fields                                                       '【書込】データ取得
            txt = txt & col.value & ","                                                 '【書込】データ取得
        Next col                                                                        '【書込】データ取得
        txt = Left(txt, Len(txt) - 1)                                                   '【書込】ラストカンマ切除
        WriteLine (txt)                                                                 '[内部ロジック使用↓]【書込】データ出力
        
        If recCount >= cutSegment Then
            CloseTextStream                                                             '[内部ロジック使用↓]【切替】CSVファイルクローズ
            now_StreamIsOpen = False                                                    '【切替】flg off
            recCount = 1                                                                '【切替】累積レコード数初期値１
            csvNumber = csvNumber + 1                                                   '【切替】CSVファイルナンバー＋１
            outputFileName = csvOutPutFolder & "\" & targetTable & Format(csvNumber, "00") & ".csv"
            OpenTextStream (outputFileName)                                             '[内部ロジック使用↓]【切替】CSVファイルオープン
            now_StreamIsOpen = True                                                     '【切替】flg on
        Else
            recCount = recCount + 1                                                     '【継続】累積レコード数＋１
        End If
        
        rs.MoveNext
        
    Loop
    
    If now_StreamIsOpen Then CloseTextStream                                            '[内部ロジック使用↓]【終了】CSVファイルが開いていたら閉じる
    
    rs.Close
    
illegalRoot:                                                                            '検問に引っかかった場合に、ロジックは未実行で終わる。
    
End Sub



'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'/_/_ここから下は　LogWritterクラス のコピペ_/_/
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/


Private Sub OpenTextStream(Optional logFileFullPath As String)

    'WriteLineLogごとにテキスト保存されるのでCloseTextStreamが働く前に不正終了してもokです

    Dim logOutPath As String                                                '出力パス（基本自分ファイルの位置に作成）
    Dim ioMode As Integer                                                   '入出力モード
    
    Const MODE_NEW_CREATE As Integer = 2                                    '新規作成モード
    Const MODE_ADDITIONAL As Integer = 8                                    '追記モード
    
    If logFileFullPath = vbNullString Then
        logOutPath = Application.CurrentProject.path & "\log.txt"           '自パスフォルダ取得
    Else
        logOutPath = logFileFullPath
    End If
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If Dir(logOutPath) <> "" Then
        ioMode = MODE_ADDITIONAL                                            'logOutPathファイルがあれば追記モード
    Else
        ioMode = MODE_NEW_CREATE                                            'logOutPathファイルがないなら新規作成
    End If
    
    Set txtStream = objFSO.OpenTextFile(logOutPath, ioMode, True)           'logOutPathファイルが存在しない場合、新しいファイルを作成（true）
    
End Sub

Private Sub CloseTextStream()
    
    txtStream.Close                                                         'ストリームオブジェクト開放
    
    Set txtStream = Nothing                                                 '後始末
    Set objFSO = Nothing                                                    '後始末

End Sub

Private Sub WriteLineLog(txt As String)

    txtStream.WriteLine Now & " " & txt                                     'タイムスタンプとともにログ出力
    
End Sub

Private Sub WriteLine(txt As String)

    txtStream.WriteLine txt                                                 'ログ出力
    
End Sub
