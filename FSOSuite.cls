Option Explicit

'フォルダへのショートカットをつくる
Function CreateFolderShortCut(targetFolder As String, destinationFolder As String)
    
    Dim objWshShell
    Set objWshShell = CreateObject("WScript.Shell")
    
    Dim objShortcut
    
    Set objShortcut = objWshShell.CreateShortCut(destinationFolder & "\" & Dir(targetFolder, vbDirectory) & " - ショートカット.lnk")
    objShortcut.TargetPath = targetFolder
    objShortcut.Save
    
End Function
        
'そのフォルダのなかで一番新しくできたファイルパスを返す
Function GetNewestFileInTheFolder(folder As String) As String
    
    Dim buf As String
    Dim d1 As Date
    Dim d2 As Date
    Dim firstFlg As Boolean
    Dim retValue As String
    
    firstFlg = True
    buf = Dir(folder & "\*")
    Do While buf <> ""
    
        If firstFlg Then
        
            'Loop初回
            d1 = FileDateTime(folder & "\" & buf)
            retValue = folder & "\" & buf
            firstFlg = False
            
        Else
        
            'dateで比較
            d2 = FileDateTime(folder & "\" & buf)
            If d2 > d1 Then
                d1 = FileDateTime(folder & "\" & buf)
                retValue = folder & "\" & buf
            End If
            
        End If
        
        buf = Dir()
        
    Loop
    
    GetNewestFileInTheFolder = retValue
    
End Function

'pathが誰かによって開かれている場合、ファイル名の末尾に年月日時分秒をつけて返す
Function GetNewNameIfTheFileIsUsing(path As String) As String

    Dim filename As String
    Dim idx As Integer
    Dim ext As String
    Dim retValue As String
    
    retValue = path
    filename = GetFileName(path)
    
    '拡張子があるときに動作条件を満たす
    idx = InStr(1, filename, ".")
    If idx > 0 Then
        
        '拡張子の決定（例：.xlsx）
        ext = Mid(filename, idx)
        
        If IsFileUsing(path) Then
            retValue = GetParentFolderName(path) & "\" & Replace(filename, ext, Format(Now(), "_yyyymmddHHMMSS")) & ext
        End If
        
    End If
    
    GetNewNameIfTheFileIsUsing = retValue

End Function

'path2つを結合します（フォルダ名の末尾の”\”を自動的に付与）
Function Combine(path1 As String, path2 As String) As String
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Combine = objFSO.BuildPath(path1, path2)
End Function

'ファイルをバイナリで開いてエラーが返ったら「ファイルが開かれている(True)」と返す
Function IsFileUsing(path As String) As Boolean

    Dim fileSequence As Integer
    Dim retValue As Boolean

    If Dir(path) <> "" Then
        fileSequence = FreeFile
        On Error Resume Next
        Open path For Binary Lock Read Write As #fileSequence
        Close #fileSequence
    End If

    If Err.Number > 0 Then
        retValue = True
        On Error GoTo 0
    End If

    IsFileUsing = retValue

End Function

Function IsFileExistsAndNotUsing(path As String) As Boolean

    Dim retValue As Boolean

    If IsFileExists(path) And Not IsFileUsing(path) Then
        retValue = True
    End If

    IsFileExistsAndNotUsing = retValue

End Function

Function GetFilelist(folderPath As String) As String()
    
    Const HIDDEN As Integer = 2
    Dim objFSO As Object
    Dim f As Object
    Dim retValue() As String

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If IsFolderExists(folderPath) Then
        For Each f In objFSO.GetFolder(folderPath).Files
            If (f.Attributes And HIDDEN) = 0 Then
                If Not IsValidArray(retValue) Then
                    ReDim retValue(0)
                Else
                    ReDim Preserve retValue(UBound(retValue) + 1)
                End If
                retValue(UBound(retValue)) = folderPath & "\" & f.Name
            End If
        Next f
    End If

    GetFilelist = retValue
    
End Function

Function IsValidArray(targetArray() As String) As Boolean
    
    'targetArrayがnullかどうかを調べる関数。
    '本来この関数はいらなくて if targetArray = null then... とやりたいところだが
    'vbaでは、配列にnullを仕込むことができないため、わざわざエラー無視の状態にして、
    '強制的に配列の第一次元目の要素数を求める
    '配列にアクセスできてなければ（エラーがhitしていれば）falseを返す（配列がnullってるってこと）
    '配列にアクセスできれば、trueを返す
    
    Dim checkValue As Variant
    Dim retValue As Boolean
    
    On Error Resume Next                                            'エラーを無効化
    
    checkValue = UBound(targetArray, 1)                             '強制的に配列の第一次元目の要素数を求める
    
    If Err.Number > 0 Then                                          'Errorが発生したら、未定義配列です。
        retValue = False
    Else                                                            'そうでないなら、1以上の要素が宣言済みです。
        retValue = True
    End If
    Err.Clear                                                       'エラーオブジェクト初期化
    
    On Error GoTo 0                                                 'エラーを有効化
    
    IsValidArray = retValue                                         '検査結果を返す

End Function

Function DeleteFile(fileFullPath As String) As String
    
    Dim objFSO As Object
    Dim retValue As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    retValue = GetFileName(fileFullPath)    '消したファイル名get
    objFSO.DeleteFile fileFullPath          'delete
    
    DeleteFile = retValue                   '消したファイル名を返す
    
End Function

Function GetFilesCount(folderPath As String) As Integer
    
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    GetFilesCount = objFSO.GetFolder(folderPath).Files.Count
    
End Function

Function GetFoldersCount(folderPath As String) As Integer
    
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    GetFoldersCount = objFSO.GetFolder(folderPath).SubFolders.Count
    
End Function

Function GetFilesCountForCSV(folderPath As String) As Integer
    
    Dim objFSO As Object
    Dim objFiles As Object
    Dim file As Variant

    Dim filesCount As Integer                                   '見つけたファイル数

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFiles = objFSO.GetFolder(folderPath).Files           '指定されたフォルダのファイルリストを取得

    For Each file In objFiles
        If Right(file.Name, 4) = ".csv" Then
            filesCount = filesCount + 1
        End If
    Next

    GetFilesCountForCSV = filesCount
    
End Function

Function GetFilesCountForMDB(folderPath As String) As Integer
    
    Dim objFSO As Object
    Dim objFiles As Object
    Dim file As Variant

    Dim filesCount As Integer                                   '見つけたファイル数

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFiles = objFSO.GetFolder(folderPath).Files           '指定されたフォルダのファイルリストを取得

    For Each file In objFiles
        If Right(file.Name, 4) = ".mdb" Then
            filesCount = filesCount + 1
        End If
    Next

    GetFilesCountForMDB = filesCount
    
End Function

Function IsFolderExists(folderPath As String) As Boolean
    
    Dim objFSO As Object
    Dim retValue As Boolean
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    retValue = objFSO.FolderExists(folderPath)
    
    IsFolderExists = retValue
    
End Function

Sub IsFolderExistsAndMakeFolder(folderPath As String)
    
    Dim objFSO As Object
    Dim retValue As Boolean
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    retValue = objFSO.FolderExists(folderPath)
    
    If retValue = False Then MkDir folderPath
    
End Sub

Function IsFileExists(filePath As String) As Boolean
    
    '▲notice
    'ファイルの存在有無を2値で返します
    'DIR関数を1ループの中で複数回使いたいときに使用します。
    'たとえば、フォルダのファイルを全て処理するルーチンで、最初にDIRを使用して
    'すべてのファイルを処理しようとする中で、さらに単ファイルの存在有無処理が必要に
    'なったときDIR関数を使うとエラーになります。そのときにこの関数で代替します。
    
    Dim objFSO As Object
    Dim retValue As Boolean
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    retValue = objFSO.FileExists(filePath)
    
    IsFileExists = retValue

End Function

Function IsValidFileForMdb(mdbPath As String) As Boolean
    
    'ファイル検査を行う
    'ファイルに異常がなければ...Trueを返す
    'ファイルが異常があるなら...Falseを返す
    
    IsValidFileForMdb = IsValidFileCore(mdbPath, ".mdb")

End Function

Function IsValidFileForCSV(csvPath As String) As Boolean
    
    'ファイル検査を行う
    'ファイルに異常がなければ...Trueを返す
    'ファイルが異常があるなら...Falseを返す
        
    IsValidFileForCSV = IsValidFileCore(csvPath, ".csv")

End Function

Private Function IsValidFileCore(path As String, compareString As String) As Boolean
    
    'ファイル検査を行う
    'ファイルに異常がなければ...Trueを返す
    'ファイルが異常があるなら...Falseを返す
    
    Dim retValue As Boolean
    retValue = True

    If IsFileExists(path) = False Then
        retValue = False                    'ファイルパスが不正だ（指定ファイルパスが存在しない）
    Else
        If Right(path, 4) <> compareString Then
            retValue = False                'ファイルパスは通っているが、拡張子が不正だ
        End If
    End If
    
    IsValidFileCore = retValue
    
End Function

Function GetFileName(filePath As String) As String
    
    '▲notice
    'フルパスから、ファイル名を取り出します
    '「C:\hoge.xls」 → 「hoge.xls」
    
    Dim objFSO As Object
    Dim retValue As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    retValue = objFSO.GetFileName(filePath)
    
    GetFileName = retValue
    
End Function

Function GetParentFolderPath(filePath As String) As String

    '▲notice
    'ファイルフルパスから、親フォルダパスを取り出します
    '「C:\hage\hoge.xls」 → 「C:\hage」
    
    Dim objFSO As Object
    Dim retValue As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    retValue = objFSO.GetParentFolderName(filePath)
    
    GetParentFolderPath = retValue

End Function

Function GetParentFolderName(filePath As String) As String
    
    '▲notice
    'ファイルパスから、親フォルダ名を取り出します
    '「C:\hage\hoge.xls」 → 「hage」
    '具体的には、ファイルパスを受け取って、上のGetParentFolderメソッドでフォルダパスを内部的に取得
    'そしてFSOで、親フォルダの名前だけを返す
    
    Dim objFSO As Object
    Dim retValue As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    retValue = objFSO.GetFolder(GetParentFolderPath(filePath)).Name
    
    GetParentFolderName = retValue
    
End Function

Function ReadHeaderLineByTextFile(textFileFullPath As String) As String
    
    'テキストファイルの1行目を読んでStringで返す
    
    '▲notice
    'このメソッドは、ファイルの存在可否検査は行わない。
    'このメソッドに入る前に、ファイルの存在可否は検査せよ。
    
    Dim objFSO As Object                                                    'FileSysObj
    Dim txtStream As Object                                                 'FileSysObj
    Dim retValue As String

    Dim ioMode As Integer                                                   '入出力モード
    Const MODE_READ_ONLY As Integer = 1                                     '読み取り専用
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")                 'FSO使用宣言

    ioMode = MODE_READ_ONLY                                                 '読み取り専用で...
    Set txtStream = objFSO.OpenTextFile(textFileFullPath, ioMode, True)     '開く。
    
    retValue = txtStream.ReadLine                                           '1行目を読みます
    
    txtStream.Close                                                         '閉じる
    Set txtStream = Nothing
    Set objFSO = Nothing
    
    ReadHeaderLineByTextFile = retValue
    
End Function

Function ReadAllByTextFile(textFileFullPath As String) As String
    
    '▲notice
    'このメソッドは、ファイルの存在可否検査は行わない。
    'このメソッドに入る前に、ファイルの存在可否は検査せよ。
    
    Dim objFSO As Object                                                    'FileSysObj
    Dim txtStream As Object                                                 'FileSysObj

    Dim ioMode As Integer                                                   '入出力モード
    Const MODE_READ_ONLY As Integer = 1                                     '読み取り専用
    
    Dim retValue As String
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")                 'FSO使用宣言
        
    ioMode = MODE_READ_ONLY                                                 '読み取り専用で...
    Set txtStream = objFSO.OpenTextFile(textFileFullPath, ioMode, True)     '開く。
    
    retValue = txtStream.ReadAll                                            '全読み
    
    txtStream.Close                                                         '閉じる
    Set txtStream = Nothing
    Set objFSO = Nothing
    
    ReadAllByTextFile = retValue

End Function

Function InvokeTextArrayByTextFile(textFileFullPath As String) As String()
    
    '▲notice
    'このメソッドは、ファイルの存在可否検査は行わない。
    'このメソッドに入る前に、ファイルの存在可否は検査せよ。

    Dim objFSO As Object                                                    'FileSysObj
    Dim txtStream As Object                                                 'FileSysObj
    Dim txt As String
    Dim retValue() As String

    Dim ioMode As Integer                                                   '入出力モード
    Const MODE_READ_ONLY As Integer = 1                                     '読み取り専用
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
        
    ioMode = MODE_READ_ONLY
    Set txtStream = objFSO.OpenTextFile(textFileFullPath, ioMode, True)
    
    Do Until txtStream.AtEndOfLine
        If txtStream.Line > 1 Then
            txt = txt & "{,}"
        End If
        txt = txt & txtStream.ReadLine
    Loop
    
    retValue = Split(txt, "{,}")
    
    txtStream.Close
    Set txtStream = Nothing
    Set objFSO = Nothing
    
    InvokeTextArrayByTextFile = retValue
    
End Function

