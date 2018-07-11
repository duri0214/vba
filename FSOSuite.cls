VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FSOSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'ファイルをバイナリで開いてエラーが返ったら「ファイルが開かれている(True)」と返す
Function IsFileUsing(path As String) As Boolean

    Dim fileSequence as integer
    Dim retValue as Boolean

    if dir(path) <> "" then
        fileSequence = FreeFile
        On Error Resume Next
        Open path For Binary Lock Read Write As #fileSequence
        close #fileSequence
    end if

    if err.number >0 then
        retValue = true
        On Error goto 0
    end if

    IsFileUsing = retValue

End Function

function IsFileExistsAndNotUsing(path as string) as Boolean

    Dim retValue as Boolean

    if IsFileExists(path) and not IsFileUsing(path)
        retValue = true
    end if

    IsFileExistsAndNotUsing = retValue

end function

function GetFilelist(folderPath as string) as string()

    dim objFSO as object
    dim f as object
    dim retValue() as string

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    if IsfolderExists(folderPath) then
        For each f in objFSO.GetFolder(folderPath).files
            if Not IsValidArray(retValue) then  
                redim retValue(0)
            Else
                redim preserve retValue(Ubound(retValue)+1)
            End if
            retValue(Ubound(retValue))= folderPath & "\" & f.name
        next f
    End if

    GetFilelist = retValue
    
end function

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

