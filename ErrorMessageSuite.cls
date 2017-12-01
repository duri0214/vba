VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorMessageSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function MessageReaction(errCode As Integer) As String
    
    Const ERR_NECESSARY_INPUT As Integer = 1
    Const ERR_FILE_NOT_EXISTS As Integer = 2
    Const ERR_FILE_SELECTION As Integer = 3
    Const ERR_FOLDER_NOT_EXISTS As Integer = 4
    
    Dim retValue As String
    
    If errCode = ERR_NECESSARY_INPUT Then
        retValue = "処理実行要件を満たしませんでした（必須項目の未入力）"
    End If
    
    If errCode = ERR_FILE_NOT_EXISTS Then
        retValue = "処理実行要件を満たしませんでした（ファイルが存在しない）"
    End If
    
    If errCode = ERR_FILE_SELECTION Then
        retValue = "処理実行要件を満たしませんでした（ファイル選択が不正です）"
    End If
    
    If errCode = ERR_FOLDER_NOT_EXISTS Then
        retValue = "処理実行要件を満たしませんでした（フォルダ選択が不正です）"
    End If
    
    MessageReaction = retValue
    
End Function
