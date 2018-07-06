VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSelectDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'▲notice
'FileSelectDialogクラスの使用時は、Microsoft Office 11.0 Object Libraryへのアーリーバインディングが必要です。

Function FileDialog(Optional title As String) As String
    
    Const mode As Integer = msoFileDialogFilePicker             'ファイルピッカー
    Dim editTitle As String
    Dim retValue As String
    
    editTitle = "対象ファイルを選択して「開く」"                'ダイアログのタイトルバーの文字
    If title <> vbNullString Then
        editTitle = title
    End If
    
    retValue = FileDialogCore(mode, editTitle)
    FileDialog = retValue
    
End Function

Function FolderDialog(Optional title As String) As String

    Const mode As Integer = msoFileDialogFolderPicker           'フォルダピッカー
    Dim editTitle As String
    Dim retValue As String
    
    editTitle = "対象フォルダの位置で「開く」"                  'ダイアログのタイトルバーの文字
    If title <> vbNullString Then
        editTitle = title
    End If
    
    retValue = FileDialogCore(mode, editTitle)
    FolderDialog = retValue
    
End Function

Private Function FileDialogCore(mode As Integer, title As String) As String

    Dim path As String
    Dim dialog As FileDialog
    Dim retValue As String
    
    Set dialog = Application.FileDialog(mode)                   'ピッカー
    dialog.initialfilename = thisworkbook.path                  'Accessの場合は自パスフォルダに変えて
    dialog.AllowMultiSelect = False                             'ファイルは複数選択できません
    dialog.title = title                                        'ダイアログのタイトルバーの文字
    
    If dialog.Show Then                                         'ダイアログ表示
        retValue = dialog.SelectedItems(1)
    End If
    
    FileDialogCore = retValue                                   '結果返し
    
End Function
