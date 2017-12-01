VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const DBL_QUOTE     As Integer = 34

Const CREATE        As Integer = 1
Const READ          As Integer = 2
Const UPDATE        As Integer = 3
Const DELETE        As Integer = 4

'    '0...null（初期値）
'    '1...Create
'    '2...Read(select)
'    '3...Update
'    '4...Delete

Public crudMode     As Integer
Dim editing         As Boolean

Public editSql      As String

Sub EditSql1_SetHeader()
    
    If Not IsValid Then
        MsgBox "CRUDモードに値を代入して下さい。", vbCritical, "クラス内メソッドを実行できません"
        Exit Sub
    End If
    
    Call ClearSql                       'ヘッダのセットなんだからSqlがクリアされてもかまわない
        
    If crudMode = CREATE Or crudMode = READ Then
        editSql = editSql & "SELECT "   'Create or Select
    End If
        
    editing = True
    
End Sub

Sub EditSql2_AddColumn(tableName As String, columnName As String)

    If Not IsValid Then
        MsgBox "CRUDモードに値を代入して下さい。", vbCritical, "クラス内メソッドを実行できません"
        Exit Sub
    End If
    
    If crudMode = CREATE Or crudMode = READ Then
        editSql = editSql & "[" & tableName & "].[" & columnName & "], "
    End If

End Sub

Sub EditSql2_AddColumn_Shukei(tableName As String, columnName As String)

    If Not IsValid Then
        MsgBox "CRUDモードに値を代入して下さい。", vbCritical, "クラス内メソッドを実行できません"
        Exit Sub
    End If
    
    If crudMode = CREATE Or crudMode = READ Then
        editSql = editSql & "SUM(" & "[" & tableName & "].[" & columnName & "]) AS " & columnName & " の合計, "
    End If
    
End Sub

Sub EditSql2_AddFreeSQL(context As String)
    
    '▲notice
    'フリーハンドSqlを追記します。文法エラー、半角空白の漏れ等に気をつけて下さい
    
    If Not IsValid Then
        MsgBox "CRUDモードに値を代入して下さい。", vbCritical, "クラス内メソッドを実行できません"
        Exit Sub
    End If
    
    editSql = editSql & context
    
End Sub

Sub EditSql3_SetFooter(tableName As String, Optional newTableName As String)
    
    If Not IsValid Then
        MsgBox "CRUDモードに値を代入して下さい。", vbCritical, "クラス内メソッドを実行できません"
        Exit Sub
    End If
    
    If crudMode = CREATE Then
        If newTableName <> vbNullString Then
            editSql = editSql & "INTO " & newTableName & " From [" & tableName & "];"
            editSql = Replace(editSql, ", INTO", " INTO")   '最終項目のカンマkill
        Else
            MsgBox "Create SQL の作成時は作成されるテーブル名を引数に代入", vbCritical, "SQL生成エラー"
        End If
    End If
    
    If crudMode = READ Then
        editSql = editSql & "From [" & tableName & "];"
        editSql = Replace(editSql, ", From", " From")       '最終項目のカンマkill
    End If
    
    editing = False
    
End Sub

Sub ClearSql()
    
    editSql = vbNullString                              '編集中のSQL文字をクリアします
    
End Sub

Private Function IsValid() As Boolean

    If crudMode = 0 Then
        IsValid = False     'クラス内メソッド実行不可能（準備不足）
    Else
        IsValid = True      'クラス内メソッド実行可能
    End If
    
End Function

