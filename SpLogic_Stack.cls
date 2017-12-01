VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpLogic_Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim ary() As String

Property Get ret() As String()
    ret = ary
End Property

Sub Push(value As String)
    
    If Not IsValidArray(ary) Then
        ReDim ary(0)
    Else
        ReDim Preserve ary(UBound(ary) + 1)
    End If
    
    ary(UBound(ary)) = value
    
End Sub

Function Pop() As String
    
    Dim retValue As String
    
    If IsValidArray(ary) Then
        retValue = ary(UBound(ary))
        If UBound(ary) = 0 Then
            Erase ary
        Else
            ReDim Preserve ary(UBound(ary) - 1)
        End If
    Else
        MsgBox "PUSHした回数を超えるPOPが実行されました。", vbCritical, "ERROR：スタック配列がNull"
        ary(0) = vbNullString                       'エラーにさせるための措置
    End If
    
    Pop = retValue
    
End Function

Private Function IsValidArray(targetArray() As String) As Boolean
    
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

