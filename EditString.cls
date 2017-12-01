VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EditString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Const DBL_QUOTE As Integer = 34

Function GetStringLengthB(str As String) As Integer

    GetStringLengthB = LenB(StrConv(str, vbFromUnicode))
    
End Function

Function StringCutterForFixed(koteichou As String, separate() As Integer) As String()
    
    Dim editing As String           '固定長文字から文字をカットするごとにどんどん文字を削っていく一時処理変数
    Dim cutElemBox() As String      'カットされたエレメント（分離された文字）が入る配列
    Dim procable As Boolean
    Dim i As Integer
    
    procable = False
    
    'SUM検査（koteichou文字数と、カットしたい文字数合計値が等しいか）
    '１０文字の固定長文字を「２，２，６」にちぎりたい。
    'では、１０＝２＋２＋６でないと、処理条件を満たさない、とする。（未定義配列を返す）
    Dim cutLengthAllAdd As Integer
    
    For i = 0 To UBound(separate)
        cutLengthAllAdd = cutLengthAllAdd + separate(i)
    Next
    
    If Len(koteichou) = cutLengthAllAdd Then
        procable = True
    End If
    
    '検査終了、処理に入る。
    If procable Then
    
        ReDim cutElemBox(UBound(separate))
        editing = koteichou
        
        For i = 0 To UBound(separate)
            cutElemBox(i) = Left(editing, separate(i))
            editing = Mid(editing, separate(i) + 1)
        Next
        
        StringCutterForFixed = cutElemBox
        
    End If
    
End Function

Function StringCutterForCSV(commaMixString As String) As String()
    
    Dim editing As String
    Dim retValue() As String
    
    editing = Replace(commaMixString, Chr(DBL_QUOTE), "")   'ダブルクォートごろし
    retValue = Split(editing, ",")
    
    StringCutterForCSV = retValue
    
End Function

Function Fit(text As String, retLengthB As Integer) As String

    Dim editText As String
    Dim procable As Boolean
    Dim mojisuuB As Integer
    Dim padSpaceOf As Integer
    
    procable = False
    editText = text
    
    If Len(text) < retLengthB Then   '返したい目的の文字数よりも小さいときにロジック発動
        procable = True
    End If
    
    '検査終了、処理に入る。
    If procable Then
        
        mojisuuB = LenB(StrConv(editText, vbFromUnicode))
        padSpaceOf = retLengthB - mojisuuB
        editText = editText & String(padSpaceOf, " ")
        
    End If
    
    Fit = editText
        
End Function

