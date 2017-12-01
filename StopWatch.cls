VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'配列NULLにするやりかたがわからねーもんで、リセット機能はなしです
'あとでわかったら追加しよう

Dim lap() As String
Dim interval() As String

Property Get LapList() As String()
    LapList = lap
End Property

Property Get LapIntervalList() As String()
    LapIntervalList = interval
End Property

Sub GetLap()
    
    If Not IsValidArray(lap) Then
        ReDim lap(0)
        ReDim interval(0)
    Else
        ReDim Preserve lap(UBound(lap) + 1)
        ReDim Preserve interval(UBound(interval) + 1)
    End If
    
    lap(UBound(lap)) = Now
    
    If UBound(lap) > 0 Then
        interval(UBound(lap)) = CalcInterval(lap(UBound(lap) - 1), lap(UBound(lap)))
    End If
    
End Sub

Private Function CalcInterval(stTime As String, enTime As String) As String
    
    Const 総 As Integer = 0
    Const 時 As Integer = 1
    Const 分 As Integer = 2
    Const 秒 As Integer = 3
    
    Dim time(3) As Long   '[0]総合計(秒)　[1]内訳(時)　[2]内訳(分)　[3]内訳(秒)
    Dim buf As Long
    Dim retValue As String
    
    buf = DateDiff("s", stTime, enTime)     '処理時間間隙取得[秒]
    
    time(総) = buf                          '総時間確定
    time(時) = buf \ 3600                   '秒の値を 3600(=1時間) で割って、内訳(時)確定
    buf = buf - (3600 * time(時))
    time(分) = buf \ 60                     '1時間未満になった値を 60(=1分) で割って、内訳(分)確定
    buf = buf - (60 * time(分))
    time(秒) = buf                          '内訳(秒)確定
    
    retValue = time(時) & "時間" & time(分) & "分" & time(秒) & "秒(" & time(総) & "秒)"
    CalcInterval = retValue
        
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

