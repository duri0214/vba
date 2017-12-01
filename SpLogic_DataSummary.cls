VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpLogic_DataSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function DataSummarize(dataBox() As String, keyColumn As String, summaryColumn As String) As String()

    '前提           ：dataBoxのレコードが昇順に並んでいること
    'dataBox        ：ブレイク処理したデータ配列。（2次元配列であることが前提）
    'keyColumn      ：キーにする列の場所    （マルチキー対応済み。マルチキーのときはカンマで区切る）
    'summaryColumn  ：集約する列の場所      （マルチ集約対応済み。マルチ集約のときはカンマで区切る）
    
    Dim summary_col_list() As String                                '集約する列のリスト
    Dim summary_array() As String                                   '結果配列
    Dim summary_zokusei() As String                                 '集約しているデータ以外の、いわゆる属性
    
    Dim boundRow As Integer
    Dim boundCol As Integer
    
    Dim firstProc As Boolean
    Dim key As String
    Dim keyModel As String
    
    Dim i As Integer
    Dim j As Integer
    Dim total() As Double
    
    boundRow = UBound(dataBox, 1)
    boundCol = UBound(dataBox, 2)
    
    summary_col_list = Split(summaryColumn, ",")                    '集約列リストを作成
    ReDim total(UBound(summary_col_list))                           '集約する列分だけTotalの要素を生成
    
    firstProc = True
    For i = 0 To boundRow
        
        key = MakeKeyString(dataBox, i, keyColumn)                  'キー文字列の作成（マルチキー対応済み）
        
        If key <> keyModel Then
        
            If firstProc = False Then
                summary_array = SummaryOut(summary_array, summary_zokusei, summary_col_list, total)
            Else
                firstProc = False
            End If
            
            For j = 0 To UBound(total)
                total(j) = 0                                        '集約値のクリア
            Next
            
            keyModel = key                                          'キーチェンジ
            ReDim summary_zokusei(boundCol)
            summary_zokusei = GetSummaringAttribute(dataBox, i)     'キーチェンジ直後に集約属性を取得
            
        End If
        
        For j = 0 To UBound(total)
            total(j) = total(j) + Val(dataBox(i, summary_col_list(j)))
        Next
        
    Next
    
    summary_array = SummaryOut(summary_array, summary_zokusei, summary_col_list, total)
    
    DataSummarize = summary_array
    
End Function

Private Function MakeKeyString(dataBox() As String, crrRow As Integer, keyColumn As String) As String
    
    'DataSummarizeロジックの子供
    'コントロールブレイクの際に使用される、キー文字列の作成（マルチキー対応済み）ルーチン
    
    Dim retValue As String
    Dim keyList() As String
    
    Dim i As Integer
    
    keyList = Split(keyColumn, ",")
    
    For i = 0 To UBound(keyList)
        retValue = retValue & dataBox(crrRow, keyList(i))
    Next
    
    MakeKeyString = retValue
    
End Function

Private Function GetSummaringAttribute(dataBox() As String, crrRow As Integer) As String()
    
    'DataSummarizeロジックの子供
    'コントロールブレイクの際に使用される、キーチェンジ直後に集約属性を取得するルーチン
    
    Dim i As Integer
    Dim retValue() As String
    Dim boundCol As Integer
    
    boundCol = UBound(dataBox, 2)                                   '元データの最大列数を取得
    
    ReDim retValue(boundCol)                                        '元データと同じ列数に拡張
    
    For i = 0 To boundCol
        retValue(i) = dataBox(crrRow, i)                            '元データから1行分のデータを取得
    Next
    
    GetSummaringAttribute = retValue
    
End Function

Private Function SummaryOut(summary_array() As String, summary_zokusei() As String, summaryColList() As String, summaryValList() As Double) As String()
    
    'DataSummarizeロジックの子供
    'サマリ後配列（結果配列）を1レコード増やして、そこにサマリデータを焼き付ける
    
    Dim retValue() As String
    Dim boundRow As Integer
    Dim boundCol As Integer
    
    Dim i As Integer
    
    retValue = summary_array
    
    If IsValidArray(retValue) = False Then                          '配列がNULLだったら、
        boundRow = 0                                                '最下行数は 0 になる
        boundCol = UBound(summary_zokusei)                          'summary_arrayがNULLなので、列数はsummary_zokuseiから取得
        ReDim retValue(boundRow, boundCol)                          '戻り配列のサイズが確定
    Else
        boundRow = UBound(summary_array, 1) + 1                     '最下行数は summary_arrayの最下行数 + 1 になる
        boundCol = UBound(summary_array, 2)                         'summary_arrayの最大列数を取得
        ReDim retValue(boundRow, boundCol)                          '戻り配列のサイズが確定
        retValue = CopyArray(summary_array, retValue)               '配列のコピー
    End If
        
    For i = 0 To boundCol
        retValue(boundRow, i) = summary_zokusei(i)                  '最終行に、現在集約している対象keyの属性を投入
    Next
    
    For i = 0 To UBound(summaryColList)
        retValue(boundRow, summaryColList(i)) = summaryValList(i)   '最終行に、集約値投入
    Next
    
    SummaryOut = retValue
    
End Function

Private Function CopyArray(summary_array() As String, new_summary_array() As String) As String()
    
    'DataSummarizeロジックの子供のSummaryOutロジックの子供
    '2次元配列のコピーを行います
    
    'preserveキーワードをつけると､最終次元目しか拡張できないので､再構築するためのルーチン｡
    'new_summary_arrayのほうが配列サイズが大きいことを前提としている、配列コピールーチンです
    
    Dim retValue() As String
    Dim boundRow As Integer
    Dim boundCol As Integer
    
    Dim i As Integer
    Dim j As Integer
    
    retValue = new_summary_array
    boundRow = UBound(summary_array, 1)                             '元データの最下行数を取得
    boundCol = UBound(summary_array, 2)                             '元データの最大列数を取得
    
    For i = 0 To boundRow
        For j = 0 To boundCol
            retValue(i, j) = summary_array(i, j)                    'COPY!
        Next
    Next
    
    CopyArray = retValue
    
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
