VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MUFG_OnlyLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'▲notice かなり汎用的になれないということです。この銀行(MUFG)のルールの上でしか役に立たない、
'         モデル的ロジックです。

Function GetPeriodFormation_pre_YYYYMM(koumokuSet() As String, period() As String) As String()

    Const prePostSw As Integer = 0      '項目名を年月文字より先頭につける
    Const formatSw As Boolean = False   '年月フォーマットをYYYYMMで
    Dim retValue() As String
    
    retValue = GetPeriodFormationCore(koumokuSet(), prePostSw, formatSw, period())
    
    GetPeriodFormation_pre_YYYYMM = retValue
    
End Function

Function GetPeriodFormation_pre_EEMM(koumokuSet() As String, period() As String) As String()

    Const prePostSw As Integer = 0      '項目名を年月文字より先頭につける
    Const formatSw As Boolean = True    '年月フォーマットをEEMMで
    Dim retValue() As String
    
    retValue = GetPeriodFormationCore(koumokuSet(), prePostSw, formatSw, period())
    
    GetPeriodFormation_pre_EEMM = retValue
    
End Function

Function GetPeriodFormation_post_YYYYMM(koumokuSet() As String, period() As String) As String()

    Const prePostSw As Integer = 1      '項目名を年月文字よりあとにつける
    Const formatSw As Boolean = False   '年月フォーマットをYYYYMMで
    Dim retValue() As String
    
    retValue = GetPeriodFormationCore(koumokuSet(), prePostSw, formatSw, period())
    
    GetPeriodFormation_post_YYYYMM = retValue
    
End Function

Function GetPeriodFormation_post_EEMM(koumokuSet() As String, period() As String) As String()

    Const prePostSw As Integer = 1      '項目名を年月文字よりあとにつける
    Const formatSw As Boolean = True    '年月フォーマットをEEMMで
    Dim retValue() As String
    
    retValue = GetPeriodFormationCore(koumokuSet(), prePostSw, formatSw, period())
    
    GetPeriodFormation_post_EEMM = retValue
    
End Function

Private Function GetPeriodFormationCore(koumokuSet() As String, prePostSw As Integer, formatSw As Boolean, period() As String) As String()
    
    'koumokuSet     :取り付ける項目
    'prePostSw      :koumokuSetにセットした各項目を年月文字より先頭（pre:0）につけるか末尾（post:1）につけるかのスイッチ
    'formatSw       :falseにすると月次項目名の日付部をYYYYMMで、trueにすると月次項目名の日付部をYYMM形式で出力
    
    'koumokuSetの単位量あたりを1レイヤとし、何レイヤー作るのか、の値
    Dim create_layers_count As Integer
    
    Dim ym As New UtilYM
    
    Dim crrPeriod As String                         '【ループ】現在の期種（A or B）
    Dim crrMonth As String                          '【ループ】現在の年月
    
    Dim stMonth As String
    Dim enMonth As String
    Dim boundFor As Integer
    
    Dim i As Integer                                '汎用ループカウンタ
    Dim j As Integer
    
    Dim col As Integer                              '【ループ】項目位置カウンタ
    
    Dim outputIdx As Integer                        '【ループ】起源月～１８ヶ月目を表現する累積値
    
    Const PRE As Integer = 0                        '項目名の出力を前にする
    Const POST As Integer = 1                       '項目名の出力を後にする
    
    Dim retValue() As String
    
    
    Dim koumokuBox() As String
    ReDim koumokuBox(UBound(koumokuSet), 1)         '引数「koumokuSet」の要素数分の行が2要素列作成される
    For i = 0 To UBound(koumokuSet)                 '0列目に引数値が入ればPREで出力されるし、1列目に入ればPOSTで出力される。
        koumokuBox(i, prePostSw) = koumokuSet(i)
    Next
    
    '■下準備開始■
    create_layers_count = ((UBound(koumokuSet) + 1) * (UBound(period) + 1)) + ((UBound(koumokuSet) + 1) * (UBound(period) + 1) * 6)
    
    '結果配列のサイズは、CREATE_LAYERS_COUNT * [計数項目数]
    '＋１をしているのは、項目の要素数を把握するため。その後に－１をしているのは、
    '配列を0ベースで開始するための補正値。
    ReDim retValue(create_layers_count - 1)
    '■下準備終了■
    
    outputIdx = 0
    
    For i = 0 To UBound(period)                                                 '指定「期」間ループ
                
        '▲処理「期」から見た期初月と期末月
        stMonth = ym.GetFirstYM_InThisPeriod(period(i))
        enMonth = ym.GetLastYM_InThisPeriod(period(i))
        
        crrPeriod = ym.GetPeriodYAB(stMonth)                                    'stMonthのYYYYABを取得（enMonthでも効果は一緒）
        
        For col = 0 To UBound(koumokuBox)
            retValue(outputIdx) = koumokuBox(col, PRE) & crrPeriod & koumokuBox(col, POST)
            outputIdx = outputIdx + 1
        Next col
        
        boundFor = ym.GetYMInterval(stMonth, enMonth)
        
        '▲算出したstMonthから見た期次処理ループ内の月次処理６ヵ月分
        For j = 0 To boundFor
        
            If formatSw = True Then
                crrMonth = ym.ConvertYM2EM(ym.GetAddYM2(stMonth, j))            '【引数分岐】EEMMフォーマット
            Else
                crrMonth = ym.GetAddYM2(stMonth, j)                             '【引数分岐】YYYYMMフォーマット
            End If
            
            For col = 0 To UBound(koumokuBox)
                retValue(outputIdx) = koumokuBox(col, PRE) & crrMonth & koumokuBox(col, POST)
                outputIdx = outputIdx + 1
            Next col
            
        Next j
    
    Next i
        
    GetPeriodFormationCore = retValue

End Function
