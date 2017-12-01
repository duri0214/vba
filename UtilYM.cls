VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UtilYM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function GetFirstYM_InThisPeriod(YYYYAB As String) As String
    
    '「期」の情報を入力し、その「期」の期初月を返す
    
    Dim retValue As String
    Dim period As String
    
    period = Right(YYYYAB, 1)               '例：2008B の B
    
    If period = "A" Then                    '上期なら
        retValue = Left(YYYYAB, 4) & "04"   '例：200804
    End If

    If period = "B" Then                    '下期なら
        retValue = Left(YYYYAB, 4) & "10"   '例：200810
    End If
    
    GetFirstYM_InThisPeriod = retValue

End Function

Function GetLastYM_InThisPeriod(YYYYAB As String) As String
    '「期」の情報を入力し、その「期」の期末月を返す
    GetLastYM_InThisPeriod = GetAddYM2(GetFirstYM_InThisPeriod(YYYYAB), 5)
End Function

Function GetTheEndOfTheMonth(YYYYMM As String) As String
    
    'YYYYMMを引き込んで、月末日を含んだ文字列を返す
    '200902 → 20090228
    
    Dim retValue As String
    
    retValue = GetFormalDate(YYYYMM)
    retValue = Replace(DateAdd("d", -1, DateAdd("m", 1, retValue)), "/", "")
    
    GetTheEndOfTheMonth = retValue
    
End Function

Function GetTheEndOfTheMonth_RetEM(YYYYMM As String) As String

    'YYYYMMを引き込んで、月末日を含んだ文字列を返す
    '200902 → 210228
    
    Dim retValue As String
    
    retValue = Right(GetTheEndOfTheMonth(YYYYMM), 2)    'DD
    retValue = ConvertYM2EM(YYYYMM) & retValue          'EEMM + DD
    
    GetTheEndOfTheMonth_RetEM = retValue
        
End Function

Function GetNowYM() As String
    
    'systemYMを返します
    
    GetNowYM = Year(Now) & Format(month(Now), "00")
    
End Function

Function GetAddYM(i As Integer) As String
    
    'systemYMに係数を加味した値を返します。
    '例:2008年01月を現在として、1と引数指定すると200802, -1と引数指定すると200712
    
    Dim kongetu As String
    
    kongetu = Left(GetNowYM, 4) & "/" & Right(GetNowYM, 2)
    GetAddYM = Left(Replace(DateAdd("m", i, kongetu), "/", ""), 6)

End Function

Function GetAddYM2(YYYYMM As String, i As Integer) As String
    
    '任意の年月に係数を加味した値を返します。
    '例:yyyymm引数に2008年01月を代入し、i引数に1と引数指定すると200802, -1と引数指定すると200712
    
    Dim retYM As String
    
    retYM = Left(YYYYMM, 4) & "/" & Right(YYYYMM, 2)
    GetAddYM2 = Left(Replace(DateAdd("m", i, retYM), "/", ""), 6)
    
End Function

Function GetYMInterval(kaishiYM As String, shuryoYM As String) As Integer
    
    '任意の年月同士の間隔を返します。ある年月からある年月までの間、ループする、
    'などという場合に便利です。
    '引数はYYYYMM形式で指定する。
    'kaishiYM引数に200801, shuryoYM引数に200812と指定すると、11が返る。
    'ループは0を始点として11回ループさせれば良い。
    
    Dim retValue As Integer
    Dim w_kaishiYM As String
    Dim w_shuryoYM As String
    
    w_kaishiYM = Left(kaishiYM, 4) & "/" & Right(kaishiYM, 2) & "/01"
    w_shuryoYM = Left(shuryoYM, 4) & "/" & Right(shuryoYM, 2) & "/01"
    
    retValue = DateDiff("m", w_kaishiYM, w_shuryoYM)
    
    GetYMInterval = retValue
    
End Function

Function GetKaikeiNendo(YYYYMM As String) As String
    
    '指定年月の会計年度を返します。
    '200701～200703 → 2006
    '200704～200803 → 2007
    'という感じに返ります。
    
    Dim retYear As String
    Dim retMonth As String
    
    retYear = Left(YYYYMM, 4)
    retMonth = Right(YYYYMM, 2)
    
    If Val(retMonth) < 4 Then
        retYear = Val(retYear) - 1
    End If
    
    GetKaikeiNendo = retYear
    
End Function

Private Function GetPeriod(YYYYMM As String) As String
    
    '引数に西暦を入力すると、期を表すアルファベットが返ります
    '例：200706といれた。
    '返ってくる値は「A」である。
    
    Dim month As Integer
    
    Const KAMIHANKI As String = "A"
    Const SIMOHANKI As String = "B"
        
    month = CInt(Right(YYYYMM, 2))
    
    GetPeriod = SIMOHANKI
    
    If (month >= 4) And (month <= 9) Then
        GetPeriod = KAMIHANKI
    End If
    
End Function

Function GetPeriodYAB(YYYYMM As String) As String
    
    '引数に西暦を入力すると、YYYYAB形式で返します
    '例：200706といれた。
    '返ってくる値は「2007A」である。
    
    Dim retValue As String
    
    retValue = GetKaikeiNendo(YYYYMM) & GetPeriod(YYYYMM)  'YYYYAB変換
    
    GetPeriodYAB = retValue

End Function

Function GetPeriodEAB(YYYYMM As String) As String
    
    '引数に西暦を入力すると、年号AB形式で返します
    '例：200706といれた。
    '返ってくる値は「19A」である。
    
    Dim kaikeiNendo As String
    Dim retValue As String
    
    kaikeiNendo = GetKaikeiNendo(YYYYMM)
    retValue = ConvertY2E(kaikeiNendo) & GetPeriod(YYYYMM)

    GetPeriodEAB = retValue
    
End Function

Function GetPeriod_Prev_RetYAB(i As Integer) As String
    '当月（systemYM - 1）から、期をステップバックした西暦期（YYYYAB）を算出
    GetPeriod_Prev_RetYAB = GetPeriodYAB(GetAddYM(-1 - (6 * i)))
End Function

Function GetPeriod_Prev_RetEAB(i As Integer) As String
    '当月（systemYM - 1）から、期をステップバックした和暦期（EEAB）を算出
    GetPeriod_Prev_RetEAB = GetPeriodEAB(GetAddYM(-1 - (6 * i)))
End Function

Function GetPeriod_Prev_RetFirstYM(i As Integer) As String

    '当月（systemYM - 1）から、期をステップバックした西暦期（YYYYAB）を算出し、
    '算出した期の期初月次を返す
    Dim retValue As String
    
    retValue = GetPeriod_Prev_RetYAB(i)
    retValue = Replace(retValue, "A", "04")
    retValue = Replace(retValue, "B", "10")
    
    GetPeriod_Prev_RetFirstYM = retValue
    
End Function

Function GetPeriod_Prev_RetFirstEM(i As Integer) As String

    '当月（systemYM - 1）から、期をステップバックした和暦期（EEAB）を算出し、
    '算出した期の期初月次を返す
    Dim retValue As String
    
    retValue = GetPeriod_Prev_RetEAB(i)
    retValue = Replace(retValue, "A", "04")
    retValue = Replace(retValue, "B", "10")
    
    GetPeriod_Prev_RetFirstEM = retValue
    
End Function

Function GetPeriod_Prev_RetLastYM(i As Integer) As String

    '当月（systemYM - 1）から、期をステップバックした西暦期（YYYYAB）を算出し、
    '算出した期の期末月次を返す
    Dim retValue As String
    
    retValue = GetAddYM2(GetPeriod_Prev_RetFirstYM(i), 5)
    
    GetPeriod_Prev_RetLastYM = retValue
    
End Function

Function GetPeriod_Prev_RetLastEM(i As Integer) As String

    '当月（systemYM - 1）から、期をステップバックした和暦期（EEAB）を算出し、
    '算出した期の期末月次を返す
    Dim retValue As String
    
    retValue = ConvertYM2EM(GetAddYM2(GetPeriod_Prev_RetFirstYM(i), 5))
    
    GetPeriod_Prev_RetLastEM = retValue
    
End Function

Function GetFormalDate(YYYYMM As String) As String
    
    '引数値YYYYMMフォーマットの年月を、YYYY/MM/01のフォーマットに変換する。
    '年月同士をコンペア関数にひっかけたりするときに使用する
    Dim retValue As String
    retValue = Left(YYYYMM, 4) & "/" & Right(YYYYMM, 2) & "/01"
    GetFormalDate = retValue
    
End Function

Function GetFormalDate2(YYYYMMDD As String) As String
    
    '引数値YYYYMMDDフォーマットの年月を、YYYY/MM/DDのフォーマットに変換する。
    '年月同士をコンペア関数にひっかけたりするときに使用する
    Dim retValue As String
    retValue = Left(YYYYMMDD, 4) & "/" & Mid(YYYYMMDD, 5, 2) & "/" & Right(YYYYMMDD, 2)
    GetFormalDate2 = retValue
    
End Function

Function GetPeriodListYAB(kaishiYM As String, shuryoYM As String) As String()
    
    '配列をnull扱いにできないので、少々複雑化せざるを得ない
    '部分的にエラー無視状態にして、配列の0番目にアクセスし、エラーが発生したら、
    '配列を1要素分として初期化する。
    
    '200504～200504までを指定すると、「2005A」という配列を返す
    '200504～200712までを指定すると、「2005A,2005B,2006A,2006B,2007A,2007B」という配列を返す
    
    Dim crrYM As String
    Dim crrPeriod As String
    Dim compPeriod As String
    Dim boundFor As Integer                                             'ループの行き止まり
    
    Dim retValue() As String
    Dim firstProc As Boolean
    
    Dim i As Integer
    
    firstProc = True
    compPeriod = vbNullString
    
    boundFor = GetYMInterval(kaishiYM, shuryoYM)                        'ループ終了idx算出
    
    For i = 0 To boundFor
    
        crrYM = GetAddYM2(kaishiYM, i)
        crrPeriod = GetKaikeiNendo(crrYM) & GetPeriod(crrYM)
        
        If crrPeriod <> compPeriod Then
            
            If firstProc = False Then
                If IsValidArray(retValue) = False Then
                    ReDim retValue(0)                                   'エラーしたということは、配列がnullなので初期化
                Else
                    ReDim Preserve retValue(UBound(retValue) + 1)       '配列サイズ増量
                End If
                retValue(UBound(retValue)) = compPeriod                 '配列の最後に年号の代入
            Else
                firstProc = False
            End If
                        
            compPeriod = crrPeriod                                      'compPeriodの更新
            
        End If
        
    Next i
      
    If IsValidArray(retValue) = False Then
        ReDim retValue(0)                                   'エラーしたということは、配列がnullなので初期化
    Else
        ReDim Preserve retValue(UBound(retValue) + 1)       '配列サイズ増量
    End If
    retValue(UBound(retValue)) = compPeriod                 '配列の最後に年号の代入
    
    GetPeriodListYAB = retValue
    
End Function

Function GetPeriodListEAB(kaishiYM As String, shuryoYM As String) As String()
        
    '200504～200504までを指定すると、「17A」という配列を返す
    '200504～200712までを指定すると、「17A,17B,18A,18B,19A,19B」という配列を返す
    
    Dim retValue() As String
    Dim i As Integer
    Dim crrSeireki As String
    
    retValue = GetPeriodListYAB(kaishiYM, shuryoYM)
    
    '▲notice
    'GetPeriodListから返ってきた配列要素に対して、西暦型Period（YYYYAB）の左4桁を和暦に変えているだけ。
    '2007B → 19B
    
    For i = 0 To UBound(retValue)
        crrSeireki = Left(retValue(i), 4)
        retValue(i) = Replace(retValue(i), crrSeireki, ConvertY2E(crrSeireki))
    Next
    
    GetPeriodListEAB = retValue
    
End Function

Function GetYesterday() As String
    
    '昨日営業日を取り出す関数
    '"昨日"の日付を取得して、その曜日が週末(土, 日)なら、再起的に -1日 の引き算して
    '"平日(月～金)"の日付[=昨日営業日]を取り出す YYYY/MM/DD
    '祝日未対応
    
    Dim retValue As String
    Dim 曜日 As Integer
    
    retValue = DateAdd("d", -1, Date)           '昨日
        
Looop:
    曜日 = Weekday(retValue, vbMonday)          '月曜を基点として 1:月～7:日 の値が返る
    If 曜日 >= 6 Then                           '土曜日(6)か日曜日(7)なら
        retValue = DateAdd("d", -1, retValue)   '再起的に -1日 の引き算
        GoTo Looop
    End If
    
    GetYesterday = retValue
    
End Function

Private Function IsValidArray(targetArray() As String) As Boolean
    
    'targetArrayがnullかどうかを調べる関数。
    'vbaでは、配列にnullを仕込むことができないため
    '本来この関数はいらなくて if targetArray = null then... とやりたいところだが
    'それができなくて、わざわざエラー無視の状態にして、
    '配列の第一要素目にアクセス。
    '配列にアクセスできてなければ（エラーがhitしていれば）falseを返す（配列がnullってるってこと）
    '配列にアクセスできれば、trueを返す
    
    Dim checkValue As Variant
    Dim retValue As Boolean
    
    On Error Resume Next            'エラーを無効化
    
    checkValue = targetArray(0)     '配列の第一要素にアクセス
    
    If checkValue = Empty Then      '検査値がEmptyを示すなら、未定義配列です。
        retValue = False
    Else                            'そうでないなら、1以上の要素が宣言済みです。
        retValue = True
    End If
    Err.Clear                       'エラーオブジェクト初期化
    
    On Error GoTo 0                 'エラーを有効化
    
    IsValidArray = retValue         '検査結果を返す

End Function

Function ConvertYM2EM(YYYYMM As String) As String

    '例：「200604」→「1804」と変換
    ConvertYM2EM = Format(Left(YYYYMM, 4) & "/" & Right(YYYYMM, 2), "e") & Right(YYYYMM, 2)
    
End Function

Function ConvertY2E(yyyy As String) As String
    
    '例：「2006」→「18」と変換
    ConvertY2E = Format(yyyy & "/01", "e")
    
End Function
