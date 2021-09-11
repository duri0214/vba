Option Explicit

Sub sort_for_graphical_display()
    
    '「元データ」シートを選択する
    Sheet1.Select
    
    '「元データ」シートのデータを並べ替え
    Dim r As Range
    Set r = ActiveSheet.Range("A1")                     'A1を取得
    Set r = ActiveSheet.Range(r, r.End(xlDown))         'rを基準に最終行までを取得
    Set r = ActiveSheet.Range(r, r.End(xlToRight))      'rを基準に最終列までを取得
    With ActiveSheet.Sort
        .SortFields.Clear                               '前回のソート条件を初期化
        .SortFields.Add Key:=ActiveSheet.Range("D1")    '圃場名でソート
        .SortFields.Add Key:=ActiveSheet.Range("G1")    '圃場内位置でソート
        .SortFields.Add Key:=ActiveSheet.Range("H1")    '圃場内位置2でソート
        .SetRange r                                     '並べ替え範囲をセットする
        .header = xlYes                                 '並べ替え範囲の1行目をヘッダーとして認識し、並べ替えから除外する
        .Apply                                          'ソートの実行
    End With
    
    'パラメータC,D,Eの中央値を求める
    Const C As Integer = 0
    Const D As Integer = 1
    Const E As Integer = 2
    Dim param(2) As Double                          '小数点の配列3要素を宣言
    
    Set r = ActiveSheet.Range("DA2")
    Set r = ActiveSheet.Range(r, r.End(xlDown))     '変数rにパラメータCのデータ範囲を取得した
    param(C) = WorksheetFunction.Median(r)          '配列：param(C)にパラメータCの中央値を代入する
    
    Set r = ActiveSheet.Range("DB2")
    Set r = ActiveSheet.Range(r, r.End(xlDown))     '変数rにパラメータDのデータ範囲を取得した
    param(D) = WorksheetFunction.Median(r)          '配列：param(D)にパラメータDの中央値を代入する
    
    Set r = ActiveSheet.Range("DC2")
    Set r = ActiveSheet.Range(r, r.End(xlDown))     '変数rにパラメータEのデータ範囲を取得した
    param(E) = WorksheetFunction.Median(r)          '配列：param(E)にパラメータEの中央値を代入する
    
    '雛形グラフの原点をパラメータの中央値にする
    グラフ1.Axes(xlValue).CrossesAt = param(E)      '特性深度×緩衝因子のグラフの縦軸交点にパラメータEの中央値を代入
    グラフ1.Axes(xlCategory).CrossesAt = param(D)   '特性深度×緩衝因子のグラフの横軸交点にパラメータDの中央値を代入
    グラフ2.Axes(xlValue).CrossesAt = param(C)      '特性深度×最大硬度のグラフの縦軸交点にパラメータCの中央値を代入
    グラフ2.Axes(xlCategory).CrossesAt = param(D)   '特性深度×最大硬度のグラフの横軸交点にパラメータDの中央値を代入
        
    'exceldbを使って、sheet3"並べ替え②"のデータ範囲を特定する
    Dim db As New exceldb
    Dim hdr As Range
    Dim data As Range
    
    'exceldbを使うための下準備（ヘッダー範囲とデータ範囲を定義）
    Set hdr = ActiveSheet.Range("A1:CX1")
    Set data = ActiveSheet.Range("A2:CX2")
    Set data = ActiveSheet.Range(data, data.End(xlDown))
    data.Select
   
    '圃場1つを選択（ドロップダウンから持ってくる）
    Dim point_hojou As String
    point_hojou = "中原-開パイ下"

    '重複削除用の作業シートを作成
    Dim sh_keys As Worksheet
    Set sh_keys = CreateSheet("keys")
    sh_keys.Cells.ClearContents

    '圃場1つのデータ範囲を取得して、重複削除用の作業シートに貼り付けて、重複を削除したらループ回数がわかる
    Dim rows As Long
    db.SetInit hdr, data
    Set r = db.GetCurser_r("圃場名", point_hojou)
    r(1, 7).Resize(r.rows.Count, 2).Copy Destination:=sh_keys.Range("A1")
    sh_keys.Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2), header:=xlNo
    rows = sh_keys.Range("A1").CurrentRegion.rows.Count

    'グラフタイトル（作物名_圃場_日付）を作ってセットする e.g. レタス_中原-開パイ下_21.07.25
    Dim graphTitle As String
    graphTitle = ConcatenateByRange(r(1, 3).Resize(, 3))
    グラフ1.ChartTitle.Text = graphTitle
    グラフ2.ChartTitle.Text = graphTitle
    グラフ6.ChartTitle.Text = graphTitle

    '最大硬度×最大深度 のグラフは散布図なので元データから設定できる
    グラフ3.ChartTitle.Text = graphTitle
    グラフ3.SetSourceData Source:=ActiveSheet.Range(GetCellAddressAfterRangeUnion(ActiveSheet.Range("DA2"), ActiveSheet.Range("DE2")))

    '平均集計シートの下準備
    Dim sh_agg As Worksheet
    Set sh_agg = CreateSheet("agg")
    sh_agg.Cells.ClearContents
    sh_agg.Range("A1").Value = "深度"
    sh_agg.Range("B1").Value = 1
    sh_agg.Range("B1").AutoFill Destination:=sh_agg.Range("B1").Resize(, 60), Type:=xlFillSeries

    '5点測点データ(60ヶ所)を集計(平均)して、1/1000しながら平均集計シートに転記します
    Dim point_alphabet As String
    Dim point_number As String
    Dim point_name As String
    Dim row As Long
    For row = 1 To rows
        point_alphabet = sh_keys.Cells(row, "A").Value
        point_number = sh_keys.Cells(row, "B").Value
        point_name = "硬度" & point_alphabet & point_number 'e.g. 硬度A1
        db.SetInit hdr, data
        Set r = db.GetCurser_r("圃場名", point_hojou)
        Set r = db.GetCurser_r("圃場内位置", point_alphabet)
        Set r = db.GetCurser_r("圃場内位置2", point_number)
        CopyAfterAggregate sh_agg, point_name, r(, 13).Resize(r.rows.Count, 60)
    Next row
    
End Sub
    
Function ConcatenateByRange(rgs As Range, Optional separate As String) As String
    '範囲形式で参照される値を連結する
    Dim buf As String
    Dim r As Range
    For Each r In rgs
        If separate <> vbNullString And buf <> vbNullString Then
            buf = buf & separate
        End If
        buf = buf & r.Value
    Next
    ConcatenateByRange = buf
End Function

Function SheetExists(search_name As String) As Boolean
    'シートが存在するかを調べる
    Dim sh As Worksheet
    SheetExists = False
    For Each sh In Worksheets
        If sh.Name = search_name Then
            SheetExists = True
        End If
    Next sh
End Function

Function CreateSheet(sheetName As String) As Worksheet
    'シートが存在するかを調べて、存在しなければシートを新規作成する
    If SheetExists(sheetName) Then
        Set CreateSheet = Worksheets(sheetName)
    Else
        Set CreateSheet = Worksheets.Add(After:=ActiveSheet)
        CreateSheet.Name = sheetName
    End If
End Function

Sub CopyAfterAggregate(toCopySheet As Worksheet, point_name As String, rng As Range)
    '5点測点データ(60ヶ所)を集計(平均)して、1/1000して toCopySheet に転記します。A列にkey情報 point_name を入れます
    Dim r As Range
    Dim i As Integer
    Set r = toCopySheet.Cells(toCopySheet.rows.Count, "A").End(xlUp).Offset(1)
    r.Value = point_name
    For i = 1 To rng.Columns.Count
        r.Offset(, i).Value = WorksheetFunction.Average(rng(i).Resize(rng.rows.Count)) / 1000
    Next i
End Sub

Function GetCellAddressAfterRangeUnion(rng1 As Range, rng2 As Range) As String
    '=元データ!$DA$2:$DA$46,元データ!$DE$2:$DE$46 のような値を返します
    Set rng1 = rng1.Parent.Range(rng1, rng1.End(xlDown))
    Set rng2 = rng1.Parent.Range(rng2, rng2.End(xlDown))
    GetCellAddressAfterRangeUnion = rng1.Address(External:=True) & "," & rng2.Address(External:=True)
End Function


