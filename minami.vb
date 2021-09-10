Option Explicit

Sub sort_for_graphical_display()
    
    'TODO: トリガーボタンはどのシートにありますか？Sheet1ですか？どこかのメニューからこのルーチンに来るのか、

    '「元データ」シートを選択する
    Sheet1.Select
    
    'STEP01: 「元データ」シートのデータを並べ替え
    Dim r As Range
    Set r = ActiveSheet.Range("A1")                     'A1を取得
    Set r = ActiveSheet.Range(r, r.End(xlDown))         'rを基準に最終行までを取得
    Set r = ActiveSheet.Range(r, r.End(xlToRight))      'rを基準に最終列までを取得
    With ActiveSheet.Sort
        .SortFields.Clear                               '前回のソート条件を初期化
        .SortFields.Add Key:=ActiveSheet.Cells(1, "D")  '圃場名でソート
        .SortFields.Add Key:=ActiveSheet.Cells(1, "G")  '圃場内位置でソート
        .SortFields.Add Key:=ActiveSheet.Cells(1, "H")  '圃場内位置2でソート
        .SetRange r                                     '並べ替え範囲をセットする
        .header = xlYes                                 '並べ替え範囲の1行目をヘッダーとして認識し、並べ替えから除外する
        .Apply                                          'ソートの実行
    End With

    'STEP02: 並べ替えたデータを別Sheet「並べ替え①」に転記する（各パラメータの中央値を出す用）
    r.Copy Destination:=Worksheets("並べ替え①").Range("A1")  '「元データ」で並べ替えたデータ範囲をを"並べ替え①"のA1を基準に転記する
    Worksheets("並べ替え①").Columns("I:CX").Hidden = True
    
    'グラフの交点になる縦軸・横軸の中央値を求める
    Dim param(2) As Double   '小数点のある数値の変数として配列：param(0)～param(2)を宣言
    
    Set r = Sheet2.Range("DA2")
    Set r = Sheet2.Range(r, r.End(xlDown))   '変数RにDA2のパラメータCのデータ範囲を取得した
    param(0) = WorksheetFunction.Median(r)   '配列：param(0)にDA2のデータ範囲の中央値を代入する
    
    Set r = Sheet2.Range("DB2")
    Set r = Sheet2.Range(r, r.End(xlDown))   '変数RにDB2のパラメータCのデータ範囲を取得した
    param(1) = WorksheetFunction.Median(r)   '配列：param(1)にDB2のデータ範囲の中央値を代入する
    
    Set r = Sheet2.Range("DC2")
    Set r = Sheet2.Range(r, r.End(xlDown))   '変数RにDC2のパラメータCのデータ範囲を取得した
    param(2) = WorksheetFunction.Median(r)   '配列：param(2)にDC2のデータ範囲の中央値を代入する
    
    グラフ1.Axes(xlValue).CrossesAt = param(2)   '特性深度×緩衝因子のグラフの縦軸交点にparam(2)を代入
    グラフ1.Axes(xlCategory).CrossesAt = param(1)   '特性深度×緩衝因子のグラフの横軸交点にparam(1)を代入
    
    グラフ2.Axes(xlValue).CrossesAt = param(0)   '特性深度×緩衝因子のグラフの縦軸交点にparam(0)を代入
    グラフ2.Axes(xlCategory).CrossesAt = param(1)   '特性深度×緩衝因子のグラフの横軸交点にparam(1)を代入
    
    
    'グラフ1～3のタイトルをsheet2"並べ替え①"の複数セルの値を組み合わせて作成し、代入する・・前処理：変数
    
    Dim s As String  '「品目」
    Dim f As String  '「圃場名」
    Dim t As String     '「収集日時」

    s = Worksheets("並べ替え①").Range("C2")
    f = Worksheets("並べ替え①").Range("D2")
    t = Worksheets("並べ替え①").Range("E2")
 
    'グラフ1のタイトルを代入する
    
    With グラフ1
    
        .HasTitle = True                    'タイトルをグラフに表示する
        .ChartTitle.Formula = s + "_" + f + "_" + t 'タイトルの文字列を指定する
        .ChartTitle.Top = 5                 'TOP位置
        .ChartTitle.Left = 100               'Left位置
    
        With .ChartTitle.Format.TextFrame2.TextRange.Font
    
            .Size = 16 '文字のサイズ
            .Fill.ForeColor.ObjectThemeColor = 2 '色を指定する
        
        End With
    
    End With
    
    'グラフ2のタイトルを代入する
    
    With グラフ2
    
        .HasTitle = True                    'タイトルをグラフに表示する
        .ChartTitle.Formula = s + "_" + f + "_" + t 'タイトルの文字列を指定する
        .ChartTitle.Top = 5                 'TOP位置
        .ChartTitle.Left = 260               'Left位置
    
        With .ChartTitle.Format.TextFrame2.TextRange.Font
    
            .Size = 14 '文字のサイズ
            .Fill.ForeColor.ObjectThemeColor = 2 '色を指定する
        
        End With
    
    End With
    
     'グラフ2のタイトルを代入する
    
    With グラフ3
    
        .HasTitle = True                    'タイトルをグラフに表示する
        .ChartTitle.Formula = s + "_" + f + "_" + t 'タイトルの文字列を指定する
        .ChartTitle.Top = 5                 'TOP位置
        .ChartTitle.Left = 260               'Left位置
    
        With .ChartTitle.Format.TextFrame2.TextRange.Font
    
            .Size = 14 '文字のサイズ
            .Fill.ForeColor.ObjectThemeColor = 2 '色を指定する
        
        End With
    
    End With
    
     'グラフ6のタイトルを代入する
    
    With グラフ6
    
        .HasTitle = True                    'タイトルをグラフに表示する
        .ChartTitle.Formula = s + "_" + f + "_" + t 'タイトルの文字列を指定する
        .ChartTitle.Top = 5                 'TOP位置
        .ChartTitle.Left = 260               'Left位置
    
        With .ChartTitle.Format.TextFrame2.TextRange.Font
    
            .Size = 14 '文字のサイズ
            .Fill.ForeColor.ObjectThemeColor = 2 '色を指定する
        
        End With
    
    End With
    
    '標準モジュール"exceldb"を使って、sheet3"並べ替え②"のデータ範囲を特定する
    
    Dim db As New exceldb
    Dim hdr As Range
    Dim data As Range
    
    Dim ret As Variant
        
    Set hdr = Sheet1.Range("A1:CX1")
    Set data = Sheet1.Range("A2:CX2")
    Set data = Sheet1.Range(data, data.End(xlDown))
    
    data.Select
    
    
   
   'ループ処理・・変数
    
    Dim locations_col(2) As String
    Dim locations_row(2) As String
        locations_col(0) = "A"
        locations_col(1) = "B"
        locations_col(2) = "C"
        locations_row(0) = "1"
        locations_row(1) = "2"
        locations_row(2) = "3"
    Dim col As Integer
    Dim row As Integer
    Dim destination_idx As Integer
    
    
    'ループ処理・・データ範囲の検索と"並べ替え②"の所定範囲へ転記
    
    For col = LBound(locations_col) To UBound(locations_col)  '変数locations_colの0～2まで
               
        For row = LBound(locations_row) To UBound(locations_row)    '変数locations_rowの0～2まで
            db.SetInit hdr, data    'Search #Init
            Set r = db.GetCurser_r("圃場内位置", locations_col(col))
            Set r = db.GetCurser_r("圃場内位置2", locations_row(row))
            Set r = r.Resize(, 60 + 12)
            r.Select
            r.Copy Sheet3.Range("A3").Offset(destination_idx)
            destination_idx = destination_idx + 48    'A1に該当するデータをsheet3のA3から48行づつずらした範囲に転記する

        Next row
       
    Next col
               

End Sub
    
