Attribute VB_Name = "クラスにしにくいやつら"
Option Compare Database
Option Explicit

'▲notice
'クラスにしにくいやっちゃなー・・・utilクラスでも作るか？

Sub テーブル集約検査君()
    
    Const TBL_PREFIX As String = "W_業績評価_"
    
    Dim ym As New UtilYM
    Dim stMonth As String
    Dim enMonth As String
    Dim crrYM As String
    Dim boundFor As Integer
    
    Dim sql As String
    Dim editSql As String
    
    Dim i As Integer
    
    sql = "SELECT " & _
            "Sum({TABLE_NAME}{YYYYMM}.国内決済収益{YYYYMM}) AS 国内決済収益{YYYYMM}の合計, " & _
            "Sum({TABLE_NAME}{YYYYMM}.為替手数料{YYYYMM}) AS 為替手数料{YYYYMM}の合計, " & _
            "Sum({TABLE_NAME}{YYYYMM}.銀行間受入手数料{YYYYMM}) AS 銀行間受入手数料{YYYYMM}の合計, " & _
            "Sum({TABLE_NAME}{YYYYMM}.銀行間支払手数料{YYYYMM}) AS 銀行間支払手数料{YYYYMM}の合計, " & _
            "Sum({TABLE_NAME}{YYYYMM}.ＥＢ関係手数料{YYYYMM}) AS ＥＢ関係手数料{YYYYMM}の合計, " & _
            "Sum({TABLE_NAME}{YYYYMM}.口座振替手数料{YYYYMM}) AS 口座振替手数料{YYYYMM}の合計 " & _
            "FROM {TABLE_NAME}{YYYYMM}; "
    
    stMonth = ym.GetPeriod_Prev_RetFirstYM(2)                       '開始月確定（2期前の期初月次）
    enMonth = ym.GetAddYM(-1)                                       '終了月確定（当月）
    boundFor = ym.GetYMInterval(stMonth, enMonth)
            
    For i = 0 To boundFor
    
        crrYM = ym.GetAddYM2(stMonth, i)
        editSql = Replace(sql, "{TABLE_NAME}", TBL_PREFIX)
        editSql = Replace(editSql, "{YYYYMM}", crrYM)
        
        SummaryResultOutput editSql
    Next i
    
End Sub

Sub SummaryResultOutput(editSql As String)
    
    'editSqlに投入されるクエリは集約クエリであること
    '出力ファイルは自パス直下にsummary_result.csvとして保存される
    '拡張力はないが、まぁもともとmodel的ロジックなんで
    
    Dim mdb As New MDBManipulator
    Dim rs As New ADODB.Recordset                                       'レコードセット

    Dim report As New LogWritter                                        '出力ファイル

    Dim msg As String
    Dim col As Variant

    report.OpenTextStream mdb.GetOwnFolderPath & "\summary_result.csv"  '出力先ファイルのオープン
    rs.ActiveConnection = CurrentProject.Connection                     'RSコネクション確立
    rs.Open (editSql)                                                   'RSオープン

    msg = vbNullString
    For Each col In rs.fields                                           'ヘッダ出力内容をくるくる集める
        msg = msg & col.Name & ","
    Next col
    msg = Left(msg, Len(msg) - 1)                                       '最後尾のカンマkill
    report.WriteLine (msg)                                              'CSV出力

    msg = vbNullString
    For Each col In rs.fields                                           'データ出力内容をくるくる集める
        msg = msg & Nz(col.value, 0) & ","
    Next col
    msg = Left(msg, Len(msg) - 1)                                       '最後尾のカンマkill
    report.WriteLine (msg)                                              'CSV出力

    rs.Close                                                            'RSクローズ
    report.CloseTextStream                                              '出力先ファイルのクローズ
    
End Sub

