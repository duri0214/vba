VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_リモコン"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub DisplayedInfo(txt As String)
    
    Form_リモコン.INFO_TXTBOX.value = txt
    DoEvents

End Sub

Private Function GetStMonth() As String
    
    Dim retValue As String
    retValue = Nz(Form_リモコン.ST_MONTH_TXTBOX.value)
    
    GetStMonth = retValue
    
End Function

Private Function GetEnMonth() As String
    
    Dim retValue As String
    retValue = Nz(Form_リモコン.EN_MONTH_TXTBOX.value)
    
    GetEnMonth = retValue
    
End Function

Private Sub startBtn_Click()
    
    Dim ym As New UtilYM
    
    Dim stMonth As String
    Dim enMonth As String
    Dim boundFor As Integer
    
    stMonth = GetStMonth()
    enMonth = GetEnMonth()
    
    boundFor = ym.GetYMInterval(stMonth, enMonth)       '期間係数確定
    If MsgBox(stMonth & "～" & enMonth & "まで処理します。よろしいですか？", vbYesNo, "処理確認") = vbNo Then
        MsgBox "処理は中断しました。"
        End
    End If
    
    DisplayedInfo "hoge作成中..."
    
    'Call hoge(stMonth, boundFor)
        
    DisplayedInfo "処理終了"
    Application.RefreshDatabaseWindow
    MsgBox ("処理終了")
                
End Sub
