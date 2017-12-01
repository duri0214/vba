VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_プログレスメッセージ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub indicatorBtn_Click()

    Dim rsCount As Long
    Dim i, j As Long
    
    Const DBG_MAXVALUE = 10000
    
    DoCmd.Hourglass True
        
    rsCount = DBG_MAXVALUE
    
    For i = 1 To rsCount
        
        textBox.value = i & "件目 処理実行中です[" & rsCount & "件]"
        DoEvents
        
        '★レコードを編集する処理
'        For j = 1 To 1000
'        Next j

    Next i
    
    DoCmd.Hourglass False

    textBox.value = ""
    MsgBox "処理終了"
        
End Sub
