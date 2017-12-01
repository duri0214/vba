VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_プログレスバー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub TreeView3_Updated(Code As Integer)

End Sub

Private Sub コマンド1_Click()

    Me.pb.Max = 100000   'プログレスバーの最大値
    Me.pb.min = 0    'プログレスバーの最小値

    pb.value = 0
    For i = 0 To 100000
        pb.value = i
    Next
    
    pb.value = 0
    MsgBox "処理終了"
    
End Sub
