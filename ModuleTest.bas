Attribute VB_Name = "ModuleTest"
Option Compare Database
Option Explicit

Sub エクスポート()

    Dim mdb As New MDBManipulator
    mdb.ExportForAcTable mdb.GetOwnFolderPath & "\こっち.mdb", "T_業績考課_振込件数店別_法業一"
    mdb.ExportForAcTableAndNameEdit mdb.GetOwnFolderPath & "\こっち.mdb", "T_業績考課_振込件数店別_法業一", "名前変更後"
    
End Sub
