VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crLaporanTagihanPenjamin 
   ClientHeight    =   9900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15480
   OleObjectBlob   =   "crLaporanTagihanPenjamin.dsx":0000
End
Attribute VB_Name = "crLaporanTagihanPenjamin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Section11_Format(ByVal pFormattingInfo As Object)
    'Dim X As Double
    'X = Round(ucTotalTagihan.Value)
    'txtTerbilang.SetText "# " & TERBILANG(X) & " #"
End Sub

Private Sub Section2_Format(ByVal pFormattingInfo As Object)
    Dim X As Double
    X = Round(ucTotalTagihan.Value)
    txtPembulatan.SetText Format(X, "##,##0.00")
    txtTerbilang.SetText "# " & TERBILANG(X) & " #"
End Sub
