VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crKuitansiPiutangPenjamin 
   ClientHeight    =   9780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15345
   OleObjectBlob   =   "crKuitansiPiutangPenjamin.dsx":0000
End
Attribute VB_Name = "crKuitansiPiutangPenjamin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Section4_Format(ByVal pFormattingInfo As Object)
    'Dim X As Double
  '  X = Round(ucTotalPenjamin.Value)
   ' txtTerbilang.SetText "# " & TERBILANG(X) & " #"
    
    Dim X As Double
    X = Round(ucJumlah.Value + ucMaterai.Value)
    txtPembulatan.SetText Format(X, "##,##0.00")
    txtTerbilang.SetText "# " & TERBILANG(X) & " #"
End Sub
