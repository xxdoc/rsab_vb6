VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crRincianBiayaPelayanan 
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16035
   OleObjectBlob   =   "crRincianBiayaPelayanan.dsx":0000
End
Attribute VB_Name = "crRincianBiayaPelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Section11_Format(ByVal pFormattingInfo As Object)
'kata bo rose tgl 26oktober 2017, pasien tidak mendapatkan billing
'sehingga terbilang ke tagihan perusahaan

    Dim X As Double
    X = Round(ucJumlahBill.Value)
    txtPembulatan.SetText Format(X, "##,##0.00")
    
    X = Round(CDbl(ucDitanggungPerusahaan.Value))
    a.SetText Format(X, "##,##0.00")
    X = Round(CDbl(ucDitanggungRS.Value))
    b.SetText Format(X, "##,##0.00")
    X = Round(CDbl(ucDitanggungSendiri.Value))
    c.SetText Format(X, "##,##0.00")
    X = Round(CDbl(ucSurplusMinusRS.Value))
    d.SetText Format(X, "##,##0.00")
    
    If usTipe.Value = "Umum/Pribadi" Then
        txtTerbilang.SetText "# " & TERBILANG(txtPembulatan.Text) & " #"
    Else
        txtTerbilang.SetText "# " & TERBILANG(ucDitanggungPerusahaan.Value) & " #"
    End If

'    ucJumlahBill.Value = Replace(txtPembulatan.Text, ".", ",")
'    txtTerbilang.SetText "# " & TerbilangDesimal(Replace(txtPembulatan.Text, ".", ",")) & " #"
    
End Sub
