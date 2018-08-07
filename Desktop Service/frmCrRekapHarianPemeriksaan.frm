VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCrRekapHarianPemeriksaan 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmCrRekapHarianPemeriksaan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   5820
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOption 
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdCetak 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cboPrinter 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   -1  'True
   End
   Begin VB.TextBox txtNamaFormPengirim 
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmCrRekapHarianPemeriksaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crRekapPemeriksaanRehabMedik
Dim Reports As New crRekapPemeriksaanRehabMedikInap
Dim Reportobat As New crRekapPendapatanObatRehabMedik
'Dim bolSuppresDetailSection10 As Boolean
'Dim ii As Integer
'Dim tempPrint1 As String
'Dim p As Printer
'Dim p2 As Printer
Dim strDeviceName As String
Dim strDriverName As String
Dim strPort As String

Private Sub cmdCetak_Click()
    Report.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
    'PrinterNama = cboPrinter.Text
    Report.PrintOut False
    Reports.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
    Reportobat.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
    'PrinterNama = cboPrinter.Text
    Reports.PrintOut False
    Reportobat.PrintOut False
End Sub

Private Sub CmdOption_Click()
    Report.PrinterSetup Me.hWnd
    Reports.PrinterSetup Me.hWnd
    Reportobat.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub Form_Load()
    Dim p As Printer
    cboPrinter.Clear
    For Each p In Printers
        cboPrinter.AddItem p.DeviceName
    Next
    cboPrinter.Text = GetTxt("Setting.ini", "Printer", "LaporanPenerimaan")
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCRLaporanPenerimaan = Nothing
End Sub

Public Sub Cetak(idKasir As String, tglAwal As String, idDepartemen As String, namaKasir As String, view As String)
'On Error GoTo errLoad
'On Error Resume Next

Set frmCrRekapHarianPemeriksaan = Nothing
Dim adocmd As New ADODB.Command
    Dim Tgl As String
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim Tanggal As Date
    
    

    If idDepartemen <> "" Then
        If idDepartemen = 18 Then
            str3 = " and ru.objectdepartemenfk in (18,28) and ru2.objectdepartemenfk in (28)"
        ElseIf idDepartemen = 16 Then
            str3 = " and ru.objectdepartemenfk in (16) and ru2.objectdepartemenfk in (28)"
        End If
    End If
    
    Tgl = Format(tglAwal, "yyyy-MM-dd")
    str1 = Format(tglAwal, "yyyy-MM-01 00:00")
    
    ReadRs2 "SELECT (date_trunc('month', tanggal::date) + interval '1 month' - interval '1 day')::date ||' 23:59' " & _
            "AS end_of_month from kalender_s where tanggal= '" & Tgl & "'"
            
    If (RS2.EOF = False) Then
        str2 = (RS2!end_of_month)
    End If
     
    
Set Report = New crRekapPemeriksaanRehabMedik
    strSQL = "SELECT " & _
                " pp.tglpelayanan,dp.namadepartemen,kps.kelompokpasien,pr.id,pr.namaproduk, " & _
                "pp.hargajual, pp.jumlah, pp.hargajual*pp.jumlah as subtotal " & _
                "from pasiendaftar_t as pd " & _
                "left join antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
                "left join pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
                "left join ruangan_m as ru on ru.id=pd.objectruanganlastfk " & _
                "left join departemen_m dp on dp.id=ru.objectdepartemenfk " & _
                "left join ruangan_m ru2 on ru2.id=apd.objectruanganfk " & _
                "left join departemen_m dp2 on dp2.id=ru2.objectdepartemenfk " & _
                "left join produk_m as pr on pr.id=pp.produkfk left join detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
                "left join kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
                "left join strukpelayanan_t as sp on sp.norec=pp.strukfk " & _
                " Where " & _
                " pp.tglpelayanan BETWEEN '" & str1 & "' and '" & str2 & "' and sp.statusenabled is null and pr.objectdepartemenfk=28 and pr.id <> 395 " & _
                str3


   ReadRs3 "SELECT " & _
                " pp.tglpelayanan,dp.namadepartemen,kps.kelompokpasien,pr.id,pr.namaproduk, " & _
                " case when kps.id in (2,4) then pp.hargajual else 0 end as tarif, " & _
                " case when kps.id in (2,4) then pp.jumlah else 0 end as qtybpjs, " & _
                " case when kps.id in (2,4) then pp.jumlah * pp.hargajual else 0 end as totalbpjs, " & _
                " case when kps.id in (1,3,5) then pp.jumlah else 0 end as qtynonbpjs, " & _
                " case when kps.id in (1,3,5) then pp.jumlah * pp.hargajual else 0 end as totalnonbpjs " & _
                "from pasiendaftar_t as pd " & _
                "left join antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
                "left join pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
                "left join ruangan_m as ru on ru.id=pd.objectruanganlastfk " & _
                "left join departemen_m dp on dp.id=ru.objectdepartemenfk " & _
                "left join ruangan_m ru2 on ru2.id=apd.objectruanganfk " & _
                "left join departemen_m dp2 on dp2.id=ru2.objectdepartemenfk " & _
                "left join produk_m as pr on pr.id=pp.produkfk left join detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
                "left join kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
                "left join strukpelayanan_t as sp on sp.norec=pp.strukfk " & _
                " Where " & _
                " pp.tglpelayanan BETWEEN '" & str1 & "' and '" & str2 & "' and sp.statusenabled is null and pr.id = 395" & _
                str3
    
    Dim tqtybpjs, ttbpjs, ttrf, tqtynonbpjs, ttnonbpjs As Double
    Dim i As Integer
    
    tqtybpjs = 0
    ttbpjs = 0
    tqtynonbpjs = 0
    ttnonbpjs = 0
    
    If (RS3.EOF = False) Then
        ttrf = CDbl(IIf(IsNull(RS3!tarif), 0, RS3!tarif))
    End If

    For i = 0 To RS3.RecordCount - 1
        tqtybpjs = tqtybpjs + CDbl(IIf(IsNull(RS3!qtybpjs), 0, RS3!qtybpjs))
        ttbpjs = ttbpjs + CDbl(IIf(IsNull(RS3!totalbpjs), 0, RS3!totalbpjs))
        tqtynonbpjs = tqtynonbpjs + CDbl(IIf(IsNull(RS3!qtynonbpjs), 0, RS3!qtynonbpjs))
        ttnonbpjs = ttnonbpjs + CDbl(IIf(IsNull(RS3!totalnonbpjs), 0, RS3!totalnonbpjs))
        RS3.MoveNext
        
    Next

            
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
            If idDepartemen = 18 Then
                .txtDepartemen.SetText "Rawat Jalan"
            ElseIf idDepartemen = 16 Then
                .txtDepartemen.SetText "Rawat Inap"
            End If
            .txtNamaKasir.SetText namaKasir
            .txtPeriode.SetText "Bulan " & Format(tglAwal, "MMMM")
            .usJenisTindakan.SetUnboundFieldSource ("{ado.namaproduk}")
            .usKelompokPasien.SetUnboundFieldSource ("{ado.kelompokpasien}")
            .unQty.SetUnboundFieldSource ("{ado.jumlah}")
            .unTarif.SetUnboundFieldSource ("{ado.hargajual}")
            .unTotal.SetUnboundFieldSource ("{ado.subtotal}")
            .unQtyBPJS.SetUnboundFieldSource tqtybpjs
            .unTarifBPJS.SetUnboundFieldSource ttrf
            .unTotalBPJS.SetUnboundFieldSource ttbpjs
            .unQtyNonBPJS.SetUnboundFieldSource tqtynonbpjs
            .unTarifnNonBPJS.SetUnboundFieldSource ttrf
            .unTotalNonBPJS.SetUnboundFieldSource ttnonbpjs
            
            
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "LaporanPenerimaan")
                .SelectPrinter "winspool", strPrinter, "Ne00:"
                .PrintOut False
                Unload Me
            Else
                With CRViewer1
                    .ReportSource = Report
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
            End If
        'End If
    End With
Exit Sub
errLoad:
End Sub
Public Sub cetakinap(idKasir As String, tglAwal As String, idDepartemen As String, namaKasir As String, view As String)
'On Error GoTo errLoad
'On Error Resume Next

Set frmCrRekapHarianPemeriksaan = Nothing
Dim adocmd As New ADODB.Command
    Dim Tgl As String
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim Tanggal As Date
    
    

    If idDepartemen <> "" Then
        If idDepartemen = 18 Then
            str3 = " and ru.objectdepartemenfk in (18,28) and ru2.objectdepartemenfk in (28)"
        ElseIf idDepartemen = 16 Then
            str3 = " and ru.objectdepartemenfk in (16) and ru2.objectdepartemenfk in (28)"
        End If
    End If
    
    Tgl = Format(tglAwal, "yyyy-MM-dd")
    str1 = Format(tglAwal, "yyyy-MM-01 00:00")
    
    ReadRs2 "SELECT (date_trunc('month', tanggal::date) + interval '1 month' - interval '1 day')::date ||' 23:59' " & _
            "AS end_of_month from kalender_s where tanggal= '" & Tgl & "'"
            
    If (RS2.EOF = False) Then
        str2 = (RS2!end_of_month)
    End If
     
    
Set Reports = New crRekapPemeriksaanRehabMedikInap
'    strSQL = "SELECT " & _
'                " pp.tglpelayanan,dp.namadepartemen,kps.kelompokpasien,pr.id,pr.namaproduk, " & _
'                "pp.hargajual, pp.jumlah, pp.hargajual*pp.jumlah as subtotal " & _
'                "from pasiendaftar_t as pd " & _
'                "left join antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
'                "left join pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
'                "left join ruangan_m as ru on ru.id=pd.objectruanganasalfk " & _
'                "left join departemen_m dp on dp.id=ru.objectdepartemenfk " & _
'                "left join ruangan_m ru2 on ru2.id=apd.objectruanganfk " & _
'                "left join departemen_m dp2 on dp2.id=ru2.objectdepartemenfk " & _
'                "left join produk_m as pr on pr.id=pp.produkfk left join detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
'                "left join kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
'                "left join strukpelayanan_t as sp on sp.norec=pp.strukfk " & _
'                " Where " & _
'                " pp.tglpelayanan BETWEEN '" & str1 & "' and '" & str2 & "' and sp.statusenabled is null and pr.objectdepartemenfk=28 and pr.id <> 395 " & _
'                str3
        strSQL = "SELECT " & _
                " pp.tglpelayanan,dp.namadepartemen,kps.kelompokpasien,pr.id,pr.namaproduk, " & _
                "pp.hargajual, pp.jumlah, pp.hargajual*pp.jumlah as subtotal " & _
                "from pasiendaftar_t as pd " & _
                "left join antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
                "left join pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
                "left join ruangan_m as ru on ru.id=pd.objectruanganlastfk " & _
                "left join departemen_m dp on dp.id=ru.objectdepartemenfk " & _
                "left join ruangan_m ru2 on ru2.id=apd.objectruanganfk " & _
                "left join departemen_m dp2 on dp2.id=ru2.objectdepartemenfk " & _
                "left join produk_m as pr on pr.id=pp.produkfk left join detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
                "left join kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
                "left join strukpelayanan_t as sp on sp.norec=pp.strukfk " & _
                " Where " & _
                " pp.tglpelayanan BETWEEN '" & str1 & "' and '" & str2 & "' and sp.statusenabled is null and pr.objectdepartemenfk=28 and pr.id <> 395 " & _
                str3

            
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Reports
        .database.AddADOCommand CN_String, adocmd
            If idDepartemen = 18 Then
                .txtDepartemen.SetText "Rawat Jalan"
            ElseIf idDepartemen = 16 Then
                .txtDepartemen.SetText "Rawat Inap"
            End If
            .txtNamaKasir.SetText namaKasir
            .txtPeriode.SetText "Bulan " & Format(tglAwal, "MMMM")
            .usJenisTindakan.SetUnboundFieldSource ("{ado.namaproduk}")
            .usKelompokPasien.SetUnboundFieldSource ("{ado.kelompokpasien}")
            .unQty.SetUnboundFieldSource ("{ado.jumlah}")
            .unTarif.SetUnboundFieldSource ("{ado.hargajual}")
            .unTotal.SetUnboundFieldSource ("{ado.subtotal}")
            
            
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "LaporanPenerimaan")
                .SelectPrinter "winspool", strPrinter, "Ne00:"
                .PrintOut False
                Unload Me
            Else
                With CRViewer1
                    .ReportSource = Reports
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
            End If
        'End If
    End With
Exit Sub
errLoad:
End Sub
Public Sub cetakobat(idKasir As String, tglAwal As String, namaKasir As String, view As String)
'On Error GoTo errLoad
'On Error Resume Next

Set frmCrRekapHarianPemeriksaan = Nothing
Dim adocmd As New ADODB.Command
    Dim Tgl As String
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim Tanggal As Date
    
    
    
    Tgl = Format(tglAwal, "yyyy-MM-dd")
    str1 = Format(tglAwal, "yyyy-MM-01 00:00")
    
    ReadRs2 "SELECT (date_trunc('month', tanggal::date) + interval '1 month' - interval '1 day')::date ||' 23:59' " & _
            "AS end_of_month from kalender_s where tanggal= '" & Tgl & "'"
            
    If (RS2.EOF = False) Then
        str2 = (RS2!end_of_month)
    End If
     
    
Set Reportobat = New crRekapPendapatanObatRehabMedik
    strSQL = "SELECT " & _
                " pp.tglpelayanan,dp.namadepartemen,pr.id,pr.namaproduk, " & _
                "pp.hargajual, pp.jumlah, pp.hargajual*pp.jumlah as subtotal " & _
                "from pasiendaftar_t as pd " & _
                "left join antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
                "left join pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
                "left join ruangan_m as ru on ru.id=pd.objectruanganasalfk " & _
                "left join departemen_m dp on dp.id=ru.objectdepartemenfk " & _
                "left join ruangan_m ru2 on ru2.id=apd.objectruanganfk " & _
                "left join departemen_m dp2 on dp2.id=ru2.objectdepartemenfk " & _
                "left join produk_m as pr on pr.id=pp.produkfk left join detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
                "left join strukpelayanan_t as sp on sp.norec=pp.strukfk " & _
                " Where " & _
                " pp.tglpelayanan BETWEEN '" & str1 & "' and '" & str2 & "' and sp.statusenabled is null and djp.objectjenisprodukfk=97 and dp2.id=28 "

            
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Reportobat
        .database.AddADOCommand CN_String, adocmd
            .txtNamaKasir.SetText namaKasir
            .txtPeriode.SetText "Bulan " & Format(tglAwal, "MMMM")
            .usJenisTindakan.SetUnboundFieldSource ("{ado.namaproduk}")
            .unQty.SetUnboundFieldSource ("{ado.jumlah}")
            .unTarif.SetUnboundFieldSource ("{ado.hargajual}")
            .unTotal.SetUnboundFieldSource ("{ado.subtotal}")
            
            
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "LaporanPenerimaan")
                .SelectPrinter "winspool", strPrinter, "Ne00:"
                .PrintOut False
                Unload Me
            Else
                With CRViewer1
                    .ReportSource = Reportobat
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
            End If
        'End If
    End With
Exit Sub
errLoad:
End Sub


