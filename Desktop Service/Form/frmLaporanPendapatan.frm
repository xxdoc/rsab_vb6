VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCRLaporanPendapatan 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmLaporanPendapatan.frx":0000
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
Attribute VB_Name = "frmCRLaporanPendapatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crLaporanPendapatan
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
End Sub

Private Sub CmdOption_Click()
    Report.PrinterSetup Me.hWnd
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

    Set frmCRLaporanPendapatan = Nothing
End Sub

Public Sub CetakLaporanPendapatan(idKasir As String, tglAwal As String, tglAkhir As String, idRuangan As String, idDokter As String, namaPrinted As String, view As String)
'On Error GoTo errLoad
'On Error Resume Next

Set frmCRLaporanPendapatan = Nothing
Dim adocmd As New ADODB.Command

    Dim str1 As String
    Dim str2 As String
    
    If idDokter <> "" Then
        str1 = "and apd.objectpegawaifk=" & idDokter & " "
    End If
    If idRuangan <> "" Then
        str2 = " and apd.objectruanganfk=" & idRuangan & " "
    End If
    
Set Report = New crLaporanPendapatan
    strSQL = "select  apd.objectruanganfk,ru.namaruangan, apd.objectpegawaifk,pg.namalengkap,ps.nocm , " & _
             "upper(ps.namapasien) as namapasien, " & _
             "case when pr.id =395 then pp.hargajual* pp.jumlah else 0 end as karcis, " & _
             "case when pr.id =10013116  then pp.hargajual* pp.jumlah else 0 end as embos, " & _
             "case when kp.id = 26 then pp.hargajual* pp.jumlah else 0 end as konsul, " & _
             "case when kp.id in (1,2,3,4,8,9,10,11,13,14) then pp.hargajual* pp.jumlah else 0 end as tindakan, " & _
             "(case when pp.hargadiscount is null then 0 else pp.hargadiscount end)* pp.jumlah as diskon, " & _
             "pd.noregistrasi,kps.kelompokpasien, " & _
             "case when cb.id = 1 or cb.id is null then '-' else 'v' end as cc, case when pd.objectkelompokpasienlastfk = 1 then '-' else 'v' end as pj , cb.id, " & _
             "case when sp.norec is null then '-' else 'v' end as verif, " & _
             "case when sbm.norec is null then '-' else 'v' end as sbm " & _
             "from pasiendaftar_t as pd " & _
             "inner JOIN antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec left JOIN pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
             "left JOIN pegawai_m as pg on pg.id=apd.objectpegawaifk left JOIN ruangan_m as ru on ru.id=apd.objectruanganfk " & _
             "inner JOIN produk_m as pr on pr.id=pp.produkfk inner JOIN detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
             "inner JOIN jenisproduk_m as jp on jp.id=djp.objectjenisprodukfk inner JOIN kelompokproduk_m as kp on kp.id=jp.objectkelompokprodukfk " & _
             "inner JOIN pasien_m as ps on ps.id=pd.nocmfk left JOIN kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
             "left JOIN strukpelayanan_t as sp  on sp.noregistrasifk=pd.norec left JOIN strukbuktipenerimaan_t as sbm  on sbm.norec=sp.nosbmlastfk " & _
             "left JOIN strukbuktipenerimaancarabayar_t as sbmc  on sbmc.nosbmfk=sbm.norec left JOIN carabayar_m as cb  on cb.id=sbmc.objectcarabayarfk " & _
             "where pd.tglregistrasi between '" & tglAwal & " 00:00' and '" & tglAkhir & " 23:59' " & _
             str2 & _
             str1 & _
             "order by pd.noregistrasi"

   ReadRs2 "select " & _
           "sum(case when sbmc.objectcarabayarfk is not null and cb.id=1 then (pp.hargajual-(case when pp.hargadiscount is null then 0 else pp.hargadiscount end ))*pp.jumlah else 0 end) as cash, " & _
           "sum(case when sbmc.objectcarabayarfk is not null and cb.id>1 then (pp.hargajual-(case when pp.hargadiscount is null then 0 else pp.hargadiscount end ))*pp.jumlah else 0 end) as kk, " & _
           "sum(case when pd.objectkelompokpasienlastfk >1 then (pp.hargajual-(case when pp.hargadiscount is null then 0 else pp.hargadiscount end ))*pp.jumlah else 0 end) as jm " & _
           "from pasiendaftar_t pd " & _
           "INNER JOIN antrianpasiendiperiksa_t apd on apd.noregistrasifk=pd.norec " & _
           "INNER JOIN pelayananpasien_t pp on pp.noregistrasifk=apd.norec " & _
           "INNER JOIN pasien_m ps on ps.id=pd.nocmfk " & _
           "left JOIN strukpelayanan_t sp on sp.noregistrasifk=pd.norec " & _
           "left JOIN strukbuktipenerimaan_t sbm on sbm.norec=sp.nosbmlastfk " & _
           "left JOIN strukbuktipenerimaancarabayar_t sbmc on sbm.norec=sbmc.nosbmfk " & _
           "left JOIN carabayar_m cb on cb.id=sbmc.objectcarabayarfk " & _
           "left JOIN ruangan_m ru on ru.id=pd.objectruanganlastfk " & _
           "left JOIN pegawai_m pg on pg.id=apd.objectpegawaifk " & _
             "where pd.tglregistrasi between '" & tglAwal & " 00:00' and '" & tglAkhir & " 23:59' " & _
             "" & str1 & " " & str2
    ReadRs3 "select sum((ppd.hargajual-(case when ppd.hargadiscount is null then 0 else ppd.hargadiscount end))*ppd.jumlah) as total " & _
            "from pasiendaftar_t pd " & _
            "INNER JOIN antrianpasiendiperiksa_t apd on apd.noregistrasifk=pd.norec " & _
            "INNER JOIN pelayananpasien_t pp on pp.noregistrasifk=apd.norec " & _
            "INNER JOIN pelayananpasiendetail_t ppd on ppd.pelayananpasien=pp.norec " & _
             "where pd.tglregistrasi between '" & tglAwal & " 00:00' and '" & tglAkhir & " 23:59' and ppd.komponenhargafk=35 " & _
             "" & str1 & " " & str2
    ReadRs4 "select sum((ppd.hargajual-(case when ppd.hargadiscount is null then 0 else ppd.hargadiscount end))*ppd.jumlah) as total " & _
            "from pasiendaftar_t pd " & _
            "INNER JOIN antrianpasiendiperiksa_t apd on apd.noregistrasifk=pd.norec " & _
            "INNER JOIN pelayananpasien_t pp on pp.noregistrasifk=apd.norec " & _
            "INNER JOIN pelayananpasiendetail_t ppd on ppd.pelayananpasien=pp.norec " & _
             "where pd.tglregistrasi between '" & tglAwal & " 00:00' and '" & tglAkhir & " 23:59' and ppd.komponenhargafk=25 " & _
             "" & str1 & " " & str2
    Dim tCash, tKk, tPj, tJm, tPm As Double
    Dim i As Integer
    
    tCash = RS2!cash
    tKk = IIf(IsNull(RS2!kk), 0, RS2!kk)
    tPj = IIf(IsNull(RS2!jm), 0, RS2!jm)
    tJm = IIf(IsNull(RS3!total), 0, RS3!total)
    tPm = IIf(IsNull(RS4!total), 0, RS4!total)
    
    Dim tAdmCc, tB3, tBPajak, tB5 As Double
    
    tAdmCc = (tKk * 3) / 100
    tB3 = tJm
    tBPajak = (tJm * 7.5) / 100
    tB5 = tB3 - tBPajak
    
    Dim tC3, tCPajak, tC5, tC7 As Double
    
    tC3 = tPm
    tCPajak = (tPm * 7.5) / 100
    tC5 = tC3 - tCPajak
    tC7 = tC5
            
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
            .txtNamaKasir.SetText namaPrinted
            .txtPeriode.SetText "Periode : " & tglAwal & " 00:00 s/d " & tglAkhir & " 23:59' "
'            .usNamaKasir.SetUnboundFieldSource ("{ado.kasir}")
            .usNamaRuangan.SetUnboundFieldSource ("{ado.namaruangan}")
            .usNamaDokter.SetUnboundFieldSource ("{ado.namalengkap}")
            .usNoCM.SetUnboundFieldSource ("{ado.nocm}")
            .usNoReg.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .ucKarcis.SetUnboundFieldSource ("{ado.karcis}")
            .ucEmbos.SetUnboundFieldSource ("{ado.embos}")
            .ucKonsul.SetUnboundFieldSource ("{ado.konsul}")
            .ucTindakan.SetUnboundFieldSource ("{ado.tindakan}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
'            .ucTotal.SetUnboundFieldSource ("{ado.kasir}")
            .usCC.SetUnboundFieldSource ("{ado.cc}")
            .usPJ.SetUnboundFieldSource ("{ado.pj}")
            .usKelompokPasien.SetUnboundFieldSource ("{ado.kelompokpasien}")
            .usVR.SetUnboundFieldSource ("{ado.verif}")
            .usSBM.SetUnboundFieldSource ("{ado.sbm}")
            
            .txtA1.SetText Format(tCash, "##,##0.00")
            .txtA2.SetText Format(tKk, "##,##0.00")
            .txtA3.SetText Format(tAdmCc, "##,##0.00")
            .txtA4.SetText Format(tPj, "##,##0.00")
            
            .txtB1.SetText Format(tJm, "##,##0.00")
            .txtB2.SetText Format(0, "##,##0.00")
            .txtB3.SetText Format(tB3, "##,##0.00")
            .txtB4.SetText Format(tBPajak, "##,##0.00")
            .txtB5.SetText Format(tB5, "##,##0.00")
            
            .txtC1.SetText Format(tPm, "##,##0.00")
            .txtC2.SetText Format(0, "##,##0.00")
            .txtC3.SetText Format(tC3, "##,##0.00")
            .txtC4.SetText Format(tCPajak, "##,##0.00")
            .txtC5.SetText Format(tC5, "##,##0.00")
            .txtC6.SetText Format(0, "##,##0.00")
            .txtC7.SetText Format(tC7, "##,##0.00")
            
            
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
