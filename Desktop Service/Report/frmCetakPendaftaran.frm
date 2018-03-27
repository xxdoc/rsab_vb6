VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakPendaftaran 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "frmCetakPendaftaran.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   9075
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
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9045
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
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
Attribute VB_Name = "frmCetakPendaftaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New Cr_cetakBuktiPendaftaran
Dim ReportTracer As New Cr_cetakLabelTracer
Dim reportSep As New crCetakSJP
Dim reportBuktiLayanan As New Cr_cetakbuktilayanan
Dim reportBuktiLayananRuangan As New Cr_cetakbuktilayananruangan
'Dim reportLabel As New Cr_cetakLabel 'LAMA
Dim reportLabel As New Cr_cetakLabel_2
Dim reportLabelZebra As New Cr_cetakLabelZebra
Dim reportSumList As New Cr_cetakSummaryList
Dim reportRmk As New Cr_cetakRMK
Dim reportLembarGC As New Cr_cetakLembarGC
Dim reportBuktiLayananRuanganPerTindakan As New Cr_cetakbuktilayananruanganpertindakan
Dim reportBuktiLayananJasa As New Cr_cetakbuktilayananruanganpertindakanJasa
Dim reportBuktiLayananRuanganBedah As New Cr_cetakbuktilayananruanganbedah

'Private fso As New Scripting.FileSystemObject
Dim reportKartuPasien As New Cr_cetakKartuPasien
'Dim WithEvents sect As CRAXDRT.Section


Dim ii As Integer
Dim tempPrint1 As String
Dim p As Printer
Dim p2 As Printer
Dim strDeviceName As String
Dim strDriverName As String
Dim strPort As String

Dim bolBuktiPendaftaran As Boolean
Dim bolBuktiLayanan  As Boolean
Dim bolBuktiLayananRuangan  As Boolean
Dim bolBuktiLayananRuanganPerTindakan  As Boolean
Dim bolBuktiLayananJasa  As Boolean
Dim bolcetakSep  As Boolean
Dim bolTracer1  As Boolean
Dim bolKartuPasien  As Boolean
Dim boolLabelPasien  As Boolean
Dim boolLabelPasienZebra  As Boolean
Dim boolSumList  As Boolean
Dim boolLembarRMK As Boolean
Dim boolLembarPersetujuan As Boolean
Dim boolBuktiLayananJasa As Boolean
Dim bolBuktiLayananRuanganBedah  As Boolean


Dim strPrinter As String
Dim strPrinter1 As String
Dim PrinterNama As String

Dim adoReport As New ADODB.Command

Private Sub cmdCetak_Click()
  If cboPrinter.Text = "" Then MsgBox "Printer belum dipilih", vbInformation, ".: Information": Exit Sub
    If bolBuktiPendaftaran = True Then
        Report.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        Report.PrintOut False
    ElseIf bolBuktiLayanan = True Then
        reportBuktiLayanan.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportBuktiLayanan.PrintOut False
    ElseIf bolBuktiLayananRuangan = True Then
        reportBuktiLayananRuangan.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportBuktiLayananRuangan.PrintOut False
    ElseIf bolBuktiLayananRuanganPerTindakan = True Then
        reportBuktiLayananRuanganPerTindakan.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportBuktiLayananRuanganPerTindakan.PrintOut False
    ElseIf bolcetakSep = True Then
        reportSep.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportSep.PrintOut False
    ElseIf bolTracer1 = True Then
        ReportTracer.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        ReportTracer.PrintOut False
    ElseIf bolKartuPasien = True Then
        reportKartuPasien.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportKartuPasien.PrintOut False
    ElseIf boolLabelPasien = True Then
        reportLabel.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportLabel.PrintOut False
    ElseIf boolLabelPasienZebra = True Then
        reportLabelZebra.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportLabelZebra.PrintOut False
    ElseIf boolSumList = True Then
        reportSumList.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportSumList.PrintOut False
    ElseIf boolLembarRMK = True Then
        reportRmk.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportRmk.PrintOut False
    ElseIf boolBuktiLayananJasa = True Then
        reportBuktiLayananJasa.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportBuktiLayananJasa.PrintOut False
    ElseIf boolLembarPersetujuan = True Then
        reportLembarGC.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportLembarGC.PrintOut False
    ElseIf bolBuktiLayananRuanganBedah = True Then
        reportBuktiLayananRuanganBedah.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        reportBuktiLayananRuanganBedah.PrintOut False
    End If
End Sub

Private Sub CmdOption_Click()
    
    If bolBuktiPendaftaran = True Then
        Report.PrinterSetup Me.hWnd
    ElseIf bolBuktiLayanan = True Then
        reportBuktiLayanan.PrinterSetup Me.hWnd
    ElseIf bolBuktiLayananRuangan = True Then
        reportBuktiLayananRuangan.PrinterSetup Me.hWnd
    ElseIf bolBuktiLayananRuanganPerTindakan = True Then
        reportBuktiLayananRuanganPerTindakan.PrinterSetup Me.hWnd
    ElseIf bolBuktiLayananJasa = True Then
        reportBuktiLayananJasa.PrinterSetup Me.hWnd
    ElseIf bolcetakSep = True Then
        reportSep.PrinterSetup Me.hWnd
    ElseIf bolTracer1 = True Then
        ReportTracer.PrinterSetup Me.hWnd
    ElseIf bolKartuPasien = True Then
        reportKartuPasien.PrinterSetup Me.hWnd
    ElseIf boolLabelPasien = True Then
         reportLabel.PrinterSetup Me.hWnd
    ElseIf boolLabelPasienZebra = True Then
         reportLabelZebra.PrinterSetup Me.hWnd
    ElseIf boolSumList = True Then
         reportSumList.PrinterSetup Me.hWnd
    ElseIf boolLembarRMK = True Then
         reportRmk.PrinterSetup Me.hWnd
    ElseIf boolBuktiLayananJasa = True Then
         reportBuktiLayananJasa.PrinterSetup Me.hWnd
    ElseIf boolLembarPersetujuan = True Then
         reportLembarGC.PrinterSetup Me.hWnd
    ElseIf bolBuktiLayananRuanganBedah = True Then
        reportBuktiLayananRuanganBedah.PrinterSetup Me.hWnd
    End If
    
    CRViewer1.Refresh
End Sub

Private Sub Form_Load()
    
    Dim p As Printer
    cboPrinter.Clear
    For Each p In Printers
        cboPrinter.AddItem p.DeviceName
    Next
    strPrinter = strPrinter1
    
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCetakPendaftaran = Nothing
'    fso.DeleteFile (App.Path & "\tempbitmap.bmp")
'    Set sect = Nothing

End Sub

Public Sub cetakBuktiPendaftaran(strNorec As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String

bolBuktiPendaftaran = True
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    With Report
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT pd.noregistrasi,ps.nocm,ps.tgllahir,ps.namapasien, " & _
                       " pd.tglregistrasi,jk.reportdisplay as jk,ap.alamatlengkap,ap.mobilephone2, " & _
                        " ru.namaruangan as ruanganPeriksa,pp.namalengkap as namadokter,kp.kelompokpasien, " & _
                        " apdp.noantrian From  pasiendaftar_t pd " & _
                        " INNER JOIN pasien_m ps ON pd.nocmfk = ps.id " & _
                        " INNER JOIN alamat_m ap ON ap.nocmfk = ps.id " & _
                        " INNER JOIN jeniskelamin_m jk ON ps.objectjeniskelaminfk = jk.id " & _
                        " INNER JOIN ruangan_m ru ON pd.objectruanganlastfk = ru.id " & _
                        " LEFT JOIN pegawai_m pp ON pd.objectpegawaifk = pp.id " & _
                        " INNER JOIN kelompokpasien_m kp ON pd.objectkelompokpasienlastfk = kp.id " & _
                        " INNER JOIN antrianpasiendiperiksa_t apdp ON apdp.noregistrasifk = pd.norec" & _
                        " where pd.noregistrasi ='" & strNorec & "' "
            
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            .database.AddADOCommand CN_String, adoReport
            .usnoantri.SetUnboundFieldSource ("{ado.noantrian}")
            .udtgl.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usnodft.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usnmpasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")
            .udTglLahir.SetUnboundFieldSource ("{ado.tgllahir}")
            .usAlamat.SetUnboundFieldSource ("{ado.alamatlengkap}")
            .usNoTelpon.SetUnboundFieldSource ("{ado.mobilephone2}")

            .usPenjamin.SetUnboundFieldSource ("{ado.kelompokpasien}")
            .usruangperiksa.SetUnboundFieldSource ("{ado.ruanganPeriksa}")
            .usNamaDokter.SetUnboundFieldSource ("{ado.namadokter}")

            If view = "false" Then
            '    Dim strPrinter As String

                strPrinter1 = GetTxt("Setting.ini", "Printer", "BuktiPendaftaran")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = Report
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:

    MsgBox Err.Number & " " & Err.Description
End Sub


Public Sub cetakTracer(strNorec As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String

bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = True
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    With ReportTracer
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT pd.noregistrasi,ps.nocm,ps.tgllahir,ps.namapasien, " & _
                       " pd.tglregistrasi,jk.reportdisplay as jk,ap.alamatlengkap,ap.mobilephone2, " & _
                        " ru.namaruangan as ruanganPeriksa,pp.namalengkap as namadokter,kp.kelompokpasien, " & _
                        " apdp.noantrian,pd.statuspasien,ps.namaayah  From  pasiendaftar_t pd " & _
                        " INNER JOIN pasien_m ps ON pd.nocmfk = ps.id " & _
                        " LEFT JOIN alamat_m ap ON ap.nocmfk = ps.id " & _
                        " INNER JOIN jeniskelamin_m jk ON ps.objectjeniskelaminfk = jk.id " & _
                        " INNER JOIN ruangan_m ru ON pd.objectruanganlastfk = ru.id " & _
                        " LEFT JOIN pegawai_m pp ON pd.objectpegawaifk = pp.id " & _
                        " INNER JOIN kelompokpasien_m kp ON pd.objectkelompokpasienlastfk = kp.id " & _
                        " INNER JOIN antrianpasiendiperiksa_t apdp ON apdp.noregistrasifk = pd.norec" & _
                        " where pd.noregistrasi ='" & strNorec & "' "
            
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            .database.AddADOCommand CN_String, adoReport
            .usnoantri.SetUnboundFieldSource ("{ado.noantrian}")
'            .udtgl.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usnodft.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usnmpasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")
            .usStatusPasien.SetUnboundFieldSource ("{ado.statuspasien}")
            .udTglReg.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usNamaDokter.SetUnboundFieldSource ("{ado.namadokter}")
            .usNamaKel.SetUnboundFieldSource ("{ado.namaayah}")
            .usruangperiksa.SetUnboundFieldSource ("{ado.ruanganPeriksa}")

            If view = "false" Then
               

                strPrinter1 = GetTxt("Setting.ini", "Printer", "Tracer1")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = ReportTracer
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:

    MsgBox Err.Number & " " & Err.Description
End Sub


Public Sub cetakSep(strNorec As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String

bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = True
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    With reportSep
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "select pa.nosep,pa.tanggalsep,pa.nokepesertaan , pi.nocm,pd.noregistrasi ,pa.norujukan,ap.namapeserta,pi.tgllahir,jk.jeniskelamin," & _
                       " rp.namaruangan,rp.kodeexternal as namapoliBpjs,pa.ppkrujukan, " & _
                       " (CASE WHEN rp.objectdepartemenfk=16 then 'Rawat Inap' else 'Rawat Jalan' END) as jenisrawat," & _
                       " dg.kddiagnosa, (case when dg.namadiagnosa is null then '-' else dg.namadiagnosa end) as namadiagnosa , " & _
                       " pi.nocm, ap.jenispeserta,ap.kdprovider,ap.nmprovider,kls.namakelas, pa.catatan from pemakaianasuransi_t pa " & _
                       " LEFT JOIN asuransipasien_m ap on pa.objectasuransipasienfk= ap.id " & _
                       " LEFT JOIN pasiendaftar_t pd on pd.norec=pa.noregistrasifk " & _
                       " LEFT JOIN pasien_m pi on pi.id=pd.nocmfk " & _
                       " LEFT JOIN jeniskelamin_m jk on jk.id=pi.objectjeniskelaminfk " & _
                       " LEFT JOIN ruangan_m rp on rp.id=pd.objectruanganlastfk " & _
                       " LEFT JOIN diagnosa_m dg on pa.diagnosisfk=dg.id" & _
                       " LEFT JOIN kelas_m kls on kls.id=ap.objectkelasdijaminfk " & _
                       " where pd.noregistrasi ='" & strNorec & "' "
            
            ReadRs strSQL
            
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            .database.AddADOCommand CN_String, adoReport

             If Not RS.EOF Then
              .txtnosjp.SetText IIf(IsNull(RS("nosep")), "-", RS("nosep")) 'RS("nosep")
              .txtTglSep.SetText Format(RS("tanggalsep"), "dd/MM/yyyy")
              .txtNomorKartuAskes.SetText IIf(IsNull(RS("nokepesertaan")), "-", RS("nokepesertaan"))
              .txtNamaPasien.SetText IIf(IsNull(RS("namapeserta")), "-", RS("namapeserta")) 'RS("namapeserta")
              .txtkelamin.SetText IIf(IsNull(RS("jeniskelamin")), "-", RS("jeniskelamin")) 'RS("jeniskelamin")
              .txtTanggalLahir.SetText IIf(IsNull(RS("tgllahir")), "-", Format(RS("tgllahir"), "dd/MM/yyyy")) 'Format(RS("tgllahir"), "dd/mm/yyyy")
              .txtTujuan.SetText RS("namapoliBpjs") & " / " & RS("namaruangan")
              .txtAsalRujukan.SetText IIf(IsNull(RS("nmprovider")), "-", RS("nmprovider"))
              .txtPeserta.SetText IIf(IsNull(RS("jenispeserta")), "-", RS("jenispeserta"))
              .txtJenisrawat.SetText IIf(IsNull(RS("jenisrawat")), "-", RS("jenisrawat")) 'RS("jenisrawat")
              .txtNoCM2.SetText IIf(IsNull(RS("nocm")), "-", RS("nocm")) 'RS("nocm")
              .txtDiagnosa.SetText IIf(IsNull(RS("namadiagnosa")), "-", RS("namadiagnosa")) 'RS("namadiagnosa")
              .txtKelasrawat.SetText IIf(IsNull(RS("namakelas")), "-", RS("namakelas")) 'RS("namakelas")
              .txtCatatan.SetText IIf(IsNull(RS("catatan")), "-", RS("catatan"))
              .txtNoCM2.SetText IIf(IsNull(RS("nocm")), "-", RS("nocm"))
              .txtNoPendaftaran2.SetText IIf(IsNull(RS("noregistrasi")), "-", RS("noregistrasi"))
             End If

            If view = "false" Then
               
                strPrinter1 = GetTxt("Setting.ini", "Printer", "CetakSep")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportSep
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:

    MsgBox Err.Number & " " & Err.Description
End Sub


Public Sub cetakBuktiLayanan(strNorec As String, strIdPegawai As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
Dim umur As String
    
bolBuktiPendaftaran = False
bolBuktiLayanan = True
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    With reportBuktiLayanan
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
'            strSQL = "SELECT pd.noregistrasi,ps.nocm,ps.tgllahir,ps.namapasien, " & _
'                       " pd.tglregistrasi,jk.reportdisplay AS jk,ru2.namaruangan AS ruanganperiksa,ru.namaruangan AS ruangakhir, " & _
'                       " (select pg.namalengkap from pegawai_m as pg INNER JOIN pelayananpasienpetugas_t p3 on p3.objectpegawaifk=pg.id " & _
'                       "where p3.pelayananpasien=tp.norec and p3.objectjenispetugaspefk=4 limit 1) AS namadokter,kp.kelompokpasien,tp.produkfk, " & _
'                       " pro.namaproduk,tp.jumlah,CASE WHEN tp.hargasatuan is null then tp.hargajual else tp.hargasatuan END as hargasatuan,ks.namakelas,ar.asalrujukan, " & _
'                       " CASE WHEN rek.namarekanan is null then '-' else rek.namarekanan END as namapenjamin,kmr.namakamar " & _
'                       " FROM pasiendaftar_t AS pd INNER JOIN pasien_m AS ps ON pd.nocmfk = ps.id " & _
'                       " INNER JOIN jeniskelamin_m AS jk ON ps.objectjeniskelaminfk = jk.id " & _
'                       " INNER JOIN ruangan_m AS ru ON pd.objectruanganlastfk = ru.id " & _
'                       " LEFT JOIN pegawai_m AS pp ON pd.objectpegawaifk = pp.id " & _
'                       " INNER JOIN kelompokpasien_m AS kp ON pd.objectkelompokpasienlastfk = kp.id " & _
'                       " INNER JOIN antrianpasiendiperiksa_t AS apdp ON apdp.noregistrasifk = pd.norec " & _
'                        " LEFT JOIN pelayananpasien_t AS tp ON tp.noregistrasifk = apdp.norec " & _
'                       " LEFT JOIN produk_m AS pro ON tp.produkfk = pro.id " & _
'                       " LEFT JOIN kelas_m AS ks ON apdp.objectkelasfk = ks.id " & _
'                       " LEFT JOIN asalrujukan_m AS ar ON apdp.objectasalrujukanfk = ar.id " & _
'                       " left JOIN rekanan_m AS rek ON rek.id= pd.objectrekananfk " & _
'                       " left JOIN kamar_m as kmr on apdp.objectkamarfk=kmr.id " & _
'                       " INNER join ruangan_m  as ru2 on ru2.id=apdp.objectruanganfk " & _
'                       " where pd.noregistrasi ='" & strNorec & "' and pro.id <> 402611  ORDER BY tp.tglpelayanan "
                       
             strSQL = "SELECT pd.noregistrasi,ps.nocm,ps.tgllahir,ps.namapasien, " & _
                       " pd.tglregistrasi,jk.reportdisplay AS jk,ru2.namaruangan AS ruanganperiksa,ru.namaruangan AS ruangakhir, " & _
                       " pp.namalengkap AS namadokter,kp.kelompokpasien,tp.produkfk, " & _
                       " pro.namaproduk,tp.jumlah,CASE WHEN tp.hargasatuan is null then tp.hargajual else tp.hargasatuan END as hargasatuan,ks.namakelas,ar.asalrujukan, " & _
                       " CASE WHEN rek.namarekanan is null then '-' else rek.namarekanan END as namapenjamin,kmr.namakamar " & _
                       " FROM pasiendaftar_t AS pd INNER JOIN pasien_m AS ps ON pd.nocmfk = ps.id " & _
                       " INNER JOIN jeniskelamin_m AS jk ON ps.objectjeniskelaminfk = jk.id " & _
                       " INNER JOIN ruangan_m AS ru ON pd.objectruanganlastfk = ru.id " & _
                       " INNER JOIN kelompokpasien_m AS kp ON pd.objectkelompokpasienlastfk = kp.id " & _
                       " INNER JOIN antrianpasiendiperiksa_t AS apdp ON apdp.noregistrasifk = pd.norec " & _
                       " LEFT JOIN pegawai_m AS pp ON apdp.objectpegawaifk = pp.id " & _
                        " LEFT JOIN pelayananpasien_t AS tp ON tp.noregistrasifk = apdp.norec " & _
                       " LEFT JOIN produk_m AS pro ON tp.produkfk = pro.id " & _
                       " LEFT JOIN kelas_m AS ks ON apdp.objectkelasfk = ks.id " & _
                       " LEFT JOIN asalrujukan_m AS ar ON apdp.objectasalrujukanfk = ar.id " & _
                       " left JOIN rekanan_m AS rek ON rek.id= pd.objectrekananfk " & _
                       " left JOIN kamar_m as kmr on apdp.objectkamarfk=kmr.id " & _
                       " INNER join ruangan_m  as ru2 on ru2.id=apdp.objectruanganfk " & _
                       " where pd.noregistrasi ='" & strNorec & "' and pro.id <> 402611  ORDER BY tp.tglpelayanan "
            
            ReadRs strSQL
            
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            
            .database.AddADOCommand CN_String, adoReport
            If RS.BOF Then
                .txtUmur.SetText "-"
            Else
                .txtUmur.SetText hitungUmur(Format(RS!tgllahir, "yyyy/MM/dd"), Format(Now, "yyyy/MM/dd"))
            End If


            .udtgl.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usnmpasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")
            
            .usUnitLayanan.SetUnboundFieldSource ("{ado.ruanganperiksa}")
            .usTipe.SetUnboundFieldSource ("{ado.kelompokpasien}")
            
            .usRujukan.SetUnboundFieldSource ("{ado.asalrujukan}")
            .usruangperiksa.SetUnboundFieldSource ("{ado.ruanganPeriksa}")
            '.usQty.SetUnboundFieldSource ("{ado.jumlah}")
            .usKamar.SetUnboundFieldSource ("if isnull({ado.namakamar}) then "" - "" else {ado.namakamar} ")
            .usKelas.SetUnboundFieldSource ("{ado.namakelas}")
            .usPenjamin.SetUnboundFieldSource ("{ado.namapenjamin}")
            
            .usDokter.SetUnboundFieldSource ("{ado.namadokter}")

            .usPelayanan.SetUnboundFieldSource ("{ado.namaproduk}")
            .ucTarif.SetUnboundFieldSource ("{ado.hargasatuan}")

            ReadRs2 "SELECT namalengkap FROM pegawai_m where id='" & strIdPegawai & "' "
            If RS2.BOF Then
                .txtuser.SetText "-"
            Else
                .txtuser.SetText UCase(IIf(IsNull(RS2("namalengkap")), "-", RS2("namalengkap")))
            End If
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "BuktiLayanan")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportBuktiLayanan
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:
    MsgBox Err.Number & " " & Err.Description
    
End Sub
Public Sub cetakBuktiLayananNorec_apd(strNorec As String, strIdPegawai As String, strruangan As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
Dim umur As String
    
bolBuktiPendaftaran = False
bolBuktiLayanan = True
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    Dim strarr() As String
    Dim norec_apc As String
    Dim i As Integer
    
    
    strarr = Split(strNorec, "|")
    For i = 0 To UBound(strarr)
       norec_apc = norec_apc + "'" & strarr(i) & "',"
    Next
    norec_apc = Left(norec_apc, Len(norec_apc) - 1)
    With reportBuktiLayanan
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT pd.noregistrasi,ps.nocm,ps.tgllahir,ps.namapasien, " & _
                       " pd.tglregistrasi,jk.reportdisplay AS jk,ru2.namaruangan AS ruanganperiksa,ru.namaruangan AS ruangakhir, " & _
                       " pp.namalengkap AS namadokter,kp.kelompokpasien,tp.produkfk, " & _
                       " pro.namaproduk,tp.jumlah,CASE WHEN tp.hargasatuan is null then tp.hargajual else tp.hargasatuan END as hargasatuan,ks.namakelas,ar.asalrujukan, " & _
                       " CASE WHEN rek.namarekanan is null then '-' else rek.namarekanan END as namapenjamin,kmr.namakamar " & _
                       " FROM pasiendaftar_t AS pd INNER JOIN pasien_m AS ps ON pd.nocmfk = ps.id " & _
                       " INNER JOIN jeniskelamin_m AS jk ON ps.objectjeniskelaminfk = jk.id " & _
                       " INNER JOIN ruangan_m AS ru ON pd.objectruanganlastfk = ru.id " & _
                       " LEFT JOIN pegawai_m AS pp ON pd.objectpegawaifk = pp.id " & _
                       " INNER JOIN kelompokpasien_m AS kp ON pd.objectkelompokpasienlastfk = kp.id " & _
                       " INNER JOIN antrianpasiendiperiksa_t AS apdp ON apdp.noregistrasifk = pd.norec " & _
                        " LEFT JOIN pelayananpasien_t AS tp ON tp.noregistrasifk = apdp.norec " & _
                       " LEFT JOIN produk_m AS pro ON tp.produkfk = pro.id " & _
                       " LEFT JOIN kelas_m AS ks ON apdp.objectkelasfk = ks.id " & _
                       " LEFT JOIN asalrujukan_m AS ar ON apdp.objectasalrujukanfk = ar.id " & _
                       " left JOIN rekanan_m AS rek ON rek.id= pd.objectrekananfk " & _
                       " left JOIN kamar_m as kmr on apdp.objectkamarfk=kmr.id " & _
                       " INNER join ruangan_m  as ru2 on ru2.id=apdp.objectruanganfk " & _
                       " where tp.norec  in (" & norec_apc & ") and pro.id <> 402611  ORDER BY tp.tglpelayanan "
                       
            
            ReadRs strSQL
            
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            
            .database.AddADOCommand CN_String, adoReport
            If RS.BOF Then
                .txtUmur.SetText "-"
            Else
                .txtUmur.SetText hitungUmur(Format(RS!tgllahir, "yyyy/MM/dd"), Format(Now, "yyyy/MM/dd"))
            End If


            .udtgl.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usnmpasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")
            
            .usUnitLayanan.SetUnboundFieldSource ("{ado.ruanganperiksa}")
            .usTipe.SetUnboundFieldSource ("{ado.kelompokpasien}")
            
            .usRujukan.SetUnboundFieldSource ("{ado.asalrujukan}")
            .usruangperiksa.SetUnboundFieldSource ("{ado.ruanganPeriksa}")
            '.usQty.SetUnboundFieldSource ("{ado.jumlah}")
            .usKamar.SetUnboundFieldSource ("if isnull({ado.namakamar}) then "" - "" else {ado.namakamar} ")
            .usKelas.SetUnboundFieldSource ("{ado.namakelas}")
            .usPenjamin.SetUnboundFieldSource ("{ado.namapenjamin}")
            
            .usDokter.SetUnboundFieldSource ("{ado.namadokter}")

            .usPelayanan.SetUnboundFieldSource ("{ado.namaproduk}")
            .ucTarif.SetUnboundFieldSource ("{ado.hargasatuan}")

            ReadRs2 "SELECT namalengkap FROM pegawai_m where id='" & strIdPegawai & "' "
            If RS2.BOF Then
                .txtuser.SetText "-"
            Else
                .txtuser.SetText UCase(IIf(IsNull(RS2("namalengkap")), "-", RS2("namalengkap")))
            End If
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "BuktiLayanan")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportBuktiLayanan
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:
    MsgBox Err.Number & " " & Err.Description
    
End Sub

Public Sub cetakBuktiLayananRuangan(strNorec As String, strIdPegawai As String, strIdRuangan As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
Dim umur As String
Dim strFilter As String

bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = True
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    strSQL = ""
    strFilter = ""
    If Left(strIdRuangan, 14) = "ORDERRADIOLOGI" Then
        strIdRuangan = Replace(strIdRuangan, "ORDERRADIOLOGI", "")
        strFilter = " AND apdp.norec = '" & strIdRuangan & "' "
    Else
        If strIdRuangan <> "" Then strFilter = " AND ru2.id = '" & strIdRuangan & "' "
    End If
    strFilter = strFilter & " and pro.id <> 402611 ORDER BY tp.tglpelayanan "
    With reportBuktiLayananRuangan
    
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT pd.noregistrasi,ps.nocm,ps.tgllahir,ps.namapasien, " & _
                       " apdp.tglregistrasi,jk.reportdisplay AS jk,ru2.namaruangan AS ruanganperiksa,ru.namaruangan AS ruangakhir, " & _
                       " (select pg.namalengkap from pegawai_m as pg INNER JOIN pelayananpasienpetugas_t p3 on p3.objectpegawaifk=pg.id " & _
                       "where p3.pelayananpasien=tp.norec and p3.objectjenispetugaspefk=4 limit 1) AS namadokter,kp.kelompokpasien,tp.produkfk, " & _
                       " pro.namaproduk,tp.jumlah,CASE WHEN tp.hargasatuan is null then tp.hargajual else tp.hargasatuan END as hargasatuan," & _
                       " (case when tp.hargadiscount is null then 0 else tp.hargadiscount end)* tp.jumlah as diskon, " & _
                       " hargasatuan*tp.jumlah as total, ks.namakelas,ar.asalrujukan, " & _
                       " CASE WHEN rek.namarekanan is null then '-' else rek.namarekanan END as namapenjamin,kmr.namakamar " & _
                       " FROM pasiendaftar_t AS pd INNER JOIN pasien_m AS ps ON pd.nocmfk = ps.id " & _
                       " LEFT JOIN jeniskelamin_m AS jk ON ps.objectjeniskelaminfk = jk.id " & _
                       " LEFT JOIN kelompokpasien_m AS kp ON pd.objectkelompokpasienlastfk = kp.id " & _
                       " LEFT JOIN antrianpasiendiperiksa_t AS apdp ON apdp.noregistrasifk = pd.norec " & _
                       " LEFT JOIN ruangan_m AS ru ON pd.objectruanganasalfk = ru.id " & _
                       " LEFT JOIN pelayananpasien_t AS tp ON tp.noregistrasifk = apdp.norec " & _
                       " LEFT JOIN pegawai_m AS pp ON apdp.objectpegawaifk = pp.id " & _
                       " LEFT JOIN produk_m AS pro ON tp.produkfk = pro.id " & _
                       " LEFT JOIN kelas_m AS ks ON apdp.objectkelasfk = ks.id " & _
                       " LEFT JOIN asalrujukan_m AS ar ON apdp.objectasalrujukanfk = ar.id " & _
                       " left JOIN rekanan_m AS rek ON rek.id= pd.objectrekananfk " & _
                       " left JOIN kamar_m as kmr on apdp.objectkamarfk=kmr.id " & _
                       " INNER join ruangan_m  as ru2 on ru2.id=apdp.objectruanganfk " & _
                       " where pd.noregistrasi ='" & strNorec & "'" & strFilter
            
            ReadRs strSQL
            
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            
            .database.AddADOCommand CN_String, adoReport
            If RS.BOF Then
                .txtUmur.SetText "-"
            Else
                .txtUmur.SetText hitungUmur(Format(RS!tgllahir, "yyyy/MM/dd"), Format(Now, "yyyy/MM/dd"))
            End If
            
            .udtgl.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usnmpasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")

            .usUnitLayanan.SetUnboundFieldSource ("{ado.ruanganperiksa}")
            .usTipe.SetUnboundFieldSource ("{ado.kelompokpasien}")

            .usRujukan.SetUnboundFieldSource ("{ado.asalrujukan}")
            .usruangperiksa.SetUnboundFieldSource ("{ado.ruangakhir}")
            .usQty.SetUnboundFieldSource ("{ado.jumlah}")
            .usKamar.SetUnboundFieldSource ("if isnull({ado.namakamar}) then "" - "" else {ado.namakamar} ")
            .usKelas.SetUnboundFieldSource ("if isnull({ado.namakelas}) then "" - "" else {ado.namakelas} ") '("{ado.namakelas}")
            .usPenjamin.SetUnboundFieldSource ("{ado.namapenjamin}")

            .usDokter.SetUnboundFieldSource ("{ado.namadokter}")

            .usPelayanan.SetUnboundFieldSource ("{ado.namaproduk}") '
            .ucTarif.SetUnboundFieldSource ("{ado.hargasatuan}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
            .usJumlah.SetUnboundFieldSource ("{ado.jumlah}")
'            .ucTotal.SetUnboundFieldSource ("{ado.total}")

            ReadRs2 "SELECT namalengkap FROM pegawai_m where id='" & strIdPegawai & "' "
            If RS2.BOF Then
                .txtuser.SetText "-"
            Else
                .txtuser.SetText UCase(IIf(IsNull(RS2("namalengkap")), "-", RS2("namalengkap")))
            End If
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "BuktiLayananRuangan")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportBuktiLayananRuangan
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:

    MsgBox Err.Number & " " & Err.Description
End Sub

Public Sub cetakKartuPasien(strNocm As String, strNamaPasien As String, strTglLahir As String, strJk As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
    
bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = True
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    With reportKartuPasien
'            Set adoReport = New ADODB.Command
'            adoReport.ActiveConnection = CN_String
'            adoReport.CommandText = strSQL
'            adoReport.CommandType = adCmdUnknown
'            .database.AddADOCommand CN_String, adoReport

'      Set sect = .Sections.Item("Section8")

        .txtNamaPas.SetText strNamaPasien & "(" & strJk & ")"

        .txtTgl.SetText strTglLahir
        .txtnocm.SetText strNocm
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "KartuPasien")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportKartuPasien
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:

    MsgBox Err.Number & " " & Err.Description
End Sub

Public Sub cetakLabelPasien(strNorec As String, view As String, qty As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
Dim i As Integer
Dim str As String
Dim jml As Integer
    
    
bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = True
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    With reportLabel
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "select pd.noregistrasi,pd.tglregistrasi,p.nocm, " & _
                    "upper(p.namapasien) as namapasien, jk.reportdisplay as jk, " & _
                        "p.tgllahir from pasiendaftar_t pd " & _
                      " INNER JOIN pasien_m p on pd.nocmfk=p.id " & _
                       " INNER JOIN jeniskelamin_m jk on jk.id=p.objectjeniskelaminfk " & _
                       " where pd.noregistrasi ='" & strNorec & "' "
            
'            jml = qty - 1
            
            str = ""
            If Val(qty) - 1 = 0 Then
                adoReport.CommandText = strSQL
             Else
                For i = 1 To Val(qty) - 1
                    str = strSQL & " union all " & str
                Next
                
                adoReport.CommandText = str & strSQL
           
           End If
           
            adoReport.CommandType = adCmdUnknown
            
            
            .database.AddADOCommand CN_String, adoReport


            .udtgl.SetUnboundFieldSource ("{ado.tgllahir}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")
    
            .udtgl1.SetUnboundFieldSource ("{ado.tgllahir}")
            .usNoreg1.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNocm1.SetUnboundFieldSource ("{ado.nocm}")
            .usNp1.SetUnboundFieldSource ("{ado.namapasien}")
            .usjk1.SetUnboundFieldSource ("{ado.jk}")
   
            .udtgl2.SetUnboundFieldSource ("{ado.tgllahir}")
            .usNoreg2.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNocm2.SetUnboundFieldSource ("{ado.nocm}")
            .usNp2.SetUnboundFieldSource ("{ado.namapasien}")
            .usjk2.SetUnboundFieldSource ("{ado.jk}")
            
            .udtgl3.SetUnboundFieldSource ("{ado.tgllahir}")
            .usNoreg3.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNocm3.SetUnboundFieldSource ("{ado.nocm}")
            .usNp3.SetUnboundFieldSource ("{ado.namapasien}")
            .usJk3.SetUnboundFieldSource ("{ado.jk}")
            
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "LabelPasien")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportLabel
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:

    MsgBox Err.Number & " " & Err.Description
End Sub

Public Sub cetakLabelPasienZebra(strNorec As String, view As String, qty As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
Dim i As Integer
Dim str As String
Dim jml As Integer
    
    
bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolLabelPasienZebra = True
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False

    With reportLabel
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "select pd.noregistrasi,pd.tglregistrasi,p.nocm, " & _
                    "upper(p.namapasien) as namapasien, jk.reportdisplay as jk, " & _
                        "p.tgllahir from pasiendaftar_t pd " & _
                      " INNER JOIN pasien_m p on pd.nocmfk=p.id " & _
                       " INNER JOIN jeniskelamin_m jk on jk.id=p.objectjeniskelaminfk " & _
                       " where pd.noregistrasi ='" & strNorec & "' "
            
'            jml = qty - 1
            
            str = ""
            If Val(qty) - 1 = 0 Then
                adoReport.CommandText = strSQL
             Else
                For i = 1 To Val(qty) - 1
                    str = strSQL & " union all " & str
                Next
                
                adoReport.CommandText = str & strSQL
           
           End If
           
            adoReport.CommandType = adCmdUnknown
            
            
            .database.AddADOCommand CN_String, adoReport


            .udtgl.SetUnboundFieldSource ("{ado.tgllahir}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")
    
            .udtgl1.SetUnboundFieldSource ("{ado.tgllahir}")
            .usNoreg1.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNocm1.SetUnboundFieldSource ("{ado.nocm}")
            .usNp1.SetUnboundFieldSource ("{ado.namapasien}")
            .usjk1.SetUnboundFieldSource ("{ado.jk}")
   
            
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "LabelPasien")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportLabel
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:

    MsgBox Err.Number & " " & Err.Description
End Sub
Public Sub cetakSummaryList(strNorec As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
   
    
bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = True
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    With reportSumList
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT  ps.nocm,ps.namapasien,ps.namaayah, case when ps.namakeluarga is null then '-' else ps.namakeluarga end as namakeluarga,ps.tempatlahir,ps.tgllahir, " & _
                       " jk.jeniskelamin,ps.noidentitas,ag.agama,pk.pekerjaan,kb.name as kebangsaan, " & _
                       " case when al.alamatlengkap is null then '-' else al.alamatlengkap end as alamatlengkap  , " & _
                       " case when al.kotakabupaten is null then '-' else al.kotakabupaten end as kotakabupaten  , " & _
                       " case when al.kecamatan is null then '-' else al.kecamatan end as kecamatan  , " & _
                       " case when al.namadesakelurahan is null then '-' else al.namadesakelurahan end as namadesakelurahan  , " & _
                       " ps.notelepon as mobilephone1, " & _
                       " sp.statusperkawinan from pasien_m ps " & _
                       " left JOIN jeniskelamin_m jk on jk.id=ps.objectjeniskelaminfk " & _
                       " left JOIN alamat_m al on ps.id=al.nocmfk " & _
                       " left JOIN agama_m ag on ps.objectagamafk=ag.id " & _
                       " left JOIN pekerjaan_m pk on pk.id=ps.objectpekerjaanfk " & _
                       " LEFT JOIN kebangsaan_m kb on kb.id=ps.objectkebangsaanfk " & _
                       " left JOIN statusperkawinan_m sp on sp.id=ps.objectstatusperkawinanfk " & _
                       " where ps.nocm ='" & strNorec & "' "
            
            ReadRs strSQL
                
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            .database.AddADOCommand CN_String, adoReport

            If RS.BOF Then
                .txtUmur.SetText "-"
            Else
                .txtUmur.SetText hitungUmur(Format(RS!tgllahir, "yyyy/MM/dd"), Format(Now, "yyyy/MM/dd"))
            End If
            .txtlTglLahir.SetText Format(RS!tgllahir, "yyyy/MM/dd")
             

            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usNamaKeuarga.SetUnboundFieldSource ("{ado.namakeluarga}")
            .udTglLahir.SetUnboundFieldSource ("{ado.tglLahir}")
            .usJK.SetUnboundFieldSource ("{ado.jeniskelamin}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usAlamat.SetUnboundFieldSource ("{ado.alamatlengkap}")
            .usKota.SetUnboundFieldSource ("{ado.kotakabupaten}")
           
           
           .usKel.SetUnboundFieldSource ("{ado.namadesakelurahan}")
           .usKec.SetUnboundFieldSource ("{ado.kecamatan}")
           .usHp.SetUnboundFieldSource ("{ado.mobilephone1}")
           .usTL.SetUnboundFieldSource ("{ado.tempatlahir}")
           .usAgama.SetUnboundFieldSource ("{ado.agama}")
           .usKebgsaan.SetUnboundFieldSource ("{ado.kebangsaan}")
           .usPekerjaan.SetUnboundFieldSource ("{ado.pekerjaan}")
           .usStatusPerkawinan.SetUnboundFieldSource ("{ado.statusperkawinan}")
           .usKatp.SetUnboundFieldSource ("{ado.noidentitas}")
'           .usKatp.SetUnboundFieldSource ("{noidentitas}")
           
           ' .usNoTlpn.SetUnboundFieldSource ("{ado.kotakabupaten}")
           ' .usRt.SetUnboundFieldSource ("{ado.Rt}")
           ' .usRw.SetUnboundFieldSource ("{ado.rw}")
           ' .usSuku.SetUnboundFieldSource ("{ado.agama}")
          
           ' .usJenisPembayaran.SetUnboundFieldSource ("{noidentitas}")
           ' .usAlergi.SetUnboundFieldSource ("{noidentitas}")
           
          
           
           
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "SummaryList")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportSumList
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:
    MsgBox Err.Number & " " & Err.Description
End Sub

'cetakLembarMasuk
Public Sub cetakLembarMasuk(strNorec As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
   
    
bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = True
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    With reportRmk
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT pd.noregistrasi, ps.nocm, upper(ps.namapasien) as namapasien, upper(case when ps.namakeluarga is null then '-' else ps.namakeluarga end) as namakeluarga," & _
                       " upper(ps.namaayah) as namaayah,upper(case when ps.tempatlahir is null then '-' else ps.tempatlahir end || ', ' || TO_CHAR(ps.tgllahir, 'dd Month YYYY')) || ' Jam: ' || TO_CHAR(ps.jamlahir, 'HH24:MI') as tempatlahir,ps.tgllahir,jk.jeniskelamin, ps.noidentitas, " & _
                       " ag.agama, pk.pekerjaan, kb.name AS kebangsaan,upper(al.alamatlengkap) as alamatlengkap,upper(al.kotakabupaten) as kotakabupaten, " & _
                       " al.kecamatan, al.namadesakelurahan, al.mobilephone1,sp.statusperkawinan, " & _
                       " (kmr.namakamar || ' - ' || kls.namakelas ) as namakamar,(tt.reportdisplay || ' - ' ||tt.nomorbed ) AS nomorbed, " & _
                       " TO_CHAR(pd.tglregistrasi, 'dd Mon YYYY') as tglregistrasi, TO_CHAR(pd.tglpulang, 'dd Mon YYYY') as tglpulang, ps.namaibu, '-' as ttlSuami, " & _
                       " COALESCE(ps.namasuamiistri,'-') as namasuamiistri, pg.namalengkap as namadokterpj, kp.kelompokpasien, " & _
                       " '-' as alamatPekerjaan,'-' as keldihubungi  ,'-' as Hubungan , '-' as alamatKeluarga, " & _
                       " '-' as NohpKeluarga,ps.notelepon, case when ddp.keterangan is null then ' - ' else ddp.keterangan end as namadiagnosa" & _
                       " FROM pasiendaftar_t pd  " & _
                       " INNER JOIN antrianpasiendiperiksa_t apdp on pd.norec=apdp.noregistrasifk  " & _
                       " INNER JOIN pasien_m ps on pd.nocmfk=ps.id " & _
                       " INNER JOIN jeniskelamin_m jk on jk.id=ps.objectjeniskelaminfk " & _
                       " INNER JOIN alamat_m al on ps.id=al.nocmfk  INNER JOIN agama_m ag on ps.objectagamafk=ag.id " & _
                       " left JOIN pekerjaan_m pk on pk.id=ps.objectpekerjaanfk " & _
                       " LEFT JOIN kebangsaan_m kb on kb.id=ps.objectkebangsaanfk " & _
                       " left JOIN statusperkawinan_m sp on sp.id=ps.objectstatusperkawinanfk " & _
                       " INNER JOIN ruangan_m ru on apdp.objectruanganfk=ru.id " & _
                       " left JOIN kamar_m kmr on apdp.objectkamarfk=kmr.id " & _
                       " left JOIN tempattidur_m tt on apdp.nobed=tt.id " & _
                       " LEFT JOIN pegawai_m pg on pd.objectpegawaifk=pg.id " & _
                       " INNER JOIN kelompokpasien_m kp on pd.objectkelompokpasienlastfk=kp.id " & _
                       " INNER JOIN kelas_m kls on apd.objectkelasfk=kls.id " & _
                       " left JOIN detaildiagnosapasien_t as ddp on ddp.id=apd.noregistrasifk left join diagnosa_m as dg on dg.id=ddp.objectdiagnosafk left JOIN jenisdiagnosa_t as jd on jd.id=ddp.objectjenisdiagosa " & _
                       " where apd.norec ='" & strNorec & "' and jd.id=5 "
            
            ReadRs strSQL
                
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            .database.AddADOCommand CN_String, adoReport

            If RS.BOF Then
                .txtUmur.SetText "Umur -"
            Else
                .txtUmur.SetText "Umur " & hitungUmur(Format(RS!tgllahir, "yyyy/MM/dd"), Format(RS!tglregistrasi, "yyyy/MM/dd"))
                .txtTglMasuk.SetText Format(RS!tglregistrasi, "dd MMM yyyy")
                .txtJamMasuk.SetText Format(RS!jamregistrasi, "HH:MM:ss")
                .txtTglPlng.SetText IIf(RS!tglpulang = "Null", "-", Format(RS!tglpulang, "dd MMM yyyy"))
                .txtJamPlng.SetText IIf(RS!tglpulang = "Null", "-", Format(RS!jampulang, "HH:MM:ss"))
            End If
            
            .usDokter.SetUnboundFieldSource ("{ado.namadokterpj}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
                
            .usKamar.SetUnboundFieldSource ("{ado.namakamar}")
            .usTempatTidur.SetUnboundFieldSource ("{ado.nomorbed}")
            
            .usNamaKeuarga.SetUnboundFieldSource ("{ado.namakeluarga}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            
            .usAlamat.SetUnboundFieldSource ("{ado.alamatlengkap}")
            .usNoTlpn.SetUnboundFieldSource ("{ado.notelepon}")
            
            .usTL.SetUnboundFieldSource ("{ado.tempatlahir}")
'            .udTglLahir.SetUnboundFieldSource ("{ado.tglLahir}")
            .usJK.SetUnboundFieldSource ("{ado.jeniskelamin}")
            
            .usStatusPerkawinan.SetUnboundFieldSource ("{ado.statusperkawinan}")
            .usAgama.SetUnboundFieldSource ("{ado.agama}")
            .usPekerjaan.SetUnboundFieldSource ("{ado.pekerjaan}")
            
            .usNamaIbu.SetUnboundFieldSource ("{ado.namaibu}")
            .usNamaAyah.SetUnboundFieldSource ("{ado.namaayah}")
            .usNamaSuami.SetUnboundFieldSource ("{ado.namasuamiistri}")
            .usTTLSuami.SetUnboundFieldSource ("{ado.ttlSuami}")
            
            .usAlamatPekerjaanAyah.SetUnboundFieldSource ("{ado.alamatPekerjaan}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            
            .usOrgDpDihubungi.SetUnboundFieldSource ("{ado.keldihubungi}")
            .usHubungan.SetUnboundFieldSource ("{ado.Hubungan}")
            .usAlamatKeluarga.SetUnboundFieldSource ("{ado.alamatlengkap}")
            .usHp.SetUnboundFieldSource ("{ado.NohpKeluarga}")
            
'            .udTglMasuk.SetUnboundFieldSource ("{ado.tglregistrasi}")
'            '.udJamMasuk.SetUnboundFieldSource ("{ado.tglregistrasi}")
'            .udTglKeluar.SetUnboundFieldSource ("{ado.tglpulang}")
'            '.udJamKeluar.SetUnboundFieldSource ("{ado.tglpulang}")
            
            .usJenisPembayaran.SetUnboundFieldSource ("{ado.kelompokpasien}")
'            .usDiagnosa.SetUnboundFieldSource ("{ado.namadiagnosa}")
           
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "CetakRMK")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportRmk
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:
    MsgBox Err.Number & " " & Err.Description
'    MsgBox Err.Description, vbInformation
End Sub

Public Sub cetakLembarMasukByNorec_APD(strNorec As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
   
    
bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = True
boolLembarPersetujuan = False

    With reportRmk
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
             
             strSQL = "SELECT distinct pd.noregistrasi, ps.nocm, upper(ps.namapasien) as namapasien, upper(case when ps.namakeluarga is null then '-' else ps.namakeluarga end) as namakeluarga," & _
                       " upper(ps.namaayah) as namaayah,upper(case when ps.tempatlahir is null then '-' else ps.tempatlahir || ', ' end) as tempatlahir,  TO_CHAR(ps.tgllahir, 'dd Month YYYY')as tgllahir, ' Jam: ' || TO_CHAR(ps.jamlahir, 'HH24:MI') as jamlahir,jk.jeniskelamin, ps.noidentitas, " & _
                       " ag.agama, pk.pekerjaan, kb.name AS kebangsaan,upper(al.alamatlengkap) as alamatlengkap,upper(al.kotakabupaten) as kotakabupaten, " & _
                       " al.kecamatan, al.namadesakelurahan, al.mobilephone1,sp.statusperkawinan, " & _
                       " (kmr.namakamar || ' - ' || kls.namakelas ) as namakamar,(tt.reportdisplay || ' - ' ||tt.nomorbed ) AS nomorbed, " & _
                       " TO_CHAR(pd.tglregistrasi, 'dd Mon YYYY') as tglregistrasi,TO_CHAR(pd.tglregistrasi, 'HH24:MI') as jamregistrasi, TO_CHAR(pd.tglpulang, 'DD Mon YYYY') as tglpulang,TO_CHAR(pd.tglpulang, 'HH24:MI') as jampulang, TO_CHAR(pd.tglpulang, 'dd Mon YYYY') as tglpulang, ps.namaibu, '-' as ttlSuami, " & _
                       " COALESCE(ps.namasuamiistri,'-') as namasuamiistri, pg.namalengkap as namadokterpj, kp.kelompokpasien, " & _
                       " '-' as alamatPekerjaan,'-' as keldihubungi  ,'-' as Hubungan , '-' as alamatKeluarga, " & _
                       " '-' as NohpKeluarga,ps.notelepon, case when ddp.keterangan is null then ' - ' else ddp.keterangan end as namadiagnosa" & _
                       " FROM pasiendaftar_t pd  " & _
                       " INNER JOIN antrianpasiendiperiksa_t apdp on pd.norec=apdp.noregistrasifk  " & _
                       " INNER JOIN pasien_m ps on pd.nocmfk=ps.id " & _
                       " INNER JOIN jeniskelamin_m jk on jk.id=ps.objectjeniskelaminfk " & _
                       " INNER JOIN alamat_m al on ps.id=al.nocmfk  INNER JOIN agama_m ag on ps.objectagamafk=ag.id " & _
                       " left JOIN pekerjaan_m pk on pk.id=ps.objectpekerjaanfk " & _
                       " LEFT JOIN kebangsaan_m kb on kb.id=ps.objectkebangsaanfk " & _
                       " left JOIN statusperkawinan_m sp on sp.id=ps.objectstatusperkawinanfk " & _
                       " INNER JOIN ruangan_m ru on apdp.objectruanganfk=ru.id " & _
                       " left JOIN kamar_m kmr on apdp.objectkamarfk=kmr.id " & _
                       " left JOIN tempattidur_m tt on apdp.nobed=tt.id " & _
                       " LEFT JOIN pegawai_m pg on pd.objectpegawaifk=pg.id " & _
                       " INNER JOIN kelompokpasien_m kp on pd.objectkelompokpasienlastfk=kp.id " & _
                       " INNER JOIN kelas_m kls on apdp.objectkelasfk=kls.id " & _
                       " left JOIN detaildiagnosapasien_t as ddp on ddp.noregistrasifk=apdp.norec left join diagnosa_m as dg on dg.id=ddp.objectdiagnosafk left JOIN jenisdiagnosa_m as jd on jd.id=ddp.objectjenisdiagnosafk " & _
                       " where apdp.norec ='" & strNorec & "' and jd.id=5 "
             
'             strSQL = "SELECT distinct pd.noregistrasi, ps.nocm, upper(ps.namapasien) as namapasien, upper(case when ps.namakeluarga is null then '-' else ps.namakeluarga end) as namakeluarga," & _
'                       " upper(ps.namaayah) as namaayah,upper(case when ps.tempatlahir is null then '-' else ps.tempatlahir end || ', ' || TO_CHAR(ps.tgllahir, 'dd Month YYYY')) || ' Jam: ' || TO_CHAR(ps.jamlahir, 'HH24:MI') as tempatlahir,ps.tgllahir,jk.jeniskelamin, ps.noidentitas, " & _
'                       " ag.agama, pk.pekerjaan, kb.name AS kebangsaan,upper(al.alamatlengkap) as alamatlengkap,upper(al.kotakabupaten) as kotakabupaten, " & _
'                       " al.kecamatan, al.namadesakelurahan, al.mobilephone1,sp.statusperkawinan, " & _
'                       " (kmr.namakamar || ' - ' || kls.namakelas ) as namakamar,(tt.reportdisplay || ' - ' ||tt.nomorbed ) AS nomorbed, " & _
'                       " TO_CHAR(pd.tglregistrasi, 'dd Mon YYYY') as tglregistrasi,TO_CHAR(pd.tglregistrasi, 'HH24:MI') as jamregistrasi, TO_CHAR(pd.tglpulang, 'DD Mon YYYY') as tglpulang,TO_CHAR(pd.tglpulang, 'HH24:MI') as jampulang, TO_CHAR(pd.tglpulang, 'dd Mon YYYY') as tglpulang, ps.namaibu, '-' as ttlSuami, " & _
'                       " COALESCE(ps.namasuamiistri,'-') as namasuamiistri, pg.namalengkap as namadokterpj, kp.kelompokpasien, " & _
'                       " '-' as alamatPekerjaan,'-' as keldihubungi  ,'-' as Hubungan , '-' as alamatKeluarga, " & _
'                       " '-' as NohpKeluarga,ps.notelepon, case when ddp.keterangan is null then ' - ' else ddp.keterangan end as namadiagnosa" & _
'                       " FROM pasiendaftar_t pd  " & _
'                       " INNER JOIN antrianpasiendiperiksa_t apdp on pd.norec=apdp.noregistrasifk  " & _
'                       " INNER JOIN pasien_m ps on pd.nocmfk=ps.id " & _
'                       " INNER JOIN jeniskelamin_m jk on jk.id=ps.objectjeniskelaminfk " & _
'                       " INNER JOIN alamat_m al on ps.id=al.nocmfk  INNER JOIN agama_m ag on ps.objectagamafk=ag.id " & _
'                       " left JOIN pekerjaan_m pk on pk.id=ps.objectpekerjaanfk " & _
'                       " LEFT JOIN kebangsaan_m kb on kb.id=ps.objectkebangsaanfk " & _
'                       " left JOIN statusperkawinan_m sp on sp.id=ps.objectstatusperkawinanfk " & _
'                       " INNER JOIN ruangan_m ru on apdp.objectruanganfk=ru.id " & _
'                       " left JOIN kamar_m kmr on apdp.objectkamarfk=kmr.id " & _
'                       " left JOIN tempattidur_m tt on apdp.nobed=tt.id " & _
'                       " LEFT JOIN pegawai_m pg on pd.objectpegawaifk=pg.id " & _
'                       " INNER JOIN kelompokpasien_m kp on pd.objectkelompokpasienlastfk=kp.id " & _
'                       " INNER JOIN kelas_m kls on apdp.objectkelasfk=kls.id " & _
'                       " left JOIN detaildiagnosapasien_t as ddp on ddp.noregistrasifk=apdp.norec left join diagnosa_m as dg on dg.id=ddp.objectdiagnosafk left JOIN jenisdiagnosa_m as jd on jd.id=ddp.objectjenisdiagnosafk " & _
'                       " where apdp.norec ='" & strNorec & "' and jd.id=5 "
            
'            strSQL = "SELECT pd.noregistrasi, ps.nocm, upper(ps.namapasien) as namapasien, upper(case when ps.namakeluarga is null then '-' else ps.namakeluarga end) as namakeluarga," & _
'                       " upper(ps.namaayah) as namaayah,upper(ps.tempatlahir || ', ' || TO_CHAR(ps.tgllahir, 'DD Month YYYY')) || ' Jam: ' || TO_CHAR(ps.tgllahir, 'HH24:MI') as tempatlahir,ps.tgllahir,jk.jeniskelamin, ps.noidentitas, " & _
'                       " ag.agama, pk.pekerjaan, kb.name AS kebangsaan,upper(al.alamatlengkap) as alamatlengkap,upper(al.kotakabupaten) as kotakabupaten, " & _
'                       " al.kecamatan, al.namadesakelurahan, al.mobilephone1,sp.statusperkawinan, " & _
'                       " (kmr.namakamar || ' - ' || kls.namakelas ) as namakamar,(tt.reportdisplay || ' - ' ||tt.nomorbed ) AS nomorbed, " & _
'                       " TO_CHAR(pd.tglregistrasi, 'DD Mon YYYY') as tglregistrasi,TO_CHAR(pd.tglregistrasi, 'HH24:MI') as jamregistrasi, TO_CHAR(pd.tglpulang, 'DD Mon YYYY') as tglpulang,TO_CHAR(pd.tglpulang, 'HH24:MI') as jampulang, ps.namaibu, '-' as ttlSuami, " & _
'                       " COALESCE(ps.namasuamiistri,'-') as namasuamiistri, pg.namalengkap as namadokterpj, kp.kelompokpasien, " & _
'                       " '-' as alamatPekerjaan,'-' as keldihubungi  ,'-' as Hubungan , '-' as alamatKeluarga, " & _
'                       " '-' as NohpKeluarga,ps.notelepon " & _
'                       " FROM pasiendaftar_t pd  " & _
'                       " INNER JOIN antrianpasiendiperiksa_t apdp on pd.norec=apdp.noregistrasifk  " & _
'                       " INNER JOIN pasien_m ps on pd.nocmfk=ps.id " & _
'                       " INNER JOIN jeniskelamin_m jk on jk.id=ps.objectjeniskelaminfk " & _
'                       " INNER JOIN alamat_m al on ps.id=al.nocmfk  INNER JOIN agama_m ag on ps.objectagamafk=ag.id " & _
'                       " left JOIN pekerjaan_m pk on pk.id=ps.objectpekerjaanfk " & _
'                       " LEFT JOIN kebangsaan_m kb on kb.id=ps.objectkebangsaanfk " & _
'                       " left JOIN statusperkawinan_m sp on sp.id=ps.objectstatusperkawinanfk " & _
'                       " INNER JOIN ruangan_m ru on apdp.objectruanganfk=ru.id " & _
'                       " left JOIN kamar_m kmr on apdp.objectkamarfk=kmr.id " & _
'                       " left JOIN tempattidur_m tt on apdp.nobed=tt.id " & _
'                       " LEFT JOIN pegawai_m pg on pd.objectpegawaifk=pg.id " & _
'                       " INNER JOIN kelompokpasien_m kp on pd.objectkelompokpasienlastfk=kp.id " & _
'                       " INNER JOIN kelas_m kls on apdp.objectkelasfk=kls.id " & _
'                       " where apdp.norec ='" & strNorec & "' "
            
            ReadRs strSQL
                
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            .database.AddADOCommand CN_String, adoReport

            If RS.BOF Then
                .txtUmur.SetText "Umur -"
            Else
                .txtUmur.SetText "Umur " & hitungUmur(Format(RS!tgllahir, "yyyy/MM/dd"), Format(RS!tglregistrasi, "yyyy/MM/dd"))
                .txtTglMasuk.SetText Format(RS!tglregistrasi, "dd MMM yyyy")
                .txtJamMasuk.SetText Format(RS!jamregistrasi, "hh:mm:ss")
                .txtTglPlng.SetText IIf(RS!tglpulang = "Null", "-", Format(RS!tglpulang, "dd MMM yyyy"))
                .txtJamPlng.SetText IIf(RS!tglpulang = "Null", "-", Format(RS!jampulang, "hh:mm:ss"))
            End If
            
            .usDokter.SetUnboundFieldSource ("{ado.namadokterpj}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
                
            .usKamar.SetUnboundFieldSource ("{ado.namakamar}")
            .usTempatTidur.SetUnboundFieldSource ("{ado.nomorbed}")
            
            .usNamaKeuarga.SetUnboundFieldSource ("{ado.namakeluarga}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            
            .usAlamat.SetUnboundFieldSource ("{ado.alamatlengkap}")
            .usNoTlpn.SetUnboundFieldSource ("{ado.notelepon}")
            
            .usTL.SetUnboundFieldSource ("{ado.tempatlahir}")
            .usTglLahir.SetUnboundFieldSource ("{ado.tgllahir}")
            .usJamLahir.SetUnboundFieldSource ("{ado.jamlahir}")
            .usJK.SetUnboundFieldSource ("{ado.jeniskelamin}")
            
            .usStatusPerkawinan.SetUnboundFieldSource ("{ado.statusperkawinan}")
            .usAgama.SetUnboundFieldSource ("{ado.agama}")
            .usPekerjaan.SetUnboundFieldSource ("{ado.pekerjaan}")
            
            .usNamaIbu.SetUnboundFieldSource ("{ado.namaibu}")
            .usNamaAyah.SetUnboundFieldSource ("{ado.namaayah}")
            .usNamaSuami.SetUnboundFieldSource ("{ado.namasuamiistri}")
            .usTTLSuami.SetUnboundFieldSource ("{ado.ttlSuami}")
            
            .usAlamatPekerjaanAyah.SetUnboundFieldSource ("{ado.alamatPekerjaan}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            
            .usOrgDpDihubungi.SetUnboundFieldSource ("{ado.keldihubungi}")
            .usHubungan.SetUnboundFieldSource ("{ado.Hubungan}")
            .usAlamatKeluarga.SetUnboundFieldSource ("{ado.alamatlengkap}")
            .usHp.SetUnboundFieldSource ("{ado.NohpKeluarga}")
            
'            .udTglMasuk.SetUnboundFieldSource ("{ado.tglregistrasi}")
'            '.udJamMasuk.SetUnboundFieldSource ("{ado.tglregistrasi}")
'            .udTglKeluar.SetUnboundFieldSource ("{ado.tglpulang}")
'            '.udJamKeluar.SetUnboundFieldSource ("{ado.tglpulang}")
            
            .usJenisPembayaran.SetUnboundFieldSource ("{ado.kelompokpasien}")
            .usDiagnosa.SetUnboundFieldSource ("{ado.namadiagnosa}")
    
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "CetakRMK")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportRmk
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:
    MsgBox Err.Number & " " & Err.Description
'    MsgBox Err.Description, vbInformation
End Sub



Public Sub cetakPersetujuan(strNorec As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
   
    
bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = True
bolBuktiLayananRuanganBedah = False

    With reportLembarGC
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT pd.noregistrasi,pd.tglregistrasi, " & _
                       " ps.nocm, ps.namapasien, " & _
                       " ps.tgllahir,jk.reportdisplay AS jk, " & _
                       " ps.namaayah , ru.namaruangan, kls.namakelas " & _
                       " from pasiendaftar_t pd " & _
                       " INNER JOIN pasien_m ps on pd.nocmfk=ps.id " & _
                       " INNER JOIN jeniskelamin_m jk on jk.id=ps.objectjeniskelaminfk " & _
                       " INNER JOIN antrianpasiendiperiksa_t apdp on pd.norec=apdp.noregistrasifk " & _
                       " INNER JOIN ruangan_m ru on apdp.objectruanganfk=ru.id " & _
                       " INNER JOIN kelas_m kls on  apdp.objectkelasfk=kls.id " & _
                       " where ru.objectdepartemenfk in (16,35) and pd.noregistrasi ='" & strNorec & "' "
            ReadRs strSQL
                
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            .database.AddADOCommand CN_String, adoReport
            
            If RS.EOF = False Then
                If Len(RS("nocm")) <= 8 Then
                   .txtn1.SetText Mid(RS("nocm"), 1, 1)
                   .txtn2.SetText Mid(RS("nocm"), 2, 1)
                   .txtn3.SetText Mid(RS("nocm"), 3, 1)
                   .txtn4.SetText Mid(RS("nocm"), 4, 1)
                   .txtn5.SetText Mid(RS("nocm"), 5, 1)
                   .txtn6.SetText Mid(RS("nocm"), 6, 1)
                   .txtn7.SetText Mid(RS("nocm"), 7, 1)
                   .txtn8.SetText Mid(RS("nocm"), 8, 1)
                Else
                   .txtn1.SetText Mid(RS("nocm"), 1, 2)
                   .txtn2.SetText Mid(RS("nocm"), 3, 2)
                   .txtn3.SetText Mid(RS("nocm"), 5, 2)
                   .txtn4.SetText Mid(RS("nocm"), 7, 2)
                   .txtn5.SetText Mid(RS("nocm"), 9, 2)
                   .txtn6.SetText Mid(RS("nocm"), 11, 2)
                   .txtn7.SetText Mid(RS("nocm"), 13, 2)
                   .txtn8.SetText Mid(RS("nocm"), 15, 2)
                
                End If
                
                .txtNamaKeluarga.SetText RS("namaayah")
                .txtNamaPasien.SetText RS("namapasien")
                .txtRuangan.SetText RS("namaruangan")
                .txtKelas.SetText RS("namakelas")
                If RS("Jk") = "P" Then .txtL.SetText "-" Else .txtP.SetText "-"
                .txtTgllahir.SetText Format(RS!tgllahir, "yyyy/MM/dd")
            End If
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "CetakGeneral")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportLembarGC
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:
    MsgBox Err.Number & " " & Err.Description

End Sub

Public Sub cetakBuktiLayananRuanganPerTindakan(strNorec As String, strIdPegawai As String, strIdRuangan As String, strIdTindakan As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
Dim umur As String
Dim strFilter As String
Dim strFilter2 As String

bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = True
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False

    strSQL = ""
    strFilter = ""
    strFilter2 = ""
    If strIdRuangan <> "" Then strFilter = " AND ru2.id = '" & strIdRuangan & "' "
    If strIdTindakan <> "" Then strFilter2 = " AND tp.produkfk = '" & strIdTindakan & "' "
    strFilter = strFilter & strFilter2 & " ORDER BY tp.tglpelayanan "
    With reportBuktiLayananRuanganPerTindakan
    
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT pd.noregistrasi,ps.nocm,ps.tgllahir,ps.namapasien, " & _
                       " pd.tglregistrasi,jk.reportdisplay AS jk,ru2.namaruangan AS ruanganperiksa,ru.namaruangan AS ruangakhir, " & _
                       " pp.namalengkap AS namadokter,kp.kelompokpasien,tp.produkfk, " & _
                       " pro.namaproduk,tp.jumlah,CASE WHEN tp.hargasatuan is null then tp.hargajual else tp.hargasatuan END as hargasatuan," & _
                       " (case when tp.hargadiscount is null then 0 else tp.hargadiscount end)* tp.jumlah as diskon, " & _
                       " hargasatuan*tp.jumlah as total,ks.namakelas,ar.asalrujukan, " & _
                       " CASE WHEN rek.namarekanan is null then '-' else rek.namarekanan END as namapenjamin,kmr.namakamar " & _
                       " FROM pasiendaftar_t AS pd INNER JOIN pasien_m AS ps ON pd.nocmfk = ps.id " & _
                       " INNER JOIN jeniskelamin_m AS jk ON ps.objectjeniskelaminfk = jk.id " & _
                       " LEFT JOIN pegawai_m AS pp ON pd.objectpegawaifk = pp.id " & _
                       " INNER JOIN kelompokpasien_m AS kp ON pd.objectkelompokpasienlastfk = kp.id " & _
                       " INNER JOIN antrianpasiendiperiksa_t AS apdp ON apdp.noregistrasifk = pd.norec " & _
                       " INNER JOIN ruangan_m AS ru ON apdp.objectruanganfk = ru.id " & _
                       " LEFT JOIN pelayananpasien_t AS tp ON tp.noregistrasifk = apdp.norec " & _
                       " LEFT JOIN produk_m AS pro ON tp.produkfk = pro.id " & _
                       " LEFT JOIN kelas_m AS ks ON apdp.objectkelasfk = ks.id " & _
                       " LEFT JOIN asalrujukan_m AS ar ON apdp.objectasalrujukanfk = ar.id " & _
                       " left JOIN rekanan_m AS rek ON rek.id= pd.objectrekananfk " & _
                       " left JOIN kamar_m as kmr on apdp.objectkamarfk=kmr.id " & _
                       " INNER join ruangan_m  as ru2 on ru2.id=apdp.objectruanganfk " & _
                       " where pd.noregistrasi ='" & strNorec & "' and pro.id <> 402611 " & strFilter
            
            ReadRs strSQL
            
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            
            .database.AddADOCommand CN_String, adoReport
            If RS.BOF Then
                .txtUmur.SetText "-"
            Else
                .txtUmur.SetText hitungUmur(Format(RS!tgllahir, "yyyy/MM/dd"), Format(Now, "yyyy/MM/dd"))
            End If
            
            .udtgl.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usnmpasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")

            .usUnitLayanan.SetUnboundFieldSource ("{ado.ruanganperiksa}")
            .usTipe.SetUnboundFieldSource ("{ado.kelompokpasien}")

            .usRujukan.SetUnboundFieldSource ("{ado.asalrujukan}")
            .usruangperiksa.SetUnboundFieldSource ("{ado.ruangakhir}")
            .usQty.SetUnboundFieldSource ("{ado.jumlah}")
            .usKamar.SetUnboundFieldSource ("if isnull({ado.namakamar}) then "" - "" else {ado.namakamar} ")
            .usKelas.SetUnboundFieldSource ("if isnull({ado.namakelas}) then "" - "" else {ado.namakelas} ") '("{ado.namakelas}")
            .usPenjamin.SetUnboundFieldSource ("{ado.namapenjamin}")

            .usDokter.SetUnboundFieldSource ("{ado.namadokter}")

            .usPelayanan.SetUnboundFieldSource ("{ado.namaproduk}")
            .ucTarif.SetUnboundFieldSource ("{ado.hargasatuan}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
            .usJumlah.SetUnboundFieldSource ("{ado.jumlah}")
            .udTglPelayanan.Suppress = True
            
'            .ucTotal.SetUnboundFieldSource ("{ado.total}")

            ReadRs2 "SELECT namalengkap FROM pegawai_m where id='" & strIdPegawai & "' "
            If RS2.BOF Then
                .txtuser.SetText "-"
            Else
                .txtuser.SetText UCase(IIf(IsNull(RS2("namalengkap")), "-", RS2("namalengkap")))
            End If
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "BuktiLayananRuanganPerTindakan")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportBuktiLayananRuanganPerTindakan
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:
    MsgBox Err.Number & " " & Err.Description
End Sub

Public Sub cetakBuktiLayananRuanganPerTindakanByNorec(strNorec As String, strIdPegawai As String, strIdRuangan As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
Dim umur As String
Dim strFilter As String
Dim strFilter2 As String

bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = True
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False


    Dim strarr() As String
    Dim norec_apc As String
    Dim i As Integer
    
    
    strarr = Split(strNorec, "|")
    For i = 0 To UBound(strarr)
       norec_apc = norec_apc + "'" & strarr(i) & "',"
    Next
    norec_apc = Left(norec_apc, Len(norec_apc) - 1)
    
    strSQL = ""
    strFilter = ""
    strFilter2 = ""
'    If strIdRuangan <> "" Then strFilter = " AND ru2.id = '" & strIdRuangan & "' "
'    If strIdTindakan <> "" Then strFilter2 = " AND tp.produkfk = '" & strIdTindakan & "' "
    strFilter = strFilter & strFilter2 & " ORDER BY tp.tglpelayanan "
    With reportBuktiLayananRuanganPerTindakan
    
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT pd.noregistrasi,ps.nocm,ps.tgllahir,ps.namapasien, " & _
                       " apdp.tglregistrasi,jk.reportdisplay AS jk,ru2.namaruangan AS ruanganperiksa,ru.namaruangan AS ruangakhir, " & _
                       " (select pg.namalengkap from pegawai_m as pg INNER JOIN pelayananpasienpetugas_t p3 on p3.objectpegawaifk=pg.id " & _
                       "where p3.pelayananpasien=tp.norec and p3.objectjenispetugaspefk=4 limit 1) AS namadokter,kp.kelompokpasien,tp.produkfk, " & _
                       " pro.namaproduk,tp.jumlah,CASE WHEN tp.hargasatuan is null then tp.hargajual else tp.hargasatuan END as hargasatuan," & _
                       " (case when tp.hargadiscount is null then 0 else tp.hargadiscount end) as diskon, " & _
                       " hargasatuan*tp.jumlah as total,ks.namakelas,ar.asalrujukan,tp.tglpelayanan, " & _
                       " CASE WHEN rek.namarekanan is null then '-' else rek.namarekanan END as namapenjamin,kmr.namakamar " & _
                       " FROM pasiendaftar_t AS pd INNER JOIN pasien_m AS ps ON pd.nocmfk = ps.id " & _
                       " INNER JOIN jeniskelamin_m AS jk ON ps.objectjeniskelaminfk = jk.id " & _
                       " INNER JOIN kelompokpasien_m AS kp ON pd.objectkelompokpasienlastfk = kp.id " & _
                       " INNER JOIN antrianpasiendiperiksa_t AS apdp ON apdp.noregistrasifk = pd.norec " & _
                       " INNER JOIN ruangan_m AS ru ON apdp.objectruanganfk = ru.id " & _
                       " LEFT JOIN pelayananpasien_t AS tp ON tp.noregistrasifk = apdp.norec " & _
                       " LEFT JOIN pegawai_m AS pp ON apdp.objectpegawaifk = pp.id " & _
                       " LEFT JOIN produk_m AS pro ON tp.produkfk = pro.id " & _
                       " LEFT JOIN kelas_m AS ks ON apdp.objectkelasfk = ks.id " & _
                       " LEFT JOIN asalrujukan_m AS ar ON apdp.objectasalrujukanfk = ar.id " & _
                       " left JOIN rekanan_m AS rek ON rek.id= pd.objectrekananfk " & _
                       " left JOIN kamar_m as kmr on apdp.objectkamarfk=kmr.id " & _
                       " INNER join ruangan_m  as ru2 on ru2.id=apdp.objectruanganfk " & _
                       " where tp.norec  in (" & norec_apc & ") and pro.id <> 402611  " & strFilter
            
            ReadRs strSQL
            
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            
            .database.AddADOCommand CN_String, adoReport
            If RS.BOF Then
                .txtUmur.SetText "-"
            Else
                .txtUmur.SetText hitungUmur(Format(RS!tgllahir, "yyyy/MM/dd"), Format(Now, "yyyy/MM/dd"))
            End If
            
            .udtgl.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usnmpasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")

            .usUnitLayanan.SetUnboundFieldSource ("{ado.ruanganperiksa}")
            .usTipe.SetUnboundFieldSource ("{ado.kelompokpasien}")

            .usRujukan.SetUnboundFieldSource ("{ado.asalrujukan}")
            .usruangperiksa.SetUnboundFieldSource ("{ado.ruangakhir}")
            .usQty.SetUnboundFieldSource ("{ado.jumlah}")
            .usKamar.SetUnboundFieldSource ("if isnull({ado.namakamar}) then "" - "" else {ado.namakamar} ")
            .usKelas.SetUnboundFieldSource ("if isnull({ado.namakelas}) then "" - "" else {ado.namakelas} ") '("{ado.namakelas}")
            .usPenjamin.SetUnboundFieldSource ("{ado.namapenjamin}")

            .usDokter.SetUnboundFieldSource ("{ado.namadokter}")
            .udTglPelayanan.SetUnboundFieldSource ("{ado.tglpelayanan}")
            .usPelayanan.SetUnboundFieldSource ("{ado.namaproduk}")
            .ucTarif.SetUnboundFieldSource ("{ado.hargasatuan}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
            .usJumlah.SetUnboundFieldSource ("{ado.jumlah}")
'            .ucTotal.SetUnboundFieldSource ("{ado.total}")

            ReadRs2 "SELECT namalengkap FROM pegawai_m where id='" & strIdPegawai & "' "
            If RS2.BOF Then
                .txtuser.SetText "-"
            Else
                .txtuser.SetText UCase(IIf(IsNull(RS2("namalengkap")), "-", RS2("namalengkap")))
            End If
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "BuktiLayananRuanganPerTindakan")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportBuktiLayananRuanganPerTindakan
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:
    MsgBox Err.Number & " " & Err.Description
End Sub

Public Sub cetakBuktiLayananJasa(strNorec As String, strIdPegawai As String, strIdRuangan As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
Dim umur As String
Dim strFilter As String
Dim strFilter2 As String

bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolBuktiLayananJasa = True
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = False


    Dim strarr() As String
    Dim norec_apc As String
    Dim i As Integer
    
    
    strarr = Split(strNorec, "|")
    For i = 0 To UBound(strarr)
       norec_apc = norec_apc + "'" & strarr(i) & "',"
    Next
    norec_apc = Left(norec_apc, Len(norec_apc) - 1)
    
    strSQL = ""
    strFilter = ""
    strFilter2 = ""
'    If strIdRuangan <> "" Then strFilter = " AND ru2.id = '" & strIdRuangan & "' "
'    If strIdTindakan <> "" Then strFilter2 = " AND tp.produkfk = '" & strIdTindakan & "' "
    strFilter = strFilter & strFilter2 & " ORDER BY tp.tglpelayanan "
    With reportBuktiLayananJasa
    
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT pd.noregistrasi,ps.nocm,ps.tgllahir,ps.namapasien, " & _
                       " apdp.tglregistrasi,jk.reportdisplay AS jk,ru2.namaruangan AS ruanganperiksa,ru.namaruangan AS ruangakhir, " & _
                       " case when ru.objectdepartemenfk =16 then (select pg.namalengkap from pegawai_m as pg INNER JOIN pelayananpasienpetugas_t p3 on p3.objectpegawaifk=pg.id " & _
                       "where p3.pelayananpasien=tp.norec and p3.objectjenispetugaspefk=4 limit 1) else pp.namalengkap end AS namadokter,kp.kelompokpasien,tp.produkfk, " & _
                       " pro.namaproduk,tp.jumlah," & _
                       " (select case when hargajual is null then 0 else hargajual end as hargajual from pelayananpasiendetail_t where pelayananpasien=tp.norec and komponenhargafk=35 limit 1) as hargasatuan,(select case when hargadiscount is null then 0 else hargadiscount end as hargadiscount from pelayananpasiendetail_t where pelayananpasien=tp.norec and komponenhargafk=35 limit 1) as diskon, " & _
                       " ks.namakelas,ar.asalrujukan,tp.tglpelayanan, " & _
                       " CASE WHEN rek.namarekanan is null then '-' else rek.namarekanan END as namapenjamin, " & _
                       " CASE WHEN kmr.namakamar is null then '-' else kmr.namakamar END as namakamar " & _
                       " FROM pasiendaftar_t AS pd INNER JOIN pasien_m AS ps ON pd.nocmfk = ps.id " & _
                       " INNER JOIN jeniskelamin_m AS jk ON ps.objectjeniskelaminfk = jk.id " & _
                       " INNER JOIN kelompokpasien_m AS kp ON pd.objectkelompokpasienlastfk = kp.id " & _
                       " INNER JOIN antrianpasiendiperiksa_t AS apdp ON apdp.noregistrasifk = pd.norec " & _
                       " INNER JOIN ruangan_m AS ru ON apdp.objectruanganfk = ru.id " & _
                       " LEFT JOIN pelayananpasien_t AS tp ON tp.noregistrasifk = apdp.norec " & _
                       " LEFT JOIN pegawai_m AS pp ON apdp.objectpegawaifk = pp.id " & _
                       " LEFT JOIN produk_m AS pro ON tp.produkfk = pro.id " & _
                       " LEFT JOIN kelas_m AS ks ON apdp.objectkelasfk = ks.id " & _
                       " LEFT JOIN asalrujukan_m AS ar ON apdp.objectasalrujukanfk = ar.id " & _
                       " left JOIN rekanan_m AS rek ON rek.id= pd.objectrekananfk " & _
                       " left JOIN kamar_m as kmr on apdp.objectkamarfk=kmr.id " & _
                       " INNER join ruangan_m  as ru2 on ru2.id=apdp.objectruanganfk " & _
                       " where tp.norec  in (" & norec_apc & ") and pro.id <> 402611  " & strFilter
            
            ReadRs strSQL
            
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            
            .database.AddADOCommand CN_String, adoReport
            If RS.BOF Then
                .txtUmur.SetText "-"
            Else
                .txtUmur.SetText hitungUmur(Format(RS!tgllahir, "yyyy/MM/dd"), Format(Now, "yyyy/MM/dd"))
            End If
            
            .udtgl.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usnmpasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")

            .usUnitLayanan.SetUnboundFieldSource ("{ado.ruanganperiksa}")
            .usTipe.SetUnboundFieldSource ("{ado.kelompokpasien}")

            .usRujukan.SetUnboundFieldSource ("{ado.asalrujukan}")
            .usruangperiksa.SetUnboundFieldSource ("{ado.ruangakhir}")
            .usQty.SetUnboundFieldSource ("{ado.jumlah}")
            .usKamar.SetUnboundFieldSource ("if isnull({ado.namakamar}) then "" - "" else {ado.namakamar} ")
            .usKelas.SetUnboundFieldSource ("if isnull({ado.namakelas}) then "" - "" else {ado.namakelas} ") '("{ado.namakelas}")
            .usPenjamin.SetUnboundFieldSource ("{ado.namapenjamin}")

            .usDokter.SetUnboundFieldSource ("if isnull({ado.namadokter}) then "" - "" else {ado.namadokter} ") '("{ado.namadokter}")
            .udTglPelayanan.SetUnboundFieldSource ("{ado.tglpelayanan}")
            .usPelayanan.SetUnboundFieldSource ("{ado.namaproduk}")
            .ucTarif.SetUnboundFieldSource ("if isnull({ado.hargasatuan}) then 0 else {ado.hargasatuan} ") '("{ado.hargasatuan}")
            .ucDiskon.SetUnboundFieldSource ("if isnull({ado.diskon}) then 0 else {ado.diskon} ") '("{ado.diskon}")
            .usJumlah.SetUnboundFieldSource ("if isnull({ado.jumlah}) then 0 else {ado.jumlah} ") '("{ado.jumlah}")
'            .ucTotal.SetUnboundFieldSource ("{ado.total}")

            ReadRs2 "SELECT namalengkap FROM pegawai_m where id='" & strIdPegawai & "' "
            If RS2.BOF Then
                .txtuser.SetText "-"
            Else
                .txtuser.SetText UCase(IIf(IsNull(RS2("namalengkap")), "-", RS2("namalengkap")))
            End If
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "BuktiLayananRuanganPerTindakan")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportBuktiLayananJasa
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:
    MsgBox Err.Number & " " & Err.Description
End Sub
Public Sub cetakBuktiLayananRuanganBedah(strNorec As String, strIdPegawai As String, strIdRuangan As String, view As String)
On Error GoTo errLoad
Set frmCetakPendaftaran = Nothing
Dim strSQL As String
Dim umur As String
Dim strFilter As String
Dim strFilter2 As String

bolBuktiPendaftaran = False
bolBuktiLayanan = False
bolBuktiLayananRuangan = False
bolBuktiLayananRuanganPerTindakan = False
bolcetakSep = False
bolTracer1 = False
bolKartuPasien = False
boolLabelPasien = False
boolLabelPasienZebra = False
boolSumList = False
boolLembarRMK = False
boolLembarPersetujuan = False
bolBuktiLayananRuanganBedah = True


    Dim strarr() As String
    Dim norec_apc As String
    Dim i As Integer
    
    
    strarr = Split(strNorec, "|")
    For i = 0 To UBound(strarr)
       norec_apc = norec_apc + "'" & strarr(i) & "',"
    Next
    norec_apc = Left(norec_apc, Len(norec_apc) - 1)
    
    strSQL = ""
    strFilter = ""
    strFilter2 = ""
'    If strIdRuangan <> "" Then strFilter = " AND ru2.id = '" & strIdRuangan & "' "
'    If strIdTindakan <> "" Then strFilter2 = " AND tp.produkfk = '" & strIdTindakan & "' "
    strFilter = strFilter & strFilter2 & " ORDER BY tp.tglpelayanan "
    With reportBuktiLayananRuanganBedah
    
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "SELECT pd.noregistrasi,ps.nocm,ps.tgllahir,ps.namapasien, " & _
                       " apdp.tglregistrasi,jk.reportdisplay AS jk,ru2.namaruangan AS ruanganperiksa,ru.namaruangan AS ruangakhir, " & _
                       " case when ru.objectdepartemenfk =16 then (select pg.namalengkap from pegawai_m as pg INNER JOIN pelayananpasienpetugas_t p3 on p3.objectpegawaifk=pg.id " & _
                       "where p3.pelayananpasien=tp.norec and p3.objectjenispetugaspefk=4 limit 1) else pp.namalengkap end AS namadokter,kp.kelompokpasien,tp.produkfk, " & _
                       " pro.namaproduk,tp.jumlah,CASE WHEN tp.hargasatuan is null then tp.hargajual else tp.hargasatuan END as hargasatuan," & _
                       " (case when tp.hargadiscount is null then 0 else tp.hargadiscount end) as diskon, " & _
                       " hargasatuan*tp.jumlah as total,ks.namakelas,ar.asalrujukan,tp.tglpelayanan, " & _
                       " CASE WHEN rek.namarekanan is null then '-' else rek.namarekanan END as namapenjamin,kmr.namakamar " & _
                       " FROM pasiendaftar_t AS pd INNER JOIN pasien_m AS ps ON pd.nocmfk = ps.id " & _
                       " INNER JOIN jeniskelamin_m AS jk ON ps.objectjeniskelaminfk = jk.id " & _
                       " INNER JOIN kelompokpasien_m AS kp ON pd.objectkelompokpasienlastfk = kp.id " & _
                       " INNER JOIN antrianpasiendiperiksa_t AS apdp ON apdp.noregistrasifk = pd.norec " & _
                       " INNER JOIN ruangan_m AS ru ON apdp.objectruanganfk = ru.id " & _
                       " LEFT JOIN pelayananpasien_t AS tp ON tp.noregistrasifk = apdp.norec " & _
                       " LEFT JOIN pegawai_m AS pp ON apdp.objectpegawaifk = pp.id " & _
                       " LEFT JOIN produk_m AS pro ON tp.produkfk = pro.id " & _
                       " LEFT JOIN kelas_m AS ks ON apdp.objectkelasfk = ks.id " & _
                       " LEFT JOIN asalrujukan_m AS ar ON apdp.objectasalrujukanfk = ar.id " & _
                       " left JOIN rekanan_m AS rek ON rek.id= pd.objectrekananfk " & _
                       " left JOIN kamar_m as kmr on apdp.objectkamarfk=kmr.id " & _
                       " INNER join ruangan_m  as ru2 on ru2.id=apdp.objectruanganfk " & _
                       " where tp.norec  in (" & norec_apc & ") and pro.id <> 402611  " & strFilter
            
            ReadRs strSQL
            
            adoReport.CommandText = strSQL
            adoReport.CommandType = adCmdUnknown
            
            
            .database.AddADOCommand CN_String, adoReport
            If RS.BOF Then
                .txtUmur.SetText "-"
            Else
                .txtUmur.SetText hitungUmur(Format(RS!tgllahir, "yyyy/MM/dd"), Format(Now, "yyyy/MM/dd"))
            End If
            
            .udtgl.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCm.SetUnboundFieldSource ("{ado.nocm}")
            .usnmpasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jk}")

            .usUnitLayanan.SetUnboundFieldSource ("{ado.ruanganperiksa}")
            .usTipe.SetUnboundFieldSource ("{ado.kelompokpasien}")

            .usRujukan.SetUnboundFieldSource ("{ado.asalrujukan}")
            .usruangperiksa.SetUnboundFieldSource ("{ado.ruangakhir}")
            .usQty.SetUnboundFieldSource ("{ado.jumlah}")
            .usKamar.SetUnboundFieldSource ("if isnull({ado.namakamar}) then "" - "" else {ado.namakamar} ")
            .usKelas.SetUnboundFieldSource ("if isnull({ado.namakelas}) then "" - "" else {ado.namakelas} ") '("{ado.namakelas}")
            .usPenjamin.SetUnboundFieldSource ("{ado.namapenjamin}")

            .usDokter.SetUnboundFieldSource ("{ado.namadokter}")
            .udTglPelayanan.SetUnboundFieldSource ("{ado.tglpelayanan}")
            .usPelayanan.SetUnboundFieldSource ("{ado.namaproduk}")
            .ucTarif.SetUnboundFieldSource ("{ado.hargasatuan}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
            .usJumlah.SetUnboundFieldSource ("{ado.jumlah}")
'            .ucTotal.SetUnboundFieldSource ("{ado.total}")

            ReadRs3 "select " & _
                    "sum(case when ppd.komponenhargafk=38 then ppd.hargajual*ppd.jumlah end) as jasasarana, " & _
                    "sum(case when ppd.komponenhargafk=35 then ppd.hargajual*ppd.jumlah end) as jasamedis, " & _
                    "sum(case when ppd.komponenhargafk=25 then ppd.hargajual*ppd.jumlah end) as jasaparamedis, " & _
                    "sum(case when ppd.komponenhargafk=30 then ppd.hargajual*ppd.jumlah end) as jasaumum, " & _
                    "sum(case when ppd.komponenhargafk=21 then ppd.hargajual*ppd.jumlah end) as anestesidr, " & _
                    "sum(case when ppd.komponenhargafk=22 then ppd.hargajual*ppd.jumlah end) as jasaspesialis, " & _
                    "sum(case when ppd.komponenhargafk=26 then ppd.hargajual*ppd.jumlah end) as jasaperawatanastesi, " & _
                    "sum(case when ppd.komponenhargafk=27 then ppd.hargajual*ppd.jumlah end) as jasaperawatinstr " & _
                    "from pasiendaftar_t as pd " & _
                    "inner join antrianpasiendiperiksa_t as apdp on apdp.noregistrasifk = pd.norec " & _
                    "left join pelayananpasien_t as tp on tp.noregistrasifk = apdp.norec " & _
                    "left join pelayananpasiendetail_t as ppd on ppd.pelayananpasien=tp.norec " & _
                    "left join produk_m as pro on tp.produkfk = pro.id " & _
                    " where tp.norec  in (" & norec_apc & ") and pro.id <> 402611 "
            
            If RS3.BOF = False Then
                .ucJasaSarana.SetUnboundFieldSource UCase(IIf(IsNull(RS3("jasasarana")), "0.00", RS3("jasasarana")))
                .ucJasaMedis.SetUnboundFieldSource UCase(IIf(IsNull(RS3("jasamedis")), "0.00", RS3("jasamedis")))
                .ucJasaParamedis.SetUnboundFieldSource UCase(IIf(IsNull(RS3("jasaparamedis")), "0.00", RS3("jasaparamedis")))
                .ucJasaUmum.SetUnboundFieldSource UCase(IIf(IsNull(RS3("jasaumum")), "0.00", RS3("jasaumum")))
                .ucAnestesiDr.SetUnboundFieldSource UCase(IIf(IsNull(RS3("anestesidr")), "0.00", RS3("anestesidr")))
                .ucJasaSpesialis.SetUnboundFieldSource UCase(IIf(IsNull(RS3("jasaspesialis")), "0.00", RS3("jasaspesialis")))
                .ucJasaPerawatAnastesi.SetUnboundFieldSource UCase(IIf(IsNull(RS3("jasaperawatanastesi")), "0.00", RS3("jasaperawatanastesi")))
                .ucJasaPerawatInstr.SetUnboundFieldSource UCase(IIf(IsNull(RS3("jasaperawatinstr")), "0.00", RS3("jasaperawatinstr")))
            End If

            ReadRs2 "SELECT namalengkap FROM pegawai_m where id='" & strIdPegawai & "' "
            If RS2.BOF Then
                .txtuser.SetText "-"
            Else
                .txtuser.SetText UCase(IIf(IsNull(RS2("namalengkap")), "-", RS2("namalengkap")))
            End If
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "BuktiLayananRuanganPerTindakan")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportBuktiLayananRuanganBedah
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:
    MsgBox Err.Number & " " & Err.Description
End Sub
