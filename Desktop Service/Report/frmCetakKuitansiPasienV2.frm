VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCRCetakKuitansiPasienV2 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmCetakKuitansiPasienV2.frx":0000
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
Attribute VB_Name = "frmCRCetakKuitansiPasienV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crKuitansiPasien
Dim bolSuppresDetailSection10 As Boolean
Dim ii As Integer
Dim tempPrint1 As String
Dim p As Printer
Dim p2 As Printer
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
    cboPrinter.Text = GetTxt("Setting.ini", "Printer", "Kwitansi")
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCRCetakKuitansiPasienV2 = Nothing
End Sub

Public Sub CetakUlangJenisKuitansi(strNoregistrasi As String, jumlahCetak As Integer, strIdPegawai As String, STD As String, view As String)
On Error GoTo errLoad

Dim strKet As Boolean
Dim jenisKwitansi As String


    strKet = True
    
    Set frmCRCetakKuitansiPasienV2 = Nothing
    Set Report = New crKuitansiPasien
    If Len(strNoregistrasi) = 10 Then
        ReadRs "select pd.noregistrasi,sbp.totaldibayar,ps.namapasien, sbp.keteranganlainnya,pd.nocmfk,ru.namaruangan,pg.namalengkap,ps.nocm from pasiendaftar_t as pd " & _
               "inner join strukpelayanan_t as sp on sp.noregistrasifk=pd.norec " & _
               "inner join strukbuktipenerimaan_t as sbp  on sbp.nostrukfk=sp.norec " & _
               "inner join pasien_m as ps on ps.id=pd.nocmfk " & _
               "inner join ruangan_m as ru on ru.id=pd.objectruanganlastfk " & _
               "inner join loginuser_s as lu on lu.id=sbp.objectpegawaipenerimafk " & _
               "inner join pegawai_m as pg on pg.id=lu.objectpegawaifk " & _
               "where pd.noregistrasi='" & strNoregistrasi & "'"
    End If
    If Len(strNoregistrasi) = 14 Then
        ReadRs "select sp.nostruk as noregistrasi,sp.totalharusdibayar as totaldibayar,sp.namapasien_klien as namapasien,pg.namalengkap, sp.keteranganlainnya,'Tindakan Non Layanan' as namaruangan,'-' as nocm from  " & _
               " strukpelayanan_t as sp  " & _
               "inner join strukbuktipenerimaan_t as sbp  on sbp.nostrukfk=sp.norec " & _
               "inner join loginuser_s as lu on lu.id=sbp.objectpegawaipenerimafk " & _
               "inner join pegawai_m as pg on pg.id=lu.objectpegawaifk " & _
               "where sbp.nosbm='" & strNoregistrasi & "'"
    End If
    If Len(strNoregistrasi) > 14 Then
        If Left(strNoregistrasi, 7) = "DEPOSIT" Then
            strNoregistrasi = Replace(strNoregistrasi, "DEPOSIT", "")
            ReadRs "select sp.nostruk as noregistrasi,sbp.totaldibayar as totaldibayar, ps.namapasien as namapasien, " & _
                    "pg.namalengkap, sbp.keteranganlainnya,sbp.keteranganlainnya as namaruangan,ps.nocm as nocm " & _
                    "from   strukpelayanan_t as sp " & _
                    "inner join strukbuktipenerimaan_t as sbp  on sbp.nostrukfk=sp.norec " & _
                    "left join loginuser_s as lu on lu.id=sbp.objectpegawaipenerimafk " & _
                    "left join pasien_m as ps on ps.id=sp.nocmfk " & _
                    "left join pegawai_m as pg on pg.id=lu.objectpegawaifk " & _
                    "where sbp.nosbm='" & strNoregistrasi & "'"
            strKet = False
'            ReadRs "select sp.nostruk as noregistrasi,sp.totalharusdibayar as totaldibayar,sp.namapasien_klien as namapasien,pg.namalengkap, sp.keteranganlainnya,sp.keteranganlainnya as namaruangan,'-' as nocm from  " & _
'               " strukpelayanan_t as sp  " & _
'               "inner join strukbuktipenerimaan_t as sbp  on sbp.nostrukfk=sp.norec " & _
'               "left join loginuser_s as lu on lu.id=sbp.objectpegawaipenerimafk " & _
'               "left join pegawai_m as pg on pg.id=lu.objectpegawaifk " & _
'               "where sbp.nosbm='" & strNoregistrasi & "'"
        ElseIf Left(strNoregistrasi, 14) = "KEMBALIDEPOSIT" Then
            jenisKwitansi = "KEMBALIDEPOSIT"
            strNoregistrasi = Replace(strNoregistrasi, "KEMBALIDEPOSIT", "")
            ReadRs "select sbp.nosbm as noregistrasi,sbp.totaldibayar as totaldibayar,  namapasien, " & _
                    "pg.namalengkap, sbp.keteranganlainnya,sbp.keteranganlainnya as namaruangan,ps.nocm as nocm " & _
                    "from strukpelayanan_t as sp inner join strukbuktipenerimaan_t as sbp  on sbp.nostrukfk=sp.norec " & _
                    "left join loginuser_s as lu on lu.id=sbp.objectpegawaipenerimafk " & _
                    "left join pasien_m as ps on ps.id=sp.nocmfk " & _
                    "left join pegawai_m as pg on pg.id=lu.objectpegawaifk " & _
                    "where sbp.nosbm='" & strNoregistrasi & "'"
            strKet = False
        Else
            Dim noreg, nostruk As String
            noreg = Left(strNoregistrasi, 10)
            nostruk = Replace(strNoregistrasi, noreg, "")
            ReadRs "select pd.noregistrasi,sbp.totaldibayar,ps.namapasien, sbp.keteranganlainnya,pd.nocmfk,ru.namaruangan,pg.namalengkap,ps.nocm from pasiendaftar_t as pd " & _
                   "inner join strukpelayanan_t as sp on sp.noregistrasifk=pd.norec " & _
                   "inner join strukbuktipenerimaan_t as sbp  on sbp.nostrukfk=sp.norec " & _
                   "inner join pasien_m as ps on ps.id=pd.nocmfk " & _
                   "inner join ruangan_m as ru on ru.id=pd.objectruanganlastfk " & _
                   "inner join loginuser_s as lu on lu.id=sbp.objectpegawaipenerimafk " & _
                   "inner join pegawai_m as pg on pg.id=lu.objectpegawaifk " & _
                   "where pd.noregistrasi='" & noreg & "' and sp.norec='" & nostruk & "'"
        End If
    End If
    
    Dim i As Integer
    Dim jumlahDuit As Double
    Dim kembaliDeposit As Boolean
    
    For i = 0 To RS.RecordCount - 1
        jumlahDuit = jumlahDuit + CDbl(RS!totaldibayar)
        RS.MoveNext
        
    Next
    RS.MoveFirst
    
    kembaliDeposit = False
    If jumlahDuit < 0 Then
        kembaliDeposit = True
    End If
    
    With Report
        If Not RS.EOF Then
            .txtNoBKM.SetText RS("noregistrasi")
            If STD = "" Then
                .txtNamaPenyetor.SetText UCase(RS("namapasien"))
            Else
                .txtNamaPenyetor.SetText UCase(STD)
            End If
            .txtNamaPasien.SetText UCase(RS("namapasien"))
            If jenisKwitansi = "KEMBALIDEPOSIT" Then
                .txtNamaPenyetor.SetText "RSAB HARAPAN KITA"
            End If
            If strKet = True Then
                .txtKeterangan.SetText UCase("Biaya Layanan " & RS("namaruangan"))  'RS("keteranganlainnya")
            Else
                If kembaliDeposit = False Then
                    .txtKeterangan.SetText UCase(RS("namaruangan"))  'RS("keteranganlainnya")
                Else
                    .txtKeterangan.SetText Replace(UCase(RS("namaruangan")), "PEMBAYARAN", "PENGEMBALIAN")
                    jumlahDuit = jumlahDuit * (-1)
                End If
            End If
'            .txtKeterangan.SetText "Biaya Perawatan Pasien"
            .txtRp.SetText "Rp. " & Format(jumlahDuit, "##,##0.00")
'            .txtRp.SetText "Rp. " & Format(11789104, "##,##0.00")
            .txtTerbilang.SetText TerbilangDesimal(CStr(jumlahDuit))
            .txtRuangan.SetText UCase(RS("namaruangan"))
            .txtNoPen2.SetText RS("noregistrasi")
            .txtNoCM2.SetText RS("nocm")
            .txtPrintTglBKM.SetText "Jakarta, " & Format(Now(), "dd MMM yyyy")
            .txtPetugasKasir.SetText RS("namalengkap")
            .txtDesc.SetText UCase("NAMA/MR/No.REG  : " & RS("namapasien") & "/ " & RS("nocm") & "/ " & RS("noregistrasi"))
'            .txtDesc.SetText UCase("NAMA/MR/No.REG  : " & RS("namapasien") & "/ " & RS("nocm") & "/ " & "1711001100")
            .txtPetugasCetak.SetText strIdPegawai
            
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "Kwitansi")
                Report.SelectPrinter "winspool", strPrinter, "Ne00:"
                Report.PrintOut False
                Unload Me
            Else
                With CRViewer1
                    .ReportSource = Report
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
            End If
        End If
    End With
Exit Sub
errLoad:
End Sub

