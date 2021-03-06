VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakKuitansiSbk 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmCetakKuitansiSbk.frx":0000
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
Attribute VB_Name = "frmCetakKuitansiSbk"
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
    Report.PrinterSetup Me.hwnd
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

    Set frmCetakKuitansiSbk = Nothing
End Sub

Public Sub CetakUlangJenisKuitansi(strNoSbk As String, jumlahCetak As Integer, strIdPegawai As String, STD As String, view As String)
On Error GoTo errLoad

Dim strKet As Boolean
Dim jenisKwitansi As String


    strKet = True
    
    Set frmCetakKuitansiSbk = Nothing
    Set Report = New crKuitansiPasien
            ReadRs "select sp.tglstruk,sp.tglstruk,sp.nostruk,sp.nofaktur, " & _
                    "pg.id as pgid,pg.namalengkap,sbk.* " & _
                    "from strukbuktipengeluaran_t as sbk " & _
                    "inner join strukpelayanan_t as sp on sp.norec = sbk.nostrukfk " & _
                    "left join loginuser_s as lu on lu.id = sbk.objectpegawaipembayarfk " & _
                    "left join pegawai_m as pg on pg.id = lu.objectpegawaifk " & _
                    "where sbk.nosbk= '" & strNoSbk & "' "
'        End If
'    End If
    
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
            .txtNoBKM.SetText RS("nosbk")
            .txtNamaPenyetor.SetText UCase(RS("namapasien"))
            .txtNamaPasien.SetText UCase(RS("namapasien"))
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
            If jenisKwitansi = "KEMBALIDEPOSIT" Then
                .txtPetugasKasir.SetText RS("namapasien")
            End If
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

