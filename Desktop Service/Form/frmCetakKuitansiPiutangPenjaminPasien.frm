VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakKuitansiPiutangPenjaminPasien 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   6390
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7455
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6375
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
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.TextBox txtNamaFormPengirim 
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmCetakKuitansiPiutangPenjaminPasien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crKuitansiPiutangPenjaminPasien
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

    Set frmCetakKuitansiPiutangPenjaminPasien = Nothing
End Sub

Public Sub CetakKuitansiPiutangPenjamin(idKasir As String, noregistrasi As String, view As String)
'On Error GoTo errLoad

Set frmCetakKuitansiPiutangPenjaminPasien = Nothing

Dim adocmd As New ADODB.Command
        
    Dim strFilter As String
    Dim orderby As String
    
    strFilter = ""
  
    strFilter = " where kps.id in (3,5) and pd.noregistrasi = '" & noregistrasi & "' "
            
    Set Report = New crKuitansiPiutangPenjaminPasien
    
    ReadRs2 "SELECT sp.tglstruk, ps.namapasien || '/ ' || ps.nocm || '/ ' || pd.noregistrasi as pasien, ru.namaruangan, " & _
            "(case when sp.totalprekanan is null then 0 else sp.totalprekanan end) as totalppenjamin, " & _
            "case when rk.namarekanan is null then '-' else rk.namarekanan end as namarekanan, " & _
            "case when rk.alamatlengkap is null then '-' else rk.alamatlengkap end as alamat " & _
            "FROM strukpelayanan_t as sp " & _
            "left join pelayananpasien_t as pp on pp.strukfk=sp.norec " & _
            "left JOIN antrianpasiendiperiksa_t as apd on apd.norec=pp.noregistrasifk " & _
            "left JOIN pasiendaftar_t as pd on pd.norec=apd.noregistrasifk " & _
            "left JOIN rekanan_m  as rk on rk.id=pd.objectrekananfk " & _
            "INNER JOIN kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
            "inner join pasien_m as ps on ps.id = pd.nocmfk " & _
            "inner join ruangan_m as ru on ru.id = sp.objectruanganfk " & _
            strFilter
            
    Dim tPiutang, tmaterai, X As Double
    Dim pasien, rekanan, alamat, ruangan As String
    Dim i As Integer
    
    'tmaterai = 3000
    
        tPiutang = tPiutang + CDbl(IIf(IsNull(RS2!totalppenjamin), 0, RS2!totalppenjamin))
        pasien = UCase(IIf(IsNull(RS2("pasien")), "-", RS2("pasien")))
        rekanan = UCase(IIf(IsNull(RS2("namarekanan")), "-", RS2("namarekanan")))
        ruangan = UCase(IIf(IsNull(RS2("namaruangan")), "Biaya Perawatan -", "Biaya Perawatan " & RS2("namaruangan")))
        
    With Report
        If Not RS2.BOF Then
            .txtRuangan.SetText ruangan
            .txtPasien.SetText pasien
            .txtPenjamin.SetText rekanan
            .ucJumlah.SetUnboundFieldSource tPiutang
            '.ucJumlah.SetUnboundFieldSource (IIf(IsNull(RS2!totalppenjamin), 0, RS2!totalppenjamin))
            
            X = Round(tPiutang)
            .txtPembulatan.SetText Format(X, "##,##0.00")
            .txtTerbilang.SetText "# " & TERBILANG(X) & " #"
        End If
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
    End With
Exit Sub
errLoad:
End Sub

