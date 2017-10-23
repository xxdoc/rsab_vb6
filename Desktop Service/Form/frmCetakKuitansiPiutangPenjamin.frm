VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakKuitansiPiutangPenjamin 
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
Attribute VB_Name = "frmCetakKuitansiPiutangPenjamin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crKuitansiPiutangPenjamin
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

    Set frmCetakKuitansiPiutangPenjamin = Nothing
End Sub

Public Sub CetakKuitansiPiutangPenjamin(idKasir As String, tglAwal As String, tglAkhir As String, idKelompok As String, idPenjamin As String, login As String, view As String)
'On Error GoTo errLoad

Set frmCetakKuitansiPiutangPenjamin = Nothing

Dim adocmd As New ADODB.Command
   ' Dim str1 As String
    
    'If idPerusahaan <> "" Then
       ' str1 = "and rk.id=" & idPerusahaan & " "
   ' End If
    
   ' Set Report = New crKuitansiPiutangPenjamin
       ' strSQL = "SELECT (case when sp.totalprekanan is null then 0 else sp.totalprekanan end) as totalppenjamin, " & _
            '    "case when rk.namarekanan is null then '-' else rk.namarekanan end as namarekanan " & _
             '   "FROM strukpelayanan_t as sp " & _
             '   "left join pasiendaftar_t as pd on pd.norec=sp.noregistrasifk " & _
             '   "left JOIN rekanan_m  as rk on rk.id=pd.objectrekananfk " & _
              '  "INNER JOIN kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
               ' "where sp.tglstruk BETWEEN '2017-10-01' and '2017-10-19' and kps.id in (1,3,5) " & _
               ' str1

    Dim strFilter As String
    Dim orderby As String
    
    strFilter = ""
    orderby = ""

    strFilter = " where pd.tglpulang BETWEEN '" & _
    tglAwal & "' AND '" & _
    tglAkhir & "'"
'    strFilter = strFilter & " and IdRuangan like '%" & strIdRuangan & "%' and IdDepartement like '%" & strIdDepartement & "%' and IdKelompokPasien like '%" & strIdKelompokPasien & "%' and IdDokter Like '%" & strIdDokter & "%'"
    
    If idKelompok <> "" Then strFilter = strFilter & " AND pd.objectkelompokpasienlastfk = '" & idKelompok & "' "
    If idPenjamin <> "" Then strFilter = strFilter & " AND rk.id = '" & idPenjamin & "' "
            
    orderby = strFilter & "group by pd.noregistrasi, sp.totalharusdibayar,sp.totalprekanan ,pd.tglpulang " & _
            "order by pd.tglpulang"
            
    Set Report = New crKuitansiPiutangPenjamin
    
    ReadRs2 "select pd.noregistrasi, " & _
            "(case when sp.totalharusdibayar is null then 0 else sp.totalharusdibayar end) as cash, " & _
            "(case when sp.totalprekanan is null then 0 else sp.totalprekanan end) as totalpiutangpenjamin " & _
            "from strukpelayanan_t as sp " & _
            "left JOIN pelayananpasien_t as pp on pp.strukfk=sp.norec  " & _
            "inner JOIN antrianpasiendiperiksa_t as apd on apd.norec=pp.noregistrasifk  " & _
            "inner JOIN pasiendaftar_t as pd on pd.norec=apd.noregistrasifk  " & _
            "INNER JOIN kelompokpasien_m as klp on klp.id=pd.objectkelompokpasienlastfk " & _
            "left join rekanan_m  as rk on rk.id=pd.objectrekananfk " & orderby
    
    Dim tCash, tPiutang As Double
    Dim i As Integer
    
    For i = 0 To RS2.RecordCount - 1
        tCash = tCash + CDbl(IIf(IsNull(RS2!cash), 0, RS2!cash))
        tPiutang = tPiutang + CDbl(IIf(IsNull(RS2!totalpiutangpenjamin), 0, RS2!totalpiutangpenjamin))
        
        RS2.MoveNext
    Next
    
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
            .txtPrinted.SetText login
            '.txtPeriode.SetText "Periode : " & tglAwal & " s/d " & tglAkhir & ""
            '.usNamaPenjamin.SetUnboundFieldSource ("{ado.namarekanan}")
            .ucJumlah.SetUnboundFieldSource (tPiutang)
            .ucMaterai.SetUnboundFieldSource 3000
            
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

