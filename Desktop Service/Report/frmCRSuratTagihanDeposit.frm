VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCRSuratTagihanDeposit 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmCRSuratTagihanDeposit.frx":0000
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
Attribute VB_Name = "frmCRSuratTagihanDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New cr_SuratTagihanDeposit
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
    cboPrinter.Text = GetTxt("Setting.ini", "Printer", "LaporanPasienPulang")
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCRSuratTagihanDeposit = Nothing
End Sub

Public Sub Cetak(noregistrasi As String, view As String)
On Error GoTo errLoad
'On Error Resume Next

Set frmCRSuratTagihanDeposit = Nothing
Dim adocmd As New ADODB.Command
Dim strFilter, orderby As String
Set Report = New cr_SuratTagihanDeposit
        
    strSQL = "select to_char(tglregistrasi, 'dd')as hari,to_char(tglregistrasi, 'dd')as tgl, " & _
            "to_char(tglregistrasi, 'mm')as bln,to_char(tglregistrasi, 'yyyy')as thn, " & _
            "pd.tglregistrasi, ps.nocm, ps.namapasien, ru.namaruangan " & _
            "from pasiendaftar_t as pd " & _
            "INNER JOIN pasien_m as ps on ps.id=pd.nocmfk " & _
            "INNER JOIN ruangan_m as ru on ru.id=pd.objectruanganlastfk " & _
            "where pd.noregistrasi='" & noregistrasi & "'"
    
    ReadRs "select sum(((case when pp.hargajual is null then 0 else pp.hargajual  end - " & _
            "case when pp.hargadiscount is null then 0 else pp.hargadiscount end) * pp.jumlah) + " & _
            "case when pp.jasa is null then 0 else pp.jasa end) as totaltagihan " & _
            "from pasiendaftar_t as pd " & _
            "INNER JOIN antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
            "INNER JOIN pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
            "where pd.noregistrasi='" & noregistrasi & "' and pp.produkfk not in (402611)"
            
    ReadRs2 "SELECT case when pp.hargajual is null then 0 else pp.hargajual end as deposit " & _
            "from pasiendaftar_t as pd " & _
            "INNER join antrianpasiendiperiksa_t as apd on apd.noregistrasifk = pd.norec " & _
            "INNER join pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
            "where pd.noregistrasi='" & noregistrasi & "' and pp.produkfk=402611"
            
    Dim tSisa, tdeposit, ttagihan As Double
    ttagihan = RS!totaltagihan
    tdeposit = RS2!deposit
    tSisa = ttagihan - tdeposit
    If tSisa < 0 Then
        tSisa = 0
    End If
    
    ReadRs3 strSQL
    
    Dim day, hari, month, bln, tgl, thn As String
    
    If RS2.EOF Then
       day = "-"
    Else
        day = RS3!hari
        month = RS3!bln
        tgl = RS3!tgl
        thn = RS3!thn
        If day = "06" Then
            hari = "Minggu"
        ElseIf day = "07" Then
            hari = "Senin"
        ElseIf day = "01" Then
            hari = "Selasa"
        ElseIf day = "02" Then
            hari = "Rabu"
        ElseIf day = "03" Then
            hari = "Kamis"
        ElseIf day = "04" Then
            hari = "Jumat"
        ElseIf day = "01" Then
            hari = "Sabtu"
        End If
        
        If month = "01" Then
            bln = "Januari"
        ElseIf month = "02" Then
            bln = "Februari"
        ElseIf month = "03" Then
            bln = "Maret"
        ElseIf month = "04" Then
            bln = "April"
        ElseIf month = "05" Then
            bln = "Mei"
        ElseIf month = "06" Then
            bln = "Juni"
        ElseIf month = "07" Then
            bln = "Juli"
        ElseIf month = "08" Then
            bln = "Agustus"
        ElseIf month = "09" Then
            bln = "September"
        ElseIf month = "10" Then
            bln = "Oktober"
        ElseIf month = "11" Then
            bln = "November"
        ElseIf month = "12" Then
            bln = "Desember"
        Else
            bln = month
        End If
        
    End If
    
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
        'If Not RS.EOF Then
            
            .txtTglRegis.SetText hari + ", " + tgl + " " + bln + " " & thn
            .usNoMR.SetUnboundFieldSource ("{ado.nocm}")
            .usPasienHead.SetUnboundFieldSource ("{ado.namapasien}")
            .usPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usRuangRawat.SetUnboundFieldSource ("{ado.namaruangan}")
            .ucTotalTagihan.SetUnboundFieldSource ttagihan
            .ucTotalDeposit.SetUnboundFieldSource tdeposit
            .ucTotalSisa.SetUnboundFieldSource tSisa
            
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "LaporanPasienPulang")
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
