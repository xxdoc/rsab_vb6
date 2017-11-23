VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmLaporanJurnalBalik 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   5790
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
      Height          =   6975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5775
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
      EnableExportButton=   -1  'True
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
Attribute VB_Name = "frmLaporanJurnalBalik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crLaporanJurnalBalik2
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

    Set frmLaporanJurnalBalik = Nothing
End Sub

Public Sub CetakLaporanJurnalBalik(idKasir As String, tglAwal As String, tglAkhir As String, idDepartemen As String, idRuangan As String, namaPrinted As String, view As String)
On Error GoTo errLoad
'On Error Resume Next

Set frmLaporanJurnalBalik = Nothing
Dim adocmd As New ADODB.Command

    Dim str1 As String
    Dim str2 As String
    Dim strFilter, strFilter2, strFilter3 As String

    strFilter = ""
    
    strFilter = " where tgl BETWEEN '" & _
    Format(tglAwal, "yyyy-MM-dd") & "' AND '" & _
    Format(tglAkhir, "yyyy-MM-dd") & "'"
    
       If idDepartemen <> "" Then
           strFilter = strFilter & " AND departemen = '" & idDepartemen & "' "
       End If
    
    'strFilter = strFilter & " GROUP BY pd.tglregistrasi"
    
Set Report = New crLaporanJurnalBalik2

'    strSQL = "select tgl, umum, perusahaan,bpjs,diskon,total,departemen,ruangandiskon,keterangan from v_jurnal_balik " & _
'            strFilter
            
    strSQL = " SELECT v_jurnal_balik_1.tgl, v_jurnal_balik_1.umum, v_jurnal_balik_1.perusahaan, " & _
             " v_jurnal_balik_1.bpjs, v_jurnal_balik_1.diskon, v_jurnal_balik_1.total, v_jurnal_balik_1.departemen, " & _
             " v_jurnal_balik_1.ruangandiskon, v_jurnal_balik_1.keterangan" & _
             " From v_jurnal_balik_1 " & _
             " Union All " & _
             " SELECT v_jurnal_balik_2.tgl, v_jurnal_balik_2.umum, v_jurnal_balik_2.perusahaan, " & _
             " v_jurnal_balik_2.bpjs, v_jurnal_balik_2.diskon,v_jurnal_balik_2.total, " & _
             " v_jurnal_balik_2.departemen,v_jurnal_balik_2.ruangandiskon, " & _
             " v_jurnal_balik_2.keterangan" & _
             " FROM v_jurnal_balik_2  " & _
               strFilter

    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
            .txtPrinted.SetText namaPrinted
            .txtTanggal.SetText Format(tglAwal, "dd/MM/yyyy")
            .txtPeriode.SetText Format(tglAwal, "MM-yyyy")
            .ucUmum.SetUnboundFieldSource ("{ado.umum}")
            .ustgl.SetUnboundFieldSource ("{ado.tgl}")
            .ucPerusahaan.SetUnboundFieldSource ("{ado.perusahaan}")
            .ucBpjs.SetUnboundFieldSource ("{ado.bpjs}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
            .ucTotal.SetUnboundFieldSource ("{ado.total}")
            .usNamaPerkiraan.SetUnboundFieldSource ("{ado.keterangan}")
''            .usRuanganDiskon.SetUnboundFieldSource ("{ado.ruangandiskon}")
'            .txtTglDeskripsi.SetText Format(tglAwal, "dd/MM/yyyy")
''            .ucDebet.SetUnboundFieldSource ("{ado.tunai}")
''            .ucKredit.SetUnboundFieldSource ("{ado.nontunai}")
'
            '.udPeriode.SetUnboundFieldSource RS2!tgl
            '.usTanggal.SetUnboundFieldSource Format(tglAwal, "dd/MM/yyyy")
            '.SetUnboundFieldSource Format(RS2!tgl, "MM-yyyy")
            '.usTglDeskripsi.SetUnboundFieldSource Format(tglAwal, "dd/MM/yyyy")

'            .ucDetailDiskon.SetUnboundFieldSource ("{ado.detaildiskon}"

           
            
            If idDepartemen = 16 Then
                .txtDeskripsi.SetText "Pendapatan R. Inap Non Tunai Tgl " + ("{ado.tgl}")
                .txtKeterangan1.SetText "Pendapatan R.Inap"
                .txtKeterangan2.SetText "Pendapatan R.Inap"
                .txtKeterangan3.SetText "Pendapatan R.Inap"
                .txtKeterangan4.SetText "Pendapatan R.Inap"
                .txtKeterangan5.SetText "Pendapatan R.Inap"
            Else
                .txtDeskripsi.SetText "Pendapatan R. Jalan Non Tunai Tgl " + ("{ado.tgl}")
                .txtKeterangan1.SetText "Pendapatan R.Jalan"
                .txtKeterangan2.SetText "Pendapatan R.Jalan"
                .txtKeterangan3.SetText "Pendapatan R.Jalan"
                .txtKeterangan4.SetText "Pendapatan R.Jalan"
                .txtKeterangan5.SetText "Pendapatan R.Jalan"
            End If
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "LaporanJurnal")
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
    MsgBox Err.Number & " " & Err.Description
End Sub
