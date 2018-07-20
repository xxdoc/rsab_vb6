VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakSPK 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "frmCetakSPK.frx":0000
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
Attribute VB_Name = "frmCetakSPK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportResep As New crSuratPerintahKerja

Dim ii As Integer
Dim tempPrint1 As String
Dim p As Printer
Dim p2 As Printer
Dim strDeviceName As String
Dim strDriverName As String
Dim strPort As String

Dim bolStrukResep As Boolean


Dim strPrinter As String
Dim strPrinter1 As String
Dim PrinterNama As String

Dim adoReport As New ADODB.Command

Private Sub cmdCetak_Click()
  If cboPrinter.Text = "" Then MsgBox "Printer belum dipilih", vbInformation, ".: Information": Exit Sub
    If bolStrukResep = True Then
        ReportResep.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
        PrinterNama = cboPrinter.Text
        ReportResep.PrintOut False
    
    End If
End Sub

Private Sub CmdOption_Click()
    
    If bolStrukResep = True Then
        ReportResep.PrinterSetup Me.hWnd
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

    Set frmCetakSPK = Nothing

End Sub

Public Sub Cetak(strNorec As String, view As String)
'On Error GoTo errLoad
Set frmCetakSPK = Nothing
Dim strSQL As String
Dim str1, str2, str3 As String

bolStrukResep = True
    
    
        With ReportResep
            Set adoReport = New ADODB.Command
            adoReport.ActiveConnection = CN_String
            
            strSQL = "select so.keteranganlainnya,so.tglvalidasi,so.nourutlogin,so.keterangankeperluan,so.noorder, so.keterangankeperluan,so.noorderintern, so.tglorder, so.keteranganorder, " & _
                    "so.nokontrakspk,so.noorderrfq,so.namarekanansales,so.alamat,so.alamattempattujuan,so.keteranganorder || ' RSAB HK THN '|| so.noorderrfq as judul, " & _
                    "pr.namaproduk, ss.satuanstandar, op.hargasatuan, op.qtyproduk, op.hargadiscount,op.hargappn,so.tglkontrak, " & _
                    "case when op.hargadiscount <> 0 then (op.hargasatuan * op.qtyproduk) / op.hargadiscount else 0 end as persenDisc, " & _
                    "case when op.hargappn <> 0 then (op.hargasatuan * op.qtyproduk) / op.hargappn else 0 end as persenPpn, " & _
                    "(op.hargasatuan * op.qtyproduk)-(hargadiscount+hargappn)as total,pg.namalengkap,pg1.namalengkap as pegawaispk,pg.nippns, " & _
                    "pg1.nippns as nippns_spk,rk.namarekanan,rk.alamatlengkap,op.deskripsiprodukquo " & _
                    "from strukorder_t so " & _
                    "left join orderpelayanan_t op on op.strukorderfk=so.norec " & _
                    "left join produk_m pr on pr.id=op.objectprodukfk " & _
                    "left join satuanstandar_m ss on ss.id=op.objectsatuanstandarfk " & _
                    "left join pegawai_m as pg on pg.id = so.objectpegawaiorderfk " & _
                    "left join pegawai_m as pg1 on pg1.id = so.objectpegawaispkfk " & _
                    "left join rekanan_m as rk on rk.id = so.objectrekananfk " & _
                    "where so.norec = '" & strNorec & "'"
             ReadRs strSQL
             If RS.EOF = False Then
                str1 = RS!namalengkap
                str2 = RS!nippns
                str3 = RS!noorderintern
             Else
                str1 = "-"
                str2 = "-"
                str3 = "-"
             End If
             
             adoReport.CommandText = strSQL
             adoReport.CommandType = adCmdUnknown
            .database.AddADOCommand CN_String, adoReport
                
             .txtNoKontrak.SetText str3
             .usNoUsulan.SetUnboundFieldSource ("{Ado.nokontrakspk}")
             .udTglSpk.SetUnboundFieldSource ("{Ado.tglkontrak}")
             .udTglStruk.SetUnboundFieldSource ("{Ado.tglorder}")
             .usNamaRekanan.SetUnboundFieldSource ("{Ado.namarekanan}")
             .usAlamat.SetUnboundFieldSource ("{Ado.alamatlengkap}")
             .usJenisUsulan.SetUnboundFieldSource ("{Ado.keteranganlainnya}")
             .usNamaBarang.SetUnboundFieldSource ("{ado.namaproduk}")
             .usSpesifikasi.SetUnboundFieldSource ("{ado.deskripsiprodukquo}")
             .unQty.SetUnboundFieldSource ("{Ado.qtyproduk}")
             .usSatuan.SetUnboundFieldSource ("{Ado.satuanstandar}")
             .ucHargaSatuan.SetUnboundFieldSource ("{Ado.hargasatuan}")
'             .unDisc.SetUnboundFieldSource ("{Ado.persenDisc}")
             .ucPajak.SetUnboundFieldSource ("{Ado.hargappn}")
'             .ucTotal.SetUnboundFieldSource ("{Ado.total}")
'             .usQtyHari.SetUnboundFieldSource ("{Ado.nourutlogin}")
'             .Text47.SetText str1
'             .Text3.SetText str2
             
'             Dim X As Double
'             X = Round("{Ado.total}")
'            .usTerbilang.SetUnboundFieldSource "# " & TERBILANG(X) & " #"
             
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "Logistik_A4")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = ReportResep
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
        End With
Exit Sub
errLoad:

End Sub

