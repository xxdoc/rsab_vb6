VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCRPenjualanHarianFarmasi 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   5655
   WindowState     =   2  'Maximized
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
      TabIndex        =   3
      Top             =   600
      Width           =   3015
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
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
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
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5655
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
End
Attribute VB_Name = "frmCRPenjualanHarianFarmasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crPenjualanHarianFarmasi
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

    Set frmCRPenjualanHarianFarmasi = Nothing
End Sub

Public Sub CetakPenjualanHarianFarmasi(namaPrinted As String, tglAwal As String, tglAkhir As String, idRuangan As String, idKelompokPasien As String, idPegawai As String, view As String)
'On Error GoTo errLoad
'On Error Resume Next

Set frmCRPenjualanHarianFarmasi = Nothing
Dim adocmd As New ADODB.Command
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    
    If idPegawai <> "" Then
        str1 = "and sr.penulisresepfk=" & idPegawai & " "
    End If
    If idRuangan <> "" Then
        str2 = " and ru.id=" & idRuangan & " "
    End If
    If idKelompokPasien <> "" Then
        str3 = " and kp.id=" & idKelompokPasien & " "
    End If
    
Set Report = New crPenjualanHarianFarmasi
    strSQL = "select sr.tglresep, sr.noresep, pd.noregistrasi, ps.namapasien," & _
            "case when jk.jeniskelamin = 'Laki-laki' then 'L' else 'P' end as jeniskelamin, " & _
            "kp.kelompokpasien, pg.namalengkap, ru2.namaruangan, pp.jumlah, pp.hargajual,  " & _
            "(pp.jumlah)*(pp.hargajual) as subtotal," & _
            "0 as diskon, 0 as jasa, 0 as ppn, (pp.jumlah*pp.hargajual)-0-0-0 as total, " & _
            "case when sp.nosbmlastfk is null then 'N' else'P' end as statuspaid, pg2.namalengkap as kasir " & _
            "from strukresep_t as sr " & _
            "LEFT JOIN pelayananpasien_t as pp on pp.strukresepfk = sr.norec " & _
            "LEFT JOIN strukpelayanan_t as sp on sp.norec=pp.strukterimafk " & _
            "inner JOIN antrianpasiendiperiksa_t as apd on apd.norec=pp.noregistrasifk " & _
            "inner JOIN pasiendaftar_t as pd on pd.norec=apd.noregistrasifk " & _
            "inner JOIN pasien_m as ps on ps.id=pd.nocmfk " & _
            "inner join jeniskelamin_m as jk on jk.id=ps.objectjeniskelaminfk " & _
            "inner JOIN pegawai_m as pg on pg.id=sr.penulisresepfk " & _
            "left join strukbuktipenerimaan_t as sbm on sbm.norec = sp.nosbklastfk " & _
            "left join pegawai_m as pg2 on pg2.id = sbm.objectpegawaipenerimafk " & _
            "inner JOIN ruangan_m as ru on ru.id=sr.ruanganfk " & _
            "inner JOIN ruangan_m as ru2 on ru2.id=apd.objectruanganfk " & _
            "inner join kelompokpasien_m kp on kp.id=pd.objectkelompokpasienlastfk " & _
            "where sr.tglresep BETWEEN '" & tglAwal & "' and '" & tglAkhir & "' " & _
            str1 & _
            str2 & _
            str3
   
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
            .txtPrinted.SetText namaPrinted
            .txtPeriode.SetText "Periode : " & tglAwal & " s/d " & tglAkhir & ""
            .udtTglResep.SetUnboundFieldSource ("{ado.tglresep}")
            .usNoResep.SetUnboundFieldSource ("{ado.noresep}")
            .usRuangan.SetUnboundFieldSource ("{ado.namaruangan}")
            .usKelPasien.SetUnboundFieldSource ("{ado.kelompokpasien}")
            .usNoReg.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usJK.SetUnboundFieldSource ("{ado.jeniskelamin}")
            .usDokter.SetUnboundFieldSource ("{ado.namalengkap}")
            .usJumlahResep.SetUnboundFieldSource ("{ado.jumlah}")
            .ucSubTotal.SetUnboundFieldSource ("{ado.subtotal}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
            .ucJasa.SetUnboundFieldSource ("{ado.jasa}")
            .ucPPN.SetUnboundFieldSource ("{ado.ppn}")
            .ucTotal.SetUnboundFieldSource ("{ado.total}")
            .usStatusPaid.SetUnboundFieldSource ("{ado.statuspaid}")
            .usKasir.SetUnboundFieldSource ("{ado.kasir}")
            
            
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "LaporanPenjualanHarianFarmasi")
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
