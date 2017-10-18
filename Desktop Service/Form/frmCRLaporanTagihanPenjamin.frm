VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCRLaporanTagihanPenjamin 
   Caption         =   "Medifirst2000"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   6330
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
      Height          =   6735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6255
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
Attribute VB_Name = "frmCRLaporanTagihanPenjamin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crLaporanTagihanPenjamin
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

    Set frmCRLaporanTagihanPenjamin = Nothing
End Sub

Public Sub CetakLaporanTagihanPenjamin(idKasir As String, tglAwal As String, tglAkhir As String, idRuangan As String, idKelompok As String, idPenjamin As String, namaPrinted As String, view As String)
On Error GoTo errLoad
'On Error Resume Next

Set frmCRLaporanTagihanPenjamin = Nothing
Dim adocmd As New ADODB.Command

    Dim strFilter As String
    Dim orderby As String
    
    strFilter = ""
    orderby = ""

    strFilter = " where pd.tglregistrasi BETWEEN '" & _
    Format(tglAwal) & "' AND '" & _
    Format(tglAkhir) & "'"
'    strFilter = strFilter & " and IdRuangan like '%" & strIdRuangan & "%' and IdDepartement like '%" & strIdDepartement & "%' and IdKelompokPasien like '%" & strIdKelompokPasien & "%' and IdDokter Like '%" & strIdDokter & "%'"
    
    If idRuangan <> "" Then strFilter = strFilter & " AND sp.objectruanganfk = '" & idRuangan & "' "
    If idKelompok <> "" Then strFilter = strFilter & " AND pd.objectkelompokpasienlastfk = '" & idKelompok & "' "
    If idPenjamin <> "" Then strFilter = strFilter & " AND rk.id = '" & idPenjamin & "' "
  
    orderby = strFilter & "group by pd.tglregistrasi, pd.noregistrasi, ps.nocm, ps.namapasien, ru.namaruangan, " & _
            "pr.id, pp.hargajual, pp.jumlah, kp.id, pp.hargadiscount , rk.namarekanan, sp.totalharusdibayar, sp.totalprekanan " & _
            "ORDER BY pd.tglregistrasi"
            'sp.tglstruk"
    
Set Report = New crLaporanTagihanPenjamin
    strSQL = "SELECT pd.tglregistrasi, pd.noregistrasi, ps.nocm, upper(ps.namapasien) as namapasien, ru.namaruangan, " & _
            "case when pr.id =395 then pp.hargajual* pp.jumlah else 0 end as karcis, " & _
            "case when pr.id =10013116  then pp.hargajual* pp.jumlah else 0 end as embos,  " & _
            "case when kp.id = 26 then pp.hargajual* pp.jumlah else 0 end as konsul, " & _
            "case when kp.id in (1,2,3,4,8,9,10,11,13,14) then pp.hargajual* pp.jumlah else 0 end as tindakan, " & _
            "(case when pp.hargadiscount is null then 0 else pp.hargadiscount end)* pp.jumlah as diskon, " & _
            "sum(case when pr.objectdetailjenisprodukfk=474 then pp.hargajual* pp.jumlah else 0 end) as totalresep, " & _
            "(case when sp.totalprekanan is null then 0 else sp.totalharusdibayar end) as totalharusdibayar, " & _
            "(case when sp.totalprekanan is null then 0 else sp.totalprekanan end) as totalppenjamin, " & _
            "case when rk.namarekanan is null then '-' else rk.namarekanan end as namarekanan " & _
            "FROM pasiendaftar_t as pd left JOIN antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
            "left JOIN pelayananpasien_t as pp on pp.noregistrasifk=apd.norec  " & _
            "left JOIN strukpelayanan_t as sp  on sp.noregistrasifk=pd.norec " & _
            "left JOIN pasien_m as ps on ps.id=pd.nocmfk " & _
            "left JOIN ruangan_m as ru on ru.id=apd.objectruanganfk " & _
            "left JOIN produk_m as pr on pr.id=pp.produkfk " & _
            "left JOIN detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
            "left JOIN jenisproduk_m as jp on jp.id=djp.objectjenisprodukfk " & _
            "left JOIN kelompokproduk_m as kp on kp.id=jp.objectkelompokprodukfk " & _
            "left join rekanan_m  as rk on rk.id=pd.objectrekananfk " & _
            "INNER JOIN kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & orderby

            
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
            .txtNamaKasir.SetText namaPrinted
            .txtPeriode.SetText "Periode : " & tglAwal & " s/d " & tglAkhir & ""
'            .usNamaKasir.SetUnboundFieldSource ("{ado.kasir}")
            .usNamaRuangan.SetUnboundFieldSource ("{ado.namaruangan}")
            '.usNamaDokter.SetUnboundFieldSource ("{ado.namalengkap}")
            .usNoCM.SetUnboundFieldSource ("{ado.nocm}")
            .udTglRegistrasi.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .ucKarcis.SetUnboundFieldSource ("{ado.karcis}")
            .ucEmbos.SetUnboundFieldSource ("{ado.embos}")
            .ucKonsul.SetUnboundFieldSource ("{ado.konsul}")
            .ucTindakan.SetUnboundFieldSource ("{ado.tindakan}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
            .ucResep.SetUnboundFieldSource ("{ado.totalresep}")
            .ucCash.SetUnboundFieldSource ("{ado.totalharusdibayar}")
            .ucTagihan.SetUnboundFieldSource ("{ado.totalppenjamin}")
            '.usKelompokPasien.SetUnboundFieldSource ("{ado.kelompokpasien}")
            .usNamaPenjamin.SetUnboundFieldSource ("{ado.namarekanan}")
            .ucMaterai.SetUnboundFieldSource 3000
            
            
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "LaporanTagihanPenjamin")
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
