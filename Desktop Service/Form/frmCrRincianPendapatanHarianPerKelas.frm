VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCrRincianPendapatanHarianPerKelas 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   Icon            =   "frmCrRincianPendapatanHarianPerKelas.frx":0000
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
Attribute VB_Name = "frmCrRincianPendapatanHarianPerKelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crRncianPendapatanHarianPerkelas
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

    Set frmCRLaporanPendapatanInap = Nothing
End Sub

Public Sub CetakLaporanPendapatan(idKasir As String, tglAwal As String, tglAkhir As String, idDepartemen As String, namaPrinted As String, view As String)
On Error GoTo errLoad
'On Error Resume Next

Set frmCRLaporanPendapatanInap = Nothing
Dim adocmd As New ADODB.Command

    Dim str1 As String

    
    If idDepartemen <> "" Then
        str1 = " and dp.id=" & idDepartemen & " "
    End If
    
Set Report = New crLaporanPendapatanInap
    strSQL = " SELECT pasien_m.nocm, pasien_m.namapasien, pd.noregistrasi, ru.namaruangan || ' ' || kls.namakelas AS namaruangan, " & _
             " pro.namaproduk, pp.hargadiscount, kls.namakelas, case when jp.id in (99,25) then sum( pp.jumlah * pp.hargajual) else 0 end as Akomodasi, " & _
             " case when jp.id in (99,25) then sum(pp.jumlah) else 0 end as VolAkomodasi, case when jp.id=101 then sum( pp.jumlah * pp.hargajual)else 0 end as Visit, " & _
             " case when jp.id=101 then sum(pp.jumlah)else 0 end as VolVisit, " & _
             " case when jp.id=27666 then sum( pp.jumlah * pp.hargajual)else 0 end as SewaAlat, " & _
             " case when jp.id=27666 then sum( pp.jumlah)else 0 end as volSewaAlat, " & _
             " case when jp.id =102 then sum( pp.jumlah * pp.hargajual) else 0 end AS jenisproduk, " & _
             " case when jp.id =102 then sum( pp.jumlah) else 0 end AS voljenisproduk " & _
             " From pasiendaftar_t AS pd " & _
             " LEFT JOIN antrianpasiendiperiksa_t AS apd ON apd.noregistrasifk = pd.norec " & _
             " LEFT JOIN pelayananpasien_t AS pp ON pp.noregistrasifk = apd.norec " & _
             " INNER JOIN produk_m AS pro ON pro.id = pp.produkfk " & _
             " LEFT JOIN kelas_m AS kls ON kls.id = apd.objectkelasfk " & _
             " LEFT JOIN detailjenisproduk_m AS djp ON djp.id = pro.objectdetailjenisprodukfk " & _
             " LEFT JOIN jenisproduk_m AS jp ON jp.id = djp.objectjenisprodukfk " & _
             " LEFT JOIN kelompokproduk_m AS kp ON kp.id = jp.objectkelompokprodukfk " & _
             " LEFT JOIN ruangan_m AS ru ON ru.id = apd.objectruanganfk " & _
             " LEFT JOIN departemen_m AS dp ON dp.id = ru.objectdepartemenfk " & _
             " INNER JOIN pasien_m ON pd.nocmfk = pasien_m.id " & _
             " Where pp.tglpelayanan between '" & tglAwal & "' and '" & tglAkhir & "' AND djp.objectjenisprodukfk <> 97 AND " & _
             " jp.id IN (25, 99, 101, 102, 27666) " & _
             " GROUP BY jp.id,pasien_m.nocm, pasien_m.namapasien, pd.noregistrasi,pp.hargadiscount, " & _
             " ru.namaruangan || ' ' || kls.namakelas,kls.namakelas, pro.namaproduk " & _
            str1 & _
            " ORDER BY pd.noregistrasi ASC "

            
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
            .txtNamaKasir.SetText namaPrinted
            .usNamaRuangan.SetUnboundFieldSource ("{ado.namaruangan}")
            .usNoMR.SetUnboundFieldSource ("{ado.Nocm}")
            .usNoReg.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoReg.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usKelas.SetUnboundFieldSource ("{ado.namakelas}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .unVolAkomodasi.SetUnboundFieldSource ("{ado.volakomodasi}")
            .ucAkomodasi.SetUnboundFieldSource ("{ado.akomodasi}")
            .unVolVisite.SetUnboundFieldSource ("{ado.volvisite}")
            .ucVisite.SetUnboundFieldSource ("{ado.visite}")
            .unVolKonsultasi.SetUnboundFieldSource ("{ado.volkonsultasi}")
            .ucKonsultasi.SetUnboundFieldSource ("{ado.konsultasi}")
            .unVolTindakan.SetUnboundFieldSource ("{ado.voltindakan}")
            .ucTindakan.SetUnboundFieldSource ("{ado.tindakan}")
            .unVolSewaAlat.SetUnboundFieldSource ("{ado.volsewaalat}")
            .ucSewaAlat.SetUnboundFieldSource ("{ado.sewaalat}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
            .txtPeriode.SetText "Periode : " & Format(tglAwal, "dd-MM-yyyy") & "  s/d  " & Format(tglAkhir, "dd-MM-yyyy")

            If idDepartemen <> "" Then
                If idDepartemen = 16 Then
                    .txtJudul.SetText "RINCIAN PENDAPATAN HARIAN PER KELAS RAWAT INAP"
                ElseIf idDepartemen = 18 Then
                    .txtJudul.SetText "RINCIAN PENDAPATAN HARIAN PER KELAS RAWAT JALAN"
                End If
            Else
                .txtJudul.SetText "RINCIAN PENDAPATAN HARIAN PER KELAS"
            End If
            
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "LaporanPedapatan")
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
