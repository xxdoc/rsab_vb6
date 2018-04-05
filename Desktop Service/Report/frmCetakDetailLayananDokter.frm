VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakDetailLayananDokter 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmCetakDetailLayananDokter.frx":0000
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
      Width           =   5895
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
Attribute VB_Name = "frmCetakDetailLayananDokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crCetakDetailLayananDokter
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

    Set frmCetakDetailLayananDokter = Nothing
End Sub

Public Sub CetakDetailLayanan(ID As String, tglAwal As String, tglAkhir As String, strIdDepartemen As String, strIdRuangan As String, _
                                        strIdKelompokPasien As String, strIdDokter As String, view As String)
On Error GoTo errLoad
'On Error Resume Next

Set frmCetakDetailLayananDokter = Nothing
Dim adocmd As New ADODB.Command
Dim strFilter, orderby As String
Set Report = New crCetakDetailLayananDokter

    strFilter = ""
    orderby = ""

    strFilter = " where pp.tglpelayanan BETWEEN '" & _
    Format(tglAwal, "yyyy-MM-dd 00:00:00") & "' AND '" & _
    Format(tglAkhir, "yyyy-MM-dd 23:59:59") & "' and djp.objectjenisprodukfk <> 97 and sp.statusenabled is null and jpg.id=1 " ' and djp.objectjenisprodukfk <> 97 and kp.id in (1,2,3,4,8,9,10,11,13,14,26) and sp.statusenabled is null "
'    strFilter = strFilter & " and IdRuangan like '%" & strIdRuangan & "%' and IdDepartement like '%" & strIdDepartement & "%' and IdKelompokPasien like '%" & strIdKelompokPasien & "%' and IdDokter Like '%" & strIdDokter & "%'"
    
    If strIdDepartemen <> "" Then
        If strIdDepartemen = 16 Then
            strFilter = strFilter & " AND ru.objectdepartemenfk in (16,17,25,26) "
        Else
            If strIdDepartemen <> "" Then
                strFilter = strFilter & " AND ru2.objectdepartemenfk not in (16,17,25,26)  "
            End If
        End If
    End If
    If strIdRuangan <> "" Then strFilter = strFilter & " AND ru.id = '" & strIdRuangan & "' "
    If strIdKelompokPasien <> "" Then strFilter = strFilter & " AND pd.objectkelompokpasienlastfk = '" & strIdKelompokPasien & "' "
    If strIdDokter <> "" Then strFilter = strFilter & " AND pg.id = '" & strIdDokter & "' "
  
    orderby = strFilter & " order by pg.namalengkap"
    strSQL = "select DISTINCT pp.tglpelayanan, apd.objectruanganfk, ru.namaruangan,kl.namakelas, " & _
            "pg.namalengkap,pd.noregistrasi,ps.nocm, upper(ps.namapasien) as namapasien, " & _
            "case when ru.objectdepartemenfk in (16,35) then 'Y' ELSE 'N' END as inap, " & _
            "kps.kelompokpasien, case when rk.namarekanan is not null then rk.namarekanan else '-' end as namarekanan, pr.namaproduk, pp.jumlah, " & _
            "case when pp.hargajual is not null then pp.hargajual else 0 end as harga, " & _
            "case when sbm.norec is null then 'N' else 'Y' end as sbm " & _
            "from pasiendaftar_t as pd inner JOIN antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
            "left JOIN pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
            "left join pelayananpasienpetugas_t as ppp on ppp.pelayananpasien = pp.norec " & _
            "left JOIN pegawai_m as pg on pg.id=ppp.objectpegawaifk " & _
            "inner join jenispegawai_m as jpg on jpg.id=pg.objectjenispegawaifk " & _
            "inner JOIN produk_m as pr on pr.id=pp.produkfk " & _
            "inner JOIN detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
            "inner JOIN jenisproduk_m as jp on jp.id=djp.objectjenisprodukfk " & _
            "inner JOIN kelompokproduk_m as kp on kp.id=jp.objectkelompokprodukfk " & _
            "inner JOIN pasien_m as ps on ps.id=pd.nocmfk " & _
            "left JOIN kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
            "left join rekanan_m as rk on rk.id=pd.objectrekananfk " & _
            "left JOIN strukpelayanan_t as sp  on sp.noregistrasifk=pd.norec " & _
            "left JOIN strukbuktipenerimaan_t as sbm  on sbm.norec=sp.nosbmlastfk " & _
            "left JOIN ruangan_m as ru on ru.id=apd.objectruanganfk " & _
            "left join kelas_m as kl on kl.id = apd.objectkelasfk " & _
            "left join departemen_m as dp on dp.id = ru.objectdepartemenfk " & orderby
        
'    strSQL = "select DISTINCT pp.tglpelayanan, apd.objectruanganfk, ru.namaruangan,pg.namalengkap,pd.noregistrasi,ps.nocm ,  " & _
'            "upper(ps.namapasien) as namapasien, " & _
'            "case when ru.objectdepartemenfk in (16,35) then 'Y' ELSE 'N' END as inap, " & _
'            "kps.kelompokpasien, case when rk.namarekanan is not null then rk.namarekanan else '-' end as namarekanan, pr.namaproduk, pp.jumlah, " & _
'            "case when pp.hargajual * pp.jumlah is not null then pp.hargajual * pp.jumlah else 0 end as harga, " & _
'            "case when sbm.norec is null then 'N' else 'Y' end as sbm " & _
'            "from pasiendaftar_t as pd inner JOIN antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
'            "left JOIN pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
'            "left join pelayananpasienpetugas_t as ppp on ppp.pelayananpasien = pp.norec " & _
'            "left JOIN pegawai_m as pg on pg.id=ppp.objectpegawaifk " & _
'            "inner join jenispegawai_m as jpg on jpg.id=pg.objectjenispegawaifk " & _
'            "inner JOIN produk_m as pr on pr.id=pp.produkfk " & _
'            "inner JOIN detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
'            "inner JOIN jenisproduk_m as jp on jp.id=djp.objectjenisprodukfk " & _
'            "inner JOIN kelompokproduk_m as kp on kp.id=jp.objectkelompokprodukfk " & _
'            "inner JOIN pasien_m as ps on ps.id=pd.nocmfk " & _
'            "left JOIN kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
'            "left join rekanan_m as rk on rk.id=pd.objectrekananfk " & _
'            "left JOIN strukpelayanan_t as sp  on sp.noregistrasifk=pd.norec " & _
'            "left JOIN strukbuktipenerimaan_t as sbm  on sbm.norec=sp.nosbmlastfk " & _
'            "left JOIN ruangan_m as ru on ru.id=sp.objectruanganfk " & _
'            "left join departemen_m as dp on dp.id = ru.objectdepartemenfk " & orderby
            
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
        'If Not RS.EOF Then
            
            .udTglPelayanan.SetUnboundFieldSource ("{ado.tglpelayanan}")
            .usNoCM.SetUnboundFieldSource ("{ado.nocm}")
            .usPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usRuanganPelayanan.SetUnboundFieldSource ("{ado.namaruangan}")
            .usDokter.SetUnboundFieldSource ("{ado.namalengkap}")
            .usJenisPasien.SetUnboundFieldSource ("{ado.kelompokpasien}")
            '.usKelas.SetUnboundFieldSource ("{ado.namakelas}")
            .usInap.SetUnboundFieldSource ("{ado.inap}")
            .usPenjamin.SetUnboundFieldSource ("{ado.namarekanan}")
            .usLayanan.SetUnboundFieldSource ("{ado.namaproduk}")
            .ucHargaLayanan.SetUnboundFieldSource ("{ado.harga}")
            .unJumlah.SetUnboundFieldSource ("{ado.jumlah}")
            .usPaid.SetUnboundFieldSource ("{ado.sbm}")
            
        .txtTgl.SetText "TANGGAL " & Format(tglAwal, "dd-MM-yyyy") & "  s/d  " & Format(tglAkhir, "dd-MM-yyyy")
             
        ReadRs2 "SELECT namalengkap FROM pegawai_m where id='" & ID & "' "
        If RS2.BOF Then
            .txtUser.SetText "-"
        Else
            .txtUser.SetText UCase(IIf(IsNull(RS2("namalengkap")), "-", RS2("namalengkap")))
        End If
            
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
