VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCRCetakRincianBiaya 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmCRCetakRincianBiaya.frx":0000
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
Attribute VB_Name = "frmCRCetakRincianBiaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crRincianBiayaPelayanan
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
    cboPrinter.Text = GetTxt("Setting.ini", "Printer", "RincianBiayaPelayanan")
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCRCetakRincianBiaya = Nothing
End Sub

Public Sub CetakRincianBiaya(strNoregistrasi As String, strNoStruk As String, strNoKwitansi As String, strIdPegawai As String, view As String)
'On Error GoTo errLoad
'On Error Resume Next

Set frmCRCetakRincianBiaya = Nothing
Dim adocmd As New ADODB.Command
Dim strFilter As String

    strFilter = ""
    
'    strFilter = " where sp.tglstruk BETWEEN '" & _
'    Format(tglAwal, "yyyy-MM-dd 00:00:00") & "' AND '" & _
'    Format(tglAkhir, "yyyy-MM-dd 23:59:59") & "'"
''    strFilter = strFilter & " and IdRuangan like '%" & strIdRuangan & "%' and IdDepartement like '%" & strIdDepartement & "%' and IdKelompokPasien like '%" & strIdKelompokPasien & "%' and IdDokter Like '%" & strIdDokter & "%'"
'
'    If strIdRuangan <> "" Then strFilter = strFilter & " AND apd.objectruanganfk = '" & strIdRuangan & "' "
'    If strIdKelompokPasien <> "" Then strFilter = strFilter & " AND pd.objectkelompokpasienlastfk = '" & strIdKelompokPasien & "' "
   
Set Report = New crRincianBiayaPelayanan
    strSQL = "SELECT sp.tglstruk,sp.nostruk as nobilling,sbm.nosbm as nokwitansi, pd.noregistrasi,ps.nocm,(upper(ps.namapasien) || ' ( ' || jk.reportdisplay || ' )' ) as namapasienjk ,ru.namaruangan as unit,ru.objectdepartemenfk,kl.namakelas,   " & _
            "pg.namalengkap as dokterpj,pd.tglregistrasi,pd.tglpulang,case when rk.namarekanan is null then '-' else rk.namarekanan end as namarekanan,pp.tglpelayanan, ru2.namaruangan as ruangantindakan,pr.namaproduk,jp.jenisproduk, pg2.namalengkap as dokter,pp.jumlah,pp.hargajual,   " & _
            "case when pp.hargadiscount is null then 0 else pp.hargadiscount end as diskon,(pp.jumlah*(pp.hargajual-case when pp.hargadiscount is null then 0 else pp.hargadiscount end)) as total, case when kmr.namakamar is null then '-' else kmr.namakamar end as namakamar ,klp.kelompokpasien as tipepasien,   " & _
            "sp.totalharusdibayar,(case when sppj.totalppenjamin is null then 0 else sppj.totalppenjamin end) as totalppenjamin,(case when sp.totalbiayatambahan is null then 0 else sp.totalbiayatambahan end) as totalbiayatambahan, pg3.namalengkap as user " & _
            "from pelayananpasien_t as pp left JOIN strukpelayanan_t as sp on pp.strukfk=sp.norec  " & _
            "LEFT JOIN strukbuktipenerimaan_t as sbm on sp.nosbmlastfk=sbm.norec   " & _
            "LEFT JOIN strukpelayananpenjamin_t as sppj on sp.norec=sppj.nostrukfk " & _
            "LEFT JOIN strukbuktipenerimaancarabayar_t as sbmc on sbm.norec=sbmc.nosbmfk  " & _
            "left JOIN carabayar_m as cb on cb.id=sbmc.objectcarabayarfk                  " & _
            "inner JOIN antrianpasiendiperiksa_t as apd on apd.norec=pp.noregistrasifk  " & _
            "inner JOIN pasiendaftar_t as pd on pd.norec=apd.noregistrasifk   " & _
            "INNER join produk_m as pr on pr.id=pp.produkfk   " & _
            "INNER join detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk   " & _
            "INNER join jenisproduk_m as jp on jp.id=djp.objectjenisprodukfk   " & _
            "INNER join pasien_m as ps on ps.id=pd.nocmfk   " & _
            "INNER join jeniskelamin_m as jk on jk.id=ps.objectjeniskelaminfk   " & _
            "INNER join ruangan_m  as ru on ru.id=pd.objectruanganlastfk   " & _
            "INNER join ruangan_m  as ru2 on ru2.id=apd.objectruanganfk   " & _
            "left join kelas_m  as kl on kl.id=pd.objectkelasfk   " & _
            "left join pegawai_m  as pg on pg.id=pd.objectpegawaifk   " & _
            "left join pegawai_m  as pg2 on pg2.id=apd.objectpegawaifk   " & _
            "left join rekanan_m  as rk on rk.id=pd.objectrekananfk   " & _
            "left join kamar_m  as kmr on kmr.id=apd.objectkamarfk  " & _
            "INNER JOIN kelompokpasien_m as klp on klp.id=pd.objectkelompokpasienlastfk left join pegawai_m as pg3 on pg3.id=sbm.objectpegawaipenerimafk " & _
            "where pd.noregistrasi='" & strNoregistrasi & "' and pr.id<>402611 or sp.nostruk='" & strNoStruk & "' and pr.id<>402611 or sbm.nosbm='" & strNoKwitansi & "' and pr.id<>402611"
    
    ReadRs2 "select sum(hargajual) as totalDeposit from pasiendaftar_t pd " & _
            "INNER JOIN antrianpasiendiperiksa_t apd on apd.noregistrasifk=pd.norec " & _
            "INNER JOIN pelayananpasien_t pp on pp.noregistrasifk=apd.norec " & _
            "where pd.noregistrasi='" & strNoregistrasi & "' and pp.produkfk=402611 "
    
    ReadRs3 "select ppd.hargadiscount,ppd.hargajual,ppd.komponenhargafk from pasiendaftar_t pd " & _
            "INNER JOIN antrianpasiendiperiksa_t apd on apd.noregistrasifk=pd.norec " & _
            "INNER JOIN pelayananpasien_t pp on pp.noregistrasifk=apd.norec " & _
            "INNER JOIN pelayananpasiendetail_t ppd on ppd.pelayananpasien=pp.norec " & _
            "where pd.noregistrasi='" & strNoregistrasi & "' and pp.produkfk<>402611"
    
    Dim TotalDiskonMedis  As Double
    Dim TotalDiskonUmum  As Double
    Dim i As Integer
    
    
    For i = 0 To RS3.RecordCount - 1
        If RS3!komponenhargafk = 35 Then TotalDiskonMedis = TotalDiskonMedis + CDbl(RS3!hargadiscount)
        TotalDiskonUmum = TotalDiskonUmum + CDbl(RS3!hargadiscount)
        RS3.MoveNext
    Next
    
    Dim TotalDeposit As Double
    TotalDeposit = IIf(IsNull(RS2(0)), 0, RS2(0))
    
'    ReadRs2 "SELECT " & _
'            "sum((pp.jumlah*(pp.hargajual-case when pp.hargadiscount is null then 0 else pp.hargadiscount end))) as total " & _
'            "from pasiendaftar_t as pd " & _
'            "inner join antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
'            "inner join pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
'            "inner join produk_m as pr on pr.id=pp.produkfk " & _
'            "inner join detailjenisproduk_m as djp on djp.id=pr.objectdetailjenisprodukfk " & _
'            "inner join jenisproduk_m as jp on jp.id=djp.objectjenisprodukfk " & _
'            "inner join pasien_m as ps on ps.id=pd.nocmfk " & _
'            "inner join jeniskelamin_m as jk on jk.id=ps.objectjeniskelaminfk " & _
'            "inner join ruangan_m  as ru on ru.id=pd.objectruanganlastfk " & _
'            "inner join ruangan_m  as ru2 on ru2.id=apd.objectruanganfk " & _
'            "LEFT join kelas_m  as kl on kl.id=pd.objectkelasfk " & _
'            "inner join pegawai_m  as pg on pg.id=pd.objectpegawaifk " & _
'            "inner join pegawai_m  as pg2 on pg2.id=apd.objectpegawaifk " & _
'            "left join rekanan_m  as rk on rk.id=pd.objectrekananfk " & _
'            "where pd.noregistrasi='" & strNoregistrasi & "' "

   
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
        'If Not RS.EOF Then
            .usNoRegistrasi.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNoCM.SetUnboundFieldSource ("{ado.nocm}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasienjk}")
            .usRuangan.SetUnboundFieldSource ("{ado.unit}")
            .usKamar.SetUnboundFieldSource IIf(IsNull("{ado.namakamar}") = True, "-", ("{ado.namakamar}"))
            .usKelasH.SetUnboundFieldSource ("{ado.namakelas}")
            .usDokterPJawab.SetUnboundFieldSource ("{ado.dokterpj}")
            .udTglMasuk.SetUnboundFieldSource ("{ado.tglregistrasi}")
            .udTglPlng.SetUnboundFieldSource IIf(IsNull("{ado.tglpulang}") = True, "-", ("{ado.tglpulang}"))
            .UsPenjamin.SetUnboundFieldSource IIf(IsNull("{ado.namarekanan}") = True, ("-"), ("{ado.namarekanan}"))
            .usTipe.SetUnboundFieldSource ("{ado.tipepasien}")
                     
            .usJenisProduk.SetUnboundFieldSource ("{ado.jenisproduk}")
            .udTanggal.SetUnboundFieldSource ("{ado.tglpelayanan}")
            .usLayanan.SetUnboundFieldSource ("{ado.namaproduk}")
            .usKelas.SetUnboundFieldSource ("{ado.namakelas}")
            .usDokter.SetUnboundFieldSource ("{ado.dokter}")
            .unQty.SetUnboundFieldSource ("{ado.jumlah}")
            .ucTarif.SetUnboundFieldSource ("{ado.hargajual}")
            .ucDiskon.SetUnboundFieldSource ("{ado.diskon}")
            .ucTotal.SetUnboundFieldSource ("{ado.total}")
            .usRuanganTindakan.SetUnboundFieldSource ("{ado.ruangantindakan}")
            .usNoStruk.SetUnboundFieldSource ("{ado.nobilling}")
            
'            .ucAdministrasi.SetUnboundFieldSource ("0") '("{ado.administrasi}")
'            .ucMaterai.SetUnboundFieldSource ("0") '("{ado.materai}")
            
            .ucDeposit.SetUnboundFieldSource (TotalDeposit) '("{ado.deposit}")
            .ucDiskonJasaMedis.SetUnboundFieldSource (TotalDiskonMedis)
            .ucDiskonUmum.SetUnboundFieldSource (TotalDiskonUmum) '("{ado.diskonumum}")
'            .ucSisaDeposit.SetUnboundFieldSource ("0")
            
            
            .ucDitanggungPerusahaan.SetUnboundFieldSource ("{ado.totalppenjamin}")
            .ucDitanggungRS.SetUnboundFieldSource ("0") '("{ado.totalharusdibayarrs}")
'            .ucDitanggungSendiri.SetUnboundFieldSource ("{ado.totalharusdibayar}")
'            .ucSurplusMinusRS.SetUnboundFieldSource ("{ado.SurplusMinusRS}")
            .usUser.SetUnboundFieldSource ("{ado.user}")
            
'            ReadRs2 "SELECT namalengkap FROM pegawai_m where id='" & strIdPegawai & "' "
'            If RS2.BOF Then
'                .txtUser.SetText "-"
'            Else
'                .txtUser.SetText UCase(IIf(IsNull(RS2("namalengkap")), "-", RS2("namalengkap")))
'            End If
            .txtUser.SetText UCase(strIdPegawai)
            
            If view = "false" Then
                Dim strPrinter As String
'
                strPrinter = GetTxt("Setting.ini", "Printer", "RincianBiayaPelayanan")
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