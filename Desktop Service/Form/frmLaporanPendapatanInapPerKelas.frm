VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmLaporanPendapatanInapPerKelas 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   Icon            =   "frmLaporanPendapatanInapPerKelas.frx":0000
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
Attribute VB_Name = "frmLaporanPendapatanInapPerKelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crLaporanPendapatanInapPerKelas
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

    Set frmLaporanPendapatanInapPerKelas = Nothing
End Sub

Public Sub CetakLaporanPendapatan(idKasir As String, tglAwal As String, tglAkhir As String, idDepartemen As String, idRuangan As String, idKelompok As String, kelasfk As String, namaPrinted As String, view As String)
On Error GoTo errLoad
'On Error Resume Next

Set frmLaporanPendapatanInapPerKelas = Nothing
Dim adocmd As New ADODB.Command

    Dim str1, str2, str3, str4, str5 As String

    If idDepartemen <> "" Then
        If idDepartemen = 16 Then
            str1 = " and dp.id in (16,17,26)"
        ElseIf idDepartemen = 25 Then
            str1 = " and dp.id in (25)"
        Else
            If idDepartemen <> "" Then
                str1 = " and dp.id not in (16,17,26,25)"
            End If
        End If
    End If
    If idRuangan <> "" Then
        str2 = " and ru.id=" & idRuangan & " "
    End If
'    If idKelas <> "" Then
'        str3 = " and kls.id=" & idKelas & " "
'    End If
    If idKelompok <> "" Then
        str4 = " and kps.id=" & idKelompok & " "
    End If
    
    If kelasfk <> "" Then
        str5 = " and kls.id= " & kelasfk & " "
     End If
    
    
Set Report = New crLaporanPendapatanInapPerKelas
'    strSQL = "select pd.noregistrasi, ru.namaruangan || ' ' || km.namakamar as namaruangan, kps.kelompokpasien,pro.namaproduk, " & _
'            "case when (kp.id in (25) and jp.id in (99,25)) then 'Akomodasi' " & _
'            "when (kp.id in (26) and jp.id in (101,100)) then 'Visit' " & _
'            "when kp.id in (1, 2, 3, 4, 8, 9, 10, 11, 13, 14) then 'Tindakan' when kp.id in (34) then 'Sewa Alat' end as jenisproduk, " & _
'            "case when pp.hargajual is not null then pp.hargajual else 0 end as hargajual, " & _
'            "case when pp.jumlah is not null and pp.hargajual is not null then pp.jumlah else 0 end as jumlah from pasiendaftar_t as pd " & _
'            "left JOIN antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
'            "left JOIN pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
'            "inner join produk_m as pro on pro.id = pp.produkfk " & _
'            "left join kamar_m as km on km.id = apd.objectkamarfk " & _
'            "left join detailjenisproduk_m as djp on djp.id = pro.objectdetailjenisprodukfk " & _
'            "left join jenisproduk_m as jp on jp.id = djp.objectjenisprodukfk " & _
'            "left JOIN kelompokproduk_m as kp on kp.id=jp.objectkelompokprodukfk " & _
'            "left join ruangan_m as ru on ru.id = apd.objectruanganfk " & _
'            "left join departemen_m as dp on dp.id = ru.objectdepartemenfk " & _
'            "left JOIN kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
'            "where pp.tglpelayanan between '" & tglAwal & "' and '" & tglAkhir & "' and djp.objectjenisprodukfk <> 97 " & _
'            "and jp.id in (15,25,99,100,101,102,107,27666) and pro.id not in(402611) " & _
'             str1 & _
'             str2 & _
'             str4 & _
'             "order by pd.noregistrasi"
             
    
'    strSQL = "select pd.noregistrasi, ru.namaruangan || ' ' || kls.namakelas as namaruangan, kps.kelompokpasien,pro.namaproduk, " & _
'            "case when (kp.id in (25) and jp.id in (99,25)) then 'Akomodasi' " & _
'            "when (kp.id in (26) and jp.id in (101,100)) then 'Visit' " & _
'            "when kp.id in (1, 2, 3, 4, 8, 9, 10, 11, 13, 14) then 'Tindakan' when kp.id in (34) then 'Sewa Alat' end as jenisproduk, " & _
'            "case when pp.hargajual is not null then pp.hargajual else 0 end as hargajual, " & _
'            "case when pp.jumlah is not null and pp.hargajual is not null then pp.jumlah else 0 end as jumlah from pasiendaftar_t as pd " & _
'            "left JOIN antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
'            "left JOIN pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
'            "inner join produk_m as pro on pro.id = pp.produkfk " & _
'            "left join detailjenisproduk_m as djp on djp.id = pro.objectdetailjenisprodukfk " & _
'            "left join jenisproduk_m as jp on jp.id = djp.objectjenisprodukfk " & _
'            "left JOIN kelompokproduk_m as kp on kp.id=jp.objectkelompokprodukfk " & _
'            "left join ruangan_m as ru on ru.id = apd.objectruanganfk " & _
'            "left join kelas_m as kls on kls.id = pd.objectkelasfk " & _
'            "left join departemen_m as dp on dp.id = ru.objectdepartemenfk " & _
'            "left JOIN kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
'            "where pp.tglpelayanan between '" & tglAwal & "' and '" & tglAkhir & "' and djp.objectjenisprodukfk <> 97 " & _
'            "and jp.id in (15,25,99,100,101,102,107,27666) and pro.id not in(402611) " & _
'             str1 & _
'             str2 & _
'             str4 & _
'             "order by pd.noregistrasi"


        strSQL = "select pd.noregistrasi, ru.namaruangan as namaruangan, kps.kelompokpasien,kls.id as kelasfk,kls.namakelas,pro.namaproduk, " & _
            "case when (kp.id in (25) and jp.id in (99,25)) then 'Akomodasi' " & _
            "when (kp.id in (26) and jp.id in (101,100)) then 'Visit' " & _
            "when kp.id in (1, 2, 3, 4, 8, 9, 10, 11, 13, 14) then 'Tindakan' when kp.id in (34) then 'Sewa Alat' end as jenisproduk, " & _
            "case when pp.hargajual is not null then pp.hargajual else 0 end as hargajual, " & _
            "case when pp.jumlah is not null and pp.hargajual is not null then pp.jumlah else 0 end as jumlah from pasiendaftar_t as pd " & _
            "left JOIN antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
            "left JOIN pelayananpasien_t as pp on pp.noregistrasifk=apd.norec " & _
            "inner join produk_m as pro on pro.id = pp.produkfk " & _
            "left join detailjenisproduk_m as djp on djp.id = pro.objectdetailjenisprodukfk " & _
            "left join jenisproduk_m as jp on jp.id = djp.objectjenisprodukfk " & _
            "left JOIN kelompokproduk_m as kp on kp.id=jp.objectkelompokprodukfk " & _
            "left join ruangan_m as ru on ru.id = apd.objectruanganfk " & _
            "left join departemen_m as dp on dp.id = ru.objectdepartemenfk " & _
            "left JOIN kelompokpasien_m as kps on kps.id=pd.objectkelompokpasienlastfk " & _
            "inner join kelas_m as kls on kls.id = apd.objectkelasfk " & _
            "where pp.tglpelayanan between '" & tglAwal & "' and '" & tglAkhir & "' and djp.objectjenisprodukfk <> 97 " & _
            "and pro.id not in(402611) " & _
             str1 & _
             str2 & _
             str4 & _
             str5 & _
             "order by pd.noregistrasi"
             
             '"and jp.id in (15,25,26,99,100,101,102,107,27666)
            
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
            .txtNamaKasir.SetText namaPrinted
            '.txtPeriode.SetText "Periode : " & tglAwal & " s/d " & tglAkhir & ""
            .usNamaRuangan.SetUnboundFieldSource ("{ado.namaruangan}")
            '.usNoReg.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usKegiatan.SetUnboundFieldSource ("{ado.namaproduk}")
            .ucTarif.SetUnboundFieldSource ("{ado.hargajual}")
            .unJumlah.SetUnboundFieldSource ("{ado.jumlah}")
            '.ucTotal.SetUnboundFieldSource ("{ado.total}")
            .usJenisProduk.SetUnboundFieldSource ("{ado.jenisproduk}")
            .usKelas.SetUnboundFieldSource ("{ado.namakelas}")
            .usKelompok.SetUnboundFieldSource ("{ado.kelompokpasien}")
            
            .txtPeriode.SetText "Periode : " & Format(tglAwal, "dd-MM-yyyy") & "  s/d  " & Format(tglAkhir, "dd-MM-yyyy")

            If idDepartemen <> "" Then
                If idDepartemen = 16 Then
                    .txtJudul.SetText "LAPORAN VOLUME KEGIATAN DAN PENDAPATAN RAWAT INAP PER KELAS"
                ElseIf idDepartemen = 18 Then
                    .txtJudul.SetText "LAPORAN VOLUME KEGIATAN DAN PENDAPATAN RAWAT JALAN"
                End If
            Else
                .txtJudul.SetText "LAPORAN VOLUME KEGIATAN DAN PENDAPATAN PER KELAS"
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
