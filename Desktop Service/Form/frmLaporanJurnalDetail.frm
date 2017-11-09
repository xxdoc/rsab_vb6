VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmLaporanJurnalDetail 
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
      Height          =   6855
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
Attribute VB_Name = "frmLaporanJurnalDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As New crLaporanJurnalDetail
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

    Set frmLaporanJurnalDetail = Nothing
End Sub

Public Sub CetakLaporanJurnal(idKasir As String, tglAwal As String, tglAkhir As String, idDepartemen As String, idRuangan As String, namaPrinted As String, view As String)
On Error GoTo errLoad
'On Error Resume Next

Set frmLaporanJurnalDetail = Nothing
Dim adocmd As New ADODB.Command

    Dim str1 As String
    Dim str2 As String
    
    If idDepartemen <> "" Then
        If idDepartemen = 18 Then
            str1 = " AND ru.objectdepartemenfk <> 16"
        Else
            If idDepartemen <> "" Then
                str1 = " AND ru.objectdepartemenfk = '" & idDepartemen & "' "
            End If
        End If
    End If
'    If idDepartemen <> "" Then
'        str1 = "and ru.objectdepartemenfk=" & idDepartemen & " "
'    End If
    If idRuangan <> "" Then
        str2 = " and apd.objectruanganfk=" & idRuangan & " "
    End If
    
    
Set Report = New crLaporanJurnalDetail
'    strSQL = "select pd.noregistrasi || '/' || ps.nocm as regcm, ps.namapasien, ru.namaruangan, tp.produkfk as kode, " & _
'            "pro.namaproduk as layanan, tp.hargajual, tp.jumlah, " & _
'            "case when pro.id = 395 then 'Pendt. Administrasi' " & _
'            "when kp.id = 26 then 'Pendt. Konsultasi' " & _
'            "when kp.id in (1,2,3,4,8,9,10,11,13,14) then 'Pendt. Tindakan' end as namaperkiraan, " & _
'            "(sum(case when pd.objectkelompokpasienlastfk = 1 then tp.hargajual*tp.jumlah  else 0 end))+ " & _
'            "(sum(case when pd.objectkelompokpasienlastfk > 1 then tp.hargajual*tp.jumlah  else 0 end))  as total, " & _
'            "'Pendapatan R.Jalan' as keterangan " & _
'            "from pasiendaftar_t as pd left JOIN antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
'            "left join pelayananpasien_t as tp on tp.noregistrasifk = apd.norec " & _
'            "inner JOIN produk_m AS pro ON tp.produkfk = pro.id " & _
'            "inner JOIN detailjenisproduk_m as djp on djp.id=pro.objectdetailjenisprodukfk " & _
'            "inner JOIN jenisproduk_m as jp on jp.id=djp.objectjenisprodukfk " & _
'            "inner JOIN kelompokproduk_m as kp on kp.id=jp.objectkelompokprodukfk " & _
'            "inner JOIN ruangan_m as ru on ru.id=apd.objectruanganfk " & _
'            "inner join departemen_m as dp on dp.id = ru.objectdepartemenfk " & _
'            "inner JOIN pasien_m as ps on ps.id=pd.nocmfk " & _
'            "where pd.tglregistrasi between '" & tglAwal & "' and '" & tglAkhir & "' and djp.objectjenisprodukfk <> 97 " & _
'            str1 & _
'            str2 & _
'            "group by pd.noregistrasi, ps.namapasien, ps.nocm, tp.hargajual, tp.jumlah, " & _
            "ru.namaruangan, tp.produkfk, pro.namaproduk, pro.id, kp.id  " & _
            "order by ps.namapasien"
            
    strSQL = "select pd.noregistrasi || '/' || ps.nocm as regcm, ps.namapasien, ru.namaruangan, tp.produkfk as kode, pro.namaproduk as layanan, tp.hargajual, tp.jumlah, case " & _
            "when pro.id = 395                                          then'Pendt. Administrasi' || ' ' || ru.namaruangan " & _
            "when kp.id = 26 and pro.id <> 395                          then 'Pendt. Konsultasi' || ' ' || ru.namaruangan " & _
            "when kp.id in (3,4,8,9,10,11,13,14) and pro.id <> 395  then 'Pendt. Tindakan' || ' ' || ru.namaruangan " & _
            "when kp.id in (1) and pro.id <> 395  then 'Pendt. Tindakan' || ' ' || ru.namaruangan " & _
            "when kp.id in (2) and pro.id <> 395  then 'Pendt. Tindakan' || ' ' || ru.namaruangan " & _
            "when kp.id in (24) and pro.id <> 395  then 'Pendt. Tindakan Ka Instalasi Farmasi' " & _
            "ELSE 'Pendt. Tindakan' || ' ' || ru.namaruangan end  as namaperkiraan, " & _
            "sum(case when (tp.hargajual* tp.jumlah) is null then 0 else (tp.hargajual* tp.jumlah) end) as total, " & _
            "'Pendapatan R.Jalan' as keterangan " & _
            "from pasiendaftar_t as pd left JOIN antrianpasiendiperiksa_t as apd on apd.noregistrasifk=pd.norec " & _
            "left join pelayananpasien_t as tp on tp.noregistrasifk = apd.norec " & _
            "LEFT JOIN produk_m AS pro ON tp.produkfk = pro.id " & _
            "left JOIN detailjenisproduk_m as djp on djp.id=pro.objectdetailjenisprodukfk " & _
            "left JOIN jenisproduk_m as jp on jp.id=djp.objectjenisprodukfk " & _
            "left JOIN kelompokproduk_m as kp on kp.id=jp.objectkelompokprodukfk " & _
            "left JOIN ruangan_m as ru on ru.id=apd.objectruanganfk " & _
            "left join departemen_m as dp on dp.id = ru.objectdepartemenfk " & _
            "inner JOIN pasien_m as ps on ps.id=pd.nocmfk " & _
            "where pd.tglregistrasi between '" & tglAwal & "' and '" & tglAkhir & "' " & _
            str1 & _
            str2 & _
            "group by pd.noregistrasi, ps.namapasien, ps.nocm, tp.hargajual, tp.jumlah, " & _
            "ru.namaruangan, tp.produkfk, pro.namaproduk, pro.id, kp.id  " & _
            "order by ps.namapasien"

   
            
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
        
    With Report
        .database.AddADOCommand CN_String, adocmd
            .txtPrinted.SetText namaPrinted
            .txtTanggal.SetText Format(tglAwal, "dd-MM-yyyy")
            .usNmPerkiraan.SetUnboundFieldSource ("{ado.namaperkiraan}")
            '.usNamaRuangan.SetUnboundFieldSource ("{ado.namaruangan}")
            .usRuangan.SetUnboundFieldSource ("{ado.namaruangan}")
            .usRegMR.SetUnboundFieldSource ("{ado.regcm}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usLayanan.SetUnboundFieldSource ("{ado.layanan}")
            .ucTarif.SetUnboundFieldSource ("{ado.hargajual}")
            .unJumlah.SetUnboundFieldSource ("{ado.jumlah}")
            .unKode.SetUnboundFieldSource ("{ado.kode}")
            '.unDebet.SetUnboundFieldSource ("{ado.P_NonJM}")
            '.unKredit.SetUnboundFieldSource ("{ado.P_JM}")
            .ucTotal.SetUnboundFieldSource ("{ado.total}")
            
            
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
