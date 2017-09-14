VERSION 5.00
Begin VB.Form frmKasir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kasir"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function Kasir(ByVal QueryText As String) As Byte()
On Error Resume Next
    Dim Root As JNode
    Dim Param1() As String
    Dim Param2() As String
    Dim Param3() As String
    Dim Param4() As String
    Dim Param5() As String
    Dim Param6() As String
    Dim Param7() As String
    Dim arrItem() As String
    
   If CN.State = adStateClosed Then Call openConnection
    
    If Len(QueryText) > 0 Then
    
        arrItem = Split(QueryText, "&")
        Param1 = Split(arrItem(0), "=")
        Param2 = Split(arrItem(1), "=")
        Param3 = Split(arrItem(2), "=")
        Param4 = Split(arrItem(3), "=")
        Param5 = Split(arrItem(4), "=")
        Param6 = Split(arrItem(5), "=")
        Param7 = Split(arrItem(6), "=")

        Select Case Param1(0)
            Case "cetak-billing"
                Call frmCRCetakBilling.CetakBilling(Param2(1), Val(Param1(1)), Param3(1))
                Set Root = New JNode
                Root("Status") = "Cetak Billing!!"
                '127.0.0.1:1237/printvb/kasir?cetak-billing=1&noregistrasi=1707000053&view=false
            Case "cetak-kwitansi"
                Call frmCRCetakKuitansiPasien.CetakUlangJenisKuitansi(Param2(1), Val(Param1(1)), Param3(1))
                Set Root = New JNode
                Root("Status") = "Cetak Kwitansi"
                '127.0.0.1:1237/printvb/kasir?cetak-kwitansi=1&noregistrasi=1708000446&view=false
            Case "cetak-kwitansiv2"
                Param4 = Split(arrItem(3), "=")
                Call frmCRCetakKuitansiPasienV2.CetakUlangJenisKuitansi(Param2(1), Val(Param1(1)), Param3(1), Param4(1))
                Set Root = New JNode
                Root("Status") = "Cetak Kwitansi"
                '127.0.0.1:1237/printvb/kasir?cetak-kwitansiv2=1&noregistrasi=1708000446&strIdPegawai=1&view=false
            Case "cetak-LaporanPasienPulang"
                Param4 = Split(arrItem(3), "=")
                Param5 = Split(arrItem(4), "=")
                Param6 = Split(arrItem(5), "=")
                Param7 = Split(arrItem(6), "=")
                Call frmCRCetakLaporanPasienPulang.CetakLaporanPasienPulang(Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Pasien Pulang"
                '127.0.0.1:1237/printvb/kasir?cetak-LaporanPasienPulang=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdPegawai=1&view=true
            Case "cetak-LaporanPendapatanRuangan"
                Param4 = Split(arrItem(3), "=")
                Param5 = Split(arrItem(4), "=")
                Param6 = Split(arrItem(5), "=")
                Param7 = Split(arrItem(6), "=")
                Call frmCRCetakLaporanPendapatanRuangan.CetakLaporanPendapatanRuangan(Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Pendapatan Ruangan"
                '127.0.0.1:1237/printvb/kasir?cetak-LaporanPendapatanRuangan=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdPegawai=1&view=true
            Case "cetak-RincianBiaya"
                Param4 = Split(arrItem(3), "=")
                Param5 = Split(arrItem(4), "=")
                Param6 = Split(arrItem(5), "=")
                Call frmCRCetakRincianBiaya.CetakRincianBiaya(Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1))
                Set Root = New JNode
                Root("Status") = "Cetak Rincian Biaya"
                '127.0.0.1:1237/printvb/kasir?cetak-RincianBiaya=1&strNoregistrasi=1707000166&strNoStruk=S000000159&strNoKwitansi=RV-17080000002&strIdPegawai=1&view=true
                '127.0.0.1:1237/printvb/kasir?cetak-RincianBiaya=1&strNoregistrasi=&strNoStruk=S000000168&strNoKwitansi=&strIdPegawai=1&view=true
            Case "cetak-LaporanPendapatanRuanganDetail"
                Param4 = Split(arrItem(3), "=")
                Param5 = Split(arrItem(4), "=")
                Param6 = Split(arrItem(5), "=")
                Param7 = Split(arrItem(6), "=")
                Call frmCRCetakLaporanPendapatanRuanganDetail.CetakLaporanPendapatanRuanganDetail(Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Pendapatan Ruangan Detail"
                '127.0.0.1:1237/printvb/kasir?cetak-LaporanPendapatanRuanganDetail=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdPegawai=1&view=true
            
            Case "cetak-laporan-penerimaan"
                Call frmCRLaporanPenerimaan.CetakLaporanPenerimaan(Param1(1), Param2(1), Param3(1), Param4(1), Param5(1), Param6(1), Param7(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Penerimaan Kasir"
                Root("by") = "as@epic"
                'http://127.0.0.1:1237/printvb/kasir?cetak-laporan-penerimaan=403&tglAwal=2017-09-02&tglAkhir=2017-09-02&idRuangan=&idDokter=&view=true
            
            Case "cetak-rekap-penerimaan"
                Call frmRekapPenerimaan.CetakRekapPenerimaan(Param1(1), Param2(1), Param3(1), Param4(1), Param5(1), Param6(1), Param7(1))
                Set Root = New JNode
                Root("Status") = "Cetak Rekap Penerimaan Kasir"
                Root("by") = "as@epic"
                'http://127.0.0.1:1237/printvb/kasir?cetak-rekap-penerimaan=403&tglAwal=2017-09-02&tglAkhir=2017-09-02&idRuangan=&idDokter=&view=true
                
            Case "cetak-laporan-pendapatan"
                If CN.State = adStateClosed Then Call openConnection
                Call frmCRLaporanPendapatan.CetakLaporanPendapatan(Param1(1), Param2(1), Param3(1), Param4(1), Param5(1), Param6(1), Param7(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Penerimaan Kasir"
                Root("by") = "as@epic"
                '127.0.0.1:1237/printvb/kasir?cetak-laporan-pendapatan=403&tglAwal=2017-09-02&tglAkhir=2017-09-02&view=false
            
            Case "cetak-rekap-pendapatan"
                Call frmRekapPendapatan.CetakRekapPendapatan(Param1(1), Param2(1), Param3(1), Param4(1), Param5(1), Param6(1), Param7(1))
                Set Root = New JNode
                Root("Status") = "Cetak Rekap Penerimaan Kasir"
                Root("by") = "as@epic"
                '127.0.0.1:1237/printvb/kasir?cetak-rekap-pendapatan=403&tglAwal=2017-09-02&tglAkhir=2017-09-02&view=false
            
            Case Else
                Set Root = New JNode
                Root("Status") = "Error"
        End Select
    End If
    With GossRESTMain.STM
        .Open
        .Type = adTypeText
        .CharSet = "utf-8"
        .WriteText Root.JSON, adWriteChar
        .Position = 0
        .Type = adTypeBinary
        Kasir = .Read(adReadAll)
        .Close
    End With
    If CN.State = adStateOpen Then CN.Close
    Unload Me
    Exit Function
    
errLoad:
End Function

Private Sub CETAK_Billing(strNoregistrasi As String, jumlahCetak As Integer, view As String)
On Error GoTo errLoad
    Dim prn As Printer
    Dim strPrinter As String
  
    ReadRs "select pp.norec,pp.tglpelayanan,pr.id as prid, pr.namaproduk, pp.jumlah,kl.id as klid, kl.namakelas, " & _
           "ru.id as ruid,ru.namaruangan,pp.produkfk,pp.hargajual,pg.id as pgid,pg.namalengkap,sp.nostruk, " & _
           "jpp.id as jppid,jpp.jenispetugaspe from " & _
           "pasiendaftar_t As pd " & _
           "inner  join antrianpasiendiperiksa_t as apd on apd.noregistrasifk= pd.norec " & _
           "inner join pelayananpasien_t as pp on pp.noregistrasifk= apd.norec " & _
           "inner join produk_m as pr ON pr.id= pp.produkfk " & _
           "inner JOIN  kelas_m as kl ON kl.id= apd.objectkelasfk " & _
           "inner join ruangan_m as ru ON ru.id= apd.objectruanganfk " & _
           "inner join pelayananpasienpetugas_t as ptu ON ptu.pelayananpasien= pp.norec " & _
           "inner join jenispetugaspelaksana_m as jpp ON jpp.id= ptu.objectjenispetugaspefk " & _
           "inner join pegawai_m as pg ON pg.id= ptu.objectpegawaifk " & _
           "left join strukpelayanan_t as sp ON sp.norec= pp.strukfk " & _
           "Where pd.tglpulang Is Not Null " & _
           "and pd.noregistrasi='" & strNoregistrasi & "'"
    
'    Dim NoAntri As String
'    Dim jmlAntrian As Integer
'    Dim jenis As String
'
'    Set RS = Nothing
'    RS.Open "select * from antrianpasienregistrasi_t where norec ='" & strNorec & "'", CN, adOpenStatic, adLockReadOnly
'    NoAntri = RS!jenis & "-" & RS!noantrian
'    jenis = RS!jenis
'    Set RS = Nothing
'    RS.Open "select count(noantrian) as jmlAntri from antrianpasienregistrasi_t where jenis ='" & jenis & "' and statuspanggil='0'", CN, adOpenStatic, adLockReadOnly
'    jmlAntrian = RS(0)
    
    strPrinter = GetTxt("Setting.ini", "Printer", "CetakAntrian")
    'GetSetting("Jasamedika Service", "CetakAntrian", "Printer")
    If Printers.count > 0 Then
        For Each prn In Printers
            If prn.DeviceName = strPrinter Then
                Set Printer = prn
                Exit For
            End If
        Next prn
    End If
    
   '
    Printer.fontSize = 10
        Printer.Print "     RUMAH SAKIT ANAK DAN BUNDA"
        Printer.fontSize = 18
        Printer.FontBold = True
        Printer.Print "      HARAPAN KITA"
        Printer.FontBold = False
        Printer.fontSize = 10
        Printer.Print "   Jl. S. Parman Kav.87, Slipi, Jakarta Barat"
        Printer.Print "      Telp. 021-5668286, 021-5668284"
        Printer.Print "      Fax.  021-5601816, 021-5673832"
        Printer.Print "___________________________________"
        Printer.Print ""
        Printer.Print "Tanggal :" & Format(Now(), "yyyy MM dd hh:mm")
        Printer.Print ""
    For i = 0 To RS.RecordCount - 1
        'MsgBox "CETAK"

        Printer.fontSize = 12
          '1,3,,4,6,8,10,12,13,15
        Printer.Print RS(1) & " " & RS(1) & " " & RS(3) & " " & RS(4) & " " & RS(6) & " " & RS(8) & " " & RS(10) & " " & RS(12) & " "
        RS.MoveNext
    Next
    Printer.EndDoc
    Exit Sub
errLoad:
End Sub




