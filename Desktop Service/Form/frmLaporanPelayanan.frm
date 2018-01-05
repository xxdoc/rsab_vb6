VERSION 5.00
Begin VB.Form frmLaporanPelayanan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Pelayanan"
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
Attribute VB_Name = "frmLaporanPelayanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function laporanPelayanan(ByVal QueryText As String) As Byte()
'On Error GoTo errLoad
On Error Resume Next

    Dim Root As JNode
    Dim Param1() As String
    Dim Param2() As String
    Dim Param3() As String
    Dim Param4() As String
    Dim Param5() As String
    Dim Param6() As String
    Dim Param7() As String
    Dim Param8() As String
    Dim Param9() As String
    
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
        Param8 = Split(arrItem(7), "=")
        Param9 = Split(arrItem(8), "=")
        Select Case Param1(0)
            Case "cetak-BukuRegisterPelayanan"
                Call frmCRCetakBukuRegisterPelayanan.CetakBukuRegisterPelayanan(Param2(1), (Param3(1)), Param4(1), Param5(1), (Param6(1)), Param7(1), Param8(1), Param9(1))
                Set Root = New JNode
                Root("Status") = "Cetak Buku Register Pelayanan!!"
                'http://127.0.0.1:1237/printvb/laporanPelayanan?cetak-BukuRegisterPelayanan=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-02%2023:59:59&strIdRuangan=6&strIdDepartement=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
            
            Case "cetak-LaporanPendapatanPoli"
                Call frmCRCetakLaporanPendapatanPoli.CetakLaporanPendapatanPoli(Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1), Param8(1), Param9(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Pendapatan Poli"
                '127.0.0.1:1237/printvb/laporanPelayanan?cetak-LaporanPendapatanPoli=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
            
            Case "cetak-RekapLaporanPendapatanPoli"
                Call frmCRCetakRekapLaporanPendapatanPoli.CetakLaporanPendapatanPoli(Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1), Param8(1), Param8(1))
                Set Root = New JNode
                Root("Status") = "Cetak Rekap Laporan Pendapatan Poli"
                '127.0.0.1:1237/printvb/laporanPelayanan?cetak-LaporanPendapatanPoli=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
    
            Case "cetak-LaporanPendapatanInap"
                Call frmCRLaporanPendapatanInap.CetakLaporanPendapatan(Param1(1), Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1), Param8(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Volume Kegiatan dan Pendapatan"
                '127.0.0.1:1237/printvb/laporanPelayanan?cetak-LaporanPendapatanPoli=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
                
            Case "cetak-LaporanPendapatan-perkelas"
                Call frmCrRincianPendapatanHarianPerKelas.CetakLaporanPendapatan(Param1(1), Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1), Param8(1), Param9(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Volume Kegiatan dan Pendapatan Per Kelas"
                '127.0.0.1:1237/printvb/laporanPelayanan?cetak-LaporanPendapatanPoli=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
                
            Case "cetak-LaporanRekapPendapataninap"
                Call frmCrRekapPendapatanInap.CetakLaporanPendapatan(Param1(1), Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Volume Kegiatan dan Pendapatan Per Kelas"
                '127.0.0.1:1237/printvb/laporanPelayanan?cetak-LaporanPendapatanPoli=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
            
            Case "cetak-detaillayanan"
                Call frmCetakDetailLayananDokter.CetakDetailLayanan(Param1(1), Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1), Param8(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Rekap Layanan"
                '127.0.0.1:1237/printvb/laporanPelayanan?cetak-LaporanPendapatanPoli=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
                
            Case "cetak-rekaplayanan"
                Call frmCetakRekapLayananDokter.CetakRekapLayanan(Param1(1), Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1), Param8(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Rekap Layanan"
                '127.0.0.1:1237/printvb/laporanPelayanan?cetak-LaporanPendapatanPoli=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
            
            Case "cetak-rekapJasaMedisRI"
                Call frmLaporanJasaMedisRI.CetakLaporan(Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Jasa Medis"
                '127.0.0.1:1237/printvb/laporanPelayanan?cetak-LaporanPendapatanPoli=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
            
            Case "cetak-rekapLaboratorium"
                Call frmCrRekapHarianPemeriksaanLaborat.cetak(Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Rekap Pemeriksaan"
                '127.0.0.1:1237/printvb/laporanPelayanan?cetak-LaporanPendapatanPoli=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
            
            Case "cetak-rekapPemeriksaan"
                Call frmCrRekapHarianPemeriksaan.cetak(Param1(1), Param2(1), (Param3(1)), Param4(1), Param5(1), Param6(1), Param7(1))
                Set Root = New JNode
                Root("Status") = "Cetak Laporan Rekap Pemeriksaan"
                '127.0.0.1:1237/printvb/laporanPelayanan?cetak-LaporanPendapatanPoli=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-08%2023:59:59&strIdRuangan=18&strIdKelompokPasien=1&strIdDokter=3&strIdPegawai=1&view=true
            
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
        laporanPelayanan = .Read(adReadAll)
        .Close
    End With
    If CN.State = adStateOpen Then CN.Close
    Unload Me
    Exit Function
    
errLoad:
End Function

