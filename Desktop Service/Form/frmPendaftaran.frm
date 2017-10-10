VERSION 5.00
Begin VB.Form frmPendaftaran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendaftaran"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblStatus 
      Caption         =   "Cetak Antrian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmPendaftaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function Pendaftaran(ByVal QueryText As String) As Byte()
    On Error GoTo cetak   'Resume Next
    Dim Root As JNode
    Dim Param1() As String
    Dim Param2() As String
    Dim Param3() As String
    Dim Param4() As String
    Dim Param5() As String
    Dim Param6() As String
    Dim Param7() As String
    Dim Param8() As String
    Dim arrItem() As String
   
    If CN.State = adStateClosed Then Call openConnection
        
    
    If Len(QueryText) > 0 Then
        arrItem = Split(QueryText, "&")
        Param1 = Split(arrItem(0), "=")
        Param2 = Split(arrItem(1), "=")
        Param3 = Split(arrItem(2), "=")
'        Param4 = Split(arrItem(3), "=")
        Select Case Param1(0)
            Case "cek-konek"
                lblStatus.Caption = "Cek"
                Set Root = New JNode
                Root("Status") = "Ok!!"
            
            Case "cetak-kartupasien"
                lblStatus.Caption = "Cetak Kartu Pasien"
                
               Call cetak_KartuPasien(Param2(1))

                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-kartupasien=1&id=1231=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                
            Case "cetak-buktipendaftaran"
                lblStatus.Caption = "Cetak Bukti Pendaftaran"
                Call frmCetakPendaftaran.cetakBuktiPendaftaran(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-buktipendaftaran=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
                
             Case "cetak-tracer"
                lblStatus.Caption = "Cetak Tracer"
                Call frmCetakPendaftaran.cetakTracer(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-tracer=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
                
            Case "cetak-sep"
                lblStatus.Caption = "Cetak SEP"
                Call frmCetakPendaftaran.cetakSep(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-sep=1&norec=40288c835ba4c322015ba816f5d0002c&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
                
            Case "cetak-buktilayanan"
                Param4 = Split(arrItem(3), "=")
                lblStatus.Caption = "Cetak Bukti Layanan"
                Call frmCetakPendaftaran.cetakBuktiLayanan(Param2(1), Param3(1), Param4(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-buktilayanan=1&norec=1707000166&strIdPegawai=1&view=true
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
             
             Case "cetak-buktilayanan-ruangan"
                Param4 = Split(arrItem(3), "=")
                Param5 = Split(arrItem(4), "=")
                lblStatus.Caption = "Cetak Bukti Layanan Ruangan"
                Call frmCetakPendaftaran.cetakBuktiLayananRuangan(Param2(1), Param3(1), Param4(1), Param5(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-buktilayanan-ruangan=1&norec=1707000166&strIdPegawai=1&strIdRuangan=&view=true
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
               
            Case "cetak-labelpasien"
                Param4 = Split(arrItem(3), "=")
                lblStatus.Caption = "Cetak Label Pasien"
                
                Call frmCetakPendaftaran.cetakLabelPasien(Param2(1), Param3(1), Param4(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-labelpasien=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false&qty=2
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
            
            Case "cetak-summarylist"
                lblStatus.Caption = "Cetak Summary list Pasien Rawat Jalan"
                Call frmCetakPendaftaran.cetakSummaryList(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-summarylist=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
            
            Case "cetak-lembarmasukkeluar"
                lblStatus.Caption = "Cetak Lembar Masuk Keluar Pasien Rawat Inap"
                Call frmCetakPendaftaran.cetakLembarMasuk(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-lembarmasukkeluar=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
             
            Case "cetak-lembarpersetujuan"
                lblStatus.Caption = "Cetak Lembar Persetjuan Rawat Inap"
                Call frmCetakPendaftaran.cetakPersetujuan(Param2(1), Param3(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-summarylist=1&norec=2c9090ce5af40be8015af40eb1f80006&view=false
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
        
            Case "cetak-pasiendaftar"
                Param4 = Split(arrItem(3), "=")
                Param5 = Split(arrItem(4), "=")
                Param6 = Split(arrItem(5), "=")
                Param7 = Split(arrItem(6), "=")
                Param8 = Split(arrItem(7), "=")
                
                lblStatus.Caption = "Cetak Pasien Daftar"
                Call frmCRCetakDaftarPasien.CetakPasienDaftar(Param2(1), Param3(1), Param4(1), Param5(1), (Param6(1)), Param7(1), Param8(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-pasiendaftar=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-02%2023:59:59&strIdRuangan=6&strIdDepartement=18&strIdKelompokPasien=1&strIdPegawai=1&view=true
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-pasiendaftar=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-02%2023:59:59&strIdRuangan=&strIdDepartement=&strIdKelompokPasien=&strIdPegawai=1&view=true
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
                
            Case "cetak-sensusbpjs"
                Param4 = Split(arrItem(3), "=")
                Param5 = Split(arrItem(4), "=")
                Param6 = Split(arrItem(5), "=")
                Param7 = Split(arrItem(6), "=")
                Param8 = Split(arrItem(7), "=")
                
                lblStatus.Caption = "Cetak Sensus BPJS"
                Call frmCRCetakSensusBPJS.CetakSensusBPJS(Param2(1), Param3(1), Param4(1), Param5(1), (Param6(1)), Param7(1), Param8(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-sensusbpjs=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-02%2023:59:59&strIdRuangan=6&strIdDepartement=18&strIdKelompokPasien=2&strIdPegawai=1&view=true
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-sensusbpjs=1&tglAwal=2017-08-01%2000:00:00&tglAkhir=2017-09-02%2023:59:59&strIdRuangan=&strIdDepartement=&strIdKelompokPasien=2&strIdPegawai=1&view=true
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
            Case "cetak-buktilayanan-ruangan-pertindakan"
                Param4 = Split(arrItem(3), "=")
                Param5 = Split(arrItem(4), "=")
                Param6 = Split(arrItem(5), "=")
                lblStatus.Caption = "Cetak Bukti Layanan Ruangan"
                Call frmCetakPendaftaran.cetakBuktiLayananRuanganPerTindakan(Param2(1), Param3(1), Param4(1), Param5(1), Param6(1))
                'http://127.0.0.1:1237/printvb/Pendaftaran?cetak-buktilayanan-ruangan=1&norec=1707000166&strIdPegawai=1&strIdRuangan=&view=true
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "grh@epic"
             
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
        Pendaftaran = .Read(adReadAll)
        .Close
    End With
    If CN.State = adStateOpen Then CN.Close
    Unload Me
    Exit Function
cetak:
' MsgBox Err.Description
End Function

Private Sub CETAK_Billing(strNoregistrasi As String, jumlahCetak As Integer)
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
    
    
    strPrinter = GetTxt("Setting.ini", "Printer", "BuktiPendaftaran")
    If Printers.count > 0 Then
        For Each prn In Printers
            If prn.DeviceName = strPrinter Then
                Set Printer = prn
                Exit For
            End If
        Next prn
    End If
    
    For i = 0 To RS.RecordCount - 1
        'MsgBox "CETAK"
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
        Printer.fontSize = 12
          '1,3,,4,6,8,10,12,13,15
        Printer.Print RS(1) & " " & RS(1) & " " & RS(3) & " " & RS(4) & " " & RS(6) & " " & RS(8) & " " & RS(10) & " " & RS(12) & " "
        
        Printer.EndDoc
    Next
    
    Exit Sub
errLoad:
End Sub


Private Sub cetak_KartuPasien(strNocm As String)
    On Error GoTo errLoad
    Dim prn As Printer
    Dim strPrinter As String
    
    strSQL = "SELECT ps.namapasien || ' ( ' ||  jk.reportdisplay || ' ) ' as namapasien ,ps.nocm, ps.tgllahir,ps.namaayah  " & _
            " from pasien_m ps INNER JOIN jeniskelamin_m jk on jk.id=ps.objectjeniskelaminfk " & _
            " where ps.id=" & strNocm & " "
      
     ReadRs strSQL
      
    strPrinter = GetTxt("Setting.ini", "Printer", "KartuPasien")
    If Printers.count > 0 Then
        For Each prn In Printers
            If prn.DeviceName = strPrinter Then
                Set Printer = prn
                Exit For
            End If
        Next prn
    End If
    
    Dim msg As String
    Dim ayah As String
    Dim ayah2 As String
    
    If IsNull(RS!namaayah) = True Then
    ayah = ""
    Else
    ayah = RS!namaayah
    End If
    If IsNull(RS!tgllahir) = True Then
    ayah2 = ""
    Else
    ayah2 = RS!tgllahir
    End If
    'ByVal prnDriver As String, ByVal text As String, ByVal imgPath As String, _
                                ByRef msg As String, nocm As String, namapasien As String, namaayah As String, tgllahir As String
    PrintFrontSideOnly strPrinter, "", "", msg, RS!nocm, RS!namapasien, ayah, ayah2
    
'     If Not RS.EOF Then
'
'
''            Printer.Print "^XA"
''
''            Printer.Print "^CFA,22"
''            Printer.Print "^FO230,170^FD " & RS!namapasien & " ^FS"
''            Printer.Print "^FO230,210^FD " & RS!namaayah & " ^FS"
''            Printer.Print "^FO230,245^FD " & RS!tgllahir & " ^FS"
''
''            Printer.Print "^CFA,45"
''            Printer.Print "^FO50,370^FD " & RS!nocm & " ^FS"
''            Printer.Print "^FO540,270^BQN,2,5^FD    " & RS!nocm & "^FS"
''
''            Printer.Print "^XZ"
'            Printer.EndDoc
'       End If
'    If msg <> "" Then Print #LogFile, _
'                  "Error "; _
'                  CStr(ErrNumber); _
'                  " (&H"; Right$("0000000" & Hex$(ErrNumber), 8); ") "; _
'                  msg
   Exit Sub
   
errLoad:
    
End Sub

