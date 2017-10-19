VERSION 5.00
Begin VB.Form frmPendaftaran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendaftaran"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Width           =   6375
   End
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
                
            Case "RIS"
                Dim lngReturnCode As Long
                Dim strShellCommand As String
                
                
                strShellCommand = "c:\Program Files\Mozilla Firefox\firefox.exe zetta://URL=http://192.168.12.11&LID=dok&LPW=dok&LICD=003&PID=" & Param2(1) & "&VTYPE=" & Param3(1) & ""
                
                 lngReturnCode = Shell(strShellCommand, vbNormalFocus)
                Set Root = New JNode
                Root("Status") = "Sedang Dicetak!!"
                Root("by") = "as@epic"
             
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
    
'    Call DrawBarcode(Text1, Picture2)
    
    Dim msg As String
    Dim ayah As String
    Dim ayah2 As String
    
    'Special128
    If IsNull(RS!namaayah) = True Then
    ayah = ""
    Else
    ayah = RS!namaayah
    End If
    If IsNull(RS!tgllahir) = True Then
    ayah2 = ""
    Else
    ayah2 = Format(RS!tgllahir, "dd-MMM-yyyy")
    End If
'    "Free 3 of 9 Extended"
    Printer.FontName = "Tahoma"
    Printer.fontSize = 10
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.FontBold = True
    Printer.Print "                                         " & Left(ayah, 17)
    Printer.Print "                                         " & Left(RS!namapasien, 17)
    Printer.Print "                                         " & ayah2
    Printer.Print ""
    
    Printer.FontBold = False
'    Printer.FontName = "Free 3 of 9 Extended" '"Bar-Code 39"
'    Printer.fontSize = 27 '20
'    Printer.CurrentX = 2900
'    Printer.CurrentY = 2250
    Call make128(RS!nocm)
    Printer.PaintPicture Picture1.Image, 2900, 2250
    
    Printer.FontBold = True
    Printer.FontName = "Tahoma"
    Printer.fontSize = 17
    Printer.CurrentX = 300
    Printer.CurrentY = 2550
    Printer.Print RS!nocm
    Printer.EndDoc
    
    
    
'     PrintFrontSideOnly strPrinter, "", "", msg, RS!nocm, RS!namapasien, ayah, ayah2
    
   Exit Sub
   
errLoad:
    MsgBox Err.Number & " " & Err.Description
End Sub



Private Sub make128(angka As Double)
Dim x As Long, y As Long, pos As Long
Dim Bardata As String
Dim Cur As String
Dim CurVal As Long
Dim chksum As Long
Dim temp As String
Dim BC(106) As String
    'code 128 is basically the ASCII chr set.
    '4 element sizes : 1=narrowest, 4=widest
    BC(0) = "212222" '<SPC>
    BC(1) = "222122" '!
    BC(2) = "222221" '"
    BC(3) = "121223" '#
    BC(4) = "121322" '$
    BC(5) = "131222" '%
    BC(6) = "122213" '&
    BC(7) = "122312" ''
    BC(8) = "132212" '(
    BC(9) = "221213" ')
    BC(10) = "221312" '*
    BC(11) = "231212" '+
    BC(12) = "112232" ',
    BC(13) = "122132" '-
    BC(14) = "122231" '.
    BC(15) = "113222" '/
    BC(16) = "123122" '0
    BC(17) = "123221" '1
    BC(18) = "223211" '2
    BC(19) = "221132" '3
    BC(20) = "221231" '4
    BC(21) = "213212" '5
    BC(22) = "223112" '6
    BC(23) = "312131" '7
    BC(24) = "311222" '8
    BC(25) = "321122" '9
    BC(26) = "321221" ':
    BC(27) = "312212" ';
    BC(28) = "322112" '<
    BC(29) = "322211" '=
    BC(30) = "212123" '>
    BC(31) = "212321" '?
    BC(32) = "232121" '@
    BC(33) = "111323" 'A
    BC(34) = "131123" 'B
    BC(35) = "131321" 'C
    BC(36) = "112313" 'D
    BC(37) = "132113" 'E
    BC(38) = "132311" 'F
    BC(39) = "211313" 'G
    BC(40) = "231113" 'H
    BC(41) = "231311" 'I
    BC(42) = "112133" 'J
    BC(43) = "112331" 'K
    BC(44) = "132131" 'L
    BC(45) = "113123" 'M
    BC(46) = "113321" 'N
    BC(47) = "133121" 'O
    BC(48) = "313121" 'P
    BC(49) = "211331" 'Q
    BC(50) = "231131" 'R
    BC(51) = "213113" 'S
    BC(52) = "213311" 'T
    BC(53) = "213131" 'U
    BC(54) = "311123" 'V
    BC(55) = "311321" 'W
    BC(56) = "331121" 'X
    BC(57) = "312113" 'Y
    BC(58) = "312311" 'Z
    BC(59) = "332111" '[
    BC(60) = "314111" '\
    BC(61) = "221411" ']
    BC(62) = "431111" '^
    BC(63) = "111224" '_
    BC(64) = "111422" '`
    BC(65) = "121124" 'a
    BC(66) = "121421" 'b
    BC(67) = "141122" 'c
    BC(68) = "141221" 'd
    BC(69) = "112214" 'e
    BC(70) = "112412" 'f
    BC(71) = "122114" 'g
    BC(72) = "122411" 'h
    BC(73) = "142112" 'i
    BC(74) = "142211" 'j
    BC(75) = "241211" 'k
    BC(76) = "221114" 'l
    BC(77) = "413111" 'm
    BC(78) = "241112" 'n
    BC(79) = "134111" 'o
    BC(80) = "111242" 'p
    BC(81) = "121142" 'q
    BC(82) = "121241" 'r
    BC(83) = "114212" 's
    BC(84) = "124112" 't
    BC(85) = "124211" 'u
    BC(86) = "411212" 'v
    BC(87) = "421112" 'w
    BC(88) = "421211" 'x
    BC(89) = "212141" 'y
    BC(90) = "214121" 'z
    BC(91) = "412121" '{
    BC(92) = "111143" '|
    BC(93) = "111341" '}
    BC(94) = "131141" '~
    BC(95) = "114113" '<DEL>        *not used in this sub
    BC(96) = "114311" 'FNC 3        *not used in this sub
    BC(97) = "411113" 'FNC 2        *not used in this sub
    BC(98) = "411311" 'SHIFT        *not used in this sub
    BC(99) = "113141" 'CODE C       *not used in this sub
    BC(100) = "114131" 'FNC 4       *not used in this sub
    BC(101) = "311141" 'CODE A      *not used in this sub
    BC(102) = "411131" 'FNC 1       *not used in this sub
    BC(103) = "211412" 'START A     *not used in this sub
    BC(104) = "211214" 'START B
    BC(105) = "211232" 'START C     *not used in this sub
    BC(106) = "2331112" 'STOP

    Picture1.Cls
'    If Text1.Text = "" Then Exit Sub
    pos = 20
    Bardata = angka 'Text1.Text

    'Check for invalid characters, calculate check sum & build temp string
    For x = 1 To Len(Bardata)
        Cur = Mid$(Bardata, x, 1)
        If Cur < " " Or Cur > "~" Then
            Picture1.Print "Invalid Character(s)"
            Exit Sub
        End If
        CurVal = Asc(Cur) - 32
        temp = temp + BC(CurVal)
        chksum = chksum + CurVal * x
    Next
    
    'Add start, stop & check characters
    chksum = (chksum + 104) Mod 103
    temp = BC(104) & temp & BC(chksum) & BC(106)

    'Generate Barcode
    For x = 1 To Len(temp)
        If x Mod 2 = 0 Then
                'SPACE
                pos = pos + (Val(Mid$(temp, x, 1))) + 1
        Else
                'BAR
                For y = 1 To (Val(Mid$(temp, x, 1)))
                    Picture1.Line (pos, 1)-(pos, 58 - 0 * 8)
                    pos = pos + 1
                Next
        End If
    Next

    'Add Label?
'    If Check1(1).Value Then
'        Picture1.CurrentX = 30 + Len(Bardata) * (3 + 1 * 2) 'kinda center
'        Picture1.CurrentY = 50
'        Picture1.Print Bardata;
'    End If
End Sub



