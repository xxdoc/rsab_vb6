VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakLabelGizi 
   Caption         =   "Medifirst2000"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   6675
   WindowState     =   2  'Maximized
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
      TabIndex        =   3
      Top             =   600
      Width           =   2775
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
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
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
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   120
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
End
Attribute VB_Name = "frmCetakLabelGizi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim reportLabel As New crCetakLabelGizi

Dim ii As Integer
Dim tempPrint1 As String
Dim p As Printer
Dim p2 As Printer
Dim strDeviceName As String
Dim strDriverName As String
Dim strPort As String


Dim strPrinter As String
Dim strPrinter1 As String
Dim PrinterNama As String

Dim adoReport As New ADODB.Command
Private Sub cmdCetak_Click()
    reportLabel.SelectPrinter "winspool", cboPrinter.Text, "Ne00:"
    'PrinterNama = cboPrinter.Text
    reportLabel.PrintOut False
End Sub

Private Sub CmdOption_Click()
    reportLabel.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub Form_Load()
    
    Dim p As Printer
    cboPrinter.Clear
    For Each p In Printers
        cboPrinter.AddItem p.DeviceName
    Next
    strPrinter = strPrinter1
    
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmCetakLabelGizi = Nothing
'    fso.DeleteFile (App.Path & "\tempbitmap.bmp")
'    Set sect = Nothing

End Sub

Public Sub Cetak(noregistrasi As String, view As String, qty As String)
'On Error GoTo errLoad
Set frmCetakLabelGizi = Nothing
Dim strSQL As String
Dim i As Integer
Dim str As String
Dim jml As Integer
    
    

    With reportLabel
            Set adoReport = New ADODB.Command
             adoReport.ActiveConnection = CN_String
            
            strSQL = "select pd.noregistrasi, sk.nokirim, sk.qtyproduk as qtykirim,op.keteranganlainnya as keteranganlainnyakirim,cast (op.qtyprodukinuse as text) as cc  ,cast (op.jumlah as text) as volume , pd.tglregistrasi,  ps.tgllahir, " & _
            "to_date(to_char(ps.tgllahir, 'YYYY/MM/DD'), 'YYYY/MM/DD') as tgllahirNew,ps.namapasien, ps.nocm, ru.namaruangan as ruanganasal, " & _
            "jw.jeniswaktu,jd.jenisdiet,  op.qtyproduk,kls.namakelas " & _
            "from orderpelayanan_t as op " & _
            "inner join pasiendaftar_t as pd on pd.norec = op.noregistrasifk " & _
            "inner join ruangan_m as ru on ru.id = op.objectruanganfk " & _
            "inner join pasien_m as ps on ps.id = op.nocmfk " & _
            "left join jeniskelamin_m as jk on jk.id = ps.objectjeniskelaminfk " & _
            "inner join strukorder_t as so on so.norec = op.strukorderfk " & _
            "left join strukkirim_t as sk on  sk.norec = op.strukkirimfk " & _
            "inner join jeniswaktu_m as jw on jw.id = op.objectjeniswaktufk " & _
            "inner join jenisdiet_m as jd on jd.id = op.objectjenisdietfk " & _
            "inner join kategorydiet_m as kd on kd.id = op.objectkategorydietfk " & _
            "left join kelas_m as kls on kls.id = op.objectkelasfk " & _
            "where sk.norec= '" & noregistrasi & "' "
'
              ReadRs strSQL
'            jml = qty - 1
             Dim strDate
'             strDate = getBulan(Format(RS!tgllahir, "yyyy/MM/dd"))
'             strDate = Format(RS!tgllahir, "yyyy/MM/dd")
             str = ""
             If Val(qty) - 1 = 0 Then
                 adoReport.CommandText = strSQL
              Else
                 For i = 1 To Val(qty) - 1
                     str = strSQL & " union all " & str
                 Next
                 
                 adoReport.CommandText = str & strSQL
            End If
          
            
             adoReport.CommandType = adCmdUnknown
             .database.AddADOCommand CN_String, adoReport
           If RS.BOF Then
                .txtUmur.SetText "-"
            Else
                .txtUmur.SetText hitungUmur(Format(RS!tgllahirNew, "yyyy/MM/dd"), Format(Now, "yyyy/MM/dd"))
            End If
            .usTglLahir.SetUnboundFieldSource ("{ado.tgllahirNew}")
            .usNoreg.SetUnboundFieldSource ("{ado.noregistrasi}")
            .usNamaPasien.SetUnboundFieldSource ("{ado.namapasien}")
            .usNocm.SetUnboundFieldSource ("{ado.nocm}")
'            .udtTglLahir.SetUnboundFieldSource ("{ado.tgllahir}")
            .usRuangan.SetUnboundFieldSource ("{ado.ruanganasal}")
            .usKelas.SetUnboundFieldSource ("{ado.namakelas}")
            .usJenisDiet.SetUnboundFieldSource ("{ado.jenisdiet}")
            .usJenisWaktu.SetUnboundFieldSource ("{ado.jeniswaktu}")
            .usKet.SetUnboundFieldSource ("{ado.keteranganlainnyakirim}")
            .usCC.SetUnboundFieldSource ("{ado.cc}")
            .usVolume.SetUnboundFieldSource ("{ado.volume}")
            
            If view = "false" Then
                strPrinter1 = GetTxt("Setting.ini", "Printer", "LabelGizi")
                .SelectPrinter "winspool", strPrinter1, "Ne00:"
                .PrintOut False
                Unload Me
                Screen.MousePointer = vbDefault
             Else
                With CRViewer1
                    .ReportSource = reportLabel
                    .ViewReport
                    .Zoom 1
                End With
                Me.Show
                Screen.MousePointer = vbDefault
            End If
     
    End With
Exit Sub
errLoad:

    MsgBox Err.Number & " " & Err.Description
End Sub



