VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FTP Updater"
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4500
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet3 
      Left            =   2520
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   1920
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1320
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Update from FTP ........"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Dim fso As New FileSystemObject
Private Sub donlotFile()
    On Error GoTo openerror
    
    
    'Inet1.protocol = icFTP
    'Inet1.URL = "ftp://192.168.12.3"
    'Inet1.UserName = "svradmin"
    'Inet1.Password = "P@ssw0rd"
    '
    'Inet1.RequestTimeout = 40
    'Inet1.Execute , "GET /home/svradmin/desktop_app/DesktopService/DesktopServiceC.exe c:\DesktopServiceC.exe"
    'Do While Inet1.StillExecuting
    'DoEvents
    'Loop
    'Inet1.Execute , "CLOSE"
    ''MsgBox "Download Completed", vbInformation, "FTP Updater"
    '
    'Inet2.protocol = icFTP
    'Inet2.URL = "ftp://192.168.12.3"
    'Inet2.UserName = "svradmin"
    'Inet2.Password = "P@ssw0rd"
    '
    'Inet2.RequestTimeout = 40
    'Inet2.Execute , "GET /home/svradmin/desktop_app/DesktopService/DesktopServiceD.exe c:\DesktopServiceD.exe"
    'Do While Inet2.StillExecuting
    'DoEvents
    'Loop
    'Inet2.Execute , "CLOSE"
    ''MsgBox "Download Completed", vbInformation, "FTP Updater"
    
    Inet3.Protocol = icFTP
    Inet3.URL = "ftp://192.168.12.3"
    Inet3.UserName = "svradmin"
    Inet3.Password = "P@ssw0rd"
     
    Inet3.RequestTimeout = 40
    Inet3.Execute , "GET /home/svradmin/desktop_app/newversion/DesktopService.exe " & UCase(fso.GetDriveName(App.Path)) & "\DesktopService\app.exe"
    Do While Inet3.StillExecuting
    DoEvents
    Loop
    Inet3.Execute , "CLOSE"
'    MsgBox "Download Completed", vbInformation, "FTP Updater"
    Exit Sub
     
openerror:
    MsgBox "Please check your Internet conection!", vbInformation, "Testing INET"
    Exit Sub
End Sub

Private Sub TerminateProcess(app_exe As String)
    Dim Process As Object
    For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = '" & app_exe & "'")
        Process.Terminate
    Next
End Sub

Private Sub Command1_Click()
'    Dim stt As Boolean
'
    TerminateProcess ("DesktopService.exe")
'    If fso.FileExists(UCase(fso.GetDriveName(App.Path)) & "\DesktopService\app.exe") = True Then
'        fso.DeleteFile UCase(fso.GetDriveName(App.Path)) & "\app" & Format(Now(), "yyyy-MM-dd") & ".exe", True
'        fso.MoveFile UCase(fso.GetDriveName(App.Path)) & "\DesktopService\app.exe", UCase(fso.GetDriveName(App.Path)) & "\app" & Format(Now(), "yyyy-MM-dd") & ".exe"
'    End If
'
'    Call donlotFile
    
    
    Dim lngReturnCode As Long
'    lngReturnCode = Shell("git pull", vbNormalFocus)
    lngReturnCode = Shell(UCase(fso.GetDriveName(App.Path)) & "\DS\DesktopService.exe", vbNormalFocus)
    End
    
'    Label1.Caption = fso.GetFile("c:\temp\DesktopService.exe").DateLastModified
    
End Sub

Private Sub Form_Load()
    Call Command1_Click
End Sub
