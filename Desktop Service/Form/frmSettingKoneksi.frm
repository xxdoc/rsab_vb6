VERSION 5.00
Begin VB.Form frmSettingKoneksi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Koneksi"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Database :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Password :"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "User Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Port :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Host :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmSettingKoneksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'    SaveTxt "Setting.ini", "Koneksi", "host", Text1.Text
'    SaveTxt "Setting.ini", "Koneksi", "port", Text2.Text
'    SaveTxt "Setting.ini", "Koneksi", "username", Text3.Text
'    SaveTxt "Setting.ini", "Koneksi", "password", Text4.Text
'    SaveTxt "Setting.ini", "Koneksi", "database", Text5.Text

    SaveTxt "Setting.ini", "Koneksi", "a", Text1.Text
    SaveTxt "Setting.ini", "Koneksi", "b", Text2.Text
    SaveTxt "Setting.ini", "Koneksi", "c", Text3.Text
    SaveTxt "Setting.ini", "Koneksi", "d", Text4.Text
    SaveTxt "Setting.ini", "Koneksi", "e", Text5.Text
    
End Sub

Private Sub Form_Load()
On Error Resume Next
 Text1.Text = GetTxt("Setting.ini", "Koneksi", "a")
   Text2.Text = GetTxt("Setting.ini", "Koneksi", "b")
   Text3.Text = GetTxt("Setting.ini", "Koneksi", "c")
   Text4.Text = GetTxt("Setting.ini", "Koneksi", "d")
   Text5.Text = GetTxt("Setting.ini", "Koneksi", "e")
End Sub
