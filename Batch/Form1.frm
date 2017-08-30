VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim newId As Double
    Dim loopi As Double
    
    Dim str As String
    
    ReadRs "select id from pasien_m where kdprofile =0 order by id"
    ReadRs3 "select max(id) from pasien_m where kdprofile<>0"
    For loopi = 0 To RS.RecordCount - 1 Step 10
        newId = loopi + CDbl(RS3(0)) + 1
        str = str + "update pasien_m set id =" & newId & ", kdprofile=1 where id= " & RS!id & ";"
        DoEvents
        RS.MoveNext
        str = str + "update pasien_m set id =" & newId + 1 & ", kdprofile=1 where id= " & RS!id & ";"
        DoEvents
        RS.MoveNext
        str = str + "update pasien_m set id =" & newId + 2 & ", kdprofile=1 where id= " & RS!id & ";"
        DoEvents
        RS.MoveNext
        str = str + "update pasien_m set id =" & newId + 3 & ", kdprofile=1 where id= " & RS!id & ";"
        DoEvents
        RS.MoveNext
        str = str + "update pasien_m set id =" & newId + 4 & ", kdprofile=1 where id= " & RS!id & ";"
        DoEvents
        RS.MoveNext
        str = str + "update pasien_m set id =" & newId + 5 & ", kdprofile=1 where id= " & RS!id & ";"
        DoEvents
        RS.MoveNext
        str = str + "update pasien_m set id =" & newId + 6 & ", kdprofile=1 where id= " & RS!id & ";"
        DoEvents
        RS.MoveNext
        str = str + "update pasien_m set id =" & newId + 7 & ", kdprofile=1 where id= " & RS!id & ";"
        DoEvents
        RS.MoveNext
        str = str + "update pasien_m set id =" & newId + 8 & ", kdprofile=1 where id= " & RS!id & ";"
        DoEvents
        RS.MoveNext
        str = str + "update pasien_m set id =" & newId + 9 & ", kdprofile=1 where id= " & RS!id & ";"
        DoEvents
        
        
        ReadRs2 str
        RS.MoveNext
        
        
        Label1.Caption = newId + 10
    Next
'    For loopi = 0 To RS.RecordCount - 1 Step 4
'        newId = loopi + CDbl(RS3(0)) + 1
'        ReadRs2 "update pasien_m set id =" & newId & ", kdprofile=1 where id= " & RS!id
'        DoEvents
'        RS.MoveNext
'        ReadRs2 "update pasien_m set id =" & newId + 1 & ", kdprofile=1 where id= " & RS!id
'        DoEvents
'        RS.MoveNext
'        ReadRs2 "update pasien_m set id =" & newId + 2 & ", kdprofile=1 where id= " & RS!id
'        DoEvents
'        RS.MoveNext
'        ReadRs2 "update pasien_m set id =" & newId + 3 & ", kdprofile=1 where id= " & RS!id
'        DoEvents
'
'        RS.MoveNext
'
'
'        Label1.Caption = newId + 4
'    Next
End Sub

Private Sub Form_Load()
    Label1.Caption = StatusCN
End Sub
