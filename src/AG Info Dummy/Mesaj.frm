VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Applied Genetics - Mesaj"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   5655
   Icon            =   "Mesaj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label msg 
      Alignment       =   2  'Center
      Caption         =   "Applied Genetics v1.2"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo ClickError
If Command = "" Then End
If Command <> "%1" Then
Cale_proiect = Mid(Command, 1, InStrRev(Command, "\"))
tmp = Mid(Command, InStrRev(Command, "\") + 1, Len(Command))
proiectul = Mid(tmp, 1, InStrRev(tmp, ".") - 1)
msg.Caption = "Fisier resursa al proiectului: [" & proiectul & "]"
ElseIf Command <> "" Then
Else
End If

ClickError:
    If Err.Number <> 0 Then Exit Sub
End Sub

Private Sub Ok_Click()
End
End Sub
