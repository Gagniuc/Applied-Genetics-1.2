VERSION 5.00
Begin VB.Form Intro 
   BorderStyle     =   0  'None
   Caption         =   "Intro"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Intro.frx":0000
   ScaleHeight     =   4335
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer gata_load 
      Interval        =   2500
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Applied Genetics v1.2 beta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      Top             =   0
      Width           =   7095
   End
   Begin AG.aicAlphaImage aicAlphaImage1 
      Height          =   1920
      Left            =   5040
      Top             =   2280
      Width           =   1920
      _ExtentX        =   3413
      _ExtentY        =   3413
      Image           =   "Intro.frx":158DF
      Props           =   5
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Applied Genetics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Yo 
      BackStyle       =   0  'Transparent
      Caption         =   "Universitatea Bucuresti, Institutul de Genetica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   4695
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
'MsgBox Command
Me.Show
principal.Hide
Mai_mult.Hide

'For i = 1 To 64
'codonul (i)
'Next i


End Sub


Private Sub gata_load_Timer()

Me.Hide
principal.Show

gata_load.Enabled = False
End Sub

