VERSION 5.00
Begin VB.Form Generator 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "   Generator"
   ClientHeight    =   570
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox tipulNUC 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "ADN"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox catNUC 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin AG.isButton okyd 
      Height          =   300
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      Icon            =   "Generator.frx":0000
      Style           =   7
      Caption         =   "Ok"
      iNonThemeStyle  =   0
      BackColor       =   12937777
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   240
      Picture         =   "Generator.frx":001C
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "Nr baze:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Generator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
tipulNUC.AddItem "ADN"
tipulNUC.AddItem "ARN"

tipulNUC.Text = "ADN"
End Sub

Private Sub Form_Unload(Cancel As Integer)
deasupara_reprezentarii_grafice = True
End Sub

Private Sub okyd_Click()

If IsDigitsOnly(catNUC.Text) = True Then
principal.out_put.Text = GENEREAZA_NUCLEOTIDE(catNUC.Text, tipulNUC.Text)
Unload Me
Else
MsgBox "Introduceti doar numere intregi !"
End If
End Sub

Function IsDigitsOnly(Value As String) As Boolean
IsDigitsOnly = Len(Value) > 0 And _
Not Value Like "*[!0-9]*"
End Function

