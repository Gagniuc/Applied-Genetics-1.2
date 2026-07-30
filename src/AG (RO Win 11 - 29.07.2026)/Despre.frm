VERSION 5.00
Begin VB.Form Despre 
   BorderStyle     =   0  'None
   Caption         =   "Despre"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AG.jcFrames Despre_mine 
      Height          =   3975
      Left            =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7011
      BackColor       =   15783104
      FillColor       =   15783104
      Caption         =   "Universitatea Bucuresti -  Institutul de Genetica "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      Begin VB.Timer despreYO 
         Interval        =   10
         Left            =   2520
         Top             =   5160
      End
      Begin AG.isButton ok 
         Height          =   420
         Left            =   1920
         TabIndex        =   4
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   741
         Icon            =   "Despre.frx":0000
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
      Begin AG.aicAlphaImage lacat_Q 
         Height          =   1920
         Left            =   3360
         Top             =   1680
         Width           =   1920
         _ExtentX        =   3413
         _ExtentY        =   3413
         Image           =   "Despre.frx":001C
         Props           =   5
      End
      Begin VB.Label Copyright_de_Paul_Gagniuc 
         BackStyle       =   0  'Transparent
         Caption         =   "© 2008-2009"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4200
         TabIndex        =   5
         Top             =   3600
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   195
         Left            =   3960
         Picture         =   "Despre.frx":28C5
         Top             =   3600
         Width           =   180
      End
      Begin VB.Image Image2 
         Height          =   195
         Left            =   120
         Picture         =   "Despre.frx":2DA3
         Top             =   120
         Width           =   180
      End
      Begin VB.Label L3 
         BackStyle       =   0  'Transparent
         Caption         =   "Autor: Drd. Paul Gagniuc"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   3435
      End
      Begin VB.Label L2 
         BackStyle       =   0  'Transparent
         Caption         =   "----------------------------------------------------------"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   2520
         Width           =   2955
      End
      Begin VB.Label L1 
         BackStyle       =   0  'Transparent
         Caption         =   "Coordonator: Prof. Dr. Lucian Gavrila"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   2160
         Width           =   3435
      End
      Begin VB.Label textul_despre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"Despre.frx":3281
         ForeColor       =   &H00000000&
         Height          =   1485
         Left            =   600
         TabIndex        =   0
         Top             =   3960
         Width           =   4035
      End
   End
End
Attribute VB_Name = "Despre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Despre_mine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub despreYO_Timer()
If textul_despre.Top < 950 Then


If L1.Left > 240 Then
If L2.Left > 240 Then

If L3.Left > 240 Then
despreYO.Enabled = False
Else
L3.Left = L3.Left + 30
End If

Else
L2.Left = L2.Left + 30
End If

Else
L1.Left = L1.Left + 30
End If

Else
textul_despre.Top = textul_despre.Top - 30
End If
End Sub

Private Sub Form_Load()
L1.Left = -L1.Width
L2.Left = -L2.Width
L3.Left = -L3.Width
End Sub

Private Sub ok_Click()
deasupara_reprezentarii_grafice = True
Unload Me
End Sub
