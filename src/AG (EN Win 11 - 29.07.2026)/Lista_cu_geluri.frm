VERSION 5.00
Begin VB.Form Lista_cu_geluri 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Inverted repeated sequences"
   ClientHeight    =   8715
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer SRI 
      Interval        =   50
      Left            =   12240
      Top             =   7920
   End
   Begin VB.VScrollBar SI_derulare 
      Height          =   7815
      Left            =   12480
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox fer_seg1 
      BackColor       =   &H00400000&
      Height          =   7815
      Left            =   0
      ScaleHeight     =   517
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   829
      TabIndex        =   1
      Top             =   0
      Width           =   12495
      Begin VB.PictureBox fer_seg2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   7575
         Left            =   0
         ScaleHeight     =   503
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   831
         TabIndex        =   2
         Top             =   0
         Width           =   12495
         Begin VB.PictureBox secvente_inverse_nr 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            ForeColor       =   &H00404040&
            Height          =   330
            Index           =   0
            Left            =   0
            ScaleHeight     =   20
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   823
            TabIndex        =   3
            Top             =   0
            Width           =   12375
         End
      End
   End
   Begin VB.PictureBox Gel_nr_X 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3255
      Index           =   0
      Left            =   0
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   0
      Top             =   9000
      Width           =   2895
   End
   Begin AG.isButton iesire 
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   8040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Icon            =   "Lista_cu_geluri.frx":0000
      Style           =   7
      Caption         =   "Exit"
      iNonThemeStyle  =   0
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
End
Attribute VB_Name = "Lista_cu_geluri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Lista_cu_geluri.Icon = principal.Icon
End Sub

Private Sub Gel_nr_X_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 Gel_nr_X(Index).Left = Gel_nr_X(Index).Left + X
End If
End Sub

Private Sub Iesire_Click()
deasupara_reprezentarii_grafice = True
Me.Hide
End Sub

Private Sub SI_derulare_Change()
SRI.Enabled = False
End Sub

Private Sub SI_derulare_Scroll()
SRI.Enabled = True
End Sub

Private Sub SRI_Timer()
fer_seg2.Top = -(SI_derulare.Value)
End Sub

Private Sub Gel_nr_X_DblClick(Index As Integer)
MsgBox Index
End Sub
