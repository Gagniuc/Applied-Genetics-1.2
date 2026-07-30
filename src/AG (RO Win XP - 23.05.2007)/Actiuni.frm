VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCTL32.ocx"
Begin VB.Form Actiuni 
   BorderStyle     =   0  'None
   Caption         =   "Actiuni"
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AG.jcFrames Nu_conteaza 
      Height          =   2295
      Left            =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4048
      BackColor       =   15783104
      FillColor       =   15783104
      Caption         =   "Cautare"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderStyle     =   1
      GradientHeaderStyle=   3
      Begin TabDlg.SSTab actiuni_bloc_cod 
         Height          =   1575
         Left            =   0
         TabIndex        =   0
         Top             =   480
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2778
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Cauta pozitie"
         TabPicture(0)   =   "Actiuni.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "du_te_n"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cat_n"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Cauta secventa"
         TabPicture(1)   =   "Actiuni.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "G_text"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "gaseste_t"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "coloreaza"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin VB.CheckBox coloreaza 
            Caption         =   "Coloreaza secventa cautata !"
            Height          =   255
            Left            =   -74520
            TabIndex        =   3
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox gaseste_t 
            BackColor       =   &H00000040&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   -72360
            TabIndex        =   2
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox cat_n 
            BackColor       =   &H00000040&
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2640
            TabIndex        =   1
            Text            =   "1"
            Top             =   600
            Width           =   2295
         End
         Begin AG.isButton G_text 
            Height          =   420
            Left            =   -74520
            TabIndex        =   4
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   741
            Icon            =   "Actiuni.frx":0038
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
         Begin AG.isButton du_te_n 
            Height          =   420
            Left            =   480
            TabIndex        =   5
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   741
            Icon            =   "Actiuni.frx":0054
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
      End
      Begin VB.Image Image2 
         Height          =   195
         Left            =   120
         Picture         =   "Actiuni.frx":0070
         Top             =   120
         Width           =   180
      End
   End
End
Attribute VB_Name = "Actiuni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub du_te_n_Click()
principal.COD_CODON.SelStart = cat_n - 1
principal.COD_CODON.SelLength = 1
deasupara_reprezentarii_grafice = True
Unload Me
End Sub

Private Sub G_text_Click()
If InStr(principal.COD_CODON.Text, gaseste_t.Text) <> 0 Then
Else
MsgBox "Secventa cautata NU exista !"
Exit Sub
End If

If coloreaza.Value = 1 Then
HighLightWord principal, principal.COD_CODON, gaseste_t.Text, RGB(222, 0, 0), True, True
End If
principal.COD_CODON.Find gaseste_t.Text, 1, Len(principal.COD_CODON.Text)
deasupara_reprezentarii_grafice = True
Unload Me
End Sub


Private Sub Nu_conteaza_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub
