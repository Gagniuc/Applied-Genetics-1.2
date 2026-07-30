VERSION 5.00
Begin VB.Form Trial_COD 
   BorderStyle     =   0  'None
   Caption         =   "INPUT CODE"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   LinkTopic       =   "Form2"
   ScaleHeight     =   2520
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AG.jcFrames Register_AG 
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      BackColor       =   15783104
      FillColor       =   15783104
      Caption         =   "Validation - Applied Genetics"
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
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox txtGen 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   1440
         Width           =   5895
      End
      Begin AG.isButton CmdReg 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "Trial_COD.frx":0000
         Style           =   7
         Caption         =   "&Validate"
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
      Begin AG.isButton cmdQuit 
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Icon            =   "Trial_COD.frx":001C
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
      Begin AG.aicAlphaImage aicAlphaImage4 
         Height          =   345
         Left            =   0
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         Image           =   "Trial_COD.frx":0038
         Scaler          =   1
         Props           =   5
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   6015
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Trial_COD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Dim clsDS2 As New clsDS2
Private dir_ale_trial As String

'*********************************************'
'                                             '
' SimpleTrial                                 '
' Feel free to re-distrubute this code, since '
' this code is freeware :).                   '
'                                             '
' Please vote for me.                         '
'                                             '
'*********************************************'

Private Sub cmdQuit_Click()

    'Quit the form when a user decides to.
    Unload Me

End Sub

Private Sub CmdReg_Click()
     '& chr(54) & chr(57) & chr(51) & chr(52) & chr(49) & chr(48) & chr(50) & chr(53) & chr(54) & chr(51) & chr(48) & chr(50) & chr(50) & chr(56) & chr(52)
    'Check to see if the user input matches correct information.
        If KeyGen(txtUsername, Chr(54) & Chr(48) & Chr(57) & Chr(51) & Chr(52) & Chr(49) & Chr(52) & Chr(53) & Chr(54) & Chr(49) & Chr(50) & Chr(53) & Chr(49) & Chr(52) & Chr(53), 3) = txtGen Then
            MsgBox "The entered key is VALID !" & vbCrLf & " Redeschideti aplicatia !", vbInformation, "Validare cheie !"
    
    'Encrypt the file to stop people from looking at this hidden info.
        txtUsername.Text = clsDS2.EncryptString(txtUsername.Text, Chr(54) & Chr(57) & Chr(51) & Chr(52) & Chr(49) & Chr(48) & Chr(50) & Chr(53) & Chr(54) & Chr(51) & Chr(48) & Chr(50) & Chr(50) & Chr(56) & Chr(52), True)
        txtGen.Text = clsDS2.EncryptString(txtGen.Text, Chr(54) & Chr(57) & Chr(51) & Chr(52) & Chr(49) & Chr(48) & Chr(50) & Chr(53) & Chr(54) & Chr(51) & Chr(48) & Chr(50) & Chr(50) & Chr(56) & Chr(52), True)
    
    'Write the details to file, if they are correct then the software will be registered.
    Open dir_ale_trial & "ag_exp.002" For Output As #1
        Print #1, txtUsername.Text
        Print #1, txtGen.Text
    Close
    
    'Copy the temp file to the trial config file.
    FileCopy dir_ale_trial & "ag_exp.002", dir_ale_trial & "ag_exp.001"
    Kill dir_ale_trial & "ag_exp.002"
    
    Unload Me

Else
MsgBox "The entered key is NOT valid !", vbCritical, "Validare cheie !"
    End If
End Sub


Private Sub Form_Load()
dir_ale_trial = SpecialFolderPath(41&)
End Sub

Private Sub Register_AG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Variant
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub
