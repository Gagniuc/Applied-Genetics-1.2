VERSION 5.00
Begin VB.Form Trial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "      Applied Genetics - Trial"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Trial.frx":0000
   ScaleHeight     =   5415
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrMain 
      Interval        =   10000
      Left            =   5160
      Top             =   5640
   End
   Begin AG.isButton CmdEnter 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "Trial.frx":23C6
      Style           =   7
      Caption         =   "Continue"
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Informatii"
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
   Begin AG.isButton CmdEntSerial 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "Trial.frx":23E2
      Style           =   7
      Caption         =   "Validation"
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Informatii"
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
   Begin AG.isButton cmdExit 
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "Trial.frx":23FE
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
   Begin AG.isButton Buy 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Icon            =   "Trial.frx":241A
      Style           =   7
      Caption         =   "Internet"
      iNonThemeStyle  =   0
      Tooltiptitle    =   "Informatii"
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
   Begin VB.Label anunt_porniri 
      BackStyle       =   0  'Transparent
      Caption         =   "The program can be used 30 times"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   720
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      Height          =   2055
      Left            =   360
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Label Mesaj_trail 
      BackStyle       =   0  'Transparent
      Caption         =   $"Trial.frx":2436
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1695
      Left            =   480
      TabIndex        =   8
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      Height          =   4215
      Left            =   240
      Top             =   240
      Width           =   5175
   End
   Begin AG.aicAlphaImage aicAlphaImage1 
      Height          =   1920
      Left            =   3600
      Top             =   480
      Width           =   1920
      _ExtentX        =   3413
      _ExtentY        =   3413
      Image           =   "Trial.frx":2528
      Props           =   5
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "beta"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "v1.2 "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Yo 
      BackStyle       =   0  'Transparent
      Caption         =   "University of Bucharest, Institute of Genetics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Applied Genetics "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
End
Attribute VB_Name = "Trial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************************************'
'                                             '
' SimpleTrial                                 '
' Feel free to re-distrubute this code, since '
' this code is freeware :).                   '
'                                             '
' Please vote for me.                         '
'                                             '
'*********************************************'

Dim clsDS2 As New clsDS2

Private dir_ale_trial As String

Private Sub cmdEnter_Click()

If Trial_pa = True Then Exit Sub

Dim Nr_trial As Variant
Nr_trial = GetSetting("WinSys", "system", "4d3fd400f6f9b354f1ff")
Dim F As Integer
If Nr_trial = "" And FileExist(dir_ale_trial & "sys_exp.dll") = False Then
    
    F = FreeFile
    Open dir_ale_trial & "sys_exp.dll" For Binary As #F
    Put #F, , "29"
    Close #F
Call SaveSetting("WinSys", "system", "4d3fd400f6f9b354f1ff", "29")
Else
If Nr_trial >= 1 Then Call SaveSetting("WinSys", "system", "4d3fd400f6f9b354f1ff", Val(Nr_trial) - 1)
    Dim tmp_srt_fis As String
    tmp_srt_fis = Val(Nr_trial) - 1
    F = FreeFile
    Kill dir_ale_trial & "sys_exp.dll"
    Open dir_ale_trial & "sys_exp.dll" For Binary As #F
    Put #F, , tmp_srt_fis
    Close #F
End If

Intro.Show
Intro.gata_load.Enabled = True
Unload Me 'Me.Hide
End Sub

Private Sub Buy_Click()
Dim R As Variant
R = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & adresa_site_general & "cumpara.php", vbMaximizedFocus)
End Sub

Private Sub Buy_MouseEnter()
Mesaj_trail.Caption = "To get an activation key for the Applied Genetics software, press this button !"
End Sub

Private Sub Buy_MouseLeave()
Mesaj_trail.Caption = "If you have the original CD of this software, you will find the activation key inside the case. Otherwise, you can get an activation key online! If you want to test the software, press the 'Continue' button ! The program is NOT activated."
End Sub

Private Sub CmdEnter_MouseEnter()
Mesaj_trail.Caption = "To continue using the Applied Genetics software, press this button !"
End Sub

Private Sub CmdEnter_MouseLeave()
Mesaj_trail.Caption = "If you have the original CD of this software, you will find the activation key inside the case. Otherwise, you can get an activation key online! If you want to test the software, press the 'Continue' button ! The program is NOT activated."
End Sub

Private Sub CmdEntSerial_MouseEnter()
Mesaj_trail.Caption = "To enter the activation key for the Applied Genetics software, press this button !"
End Sub

Private Sub CmdEntSerial_MouseLeave()
Mesaj_trail.Caption = "If you have the original CD of this software, you will find the activation key inside the case. Otherwise, you can get an activation key online! If you want to test the software, press the 'Continue' button ! The program is NOT activated."
End Sub

Private Sub cmdExit_Click()

    'Terminate the program if the user decides to.
    Unload Principal
    Unload Mai_mult
    Unload Intro
    Unload Trial_COD
    Unload Me

End Sub


Private Sub CmdEntSerial_Click()

    'Load the details entry form.
        Trial_COD.Show
End Sub

Private Sub TmrMain_Timer()
    
    'Delay the enter button.
        CmdEnter.Enabled = True

End Sub

Private Sub Form_Load()
On Error Resume Next

'Intro.Hide
'Intro.gata_load.Enabled = False

CmdEntSerial.ToolTipText = "To enter the activation key, press here !"
Buy.ToolTipText = "To get an activation key OnLine, press here !"
CmdEnter.ToolTipText = "To continue testing the program " & vbCrLf & "software in TRIAL mode, press here !"

Trial_pa = False
dir_ale_trial = SpecialFolderPath(41&)

    Dim Line01 As String
    Dim Line02 As String
    
    'Open trial config file to check if the software is registered or not.
    Open dir_ale_trial & "ag_exp.001" For Input As #1
    
    'Grab details from config file.
    Line Input #1, Line01
    Line Input #1, Line02
    Close #1
    
    'Decrypt the text using DS2 Cipher decryption.
    Line01 = clsDS2.DecryptString(Line01, Chr(54) & Chr(57) & Chr(51) & Chr(52) & Chr(49) & Chr(48) & Chr(50) & Chr(53) & Chr(54) & Chr(51) & Chr(48) & Chr(50) & Chr(50) & Chr(56) & Chr(52), True)
    Line02 = clsDS2.DecryptString(Line02, Chr(54) & Chr(57) & Chr(51) & Chr(52) & Chr(49) & Chr(48) & Chr(50) & Chr(53) & Chr(54) & Chr(51) & Chr(48) & Chr(50) & Chr(50) & Chr(56) & Chr(52), True)
    
    'Check to see if the text matches a valid registration code.
    If KeyGen(Line01, Chr(54) & Chr(48) & Chr(57) & Chr(51) & Chr(52) & Chr(49) & Chr(52) & Chr(53) & Chr(54) & Chr(49) & Chr(50) & Chr(53) & Chr(49) & Chr(52) & Chr(53), 3) = Line02 Then
    Intro.Show
    Intro.gata_load.Enabled = True
    Unload Me
    Else
Dim Nr_trial As Variant
Nr_trial = GetSetting("WinSys", "system", "4d3fd400f6f9b354f1ff")

If Nr_trial = "" And FileExist(dir_ale_trial & "sys_exp.dll") = True Then
MsgBox "The program has already been installed on this PC !" & vbCrLf & " Using the program in TRIAL form on this PC is not possible !"
End
End If

If Nr_trial <> "" And FileExist(dir_ale_trial & "sys_exp.dll") = False Then
MsgBox "The program has already been installed on this PC !" & vbCrLf & " Using the program in TRIAL form on this PC is not possible !"
End
End If

anunt_porniri.Caption = "The program can be used " & Nr_trial & " more times"


If FileExist(dir_ale_trial & "sys_exp.dll") = True Then
  Dim Filenr As Integer
  Filenr = FreeFile
  Dim BitArray As String
  Open dir_ale_trial & "sys_exp.dll" For Binary As #Filenr
  BitArray = Space(LOF(Filenr))
  Get #Filenr, , BitArray
  Close #Filenr
Else
BitArray = ""
End If

If Nr_trial = BitArray Then
'MsgBox Nr_trial & "=" & BitArray
Else
'MsgBox Nr_trial & "=" & BitArray
MsgBox "The program is closing due to user actions! Thank you."
End
End If


If Val(Nr_trial) > 30 Then

anunt_porniri.Caption = "The number of uses has expired !"
CmdEnter.Caption = "Expired !"
CmdEnter.ToolTipForeColor = vbRed
CmdEnter.ToolTipText = "The number of uses has expired !"
Trial_pa = True
End If

If Nr_trial = "0" Then
anunt_porniri.Caption = "The number of uses has expired !"
CmdEnter.Caption = "Expired !"
CmdEnter.ToolTipForeColor = vbRed
CmdEnter.ToolTipText = "The number of uses has expired !"
Trial_pa = True
End If

If Nr_trial = "" Then anunt_porniri.Caption = "The program can be used 30 times"
    End If
    
'-------------------------------------------------------|
adresa_site_general = "http://www.google.com"

If FExist(App.Path & "\sit.dat") = True Then
  Filenr = FreeFile
  Open App.Path & "\sit.dat" For Binary As #Filenr
  BitArray = Space(LOF(Filenr))
  Get #Filenr, , BitArray
  Close #Filenr
  
  adresa_site_general = BitArray
End If
'-------------------------------------------------------|
End Sub
