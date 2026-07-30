VERSION 5.00
Begin VB.Form Interfata 
   BorderStyle     =   0  'None
   Caption         =   "Interfata"
   ClientHeight    =   7455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   Icon            =   "Interfata.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Interfata.frx":290EF
   ScaleHeight     =   7455
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.KDCButton ins 
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Appearance      =   15
      Style           =   1
      Caption         =   "Instalare"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorTop    =   16777215
      BackColorBottom =   9201961
      BorderColorTop  =   5912576
      BorderColorBottom=   6504704
   End
   Begin Project1.KDCButton ies 
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   4680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Appearance      =   15
      Style           =   1
      Caption         =   "Iesire"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorTop    =   16777215
      BackColorBottom =   9201961
      BorderColorTop  =   5912576
      BorderColorBottom=   6504704
   End
   Begin Project1.KDCButton tut 
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   3960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Appearance      =   15
      Style           =   1
      Caption         =   "Cum sa ..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorTop    =   16777215
      BackColorBottom =   9201961
      BorderColorTop  =   5912576
      BorderColorBottom=   6504704
   End
   Begin Project1.KDCButton cop 
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      Appearance      =   15
      Style           =   1
      Caption         =   "Copiaza"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorTop    =   16777215
      BackColorBottom =   9201961
      BorderColorTop  =   5912576
      BorderColorBottom=   6504704
   End
   Begin VB.Label la_suta 
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   7680
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label caledir 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   9735
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   9000
      Picture         =   "Interfata.frx":296D6
      Top             =   6480
      Width           =   720
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " (c) 2008"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   6960
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Autor: Drd. Paul Gagniuc"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   7455
      Left            =   0
      Top             =   0
      Width           =   9975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Universitatea din Bucuresti "
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   6720
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Institutul de Genetica"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   6480
      Width           =   3735
   End
   Begin VB.Label titlu 
      BackStyle       =   0  'Transparent
      Caption         =   "Applied Genetics V1.2"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   7095
   End
   Begin VB.Image coperta 
      Height          =   7530
      Left            =   0
      Picture         =   "Interfata.frx":2BB18
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "Interfata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_MEMORY = &H4
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2



Private Sub coperta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub titlu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub ies_Click()
End
End Sub

Private Sub cop_Click()
WAVPlay (Replace(App.Path & "\aferent\apasa.wav", "\\", "\"))

Dim tmpPath As String
tmpPath = txtStart ' Take the selected path from txtStart
If Len(tmpPath) > 0 Then
    If Not Right$(tmpPath, 1) <> "\" Then tmpPath = Left$(tmpPath, Len(tmpPath) - 1) ' Remove "\" if the user added
End If
txtStart = tmpPath
tmpPath = BrowseForFolder(tmpPath) ' Browse for folder
If tmpPath = "" Then
    caledir.Caption = "Nu a fost selectat nici un director !"
Else
    caledir.Caption = "Copiez in: " & tmpPath
    If FileExist(Replace(tmpPath & "\AppliedGenetics_setup.exe", "\\", "\")) Then Kill Replace(tmpPath & "\AppliedGenetics_setup.exe", "\\", "\")
    Call CopyFileAny(Replace(App.Path & "\Kit\AppliedGenetics_setup.exe", "\\", "\"), Replace(tmpPath & "\AppliedGenetics_setup.exe", "\\", "\"))
End If

End Sub

Private Sub ins_Click()
Dim R As Variant
Dim ss As String
WAVPlay (Replace(App.Path & "\aferent\apasa.wav", "\\", "\"))
ss = Replace(App.Path & "\Kit\AppliedGenetics_setup.exe", "\\", "\")
If FileExist(ss) = True Then
R = Shell(ss, vbHide)
Else
MsgBox "Interfata de instalare APPLIED GENETICS nu se afla pe CD-ul de instalare !" & vbCrLf & "Fisierul: " & ss & " nu este prezent in locatia cunoscuta de interfata !"
End If
End Sub

Private Sub tut_Click()
Dim R As Variant
Dim ss As String
WAVPlay (Replace(App.Path & "\aferent\apasa.wav", "\\", "\"))
ss = Replace(App.Path & "\video\index.htm", "\\", "\")
If FileExist(ss) = True Then
R = Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & ss, vbMaximizedFocus)
Else
MsgBox "Interfata de instalare APPLIED GENETICS nu se afla pe CD-ul de instalare !" & vbCrLf & "Fisierul: " & ss & " nu este prezent in locatia cunoscuta de interfata !"
End If
End Sub

Public Function FileExist(aFile As String) As Boolean
On Error GoTo 23

If aFile = "" Then
FileExist = False
Exit Function
End If

If Dir$(aFile) = "" Then
If Dir$(aFile, vbHidden) = "" Then
FileExist = False
Else
FileExist = True
End If
Else
FileExist = True
End If
Exit Function
23:
FileExist = False
'KillPC
End Function

Sub WAVStop()
    Call WAVPlay(" ")
End Sub


Sub WAVLoop(File)
    Dim SoundName As String
    SoundName$ = File
    wFlags% = SND_ASYNC Or SND_LOOP
    X = sndPlaySound(SoundName$, wFlags%)
End Sub


Sub WAVPlay(File)
    Dim SoundName As String
    SoundName$ = File
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound(SoundName$, wFlags%)
End Sub


Public Function CopyFileAny(currentFilename As String, newFilename As String)
    Dim a%, buffer%, temp$, fRead&, fSize&, b%
    On Error GoTo ErrHan:
    a = FreeFile
    buffer = 4048
    Open currentFilename For Binary Access Read As a
    b = FreeFile
    Open newFilename For Binary Access Write As b
    fSize = FileLen(currentFilename)
    


    While fRead < fSize

la_suta.Caption = (100 / fSize) * (fRead / 100) & "%"
        DoEvents
            If buffer > (fSize - fRead) Then buffer = (fSize - fRead)
            temp = Space(buffer)
            Get a, , temp
            Put b, , temp
            fRead = fRead + buffer
        Wend
        Close b
        Close a
        CopyFileAny = 1
        Exit Function
ErrHan:
        CopyFileAny = 0
    End Function

Public Function BrowseForFolder(selectedPath As String) As String
Dim Browse_for_folder As BROWSEINFOTYPE
Dim itemID As Long
Dim selectedPathPointer As Long
Dim tmpPath As String * 256
With Browse_for_folder
    .hOwner = Me.hWnd ' Window Handle
    .lpszTitle = "Alegeti directorul in care doriti sa copiati KIT-ul Applied Genetics V1.2 !" ' Dialog Title
    .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) ' Dialog callback function that preselectes the folder specified
    selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1) ' Allocate a string
    CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1 ' Copy the path to the string
    .lParam = selectedPathPointer ' The folder to preselect
End With
itemID = SHBrowseForFolder(Browse_for_folder) ' Execute the BrowseForFolder API
If itemID Then
    If SHGetPathFromIDList(itemID, tmpPath) Then ' Get the path for the selected folder in the dialog
        BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1) ' Take only the path without the nulls
    End If
    Call CoTaskMemFree(itemID) ' Free the itemID
End If
Call LocalFree(selectedPathPointer) ' Free the string from the memory
End Function
