VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBruteforce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer

Dim m_strCharSet() As String
Dim m_strCharSet2 As String
Dim m_strStartWord As String
Dim m_strPassword As String

Dim m_CharSetLen As Integer
Dim m_intStartLen As Integer
Dim m_intEndLen As Integer

Public sTotalCombo As String
Dim intCharLen As Integer
Dim strCharSet() As String
Dim strGuess As String

Dim lComboCount As Long
Dim lComboTime As Long

Public Event Statistics(lngTotalCombos As Long, lngCombinationsPerSec As Long, strCurrentPassword As String)
Public Event Complete()

Private Declare Function GetInputState Lib "user32" () As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long

Property Get CurrentPassword() As String
CurrentPassword = m_strPassword
End Property

Property Get CharacterSet() As String
CharacterSet = m_strCharSet2
End Property

Property Let CharacterSet(ByVal Val As String)

m_CharSetLen = Len(Val) - 1

ReDim m_strCharSet(m_CharSetLen) As String

For x = 0 To m_CharSetLen
    m_strCharSet(x) = Mid(Val, x + 1, 1)
Next
m_strCharSet2 = Val
End Property

Property Get StartWord() As String
StartWord = m_strStartWord
End Property

Property Let StartWord(ByVal Val As String)
m_strStartWord = Val
m_strPassword = m_strStartWord
End Property

Property Get EndLength() As String
EndLength = m_intEndLen
End Property

Property Let EndLength(ByVal Val As String)
m_intEndLen = Val
End Property

Property Get StartLength() As Integer
StartLength = m_intStartLen
End Property

Property Let StartLength(ByVal Val As Integer)
m_intStartLen = Val
m_strPassword = String(m_intStartLen, m_strCharSet(0))
End Property

Private Sub Class_Initialize()
CharacterSet = "acgt"
StartLength = 10
EndLength = 12
sTotalCombo = 0
End Sub

Private Sub Combinations()
    If GetInputState() = 0 Then DoEvents
    If GetTickCount - lComboTime >= 1000 Then
        lComboTime = GetTickCount
        sTotalCombo = sTotalCombo + lComboCount
        RaiseEvent Statistics(CLng(sTotalCombo), CLng(lComboCount), m_strPassword)
        lComboCount = 0
    End If
End Sub

Function BruteForce() As String
Dim intPos As Integer
Static intCount As Integer

While m_intStartLen <= m_intEndLen
    For x = intCount To m_CharSetLen
        Mid(m_strPassword, 1, 1) = m_strCharSet(x)
        BruteForce = m_strPassword
        intCount = intCount + 1
        lComboCount = lComboCount + 1
        Exit Function
    Next x
    
    intCount = 0
    
    If Mid$(m_strPassword, m_intStartLen, 1) = m_strCharSet(m_CharSetLen) Then
        Mid$(m_strPassword, m_intStartLen, 1) = m_strCharSet(0)
        m_intStartLen = m_intStartLen + 1
        m_strPassword = m_strPassword + m_strCharSet(0)
    Else
        For x = 1 To m_intStartLen - 1
            If Mid$(m_strPassword, x, 1) = m_strCharSet(m_CharSetLen) Then
                Mid$(m_strPassword, x, 1) = m_strCharSet(0)
                intPos = InStr(m_strCharSet2, Mid$(m_strPassword, x + 1, 1)) - 1
                If intPos < m_CharSetLen Then
                    Mid$(m_strPassword, x + 1, 1) = m_strCharSet(intPos + 1)
                    Combinations
                    Exit For
                End If
            End If
        Next x
    End If
Wend

RaiseEvent Complete
End Function

Sub ForceCombinationCalc()
    lComboTime = 99999
    Combinations
End Sub

Sub ResetStats()
    sTotalCombo = 0
    lComboCount = 0
End Sub
