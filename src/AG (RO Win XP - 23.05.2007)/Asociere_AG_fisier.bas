Attribute VB_Name = "Asociere_AG_fisier"
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&
Public Function AssociateMe(Ext As String, Promt As String, Path As String, ExeName As String)

    On Error GoTo AssocError
    Dim EXpath As String

    If Right(App.Path, 1) = "\" Then
     EXpath = Path + ExeName + ".exe %1"
    Else
     EXpath = Path + ExeName + ".exe %1"
    End If

    Call SaveSettingString(HKEY_CLASSES_ROOT, Ext, "", ExeName + ".Exe" + Ext)
    Call SaveSettingString(HKEY_CLASSES_ROOT, Ext & "\ShellNew", "", "")
    Call SaveSettingString(HKEY_CLASSES_ROOT, ".txt\ShellNew", "NullFile", "")
    Call SaveSettingString(HKEY_CLASSES_ROOT, ExeName + ".Exe" + Ext, "", Promt)
    Call SaveSettingString(HKEY_CLASSES_ROOT, ExeName + ".Exe" + Ext + "\Shell\Open\command", "", EXpath)
    Call SaveSettingString(HKEY_CLASSES_ROOT, ExeName + ".Exe" + Ext + "\DefaultIcon", "", App.Path & "\" & App.ExeName & ".exe,0")
    Call SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0)

AssocError:
    If Err.Number <> 0 Then Exit Function

End Function

Public Function AssociateMe2(Ext As String, Promt As String, Path As String, ExeName As String)

    On Error GoTo AssocError
    Dim EXpath As String

    If Right(App.Path, 1) = "\" Then
     EXpath = Path + ExeName + ".exe %1"
    Else
     EXpath = Path + ExeName + ".exe %1"
    End If

    Call SaveSettingString(HKEY_CLASSES_ROOT, Ext, "", ExeName + ".Exe" + Ext)
    Call SaveSettingString(HKEY_CLASSES_ROOT, Ext & "\ShellNew", "", "")
    Call SaveSettingString(HKEY_CLASSES_ROOT, ".txt\ShellNew", "NullFile", "")
    Call SaveSettingString(HKEY_CLASSES_ROOT, ExeName + ".Exe" + Ext, "", Promt)
    Call SaveSettingString(HKEY_CLASSES_ROOT, ExeName + ".Exe" + Ext + "\Shell\Open\command", "", EXpath)
    Call SaveSettingString(HKEY_CLASSES_ROOT, ExeName + ".Exe" + Ext + "\DefaultIcon", "", App.Path & "\MS.exe,0")
    Call SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0)

AssocError:
    If Err.Number <> 0 Then Exit Function

End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)

    Dim hCurKey As Long
    Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))

    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
    
End Sub
Public Function FExist(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FExist = True
    Exit Function
MakeF:
        'error, file does Not exist
        FExist = False
    Exit Function
End Function


