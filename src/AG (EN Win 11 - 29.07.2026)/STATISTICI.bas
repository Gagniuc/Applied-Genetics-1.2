Attribute VB_Name = "STATISTICI"
Private Function InStrR(Optional lStart As Long, Optional sTarget As String, Optional sFind As String, Optional iCompare As Integer) As Long
    Dim cFind As Long, i As Long
    cFind = Len(sFind)
    For i = lStart - cFind + 1 To 1 Step -1
        If StrComp(Mid$(sTarget, i, cFind), sFind, iCompare) = 0 Then
            InStrR = i
            Exit Function
        End If
    Next
End Function


Private Function Insert(ByVal outStr As String, ByVal insStr As String, ByVal lStart As Long) As String
Dim lLen As Long
 lLen& = Len(insStr$)
 Insert$ = Mid$(outStr, 1, lStart& - 1) & insStr$ & Mid$(outStr$, lStart& + lLen&)
End Function

Private Function StripHTML(sHTML As String) As String

Dim sTemp As String, lSpot1 As Long, lSpot2 As Long, lSpot3 As Long

sTemp$ = sHTML$
Do
  lSpot1& = InStr(lSpot3& + 1, sTemp$, "<")
  lSpot2& = InStr(lSpot1& + 1, sTemp$, ">")
  
    If lSpot1& = lSpot3& Or lSpot1& < 1 Then Exit Do
    If lSpot2& < lSpot1& Then lSpot2& = lSpot1& + 1
    
  sTemp$ = Left$(sTemp$, lSpot1& - 1) + Right$(sTemp$, Len(sTemp$) - lSpot2&)
  lSpot3& = lSpot1& - 1
Loop

StripHTML$ = sTemp$

End Function




Private Sub TrapUrl(ByRef s As String)
On Error Resume Next
Dim l&, K&, c&

Do
 l& = InStr(c& + 1, LCase$(s$), "http://")
 K& = InStr(l& + 1, LCase$(s$), " ")
 If K& = 0 Then K& = InStr(l& + 1, LCase$(s$), Chr(13))

  If l& = 0 Then Exit Do
  If K& <= l& Then K& = Len(s$) + 4

  s$ = Left$(s$, l& - 1) & Chr(27) & "[urlm" & Mid$(s$, l&, K& - l&) & Chr(27) & "[/urlm" & Mid$(s$, K&)
 c& = InStr(l& + 1, LCase$(s$), " ")

 If c& = 0 Then Exit Do
DoEvents
Loop
End Sub

Private Sub TrapWWW(ByRef s As String)
On Error Resume Next
Dim l&, K&, c&

Do
 l& = InStr(c& + 1, LCase$(s$), "www.")
 K& = InStr(l& + 1, LCase$(s$), " ")
 If K& = 0 Then K& = InStr(l& + 1, LCase$(s$), Chr(13))

  If l& = 0 Then Exit Do
  If K& <= l& Then K& = Len(s$) + 4

  s$ = Left$(s$, l& - 1) & Chr(27) & "[urlm" & Mid$(s$, l&, K& - l&) & Chr(27) & "[/urlm" & Mid$(s$, K&)
 c& = InStr(l& + 1, LCase$(s$), " ")

 If c& = 0 Then Exit Do
DoEvents
Loop
End Sub


Public Sub SaveScript(frm As Form, Optional ByVal SaveAs As Boolean)
On Error GoTo 1
Dim sF As String

 sF$ = frm.Tag

If SaveAs = True Or Left$(frm.Tag, 1) = "\" Then 'save as
 With mdiMain.CD
  .Filter = "VA Script File (*.vas)|*.vas|"
  .CancelError = True
  .FileName = frm.Tag
  .ShowSave
  sF$ = .FileName
 End With
End If

  Open sF$ For Output As #1
    Print #1, frm.rtbEdit.Text
  Close #1
  frm.Tag = sF$
1
End Sub




