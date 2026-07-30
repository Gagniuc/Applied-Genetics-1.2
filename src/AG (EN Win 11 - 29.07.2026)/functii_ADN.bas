Attribute VB_Name = "functii_ADN"
Function ADN_ARN_f(ByVal p_tmp As String) As String
'***
For n = 1 To Len(p_tmp)
d = LCase(Mid(p_tmp, n, 1))

If (d = "u") Then MsgBox "The selected sequence is RNA !": Exit Function
If (d = "a") Then d = "a": GoTo 2
If (d = "t") Then d = "u": GoTo 2
If (d = "g") Then d = "g": GoTo 2
If (d = "c") Then d = "c": GoTo 2

2:
p = p & UCase(d)
Next n
'***
ADN_ARN_f = p
End Function
