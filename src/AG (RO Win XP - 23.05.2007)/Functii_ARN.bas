Attribute VB_Name = "Functii_ARN"
Function ARN_ADN_f(ByVal p_tmp As String) As String
'***
For n = 1 To Len(p_tmp)
d = LCase(Mid(p_tmp, n, 1))

If (d = "t") Then MsgBox "Secventa selectata este ADN !": Exit Function
If (d = "a") Then d = "a": GoTo 2
If (d = "u") Then d = "t": GoTo 2
If (d = "g") Then d = "g": GoTo 2
If (d = "c") Then d = "c": GoTo 2

2:
p = p & UCase(d)
Next n
'***
ARN_ADN_f = p
End Function
