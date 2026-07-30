Attribute VB_Name = "Universal"
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020

Public Trial_pa As Boolean
Public amino_notatie_scurta As Boolean
Public adresa_site_general As String
Public deasupara_reprezentarii_grafice As Boolean

Public culoare1_gel_old  As Long
Public culoare2_gel_old  As Long

Function Wav_Play(ByVal fi As String)
On Error Resume Next
    soundfile$ = fi
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    HaHa = sndPlaySound(soundfile$, wFlags%)
End Function
'-------STOP WAVE SOUND-------


'Private Sub StopTheSound_Click()
'    StopTheSoundNOW = sndPlaySound(soundfile$, wFlags%)
'End Sub


Function SC_out_put(ByVal textul As Variant)
principal.out_put.Text = principal.out_put.Text & textul
'principal.out_put.SelStart = 1
'principal.out_put.SelLength = Len(principal.out_put.Text)
'principal.out_put.SelColor = vbRed
principal.out_put.SelStart = Len(principal.out_put.Text)
End Function

Function Reprezentare_grafic_x(ByVal old_X As Variant, ByVal new_X As Variant, ByVal y1 As Variant, ByVal y2 As Variant, ByVal textul As Object, ByVal cu As Long, ByVal ee As Variant)
x1 = (Lista_cu_geluri.secvente_inverse_nr(ee - 1).ScaleWidth / Len(textul.Text)) * (old_X)
x2 = (Lista_cu_geluri.secvente_inverse_nr(ee - 1).ScaleWidth / Len(textul.Text)) * (new_X)
Lista_cu_geluri.secvente_inverse_nr(ee - 1).Line (x1, y1)-(x2, y2), cu, BF
End Function


Function Reprezentare_grafic_liniar(ByVal old_X As Variant, ByVal new_X As Variant, ByVal y1 As Variant, ByVal y2 As Variant, ByVal textul As Object, ByVal cu As Long)
x1 = (principal.segment_total_gena.ScaleWidth / Len(textul.Text)) * (old_X)
x2 = (principal.segment_total_gena.ScaleWidth / Len(textul.Text)) * (new_X)
principal.segment_total_gena.Line (x1, y1)-(x2, y2), cu, BF
End Function


Function punct_lovit_enzima(ByVal old_X As Variant, ByVal new_X As Variant, ByVal textul As Object, ByVal cu As Long)
x1 = (principal.segment_total_gena.ScaleWidth / Len(textul.Text)) * (old_X)
x2 = (principal.segment_total_gena.ScaleWidth / Len(textul.Text)) * (new_X)
principal.segment_total_gena.Line (x1, 42)-(x2, 62), cu, BF
End Function


Function GENEREAZA_NUCLEOTIDE(ByVal nr As Variant, ByVal tip As String) As String
'***
Dim nucleo(1 To 5) As String
nucleo(1) = "A"
nucleo(2) = "T"
nucleo(3) = "G"
nucleo(4) = "C"
nucleo(5) = "U"

For n = 1 To nr

If (tip = "ADN") Then
c = Int(3 * Rnd(3))
p = p & nucleo(c + 1)
End If

If (tip = "ARN") Then
c = Int(4 * Rnd(4))
If (c + 1 = 2) Then c = 4
p = p & nucleo(c + 1)
End If

Next n
'***
GENEREAZA_NUCLEOTIDE = p
End Function

Function REVERSUL(ByVal p_tmp As String) As String
'00000000000000000000000000000000000000000000
p = StrReverse(p_tmp)
REVERSUL = p
'00000000000000000000000000000000000000000000
End Function

Function CATENA_PLUS(ByVal p_tmp As String) As String
'***
For n = 1 To Len(p_tmp)
d = LCase(Mid(p_tmp, n, 1))

If (d = "u") Then MsgBox "The selected sequence is RNA !": Exit Function
If (d = "a") Then d = "t": GoTo 2
If (d = "t") Then d = "a": GoTo 2
If (d = "g") Then d = "c": GoTo 2
If (d = "c") Then d = "g": GoTo 2

2:
p = p & UCase(d)
Next n
'***
CATENA_PLUS = p
End Function


Function Segment_TOTAL(ByVal textul As String)
'----------------------------------------------------|
imparte = principal.segment_total_gena.Width / 100
t = situsuri_restrictie_del_d + situsuri_restrictie_del_s + 1
'---------------------------------------------------------------------------
For i = 1 To t
DoEvents
If Ordine_ambii_algoritmi(i) = Empty Then Exit Function

new_X = Ordine_ambii_algoritmi(i)
x1 = (principal.segment_total_gena.ScaleWidth / Len(textul)) * (old_X)
x2 = (principal.segment_total_gena.ScaleWidth / Len(textul)) * (new_X)
old_X = new_X

A = A + 1
If A > 2 Then A = 1
If A = 1 Then cu = Mai_mult.start5normal.BackColor
If A = 2 Then cu = Mai_mult.start3normal.BackColor
principal.segment_total_gena.Line (x1, 2)-(x2, 20), cu, BF
Next i
'---------------------------------------------------------------------------
End Function



Function Segment_geneza(ByVal y1 As Variant, ByVal y2 As Variant, ByVal cul As Long)
x1 = 0
x2 = principal.segment_total_gena.ScaleWidth
principal.segment_total_gena.Line (x1, y1)-(x2, y2), cul, BF
End Function



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


Function Proportie_A_T_G_C_U(ByVal textul As Variant, ByVal totul As Boolean) As String
On Error Resume Next
For n = 1 To Len(textul)
DoEvents

d = LCase(Mid(textul, n, 1))
If (d = "a") Then A = A + 1
If (d = "t") Then t = t + 1
If (d = "g") Then g = g + 1
If (d = "c") Then c = c + 1
If (d = "u") Then u = u + 1
Next n

tot_txt_cod = 0
tot_txt_cod = Len(textul)

If totul = True Then
DATELE_GCATU = vbCrLf & _
"A[" & Int(A) & "b] ~ [" & Int((100 / tot_txt_cod) * A) & "%]" & vbCrLf & _
"T[" & Int(t) & "b] ~ [" & Int((100 / tot_txt_cod) * t) & "%]" & vbCrLf & _
"C[" & Int(c) & "b] ~ [" & Int((100 / tot_txt_cod) * c) & "%]" & vbCrLf & _
"G[" & Int(g) & "b] ~ [" & Int((100 / tot_txt_cod) * g) & "%]" & vbCrLf & _
"U[" & Int(u) & "b] ~ [" & Int((100 / tot_txt_cod) * u) & "%]"
Else
DATELE_GCATU = " - GC[" & Int(g + c) & "b] - [" & Int((100 / tot_txt_cod) * Int(g + c)) & "%]"
End If

Proportie_A_T_G_C_U = DATELE_GCATU
End Function


Function RTFToHTM(RTFStr As String, Title As String, Optional UseLines As Boolean) As String
    'This routine is quite fast - but only because we know and control the content
    'of the rich text. It would need modifications to be used as a general
    'RTF to HTML routine
    Dim Z As Long
    Dim temp As String
    Z = InStr(RTFStr, "\fs20")
    temp = Right(RTFStr, Len(RTFStr) - Z - 5) 'RTF minus header
    temp = Left(temp, Len(temp) - 3) 'get rid of the tail
    'Determine the first color used
    Z = InStr(RTFStr, "\cf")
    If Z > -1 Then
        If IsNumeric(Mid(RTFStr, Z + 3, 1)) Then
            Z = CInt(Mid(RTFStr, Z + 3, 1))
        Else
            Z = 0
        End If
    End If
    Select Case Z
        Case 1 'blue
            temp = "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>" & Title & "</title>" & vbCrLf & "</head>" & "<body><font color=#000080 font face=" & Chr(34) & "Courier New" & Chr(34) & " size=2>" & vbCrLf & temp
        Case Else 'black or green
            'if it was green, we still use black as we are only interested in the forecolor here
            temp = "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>" & Title & "</title>" & vbCrLf & "</head>" & "<body><font color=#000000 font face=" & Chr(34) & "Courier New" & Chr(34) & " size=2>" & vbCrLf & temp
    End Select
    temp = Replace(temp, "\pard", "") 'tail end of RTF
    temp = Replace(temp, "\cf0 ", "</font><font color=#000000 font face=" & Chr(34) & "Courier New" & Chr(34) & " size=2>") 'blue - stays as blue in HTML
    temp = Replace(temp, "\cf1 ", "</font><font color=#000080 font face=" & Chr(34) & "Courier New" & Chr(34) & " size=2>") 'black - stays as black in HTML
    temp = Replace(temp, "\cf2 ", "</font><font color=#000000 font face=" & Chr(34) & "Courier New" & Chr(34) & " size=2>") 'green uses black forecolor
    temp = Replace(temp, RemRTFLine, IIf(UseLines, "<hr>", "")) 'lines
    temp = Replace(temp, "\par ", "<br>") 'change carriage returns to breaks
    temp = Replace(temp, "\tab ", "&nbsp;&nbsp;&nbsp;&nbsp;") 'change tabs to 4 spaces
    temp = Replace(temp, "\tab", "&nbsp;&nbsp;&nbsp;&nbsp;") 'change tabs to 4 spaces - multiple "\tab"'s will not have a space after them
    'this is where we fix up the green bits - which are only green in backcolor
    temp = Replace(temp, "\highlight2", "<span style=" & Chr(34) & "background-color: #00FF00" & Chr(34) & ">")
    temp = Replace(temp, "\highlight0 ", "</span>")
    temp = temp & vbCrLf & "</font></body></html>" 'finish up
    RTFToHTM = temp
End Function


Public Function GetTextName(TextValue As String, Optional ObjectTarget As Integer = 1, Optional OutputString As String, Optional OutputTotalValue As Integer) As String
Dim NowObject As Integer 'The Text Value must be in this format ",[txt],[txt],[txt],"
Dim TextResult As String
Dim TotalDelimiter As Integer
Dim i As Integer
    
    'Get total Value
    '------------------------------------
    For i = 1 To Len(TextValue)
        If Mid(TextValue, i, 1) = "," Then TotalDelimiter = TotalDelimiter + 1
    Next i
    OutputTotalValue = TotalDelimiter - 1
    '------------------------------------
    
    For i = 1 To Len(TextValue) 'To Get a text, This Function work similar with select case
    If Mid(TextValue, i, 1) = "," Then NowObject = NowObject + 1
        If NowObject = ObjectTarget Then
           If Mid(TextValue, i + 1, 1) <> "," Then TextResult = TextResult + Mid(TextValue, i + 1, 1)
        ElseIf ObjectTarget < NowObject Then
           OutputString = TextResult
           GetTextName = TextResult
           Exit For
        End If
    Next i
End Function

Public Function DectoWebCol(lngColour As Long) As String
    '***************************************
    '     *********
    '* This function takes a decimal colour,
    '
    '* for example one returned by the CDB
    '* and converts it into a hex colour
    '* suitable for use in a web page.
    '* Copyright by Mark Bennett 2002.
    '* You may use this code for any purpose
    '     .
    '***************************************
    '     *********
    Dim strColour As String
    'Convert decimal colour to hex
    strColour = Hex(lngColour)
    'Add leading zero's


    Do While Len(strColour) < 6
        strColour = "0" & strColour
    Loop
    'Reverse the bgr string pairs to rgb
    DectoWebCol = "#" & Right$(strColour, 2) & _
    Mid$(strColour, 3, 2) & _
    Left$(strColour, 2)
End Function



Public Function JPEG2BMP(FileName As String, LoadPB As PictureBox, SavePB As PictureBox) As Boolean
On Error GoTo FileMuffUp

LoadPB = LoadPicture(FileName & ".jpg")
SavePB = LoadPB
SavePicture SavePB.Picture, FileName & ".bmp"
JPEG2BMP = True
Exit Function

FileMuffUp:
JPEG2BMP = False
End Function
