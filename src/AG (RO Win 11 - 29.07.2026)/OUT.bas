Attribute VB_Name = "ANALIZA"
'##############################################################################################
'###                                                                                        ###
'###                                SECVENTE_INVERSE [->]X[<-]                              ###
'###                                                                                        ###
'##############################################################################################
Function SECVENTE_INVERSE(ByVal obiectul As RichTextBox, ByVal enzima_normal As String, ByVal codul_ADN As String) As String
enzima_normal = UCase(Replace(enzima_normal, "*", ""))
COD_ramas = UCase(codul_ADN)

warza = 0
carnat = 0
old_carnat = 0
ee = Lista_cu_geluri.secvente_inverse_nr.Count
Lista_cu_geluri.secvente_inverse_nr(ee - 1).ToolTipText = ""
'If ee > 200 Then
'Exit Sub
'End If
'-------------------------------------------------------------------
anuntare = ""
X = 0
D_enzima_normal = enzima_normal
S_enzima_normal = StrReverse(enzima_normal)
0:
If X > 1 Then X = 0

If X = 0 Then enzima_normal = D_enzima_normal
If X = 1 Then enzima_normal = S_enzima_normal

'------------
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then
    Debug.Print "[" & Timer & "] SECVENTE_INVERSE: Stop detectat in interiorul buclei de scanare, ies (warza=" & warza & ")"
    GoTo 22
End If
'------------
poz = InStr(COD_ramas, enzima_normal) - 1
If (poz < 0) Then GoTo 22

linie = Mid$(COD_ramas, 1, poz)
carnat = carnat + poz + Len(enzima_normal)
COD_ramas = Mid$(COD_ramas, poz + Len(enzima_normal) + 1, Len(COD_ramas))
'-----------------------------------------------------------------
'If InStr(COD_ramas, D_enzima_normal) <> 0 And InStr(COD_ramas, S_enzima_normal) <> 0 Then
'-------------------------------------------------------------------
If X = 0 And InStr(COD_ramas, D_enzima_normal) <> 0 And InStr(COD_ramas, S_enzima_normal) <> 0 Then
obiectul.SelStart = carnat - Len(enzima_normal)
obiectul.SelLength = Len(enzima_normal)
obiectul.SelColor = Mai_mult.SI1_culoare.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.SI1_culoare.BackColor, False)
Call Reprezentare_grafic_liniar(carnat - Len(enzima_normal), carnat, 64, 84, obiectul, Mai_mult.SI1_culoare.BackColor)
Call Reprezentare_grafic_x(carnat - Len(enzima_normal), carnat, 0, 20, obiectul, Mai_mult.SI1_culoare.BackColor, ee)

anuntare1 = ""
anuntare1 = anuntare1 & "[" & Mid(principal.COD_CODON.Text, carnat - Len(enzima_normal) + 1, Len(enzima_normal)) & "]"
End If
'-------------------------------------------------------------------

'-------------------------------------------------------------------
11:
If X = 1 Then
obiectul.SelStart = carnat - Len(enzima_normal)
obiectul.SelLength = Len(enzima_normal)
obiectul.SelColor = Mai_mult.SI2_culoare.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.SI2_culoare.BackColor, False)
Call Reprezentare_grafic_liniar(carnat - Len(enzima_normal), carnat, 64, 84, obiectul, Mai_mult.SI2_culoare.ForeColor)
Call Reprezentare_grafic_x(carnat - Len(enzima_normal), carnat, 0, 20, obiectul, Mai_mult.SI2_culoare.ForeColor, ee)


obiectul.SelStart = old_carnat
obiectul.SelLength = carnat - old_carnat - Len(enzima_normal)
obiectul.SelColor = Mai_mult.SI_mid.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.SI_mid.BackColor, False)
Call Reprezentare_grafic_liniar(old_carnat, carnat - Len(enzima_normal), 64, 84, obiectul, Mai_mult.SI_mid.BackColor)
Call Reprezentare_grafic_x(old_carnat, carnat - Len(enzima_normal), 0, 20, obiectul, Mai_mult.SI_mid.BackColor, ee)

'l_zero_palindrom = carnat - Old_carnat - Len(enzima_normal)
'If l_zero_palindrom <= 0 Then s_e_polindrom = "[Secventa de mai jos este un palindrom !"
cap_informatie = vbCrLf & "[Nr." & warza & "]-[Pozitie:" & old_carnat - Len(enzima_normal) + 1 & "b - Lungime:" & Len(enzima_normal) & "b][Pozitie:" & old_carnat + 1 & "b - Lungime:" & carnat - old_carnat - Len(enzima_normal) & "][Pozitie:" & carnat - Len(enzima_normal) + 1 & "b - Lungime:" & Len(enzima_normal) & "b]" & vbCrLf
anuntare = anuntare & cap_informatie & anuntare1 & "[" & Mid(principal.COD_CODON.Text, old_carnat + 1, carnat - old_carnat - Len(enzima_normal)) & "]" & "[" & Mid(principal.COD_CODON.Text, carnat - Len(enzima_normal) + 1, Len(enzima_normal)) & "]" & vbCrLf
warza = warza + 1

If principal.SI_grafic_step_by_step.Value = 1 Then
Lista_cu_geluri.Show
Else
Lista_cu_geluri.Hide
End If

Lista_cu_geluri.secvente_inverse_nr(ee - 1).ToolTipText = Lista_cu_geluri.secvente_inverse_nr(ee - 1).ToolTipText & old_carnat - Len(enzima_normal) + 1 & "@" & enzima_normal & "@" & carnat - old_carnat - Len(enzima_normal) & "#"

'End If
'-------------------
End If
'-------------------------------------------------------------------
old_carnat = carnat
'-------------------------------------------------------------------
If COD_ramas = "" Then
'If X = 0 Then
'X = 1
'GoTo 11
'End If

GoTo 22
End If

X = X + 1
GoTo 0
22:
ee = Lista_cu_geluri.secvente_inverse_nr.Count
Load Lista_cu_geluri.secvente_inverse_nr(ee)
Lista_cu_geluri.secvente_inverse_nr(ee).Width = principal.segment_total_gena.ScaleWidth
Lista_cu_geluri.secvente_inverse_nr(ee).Visible = True
Lista_cu_geluri.secvente_inverse_nr(ee).Top = Lista_cu_geluri.secvente_inverse_nr(ee - 1).Top + Lista_cu_geluri.secvente_inverse_nr(ee - 1).ScaleHeight
'-------------------
'If Lista_cu_geluri.fer_seg2.Height > Lista_cu_geluri.fer_seg1.ScaleHeight Then
Lista_cu_geluri.fer_seg2.Height = Lista_cu_geluri.secvente_inverse_nr(ee - 1).Top + Lista_cu_geluri.secvente_inverse_nr(ee - 1).ScaleHeight
Lista_cu_geluri.SI_derulare.Max = Lista_cu_geluri.fer_seg2.Height - Lista_cu_geluri.fer_seg1.Height 'Lista_cu_geluri.SI_derulare.Max / ee
Lista_cu_geluri.SI_derulare.Min = 0



If (warza > 0) Then
SECVENTE_INVERSE = anuntare & vbCrLf & "Nr. Secvente inverse:[" & Val(warza) & "]" & vbCrLf & "--------------------------------------------------------------------------------------"
anuntare = Empty
Else
SECVENTE_INVERSE = "Nu exista secvente inverse !" & vbCrLf
End If
End Function
'##############################################################################################
'##############################################################################################


'##############################################################################################
'###                                                                                        ###
'###                                REPETITII [->][->][->]                                  ###
'###                                                                                        ###
'##############################################################################################
Function REPETITII(ByVal obiectul As RichTextBox, ByVal enzima_normal As String, ByVal codul_ADN As String) As String
enzima_normal = UCase(Replace(enzima_normal, "*", ""))
COD_ramas = UCase(codul_ADN)
For rt = 1 To 1000
DoEvents
poz_fragment_dreapta(rt) = Empty
Next rt
warza = 0
carnat = 0
repet_x = 1
situsuri_restrictie_del_d = 0 'asfel se fute algoritmul de la gel ... totul va fi /2
'-------------------------------------------------------------------
0:
'------------
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then
    Debug.Print "[" & Timer & "] REPETITII: Stop detectat in interiorul buclei de scanare, ies (warza=" & warza & ")"
    GoTo 22
End If
'------------
poz = InStr(COD_ramas, enzima_normal) - 1
If (poz < 0) Then GoTo 22

linie = Mid$(COD_ramas, 1, poz)
oldcarnat = carnat
carnat = carnat + poz + Len(enzima_normal)
OLD_COD_ramas = Len(COD_ramas)
COD_ramas = Mid$(COD_ramas, poz + Len(enzima_normal) + 1, Len(COD_ramas))
'-----------------------------------------------------------------
'MsgBox OLD_COD_ramas - Len(COD_ramas) & "=" & Len(enzima_normal) & "-" & enzima_normal

If OLD_COD_ramas - Len(COD_ramas) = Len(enzima_normal) Then
If repet_x = 1 Then start_rep = oldcarnat - Len(enzima_normal) '- Len(enzima_normal)
repet_x = repet_x + 1

If repet_x >= principal.rep_min.Value And repet_x <= principal.rep_max.Value Then
If start_rep < 0 Then start_rep = 0
obiectul.SelStart = start_rep
obiectul.SelLength = repet_x * Len(enzima_normal)
obiectul.SelColor = Mai_mult.start5normal.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.start5normal.BackColor, False)
Call Reprezentare_grafic_liniar(start_rep, start_rep + repet_x * Len(enzima_normal), 86, 104, obiectul, Mai_mult.start5normal.BackColor)


COD_repetitii = " pozitia:[" & start_rep & "]-" & "[" & enzima_normal & " * " & repet_x & "] " & Mid(codul_ADN, start_rep + 1, repet_x * Len(enzima_normal))

End If

Else
If repet_x >= principal.rep_min.Value And repet_x <= principal.rep_max.Value Then
obiectul.SelStart = start_rep
obiectul.SelLength = repet_x * Len(enzima_normal)
obiectul.SelColor = Mai_mult.start5normal.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.start5normal.BackColor, False)
End If

repet_x = 1
End If


'-----------------------------------------------------------------
'For u = 1 To principal.rep_min.Value
'qq = Mid(COD_ramas, u * Len(enzima_normal), Len(enzima_normal))

'If qq = enzima_normal Then repet_x = repet_x + 1 Else GoTo 5


'Next u

'If repet_x = principal.rep_min.Value Then
'For u = 1 To principal.rep_max.Value

'Next u
'End If

'5:
'-----------------------------------------------------------------
'situsuri_restrictie_del_d = situsuri_restrictie_del_d + 1
'obiectul.SelStart = carnat - Len(enzima_normal)
'obiectul.SelLength = Len(enzima_normal)
'obiectul.SelColor = Mai_mult.start5normal.ForeColor
'Call HighLightSelection(principal, obiectul, Mai_mult.start5normal.BackColor, False)
'-------------------------------------------------------------------
warza = warza + 1
'poz_fragment_dreapta(warza) = carnat
'-------------------------------------------------------------------
If COD_ramas = "" Then GoTo 22

GoTo 0
22:
If (warza > 0) Then
REPETITII = "Repetitii in " & COD_repetitii & vbCrLf
Else
REPETITII = "Nu s-au gasit repetitii !" & vbCrLf
End If
End Function
'##############################################################################################
'##############################################################################################


Function ANALIZA_CODULUI(ByVal Z As Object, ByVal q As RichTextBox, ByVal xx As String)
'00000000000000000000000000000000000000000000
q.SelStart = 0
'00000000000000000000000000000000000000000000
Cul1 = RGB(200, 0, 0)
Cul2 = RGB(0, 0, 200)
codon_coloreaza = InStr(q.Text, Z.coloreaza_acesta.Text)
If (codon_coloreaza <> 0) Then
Else
Exit Function
End If

q.SelStart = codon_coloreaza + Len(Z.coloreaza_acesta.Text)
'00000000000000000000000000000000000000000000
For R = codon_coloreaza + Len(Z.coloreaza_acesta.Text) - 1 To Len(q.Text) Step 3

If (A = Cul1) Then q.SelColor = Cul2: A = Cul2 Else q.SelColor = Cul1: A = Cul1


q.SelStart = R
q.SelLength = 3
w = w + 1
Next R
'00000000000000000000000000000000000000000000
SC_out_put ("Au fost gasiti [" & w & "] codoni" & vbCrLf)

End Function


Function codon_aminoacid(ByVal codon As String, ByVal scurt As Boolean) As String
Dim A As Variant
For i = 0 To 20


    'codon = UCase(codon)
    A = Parse(Mai_mult.codoni_aminoacidului(i).Caption, ",")

    For w = 1 To UBound(A)
    A(w) = Replace(A(w), " ", "")
    If A(w) = codon Then
    If scurt = True Then
    codon_aminoacid = Split(Mai_mult.amino_buton(i).Caption, "\")(0)
    Else
    codon_aminoacid = Split(Mai_mult.amino_buton(i).Caption, "\")(1)
    End If
    End If
    Next w



Next i
End Function

Function Gena_codon_aminoacid(ByVal codoni_n As String, ByVal scurt_x As Boolean) As String
aa = 0
max_bara_aici = Len(codoni_n)
For i = principal.FrameShift_Val.Value To Len(codoni_n) Step 3

'############
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then Exit Function
'############

DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then Exit Function
principal.bara_date.Value = (principal.bara_date.Max / max_bara_aici) * i


old_M = rez_aa
rez_aa = codon_aminoacid(Mid(codoni_n, i, 3), scurt_x)

If rez_aa <> "" Then aa = aa + 1 Else rez_aa = "-"


If rez_aa = "M" Or rez_aa = "Met" And rez_aa <> old_M Then '-------------------------]
If principal.Delimitator_AUG.Value = 1 Then linie_jos = vbCrLf & "-------------------------------------------------------------" & vbCrLf Else linie_jos = ""

principal.COD_CODON.SelStart = i - 1
principal.COD_CODON.SelLength = 3
principal.COD_CODON.SelColor = Mai_mult.Start_CUL.ForeColor
Call HighLightSelection(principal, principal.COD_CODON, Mai_mult.Start_CUL.BackColor, False)
Call Reprezentare_grafic_liniar(i - 1, i - 1 + 3, 106, 126, principal.COD_CODON, Mai_mult.Start_CUL.BackColor)

Else '-------------------------------------------------------------]
linie_jos = ""

ss = ss + 1
If ss >= 2 Then ss = 0

If ss = 0 Then
principal.COD_CODON.SelStart = i - 1
principal.COD_CODON.SelLength = 3
principal.COD_CODON.SelColor = Mai_mult.Exon_cul.ForeColor
Else
principal.COD_CODON.SelStart = i - 1
principal.COD_CODON.SelLength = 3
principal.COD_CODON.SelColor = Mai_mult.Intron_cul.ForeColor
End If


End If '-----------------------------------------------------------]

If rez_aa = "Z" Or rez_aa = "[STOP]" And rez_aa <> old_M Then '-------------------------]
principal.COD_CODON.SelStart = i - 1
principal.COD_CODON.SelLength = 3
principal.COD_CODON.SelColor = Mai_mult.Stop_cul.ForeColor
Call HighLightSelection(principal, principal.COD_CODON, Mai_mult.Stop_cul.BackColor, False)
Call Reprezentare_grafic_liniar(i - 1, i - 1 + 3, 106, 126, principal.COD_CODON, Mai_mult.Stop_cul.BackColor)
End If


If rez_aa = "M" Then rez_aa = Replace(principal.start_not.Text, "POZITIE", i - 1)
If rez_aa = "Z" Then rez_aa = Replace(principal.stop_not.Text, "POZITIE", i - 1)
proteina_x = proteina_x & linie_jos & rez_aa
Next i

principal.bara_date.Value = 0
Gena_codon_aminoacid = vbCrLf & "FRAME SHIFT [" & principal.FrameShift_Val.Value & "]-------------------------------------------------------------" & vbCrLf & proteina_x & vbCrLf & "----------------------------------------------------------------------------"
End Function

Function X_codon_aminoacid(ByVal codoni_n As String, ByVal scurt_x As Boolean) As String
aa = 0
For i = 1 To Len(codoni_n) Step 3
rez_aa = codon_aminoacid(Mid(codoni_n, i, 3), scurt_x)

If rez_aa <> "" Then aa = aa + 1 Else rez_aa = "-"
proteina_x = proteina_x & rez_aa
Next i
X_codon_aminoacid = aa & "@" & proteina_x
End Function


Public Function Parse(sIn As String, sDel As String) As Variant
    Dim i As Integer, X As Integer, s As Integer, t As Integer
    i = 1: s = 1: t = 1: X = 1
    ReDim tArr(1 To X) As Variant


    If InStr(1, sIn, sDel) <> 0 Then


        Do
            ReDim Preserve tArr(1 To X) As Variant
            tArr(i) = Mid(sIn, t, InStr(s, sIn, sDel) - t)
            t = InStr(s, sIn, sDel) + Len(sDel)
            s = t
            If tArr(i) <> "" Then i = i + 1
            X = X + 1
        Loop Until InStr(s, sIn, sDel) = 0
        ReDim Preserve tArr(1 To X) As Variant
        tArr(i) = Mid(sIn, t, Len(sIn) - t + 1)
    Else
        tArr(1) = sIn
    End If
    Parse = tArr
End Function
