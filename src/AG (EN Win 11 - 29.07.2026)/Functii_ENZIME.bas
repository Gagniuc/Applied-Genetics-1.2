Attribute VB_Name = "Functii_ENZIME"
Public poz_fragment_dreapta(1 To 1000) As Variant
Public poz_fragment_stanga(1 To 1000) As Variant
Public situsuri_restrictie_del_d As Variant
Public situsuri_restrictie_del_s As Variant
Public Ordine_ambii_algoritmi(1 To 2000) As Variant

Public del_fragment_dreapta(1 To 1000) As Variant
Public del_fragment_stanga(1 To 1000) As Variant

Public ins_fragment_dreapta(1 To 1000) As Variant
Public ins_fragment_stanga(1 To 1000) As Variant

Public secvente_rerstrictate(1 To 1000) As String

Public Gel_vizivil As Boolean
'##############################################################################################
'###                                                                                        ###
'###                            COMANDA GENERALA ALGORITMI ENZIME                           ###
'###                                                                                        ###
'##############################################################################################
Function ALEGE_ALGORITM_ENZIME(ByVal F As Form)
'############
If Mai_mult.enzime_de_test.ListCount = 0 Then MsgBox "No restriction enzyme was chosen from the list ! " & vbCrLf & " Apasati butonul " & Chr(34) & "ENZIME" & Chr(34)

If principal.aduna_LA_OUTPUT.Value = False Then
principal.out_put.Text = Empty
End If
'------------
If (F.normal_stanga.Value = True And F.normal_dreapta.Value = True) Then
principal.Gel_general.Cls
For i = 1 To Lista_cu_geluri.Gel_nr_X.Count - 1
        Unload Lista_cu_geluri.Gel_nr_X(i)
Next
End If
'############

For Y = 0 To Mai_mult.enzime_de_test.ListCount
'############
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then Exit Function
'############

DoEvents
'--------
imparte_en = Mai_mult.enzime_de_test.ListCount
If imparte_en = 0 Then imparte_en = 1
F.bara_date2.Value = (F.bara_date2.Max / imparte_en) * Y
'--------
'############
If Mai_mult.enzime_de_test.List(Y) = "" Then GoTo 1
If (verifica_corect(Mai_mult.enzime_de_test.List(Y), F.COD_CODON.Text) = False) Then Exit Function
F.enzima.Text = Mai_mult.enzime_de_test.List(Y)
tot_codul_ADN = UCase(F.COD_CODON.Text)

'=======================================
'===    DACA VREA DE LA X LA Y       ===
'=======================================
If (F.distanta_n.Value = 1) Then
'------------------
For i = Val(F.de_la_distanta.Caption) To Val(F.dista.Caption)
DoEvents
'--------
F.bara_date.Value = (F.bara_date.Max / F.dista.Caption) * i
'--------
If (F.del_dreapta.Value = True) Then
SC_out_put (enz_del_stanga(F.COD_CODON, Mai_mult.enzime_de_test.List(Y), tot_codul_ADN, i))   ' ->
End If

If (F.del_stanga.Value = True) Then
SC_out_put (enz_del_dreapta(F.COD_CODON, Mai_mult.enzime_de_test.List(Y), tot_codul_ADN, i))  ' <-
End If

If (F.ins_dreapta.Value = True) Then
SC_out_put (enz_ins_dreapta(F.COD_CODON, Mai_mult.enzime_de_test.List(Y), tot_codul_ADN, i)) ' <-
End If

If (F.ins_stanga.Value = True) Then
SC_out_put (enz_ins_stanga(F.COD_CODON, Mai_mult.enzime_de_test.List(Y), tot_codul_ADN, i)) ' <-
End If
Next i
'------------------
End If
'=======================================
'=======================================


'=======================================
'===   DACA VREA DOAR X NUCLEOTIDE   ===
'=======================================
If (F.distanta_x.Value = 1) Then
'-----------------------
i = F.dista.Caption

If (F.del_dreapta.Value = True) Then
SC_out_put (enz_del_stanga(F.COD_CODON, Mai_mult.enzime_de_test.List(Y), tot_codul_ADN, i)) ' ->
End If

If (F.del_stanga.Value = True) Then
SC_out_put (enz_del_dreapta(F.COD_CODON, Mai_mult.enzime_de_test.List(Y), tot_codul_ADN, i)) ' <-
End If

If (F.ins_dreapta.Value = True) Then
SC_out_put (enz_ins_dreapta(F.COD_CODON, Mai_mult.enzime_de_test.List(Y), tot_codul_ADN, i)) ' <-
End If

If (F.ins_stanga.Value = True) Then
SC_out_put (enz_ins_stanga(F.COD_CODON, Mai_mult.enzime_de_test.List(Y), tot_codul_ADN, i)) ' <-
End If
'-----------------------
End If
'=======================================
'=======================================


'=======================================
'===        VERIFICA NORMAL <>       ===
'=======================================
local_ACCEPT_GEL1 = False
local_ACCEPT_GEL2 = False

For rt = 1 To 2000
DoEvents
Ordine_ambii_algoritmi(rt) = Empty
Next rt


If (F.normal_dreapta.Value = True) Then
SC_out_put (enz_normal_dreapta(F.COD_CODON, Mai_mult.enzime_de_test.List(Y), tot_codul_ADN))
local_ACCEPT_GEL1 = True
End If

If (F.normal_stanga.Value = True) Then
SC_out_put (enz_normal_stanga(F.COD_CODON, Mai_mult.enzime_de_test.List(Y), tot_codul_ADN))
local_ACCEPT_GEL2 = True
End If

If principal.Arata_fragmente.Value = 1 And local_ACCEPT_GEL1 = True And local_ACCEPT_GEL2 = True Then
Call Restricteaza_COD(principal.COD_CODON.Text, Mai_mult.enzime_de_test.List(Y), Y)
End If
'=======================================
'=======================================
1:

'############
Next Y
F.bara_date.Value = 0
F.bara_date2.Value = 0
'############
End Function
'##############################################################################################
'###                                                                                        ###
'###                                                                                        ###
'###                                                                                        ###
'##############################################################################################


'##############################################################################################
'###                                                                                        ###
'###                                ENZIMA RESTRICTIE NORMALA [->]                          ###
'###                                                                                        ###
'##############################################################################################
Function enz_normal_dreapta(ByVal obiectul As RichTextBox, ByVal enzima_normal As String, ByVal codul_ADN As String) As String
enzima_normal = UCase(Replace(enzima_normal, "*", ""))
COD_ramas = UCase(codul_ADN)
For rt = 1 To 1000
DoEvents
poz_fragment_dreapta(rt) = Empty
Next rt
warza = 0
carnat = 0
situsuri_restrictie_del_d = 0 'asfel se fute algoritmul de la gel ... totul va fi /2
'-------------------------------------------------------------------
0:
'############
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then Exit Function
'############
'------------
DoEvents
'------------
poz = InStr(COD_ramas, enzima_normal) - 1
If (poz < 0) Then GoTo 22

linie = Mid$(COD_ramas, 1, poz)
carnat = carnat + poz + Len(enzima_normal)
COD_ramas = Mid$(COD_ramas, poz + Len(enzima_normal) + 1, Len(COD_ramas))
'-----------------------------------------------------------------
situsuri_restrictie_del_d = situsuri_restrictie_del_d + 1
obiectul.SelStart = carnat - Len(enzima_normal)
obiectul.SelLength = Len(enzima_normal)
obiectul.SelColor = Mai_mult.start5normal.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.start5normal.BackColor, False)
'-------------------------------------------------------------------
warza = warza + 1
poz_fragment_dreapta(warza) = carnat
'-------------------------------------------------------------------
If COD_ramas = "" Then GoTo 22

GoTo 0
22:
If (warza > 0) Then
enz_normal_dreapta = "Nr. restriction sites already existing in the 5'-3' direction:[" & warza & "]" & vbCrLf
Else
enz_normal_dreapta = "In the 5'-3' direction there are no restriction sites !" & vbCrLf
End If
End Function
'##############################################################################################
'##############################################################################################


'##############################################################################################
'###                                                                                        ###
'###                            ENZIMA RESTRICTIE NORMALA [<-]                              ###
'###                                                                                        ###
'##############################################################################################
Function enz_normal_stanga(ByVal obiectul As RichTextBox, ByVal enzima_normal As String, ByVal codul_ADN As String) As String
enzima_normal = UCase(CATENA_PLUS(Replace(enzima_normal, "*", "")))
COD_ramas = UCase(codul_ADN)
For rt = 1 To 1000
DoEvents
poz_fragment_stanga(rt) = Empty
Next rt
warza = 0
carnat = 0
situsuri_restrictie_del_s = 0
'-------------------------------------------------------------------
0:
'############
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then Exit Function
'############
'------------
DoEvents
'------------
poz = InStr(COD_ramas, enzima_normal) - 1
If (poz < 0) Then GoTo 22

linie = Mid$(COD_ramas, 1, poz)
carnat = carnat + poz + Len(enzima_normal)
COD_ramas = Mid$(COD_ramas, poz + Len(enzima_normal) + 1, Len(COD_ramas))
'-----------------------------------------------------------------
situsuri_restrictie_del_s = situsuri_restrictie_del_s + 1
obiectul.SelStart = carnat - Len(enzima_normal)
obiectul.SelLength = Len(enzima_normal)
obiectul.SelColor = Mai_mult.start3normal.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.start3normal.BackColor, False)
'-------------------------------------------------------------------
warza = warza + 1
poz_fragment_stanga(warza) = carnat
'-------------------------------------------------------------------
If COD_ramas = "" Then GoTo 22

GoTo 0
22:
If (warza > 0) Then
enz_normal_stanga = "Nr. restriction sites already existing in the 3'-5' direction:[" & warza & "]" & vbCrLf
Else
enz_normal_stanga = "In the 3'-5' direction there are no restriction sites !" & vbCrLf
End If
End Function
'##############################################################################################
'##############################################################################################


'##############################################################################################
'###                                                                                        ###
'###                          ENZIMA RESTRICTIE DELETIE [->]                                ###
'###                                                                                        ###
'##############################################################################################
Function enz_del_stanga(ByVal obiectul As RichTextBox, ByVal enzima_restrictie_del As String, ByVal codul_ADN As String, ByVal Dist As Variant) As String
Dim poz As String
Dim linie As String
Dim unu As String
Dim doi As String

ce_enzima = enzima_restrictie_del
enzima_restrictie_del = UCase(Replace(enzima_restrictie_del, "*", ""))

For i = 1 To Len(enzima_restrictie_del)
'00000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000
DoEvents
COD_ramas = UCase(codul_ADN)
doi = UCase(Mid(enzima_restrictie_del, i, Len(enzima_restrictie_del)))
unu = UCase(Mid(enzima_restrictie_del, 1, i - 1))
'-------------------------------------------------------------------
carnat = 0
0:
'############
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then Exit Function
'############
'------------
DoEvents
'------------
poz = Empty
poz_doi = InStr(COD_ramas, doi) - 1

If (poz_doi <= 0) Then GoTo 22
If (Len(unu) = 0) Then GoTo 22

linie = Mid$(COD_ramas, 1, poz_doi)
carnat = carnat + poz_doi + Len(doi)
COD_ramas = Mid$(COD_ramas, poz_doi + Len(doi) + 1, Len(COD_ramas))


If (Len(linie) = Len(unu)) Then
A = ((Len(linie) - Dist) - Len(unu))
Else
A = ((Len(linie) - Dist) - Len(unu)) + 1
End If

If (A >= 0) Then
If (A = 0) Then A = 1
poz = Mid$(linie, A, Len(unu))
End If

If (poz = unu) Then
'-----------------------------------------------------------------
situsuri_restrictie_del = situsuri_restrictie_del + 1

If ((carnat - Len(doi) - Len(unu) - Dist) > 0) Then
'------
aaa = carnat - Len(doi) - Len(unu) - Dist
obiectul.SelStart = aaa
obiectul.SelLength = Len(unu)
obiectul.SelColor = Mai_mult.start5.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.start5.BackColor, False)
'------
bbb = carnat - Len(doi) - Dist
obiectul.SelStart = bbb
obiectul.SelLength = Dist
obiectul.SelColor = Mai_mult.mijloc5.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.mijloc5.BackColor, False)
'------
ccc = carnat - Len(doi)
obiectul.SelStart = ccc
obiectul.SelLength = Len(doi)
obiectul.SelColor = Mai_mult.stop5.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.stop5.BackColor, False)
'------
Call punct_lovit_enzima(aaa, ccc, principal.COD_CODON, Mai_mult.start5.BackColor)
'-------
zuluq = zuluq + 1
out = out & vbCrLf & "[" & zuluq & "] dupa nucleotida: [" & aaa & "]"
'------
End If
'-------------------------------------------------------------------
End If
warza = warza + 1
If COD_ramas = "" Then GoTo 22

GoTo 0
22:
'00000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000
Next i

If (situsuri_restrictie_del > 0) Then
out_put_x = out_put_x & "[" & ce_enzima & "] POSSIBLE r. sites through DELETION of " & Dist & " nucleotides 5'-3': [" & situsuri_restrictie_del & "]" & out & vbCrLf & vbCrLf
Else
'out_put_x = out_put_x & "In the 5'-3' direction through DELETION of " & Dist & " nucleotides there are NO POSSIBLE r. sites!" & vbCrLf
End If

enz_del_stanga = out_put_x
End Function
'##############################################################################################
'##############################################################################################


'##############################################################################################
'###                                                                                        ###
'###                        ENZIMA RESTRICTIE DELETIE [<-]                                  ###
'###                                                                                        ###
'##############################################################################################
Function enz_del_dreapta(ByVal obiectul As RichTextBox, ByVal enzima_restrictie_del As String, ByVal codul_ADN As String, ByVal Dist As Variant) As String
Dim poz As String
Dim linie As String
Dim unu As String
Dim doi As String

ce_enzima = enzima_restrictie_del
enzima_restrictie_del = UCase(CATENA_PLUS(Replace(enzima_restrictie_del, "*", "")))

For i = 1 To Len(enzima_restrictie_del)
'00000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000
DoEvents
COD_ramas = UCase(codul_ADN)
doi = UCase(Mid(enzima_restrictie_del, i, Len(enzima_restrictie_del)))
unu = UCase(Mid(enzima_restrictie_del, 1, i - 1))
'-------------------------------------------------------------------

carnat = 0
0:
'############
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then Exit Function
'############
'------------
DoEvents
'------------
poz_doi = InStr(COD_ramas, doi) - 1

If (poz_doi <= 0) Then GoTo 22 'Len(doi) + Len(unu) + dist
If (Len(unu) = 0) Then GoTo 22

linie = Mid$(COD_ramas, 1, poz_doi)
carnat = carnat + poz_doi + Len(doi)
COD_ramas = Mid$(COD_ramas, poz_doi + Len(doi) + 1, Len(COD_ramas))
'------------------------------------------ ' doar daca e caul !
'If (Len(linie) = Len(unu)) Then
'A = ((Len(linie) - Dist) - Len(unu))
'Else
'A = ((Len(linie) - Dist) - Len(unu)) + 1
'End If
'------------------------------------------ ' doar daca e caul !
A = ((Len(linie) - Dist) - Len(unu)) + 1
If (A >= 0) Then
If (A = 0) Then A = 1
poz = Mid$(linie, A, Len(unu))
End If

If (poz = unu) Then
'-----------------------------------------------------------------
situsuri_restrictie_del_inv = situsuri_restrictie_del_inv + 1

If ((Len(linie) - Len(unu) - Dist) >= 0) Then
'-------
aaa = carnat - Len(doi) - Len(unu) - Dist
obiectul.SelStart = aaa
obiectul.SelLength = Len(unu)
obiectul.SelColor = Mai_mult.start3.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.start3.BackColor, False)
'-------
bbb = carnat - Len(doi) - Dist
obiectul.SelStart = bbb
obiectul.SelLength = Dist
obiectul.SelColor = Mai_mult.mijloc3.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.mijloc3.BackColor, False)
'obiectul.SelBold = True
'obiectul.SelUnderline = True
'-------
ccc = carnat - Len(doi)
obiectul.SelStart = ccc
obiectul.SelLength = Len(doi)
obiectul.SelColor = Mai_mult.stop3.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.stop3.BackColor, False)
'-------
Call punct_lovit_enzima(aaa, ccc, principal.COD_CODON, Mai_mult.start3.BackColor)
'-------

zuluq = zuluq + 1
out = out & vbCrLf & "[" & zuluq & "] dupa nucleotida: [" & aaa & "]"
'-------
End If
'-------------------------------------------------------------------
End If

warza = warza + 1
If COD_ramas = "" Then GoTo 22

GoTo 0
22:
'00000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000
Next i

If (situsuri_restrictie_del_inv > 0) Then
out_put_x = out_put_x & "[" & ce_enzima & "] POSSIBLE r. sites through DELETION of " & Dist & " nucleotides 3'-5': [" & situsuri_restrictie_del_inv & "]" & out & vbCrLf & vbCrLf
Else
'out_put_x = out_put_x & "In the 3'-5' direction through DELETION of " & Dist & " nucleotides there are NO POSSIBLE r. sites!" & vbCrLf
End If

enz_del_dreapta = out_put_x
End Function
'##############################################################################################
'##############################################################################################


'##############################################################################################
'###                                                                                        ###
'###                          ENZIMA RESTRICTIE INSERTIE [->]                               ###
'###                                                                                        ###
'##############################################################################################
Function enz_ins_dreapta(ByVal obiectul As RichTextBox, ByVal enzima_restrictie_del As String, ByVal codul_ADN As String, ByVal Dist As Variant) As String
Dim poz As String
Dim linie As String
Dim unu As String
Dim doi As String

If (Len(enzima_restrictie_del) - Dist < 3) Then Exit Function

ce_enzima = enzima_restrictie_del
enzima_restrictie_del = UCase(Replace(enzima_restrictie_del, "*", ""))

For i = 1 To Len(enzima_restrictie_del)
'00000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000
DoEvents
COD_ramas = UCase(codul_ADN)
doi = UCase(Mid(enzima_restrictie_del, i, Len(enzima_restrictie_del)))
unu = UCase(Mid(enzima_restrictie_del, 1, i - 1))
nuc_de_inserat = Mid(doi, 1, Dist)
doi = Mid(doi, Dist + 1, Len(doi))

enzima_normal = unu & doi
carnat = 0
'-------------------------------------------------------------------
If (Len(enzima_restrictie_del) - Len(unu) < Dist) Then GoTo 22
If (Len(doi) + Len(unu) < 3) Then GoTo 22
'-------------------------------------------------------------------
'out_put_x = out_put_x & "[" & unu & "^" & doi & "]" & vbCrLf
0:
'############
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then Exit Function
'############
'------------
DoEvents
'------------
poz = InStr(COD_ramas, enzima_normal) - 1
If (poz < 0) Then GoTo 22

linie = Mid$(COD_ramas, 1, poz)
carnat = carnat + poz + Len(enzima_normal)
COD_ramas = Mid$(COD_ramas, poz + Len(enzima_normal) + 1, Len(COD_ramas))
'-----------------------------------------------------------------
situsuri_restrictie_ins = situsuri_restrictie_ins + 1
'-------
aaa = carnat - Len(enzima_normal)
obiectul.SelStart = aaa
obiectul.SelLength = Len(unu)
obiectul.SelColor = Mai_mult.start_ins5.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.start_ins5.BackColor, False)
obiectul.SelStart = carnat - Len(enzima_normal) + Len(unu)
obiectul.SelLength = Len(doi)
obiectul.SelColor = Mai_mult.stop_ins5.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.stop_ins5.BackColor, False)
'-------
Call punct_lovit_enzima(aaa, carnat - Len(enzima_normal) + Len(unu), principal.COD_CODON, Mai_mult.start_ins5.BackColor)
'-------
zuluq = zuluq + 1
out = out & vbCrLf & "[" & zuluq & "] dupa nucleotida: [" & aaa & "]-[" & unu & "^" & doi & "] - de inserat " & Dist & "b" & " - [" & nuc_de_inserat & "]"
'-------
'-------------------------------------------------------------------
warza = warza + 1
If COD_ramas = "" Then GoTo 22

GoTo 0
22:
'00000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000
Next i

If (situsuri_restrictie_ins > 0) Then
out_put_x = out_put_x & "[" & ce_enzima & "] POSSIBLE r. sites through INSERTION of " & Dist & " nucleotides 5'-3': [" & situsuri_restrictie_ins & "]" & out & vbCrLf & vbCrLf
Else
'out_put_x = out_put_x & "In the 5'-3' direction through DELETION of " & Dist & " nucleotides there are NO POSSIBLE r. sites!" & vbCrLf
End If

enz_ins_dreapta = out_put_x
End Function
'##############################################################################################
'##############################################################################################

'##############################################################################################
'###                                                                                        ###
'###                          ENZIMA RESTRICTIE INSERTIE [<-]                               ###
'###                                                                                        ###
'##############################################################################################
Function enz_ins_stanga(ByVal obiectul As RichTextBox, ByVal enzima_restrictie_del As String, ByVal codul_ADN As String, ByVal Dist As Variant) As String
Dim poz As String
Dim linie As String
Dim unu As String
Dim doi As String

If (Len(enzima_restrictie_del) - Dist < 3) Then Exit Function

ce_enzima = enzima_restrictie_del
enzima_restrictie_del = UCase(CATENA_PLUS(Replace(enzima_restrictie_del, "*", "")))

For i = 1 To Len(enzima_restrictie_del)
'00000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000
DoEvents
COD_ramas = UCase(codul_ADN)
unu = UCase(Mid(enzima_restrictie_del, i, Len(enzima_restrictie_del)))
nuc_de_inserat = Mid(unu, 1, Dist)
doi = UCase(Mid(enzima_restrictie_del, 1, i - 1))
unu = Mid(unu, Dist + 1, Len(unu))

unu_bis = doi
doi_bis = unu
unu = unu_bis
doi = doi_bis

enzima_normal = unu & doi
carnat = 0
'-------------------------------------------------------------------
If (Len(enzima_restrictie_del) - Len(unu) < Dist) Then GoTo 22
If (Len(doi) + Len(unu) < 3) Then GoTo 22
'-------------------------------------------------------------------
'out_put_x = out_put_x & "[" & unu & "^" & doi & "]" & vbCrLf
0:
'############
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then Exit Function
'############
'------------
DoEvents
'------------
poz = InStr(COD_ramas, enzima_normal) - 1
If (poz < 0) Then GoTo 22

linie = Mid$(COD_ramas, 1, poz)
carnat = carnat + poz + Len(enzima_normal)
COD_ramas = Mid$(COD_ramas, poz + Len(enzima_normal) + 1, Len(COD_ramas))
'-----------------------------------------------------------------
situsuri_restrictie_ins = situsuri_restrictie_ins + 1
'-----------------------------------------------------------------
aaa = carnat - Len(enzima_normal)
obiectul.SelStart = aaa
obiectul.SelLength = Len(unu)
obiectul.SelColor = Mai_mult.start_ins3.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.start_ins3.BackColor, False)
obiectul.SelStart = carnat - Len(enzima_normal) + Len(unu)
obiectul.SelLength = Len(doi)
obiectul.SelColor = Mai_mult.stop_ins3.ForeColor
Call HighLightSelection(principal, obiectul, Mai_mult.stop_ins3.BackColor, False)
'-------
Call punct_lovit_enzima(aaa, carnat - Len(enzima_normal) + Len(unu), principal.COD_CODON, Mai_mult.start_ins3.BackColor)
'-------
zuluq = zuluq + 1
out = out & vbCrLf & "[" & zuluq & "] dupa nucleotida: [" & aaa & "]-[" & unu & "^" & doi & "] - de inserat " & Dist & "b" & " - [" & nuc_de_inserat & "]"
'-------
'-------------------------------------------------------------------
warza = warza + 1
If COD_ramas = "" Then GoTo 22

GoTo 0
22:
'00000000000000000000000000000000000000000000
'00000000000000000000000000000000000000000000
Next i

If (situsuri_restrictie_ins > 0) Then
out_put_x = out_put_x & "[" & ce_enzima & "] POSSIBLE r. sites through INSERTION of " & Dist & " nucleotides 3'-5': [" & situsuri_restrictie_ins & "]" & out & vbCrLf & vbCrLf
Else
'out_put_x = out_put_x & "In the 5'-3' direction through DELETION of " & Dist & " nucleotides there are NO POSSIBLE r. sites!" & vbCrLf
End If

enz_ins_stanga = out_put_x
End Function
'##############################################################################################
'##############################################################################################

'##############################################################################################
'###                                                                                        ###
'###                                     RESTRICTEAZA                                       ###
'###                                                                                        ###
'##############################################################################################
Function Restricteaza_COD(ByVal cod As String, ByVal enzima_normal As String, ByVal enzima_nr As Variant) As String
distantier_linii = "----------------------------------------------------------------------------------------"
SC_out_put (vbCrLf & distantier_linii & vbCrLf)

enzima_RES1 = UCase(Mid(enzima_normal, 1, InStr(enzima_normal, "*") - 1))
enzima_RES2 = UCase(Mid(enzima_normal, InStr(enzima_normal, "*") + 1, Len(enzima_normal)))

t = situsuri_restrictie_del_d + situsuri_restrictie_del_s + 1 ' plus 1 pt ca ciclul incepe de la 1  si nu de la 0
aa = 0
bB = 0

rev1 = 0
rev2 = 0

old_intaietate_are = 1
intaietate_are = 1

For i = 1 To t
'############
DoEvents
If principal.Testeaza_CODUL.Caption = "Start" Then Exit Function
'############
DoEvents

old_intaietate_are = intaietate_are
'########################################
If poz_fragment_dreapta(i - bB) < poz_fragment_stanga(i - aa) Then
aa = aa + 1
intaietate_are = poz_fragment_dreapta(i - bB)
'------------
eu_nou = "->"
'------------
If eu_vechi = "->" Then
rev1 = Len(enzima_RES2)
SC_out_put (vbCrLf & vbCrLf & "[Nr." & i & "][" & eu_vechi & "][" & eu_nou & "][" & enzima_RES2 & "_" & enzima_RES1 & "]")
End If
If eu_vechi = "<-" Then
rev1 = Len(enzima_RES1)
SC_out_put (vbCrLf & vbCrLf & "[Nr." & i & "][" & eu_vechi & "][" & eu_nou & "][" & StrReverse(CATENA_PLUS(enzima_RES1)) & "_" & enzima_RES1 & "]")
End If
'------------
rev2 = Len(enzima_RES2)
'------------
eu_vechi = eu_nou
Else '###################################
bB = bB + 1
intaietate_are = poz_fragment_stanga(i - aa)
'------------
eu_nou = "<-"
'------------
If eu_vechi = "->" Then
rev1 = Len(enzima_RES2)
SC_out_put (vbCrLf & vbCrLf & "[Nr." & i & "][" & eu_vechi & "][" & eu_nou & "][" & enzima_RES2 & "_" & StrReverse(CATENA_PLUS(enzima_RES2)) & "]")
End If
If eu_vechi = "<-" Then
rev1 = Len(enzima_RES1)
SC_out_put (vbCrLf & vbCrLf & "[Nr." & i & "][" & eu_vechi & "][" & eu_nou & "][" & StrReverse(CATENA_PLUS(enzima_RES1)) & "_" & enzima_RES2 & "]")
End If
'------------
rev2 = Len(enzima_RES1)
'------------
eu_vechi = eu_nou
End If
'########################################


If (old_intaietate_are = Empty Or intaietate_are = Empty) Then GoTo 1
If (old_intaietate_are = 1) Then
rev1 = 1
'rev1 = 1
End If
'If (old_intaietate_are = 1) Then old_intaietate_are = 2
SC_out_put (" The restricted fragment starts at position: [" & old_intaietate_are + 1 - rev1 & "]")

' Fix: Mid$ cere start >= 1 si lungime >= 0; cand un sit de restrictie e
' aproape de inceputul secventei, "old_intaietate_are + 1 - rev1" poate iesi
' sub 1, iar lungimea poate iesi negativa -> Run-time error 5. Le limitam
' fara sa schimbam pozitiile calculate mai sus (Ordine_ambii_algoritmi ramane neschimbat).
Dim lngStartMid As Long, lngLenMid As Long
lngStartMid = old_intaietate_are + 1 - rev1
If lngStartMid < 1 Then lngStartMid = 1
lngLenMid = (intaietate_are - old_intaietate_are) - rev2 + rev1
If lngLenMid < 0 Then lngLenMid = 0

secventa_restrictata = Mid$(cod, lngStartMid, lngLenMid)
secvente_rerstrictate(i) = secventa_restrictata
Ordine_ambii_algoritmi(i) = intaietate_are + 1 - rev1

SC_out_put (vbCrLf & distantier_linii & vbCrLf & secventa_restrictata & vbCrLf & distantier_linii & vbCrLf & "-- The restriction fragment has: [" & Len(secventa_restrictata) & "b] -- " & Proportie_A_T_G_C_U(secventa_restrictata, True) & vbCrLf & distantier_linii & vbCrLf & vbCrLf)
1:
Next i

SC_out_put (vbCrLf & distantier_linii & vbCrLf & _
"Total restriction sites existing in the 5'-3' and 3'-5' directions:[" & situsuri_restrictie_del_d + situsuri_restrictie_del_s & "]" & vbCrLf & _
distantier_linii & vbCrLf)

Call Segment_TOTAL(cod)

Call GEL(situsuri_restrictie_del_d + situsuri_restrictie_del_s, enzima_nr)
End Function
'##############################################################################################
'##############################################################################################


Function verifica_corect(ByVal enzima As String, ByVal cod As String) As Boolean

If (Len(enzima) >= Len(cod)) Then
MsgBox "The enzyme is larger than the code to be analyzed !"
verifica_corect = False
Exit Function
End If

verifica_corect = True
End Function
