Attribute VB_Name = "Functii_PROTEINE"
Public codonul(1 To 64) As Variant




Function Proteina(ByVal codon As Variant) As String

'00000000000000000000000000000000000000000000
q.SelStart = 0
'00000000000000000000000000000000000000000000
Cul1 = RGB(200, 0, 0)
Cul2 = RGB(0, 0, 200)
codon_coloreaza = InStr(q.Text, z.coloreaza_acesta.Text)
If (codon_coloreaza <> 0) Then
Else
Exit Function
End If

q.SelStart = codon_coloreaza + Len(z.coloreaza_acesta.Text)
'00000000000000000000000000000000000000000000
For r = codon_coloreaza + Len(z.coloreaza_acesta.Text) - 1 To Len(q.Text) Step 3

If (A = Cul1) Then q.SelColor = Cul2: A = Cul2 Else q.SelColor = Cul1: A = Cul1


q.SelStart = r
q.SelLength = 3
w = w + 1
Next r
'00000000000000000000000000000000000000000000

Proteina = rezultat
End Function
