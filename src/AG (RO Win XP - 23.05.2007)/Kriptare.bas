Attribute VB_Name = "Kriptare"
Function xQ(sText As String)
On Error Resume Next
Dim ekey As Long, i As Long
Dim hash As String, crbyte As String
ekey = 1234
For i = 1 To Len(sText)
hash = Asc(Mid(sText, i, 1))
crbyte = Chr(hash Xor (ekey Mod 255))
xQ = xQ & crbyte
Next i
End Function

' function x encrypts strings
Function x666(sText As String)
On Error Resume Next
Dim ekey As Long, i As Long
Dim hash As String, crbyte As String
ekey = 1234
For i = 1 To Len(sText)
hash = Asc(Mid(sText, i, 1))
crbyte = Chr(hash Xor (ekey Mod 255))
x666 = x666 & crbyte
Next i
End Function

Function SavageDown(shit)
For A = 1 To Len(shit)
 b = Asc(Mid(shit, A, Len(shit)))
 b = Chr(b + 6)
 c = c & b
Next
SavageDown = c
End Function

Function SavageUp(shit)
For A = 1 To Len(shit)
 b = Asc(Mid(shit, A, Len(shit)))
 b = Chr(b - 6)
 c = c & b
Next
SavageUp = c
End Function

Public Function Crypt(ByVal inp As String, ByVal Key As String) As String
Dim Sbox(0 To 255) As Long
Dim Sbox2(0 To 255) As Long
Dim j As Long, i As Long, t As Double
Dim K As Long, temp As Long, Outp As String
Dim X

For i = 0 To 255
Sbox(i) = i
Next i

j = 1

For i = 0 To 255
If j > Len(Key) Then j = 1
Sbox2(i) = Asc(Mid(Key, j, 1))
j = j + 1
Next i

j = 0
For i = 0 To 255
j = (j + Sbox(i) + Sbox2(i)) Mod 256
temp = Sbox(i)
Sbox(i) = Sbox(j)
Sbox(j) = temp
Next i

i = 0
j = 0
For X = 1 To Len(inp)
i = (i + 1) Mod 256
j = (j + Sbox(i)) Mod 256
temp = Sbox(i)
Sbox(i) = Sbox(j)
Sbox(j) = temp
t = (Sbox(i) + Sbox(j)) Mod 256
K = Sbox(t)

Outp = Outp + Chr(Asc(Mid(inp, X, 1)) Xor K)
Next X

Crypt = Outp
End Function


Private Function Encrypt(TextToEncrypt, ByVal intkey As Integer) As String
    Dim EncryptedText As String
    Randomize
    Dim varBlah() As String


    If intkey = 0 Then
        intkey = 1
    End If
    'create an array equal to the size of th
    '     e given key
    ReDim varBlah(intkey - 1) As String


    If TextToEncrypt <> "" Then
        Dim varCrypt As String
        Dim varTemp As Long
        Dim i As Long
        Dim j As Integer


        For i = 1 To Len(TextToEncrypt)
            'convert the string into an ascii value
            varTemp = Asc(Mid$(TextToEncrypt, i, 1))
            'convert the ascii value into a string
            varCrypt = Chr$(varTemp - i)
            EncryptedText = EncryptedText + varCrypt
            'cycle through the array and add to the
            '     encrypted string a random
            'character


            For j = LBound(varBlah()) To UBound(varBlah())
                varBlah(j) = Chr$((255 * Rnd + 1))
                EncryptedText = EncryptedText + varBlah(j)
            Next
        Next
    End If
    Encrypt = EncryptedText
End Function


Private Function Decrypt(TextToDecrypt, ByVal intkey As Integer) As String
    Dim DecryptedText As String


    If TextToDecrypt <> "" Then
        Dim varCrypt As String
        Dim varTemp As Long


        If intkey = 0 Then
            intkey = 1
        End If
        Dim i As Long
        Dim j As Integer
        j = 1
        'takes only the characters in the encryp
        '     ted string, based upon the
        'key value


        For i = 1 To Len(TextToDecrypt) Step (intkey + 1)
            varTemp = Asc(Mid$(TextToDecrypt, i, 1))
            varCrypt = Chr$(varTemp + j)
            DecryptedText = DecryptedText + varCrypt
            j = j + 1
        Next
    End If
    Decrypt = DecryptedText
End Function


