Attribute VB_Name = "RC4_Calc"
Public key As String
'I am  using this function to encode/decode
'IM text for privacy issues. It is a fairly strong
'logarithm called RC4.
'The key is set in MDIForm_Load. Change the key
'for your own personal uses

Public Function RC4(inp As String, key As String) As String
Dim S(0 To 255) As Byte, K(0 To 255) As Byte, i As Long
Dim j As Long, temp As Byte, Y As Byte, t As Long, x As Long
Dim Outp As String

For i = 0 To 255
    S(i) = i
Next

j = 1
For i = 0 To 255
    If j > Len(key) Then j = 1
    K(i) = Asc(Mid(key, j, 1))
    j = j + 1
Next i

j = 0
For i = 0 To 255
    j = (j + S(i) + K(i)) Mod 256
    temp = S(i)
    S(i) = S(j)
    S(j) = temp
Next i

i = 0
j = 0
For x = 1 To Len(inp)
    i = (i + 1) Mod 256
    j = (j + S(i)) Mod 256
    temp = S(i)
    S(i) = S(j)
    S(j) = temp
    t = (S(i) + (S(j) Mod 256)) Mod 256
    Y = S(t)
    
    Outp = Outp & Chr(Asc(Mid(inp, x, 1)) Xor Y)
Next
RC4 = Outp
End Function
