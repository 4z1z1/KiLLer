Attribute VB_Name = "Module1"
Const Clef = 235
Public Function Crypt(Texte)
b = ""
For a = 1 To Len(Texte)
Crypt = Crypt & Chr(Asc(Mid(Texte, a, 1)) Xor (129 - (a Mod 259)) Xor Clef)
Next a
End Function
