Attribute VB_Name = "Module1"
Sub PasaBinario(Num, Largo, StrBin)

a = Num
StrBin = ""
Do
    a1 = Int(a / 2)
    StrBin = Trim(Str(a - a1 * 2)) & StrBin
    a = a1
Loop Until a < 2
StrBin = Trim(Str(a)) & StrBin

For r = 1 To Largo
    StrBin = "0" & StrBin
Next r
StrBin = Right(StrBin, Largo)


End Sub
