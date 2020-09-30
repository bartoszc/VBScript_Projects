Option Explicit

Dim strCiagCyfr
Dim arrCyfry(9)
Dim intIterator
Dim intDlugoscCiagu
Dim strRezultat

arrCyfry(0) = "Zero"
arrCyfry(1) = "Jeden"
arrCyfry(2) = "Dwa"
arrCyfry(3) = "Trzy"
arrCyfry(4) = "Cztery"
arrCyfry(5) = "Piec"
arrCyfry(6) = "Szesc"
arrCyfry(7) = "Siedem"
arrCyfry(8) = "Osiem"
arrCyfry(9) = "Dziewiec"

strCiagCyfr = InputBox("Podaj ciag cyfr")

intDlugoscCiagu = Len(strCiagCyfr)

For intIterator = 1 To intDlugoscCiagu
	If (intIterator = intDlugoscCiagu) Then
		strRezultat = strRezultat & arrCyfry(Mid(strCiagCyfr, intIterator, 1))
	Else
		strRezultat = strRezultat & arrCyfry(Mid(strCiagCyfr, intIterator, 1)) & ", " 
	End If
Next

intIterator = 1
While (intIterator <= Len(strCiagCyfr))
	charCyfra = Mid(strCiagCyfr, intIterator, 1)
	strWynik = strWynik + arrCyfry(cInt(charCyfra)) & ", "
	intIterator = intIterator + 1
Wend

intIterator = 1
Do While (intIterator <= Len(strCiagCyfr))
	charCyfra = Mid(strCiagCyfr, intIterator, 1)
	strWynik = strWynik + arrCyfry(cInt(charCyfra)) & ", "
	intIterator = intIterator + 1
Loop

intIterator = 1
Do Until (intIterator > Len(strCiagCyfr))
	charCyfra = Mid(strCiagCyfr, intIterator, 1)
	strWynik = strWynik + arrCyfry(cInt(charCyfra)) & ", "
	intIterator = intIterator + 1
Loop

MsgBox strRezultat