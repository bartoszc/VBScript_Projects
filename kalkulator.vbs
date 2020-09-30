'Wymuszenie deklaracji zmiennych
Option Explicit

Dim intNumberA
Dim intNumberB
Dim intSuma
Dim intRoznica
Dim intIloczyn
Dim intIloraz

'intNumberA = 3
'intNumberB = 5

intNumberA = InputBox("Podaj liczbe A:")
intNumberB = InputBox("Podaj liczbe B:")

intNumberA = cInt(intNumberA)
intNumberB = cInt(intNumberB)

intSuma = FnKalkulator(intNumberA, intNumberB, "suma")
intRoznica = FnKalkulator(intNumberA, intNumberB, "roznica")
intIloczyn = FnKalkulator(intNumberA, intNumberB, "iloczyn")
intIloraz = FnKalkulator(intNumberA, intNumberB, "iloraz")

MsgBox "Wynik: " & vbCrLf & "Suma = " & intSuma & _
				   vbCrLf & "Roznica = " & intRoznica & _
				   vbCrLf & "Iloczyn = " & intIloczyn & _
				   vbCrLf & "Iloraz = " & intIloraz
'------------------------------------------------------
Function FnKalkulator (intA, intB, strDzialanie)
Dim intDzialanie

Select Case strDzialanie

	Case "suma"
		intDzialanie = intA + intB
	Case "roznica"
		intDzialanie = intA - intB
	Case "iloczyn"
		intDzialanie = intA * intB
	Case "iloraz"
		If (intB = 0) Then
			intDzialanie = "Nie ma dzielenia przez 0"
		Else
			intDzialanie = intA / intB
		End If
	Case Else
		intDzialanie = "Niepoprawny parametr strDzialanie = " & strDzialanie
End Select

FnKalkulator = intDzialanie
End Function