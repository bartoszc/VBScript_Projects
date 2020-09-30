Option Explicit

Dim intNumberA
Dim intSilnia
Dim intIterator

intNumberA = InputBox("Podaj wartosc liczby: ")
intSilnia = 1

For intIterator = 1 To intNumberA
	intSilnia = intSilnia * intIterator
Next

MsgBox "Silnia z liczby " & intNumberA & " wynosi: " & intSilnia