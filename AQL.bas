Sub AQL()
'
' Mechanics Macro
' Macro recorded 29.01.2011 by AlexCulea
'AQL
Range("G8").Activate 'This is done becouse buttons become incative in excel and with this line reactivates them.
Sheets("Report").Select

Dim var As Variant, result As String
var = Range("D19")
Select Case var

Case "": result = MsgBox("Te rog sa introduci la cantitatea livrata minim 1." & vbNewLine & "Nu ai trecut nici un numar.")

Case 0: result = MsgBox("Te rog sa introduci la cantitatea livrata minim 1")

Case 1: result = 1

Case 2 To 8: result = 3

Case 9 To 15: result = 3

Case 16 To 25: result = 3

Case 26 To 50: result = 5

Case 51 To 90: result = 6

Case 91 To 150: result = 7

Case 151 To 280: result = 10

Case 281 To 500: result = 11

Case 501 To 1200: result = 15

Case 1201 To 3200: result = 18

Case 3201 To 10000: result = 22

Case 10001 To 35000: result = 29

Case 35001 To 150000: result = 29

Case 150001 To 500000: result = 29

Case Else: result = "Over 500001"
 

End Select
 Range("D21").value = result
 If Range("D19").value = 0 Then Range("D21").value = 0
        Sheets("Report").Select
        Range("G22").Select
        ActiveCell.FormulaR1C1 = "4.00"
        Range("I22").Select
        ActiveCell.FormulaR1C1 = "10"
        
  If Range("D19") < 10 Then 'If Sample Size is smaller than 10 then show number from cell D21
       Range("D19").Select
       Selection.Copy
       Range("I22").Select
       Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
       Selection.HorizontalAlignment = xlCenter
  End If
    Range("C29").Select

End Sub    ' End question script
