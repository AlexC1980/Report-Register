'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX   Contacts Electronics   XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub ABi_E()              'ABB JAGUAR Electronice
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "UTILIZATOR, ai ales un proiect la electronice. TE ROG MODIFICÃ NECESARELE."
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C5").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C6").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub BARCO_E()              'BARCO Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C61").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C62").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub CINIONIC_E()              'CINIONIC Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C61").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C62").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub EMERSON_E()              'EMERSON Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C17").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C18").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub EmersonS_E()              'EMERSON SPECTRONIX Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C53").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C54").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub FLUKE_E()              'FLUKE Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C21").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C22").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub GE_E()              'GEH Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C9").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C10").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub INVENSYS_E()              'INVENSYS Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C29").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C30").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub MAREL_E()              'MAREL Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C45").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C46").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub PREH()              'Preh Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C25").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C26").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub Rational_E()              'Rational Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C41").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C42").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub PARKER_E()              'Parker Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C13").value 'Contacts range for TO
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C14").value ' Range Emails for CC
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub PHILIPS_E()              'PHILIPS Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C45").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C46").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub ROSEMOUNT_E()              'ROSEMOUNT Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C53").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C54").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub RohdeNSchwarz_E()              'Rohde & Schwarz Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C29").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C30").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub

Sub SIEMENS_E()              'SIEMENS Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C49").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C50").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub SETRA_E()              'SETRA Electronics
'La Setra se trimite email la cei de la Fluke
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C37").value 'Contacts range TO
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C38").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
Sub WEINMANN_E()              'WEINMANN Electronics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")

If ActiveCell.Offset(0, 9).Range("A1").value = "IMPACHETARE NECONFORMA" Then
strEmailSubject = Subject + " -Impachetare Neconformã-"
strEmailText = "Buna" & vbNewLine & _
""
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

'XXXXXXXXXXXXX
MPNa = ActiveCell.Offset(0, 2).Range("A1").value
MPNp = ActiveCell.Offset(0, 3).Range("A1").value
Supp = ActiveCell.Offset(0, 4).Range("A1").value
PO = ActiveCell.Offset(0, 8).Range("A1").value
qnt = ActiveCell.Offset(0, 11).Range("A1").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN PRODUS" Then
strEmailSubject = Subject + " -Lipsã MPN Produs-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are MPN pe produs." & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If
'XXXXXXXXXXXXX

If ActiveCell.Offset(0, 9).Range("A1").value = "MPN NECONFORM AGILE" Then
strEmailSubject = Subject + " -MPN Neconform-"
strEmailText = "Buna" & vbNewLine & _
"MPN diferit fata de AGILE" & vbNewLine & _
"" & vbNewLine & _
"MPN Agile   : " + MPNa & vbNewLine & _
"MPN Product: " + MPNp & vbNewLine & _
"Supplier: " + Supp & vbNewLine & _
"Manufacturer: " & vbNewLine & _
"PO: " + PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Baan Prod Line: " & Range("='[Report-Register.xlsm]Data'!$B$1")
End If

nameList = nameList & ";" & Sheets("Emails").Range("C57").value 'Contacts range
    EmailSendTo = nameList
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = Sheets("Emails").Range("C58").value
.Body = strEmailText
.Subject = strEmailSubject
.Display

End With
End Sub
