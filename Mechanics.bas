'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX    Contacts Mechanics    XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Sub ABi_M()              'ABB JAGUAR Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna," & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C3").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C4").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$63")
cc = Sheets("Emails").Range("K64").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
    
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub BARCO_M()              'BARCO Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C59").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C60").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$65")
cc = Sheets("Emails").Range("K66").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
    
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub BEI_M()              'BEI Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C35").value 'Contacts range TO
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C36").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$65")
cc = Sheets("Emails").Range("K70").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
    
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub CINIONIC_M()              'CINIONIC Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C59").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C60").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$65")
cc = Sheets("Emails").Range("K66").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
    
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub Preh_M()              'PREH Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C23").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C24").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$67")
cc = Sheets("Emails").Range("K68").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")

With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub FLUKE_M()              'FLUKE Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C19").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C20").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$69")
cc = Sheets("Emails").Range("K70").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")

With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub GE_M()              'GEH Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C7").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C8").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$71")
cc = Sheets("Emails").Range("K72").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
    
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub MAREL_M()              'MAREL Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C43").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C44").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$79")
cc = Sheets("Emails").Range("K80").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
  
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")

With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub

Sub EmersonS_M()              'EMERSON SPECTRONIX Mechanics
'Pt. emailuri la EMERSON SPECTRONIX se trimite la Rosemount
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C51").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C52").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$79")
cc = Sheets("Emails").Range("K80").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
  
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")

With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub PARKER_M()              'Parker Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C11").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C12").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$75")
cc = Sheets("Emails").Range("K76").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
    
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub PHILIPS_M()              'PHILIPS Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
strEmailSubject = Subject + " -Lipsã desen-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

nameList = nameList & ";" & Sheets("Emails").Range("C43").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C44").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$77")
cc = Sheets("Emails").Range("K78").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If
    
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub Rational_M()              'Rational Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C39").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C40").value

'Start of Lipsa desen
If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$85")
cc = Sheets("Emails").Range("K86").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If
'End of 'Start of Lipsa desen

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub ROSEMOUNT_M()              'ROSEMOUNT Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C51").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C52").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$79")
cc = Sheets("Emails").Range("K80").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub RohdeNSchwarz_M()              'Rohde & Schwarz Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C27").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C28").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$87")
cc = Sheets("Emails").Range("K88").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If

Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub SIEMENS_M()              'SIEMENS Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & SuppEmailText = "Buna" & vbNewLine & _
""

nameList = nameList & ";" & Sheets("Emails").Range("C47").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C48").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$81")
cc = Sheets("Emails").Range("K82").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
    
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub SETRA_M()              'SETRA Mechanics
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
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C35").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C36").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$69")
cc = Sheets("Emails").Range("K70").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
    
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub
Sub WEINMANN_M()              'WEINMANN Mechanics
Dim olLook As Object 'Start MS Outlook
Dim olNewEmail As Object 'New email in Outlook
Dim strEmailSubject As String 'Contact email address
Set olLook = CreateObject("Outlook.Application")
Set olNewEmail = olLook.CreateItem(0)
ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveWorkbook.Save
Subject = ActiveCell.Offset(0, 1).Range("A1")
Motiv = ActiveCell.Offset(0, 9).Range("A1")

PO = ActiveCell.Offset(0, 8).Range("A1").value
Lot = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("G9")
qnt = ActiveCell.Offset(0, 11).Range("A1").value
Supp = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("F8")

strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"" & vbNewLine & _
"" & vbNewLine & _
"Lot: " + Lot & vbNewLine & _
"PO: " & PO & vbNewLine & _
"Quantity: " & qnt & vbNewLine & _
"Supplier: " & Supp

nameList = nameList & ";" & Sheets("Emails").Range("C55").value 'Contacts range
    EmailSendTo = nameList
cc = Sheets("Emails").Range("C56").value

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA DESEN" Then
SubjectLD = Range("='[PARTURI NOK INCOMING.xlsm]Emails'!$K$83")
cc = Sheets("Emails").Range("K84").value
EmailSendTo = SubjectLD
strEmailSubject = Subject + " -" + Motiv + "-"
strEmailText = "Buna" & vbNewLine & _
"Avem nevoie de desen la partul din subiect"
End If

If ActiveCell.Offset(0, 9).Range("A1").value = "LIPSA MPN AGILE" Then
strEmailSubject = Subject + " -Lipsã MPN-"
strEmailText = "Buna" & vbNewLine & _
"Partul din subiect nu are setat MPN in Agile, vã rog sã ne ajutati."
End If
    
Dim DDate As String
DDate = Format(Date, "dd.mm.yyyy")
    
With olNewEmail 'Attach template
.To = EmailSendTo
.cc = cc
.Body = strEmailText
.Subject = strEmailSubject
.Display
On Error Resume Next
.attachments.Add ("G:\Incoming\NCR PN Neconforme\" & Subject & " " & DDate & ".docx")

End With
End Sub

