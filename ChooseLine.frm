Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

Private Sub CheckBox1_Click()
If CheckBox1 = True Then
   CheckBox2 = False
End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2 = True Then
   CheckBox1 = False
End If
End Sub
Private Sub CheckBox3_Click()
If CheckBox3 = True Then
    Sheets("PARTURI SUSPECTE INCOMING").Select
Range("T5").FormulaR1C1 = "=TEXT(RIGHT(R[-1]C,2)+1,""00"")" 'old script: Range("T5").value = Range("T5") + 1
    Range("T6").FormulaR1C1 = "=""I""&TEXT(RIGHT(R[-2]C,2)+1,""00"")"
    Range("T4").value = "i" & Range("T5")
    ActiveCell.Offset(0, 17).Range("A1").value = Range("T4")
    Sheets("Inregistrare parturi").Select
End If
'If CheckBox3 = False Then ActiveCell.Offset(0, 17).Range("A1") = ""
End Sub

Private Sub CommandButton_Click()
    Dim cnt As Integer
    For Each ctl In Me.Controls
        If TypeName(ctl) = "OptionButton" Then
            If ctl.value = True Then cnt = cnt + 1
        End If
     Next ctl
     If cnt = 0 Then
     MsgBox "Selecteaza un motiv"
        Exit Sub
     Else: GoTo Nextone
     End If
Nextone:
  If CheckBox1 = False Then
  If CheckBox2 = False Then
         MsgBox "Selecteaza Electronice sau Mecanice"
        Exit Sub
  End If
  End If
Unload Me

' XXX ELECTRONICS XXX
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "ABB JAGUAR" Then
       Call ABi_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "BARCO" Then
      Call BARCO_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "CINIONIC" Then
      Call CINIONIC_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "EMERSON SPECTRONIX" Then
      Call EmersonS_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "FLUKE" Then
      Call FLUKE_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "GE" Then
      Call GE_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "MAREL" Then
      Call MAREL_E 'Electronics
    End If
    End If
  
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "PREH" Then
      Call PREH 'Electronics
    End If
    End If
   
   
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "PARKER" Then
      Call PARKER_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "PHILIPS" Then
      Call PHILIPS_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "RATIONAL" Then
      Call Rational_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "ROSEMOUNT" Then
      Call ROSEMOUNT_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "ROHDE & SCHWARZ" Then
      Call RohdeNSchwarz_E 'Electronics
    End If
    End If
               
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "SETRA" Then
      Call SETRA_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "SIEMENS" Then
      Call SIEMENS_E 'Electronics
    End If
    End If
    
    If Me.CheckBox1.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "LOWENSTEIN MEDICAL" Then
      Call WEINMANN_E 'Electronics
    End If
    End If
    
    
' XXX MECHANICS XXX
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "ABB JAGUAR" Then
       Call ABi_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "BARCO" Then
      Call BARCO_M
    End If
    End If
            
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "BEI" Then
      Call BEI_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "CINIONIC" Then
      Call CINIONIC_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "EMERSON SPECTRONIX" Then
      Call EmersonS_M
    End If
    End If
  
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "FLUKE" Then
      Call FLUKE_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "GE" Then
      Call GE_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "MAREL" Then
      Call MAREL_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "PARKER" Then
      Call PARKER_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "PREH" Then
      Call Preh_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "PHILIPS" Then
      Call PHILIPS_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "RATIONAL" Then
      Call Rational_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "ROSEMOUNT" Then
      Call ROSEMOUNT_M
    End If
    End If
        
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "ROHDE & SCHWARZ" Then
      Call RohdeNSchwarz_M 'Electronics
    End If
    End If
        
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "SETRA" Then
      Call SETRA_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "SIEMENS" Then
      Call SIEMENS_M
    End If
    End If
    
    If Me.CheckBox2.BoundValue = True Then
    If Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8").value = "LOWENSTEIN MEDICAL" Then
      Call WEINMANN_M
    End If
    End If
    

End Sub

Private Sub CommandButton1_Click() 'Butonul Anuleaza
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Sheets("PARTURI SUSPECTE INCOMING").Select
Selection.Rows(1).EntireRow.Delete

If Not (GetKeyState(vbKeyNumlock) = 0) Then
Application.SendKeys ("{NUMLOCK}")
End If

Unload ChooseLine
Application.DisplayAlerts = False
ActiveWorkbook.Close
End Sub

Private Sub OptionButton14_Click() 'ABATERI DIMENSIONALE
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveCell.Offset(0, 9).Range("A1") = "ABATERI DIMENSIONALE"
Sheets("Inregistrare parturi").Select
'part
Range("B15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("B8") 'PN from Report Register - sheet Register

'name
Range("G11").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Report").Range("D11") 'Name&Surname from Report Register - sheet Report

'quantity
Range("G15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("H8") 'Quantity from Report Register - sheet Register
    
'reason
Range("C19").Select
    ActiveCell.Range("A1") = "ABATERI DIMENSIONALE"
'status
Range("B23").Select
    ActiveCell.Range("A1") = "Se asteapta raspuns de la QE"
   CheckBox2 = True  'Mecanice
   CheckBox1 = False 'Electronice
   TextBox1 = Clear
End Sub
Private Sub OptionButton15_Click() 'COTE NETOLERATE
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveCell.Offset(0, 9).Range("A1") = "COTE NETOLERATE"
Sheets("Inregistrare parturi").Select
'part
Range("B15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("B8") 'PN from Report Register - sheet Register

'name
Range("G11").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Report").Range("D11") 'Name&Surname from Report Register - sheet Report

'quantity
Range("G15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("H8") 'Quantity from Report Register - sheet Register

'reason
Range("C19").Select
    ActiveCell.Range("A1") = "COTE NETOLERATE"
'status
Range("B23").Select
    ActiveCell.Range("A1") = "Se asteapta raspuns de la QE"
   CheckBox2 = True  'Mecanice
   CheckBox1 = False 'Electronice
   TextBox1 = Clear
End Sub
Private Sub OptionButton16_Click() 'IMPACHETARE NECONFORMA
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveCell.Offset(0, 9).Range("A1") = "IMPACHETARE NECONFORMA"
Sheets("Inregistrare parturi").Select
'part
Range("B15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("B8") 'PN from Report Register - sheet Register

'name
Range("G11").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Report").Range("D11") 'Name&Surname from Report Register - sheet Report

'quantity
Range("G15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("H8") 'Quantity from Report Register - sheet Register

'reason
Range("C19").Select
    ActiveCell.Range("A1") = "IMPACHETARE NECONFORMA"
'status
Range("B23").Select
    ActiveCell.Range("A1") = "Se asteapta raspuns de la PE"
   CheckBox2 = True  'Mecanice
   CheckBox1 = False 'Electronice
   TextBox1 = Clear
End Sub
Private Sub OptionButton19_Click() 'LIPSA DESEN
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveCell.Offset(0, 9).Range("A1") = "LIPSA DESEN"
Sheets("Inregistrare parturi").Select
'part
Range("B15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("B8") 'PN from Report Register - sheet Register

'name
Range("G11").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Report").Range("D11") 'Name&Surname from Report Register - sheet Report

'quantity
Range("G15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("H8") 'Quantity from Report Register - sheet Register

'reason
Range("C19").Select
    ActiveCell.Range("A1") = "LIPSA DESEN"
'status
Range("B23").Select
    ActiveCell.Range("A1") = "Se asteapta raspuns de la PM; QE si PE"
    
   CheckBox2 = True  'Mecanice
   CheckBox1 = False 'Electronice
   TextBox1 = Clear
End Sub
Private Sub OptionButton18_Click() 'LIPSA MPN AGILE
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveCell.Offset(0, 9).Range("A1") = "LIPSA MPN AGILE"
Sheets("Inregistrare parturi").Select
'part
Range("B15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("B8") 'PN from Report Register - sheet Register

'name
Range("G11").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Report").Range("D11") 'Name&Surname from Report Register - sheet Report

'quantity
Range("G15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("H8") 'Quantity from Report Register - sheet Register

'reason
Range("C19").Select
    ActiveCell.Range("A1") = "LIPSA MPN AGILE"
'status
Range("B23").Select
    ActiveCell.Range("A1") = "Se asteapta raspuns de la CE"
    
   CheckBox1 = True  'Electronice
   CheckBox2 = False 'Mecanice
   TextBox1 = Clear
End Sub
Private Sub OptionButton17_Click() 'MPN NECONFORM AGILE
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveCell.Offset(0, 9).Range("A1") = "MPN NECONFORM AGILE"
Sheets("Inregistrare parturi").Select
'part
Range("B15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("B8") 'PN from Report Register - sheet Register

'name
Range("G11").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Report").Range("D11") 'Name&Surname from Report Register - sheet Report

'quantity
Range("G15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("H8") 'Quantity from Report Register - sheet Register

'reason
Range("C19").Select
    ActiveCell.Range("A1") = "MPN NECONFORM AGILE"
'status
Range("B23").Select
    ActiveCell.Range("A1") = "Se asteapta raspuns de la CE"
        
   CheckBox1 = True  'Electronice
   CheckBox2 = False 'Mecanice
   TextBox1 = Clear
End Sub
Private Sub OptionButton20_Click() 'PRODUS DIFERIT DESEN AGILE
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveCell.Offset(0, 9).Range("A1") = "PRODUS DIFERIT DESEN AGILE"
Sheets("Inregistrare parturi").Select
'part
Range("B15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("B8") 'PN from Report Register - sheet Register

'name
Range("G11").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Report").Range("D11") 'Name&Surname from Report Register - sheet Report

'quantity
Range("G15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("H8") 'Quantity from Report Register - sheet Register

'reason
Range("C19").Select
    ActiveCell.Range("A1") = "PRODUS DIFERIT DESEN AGILE"
'status
Range("B23").Select
    ActiveCell.Range("A1") = "Se asteapta raspuns de la QE"
    
   CheckBox2 = True  'Mecanice
   CheckBox1 = False 'Electronice
   TextBox1 = Clear
End Sub

Private Sub OptionButton21_Click() 'Lipsa MPN Produs
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveCell.Offset(0, 9).Range("A1") = "LIPSA MPN PRODUS"
Sheets("Inregistrare parturi").Select
'part
Range("B15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("B8") 'PN from Report Register - sheet Register

'name
Range("G11").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Report").Range("D11") 'Name&Surname from Report Register - sheet Report

'quantity
Range("G15").Select
    ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("H8") 'Quantity from Report Register - sheet Register

'reason
Range("C19").Select
    ActiveCell.Range("A1") = "LIPSA MPN PRODUS"
'status
Range("B23").Select
    ActiveCell.Range("A1") = "Se asteapta raspuns de la QE"
        
   CheckBox1 = True  'Electronice
   CheckBox2 = False 'Mecanice
   TextBox1 = Clear
End Sub

Private Sub OptionButton22_Click() 'Alt Motiv

If Me.OptionButton22 = True Then
If Me.TextBox1 = vbNullString Then
    MsgBox "Te rog tasteaza ceva."
    OptionButton22 = False
    TextBox1.SetFocus
   Else: GoTo Ende
End If
Exit Sub
Ende:
End If

Sheets("PARTURI SUSPECTE INCOMING").Select
ActiveCell.Offset(0, 9).Range("A1") = TextBox1
Sheets("Inregistrare parturi").Select
'reason
Range("C19").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "General"
    ActiveCell.Range("A1") = TextBox1
'name
Range("G11").Select
ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Report").Range("D11")

'PN
Range("B15").Select
ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("B8")

'quantity
Range("G15").Select
ActiveCell.Formula = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("H8")
  
End Sub
Private Sub TextBox1_Change()
TextBox1.Text = UCase(TextBox1.Text)
End Sub
