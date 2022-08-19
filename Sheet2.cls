Private Sub CheckBox1_Click()

If CheckBox1.value = False Then
 Sheets("Report").Range("G18").FormulaR1C1 = ""
 Sheets("Register").CheckBox1.value = False
 End If

If CheckBox1.value = True Then
    Q = Application.InputBox _
             (Prompt:="Introdu Asset Nr. Subler", _
                    Title:="Asset Nr#", Type:=2) 'Default:="AM0024",
    
 If Q = 0 Then
 Sheets("Register").CheckBox1.value = False
 Sheets("Report").Range("G18").FormulaR1C1 = ""
 GoTo EndNOW
 End If
    
    Sheets("Report").Range("G18").FormulaR1C1 = "Subler/" + Q
    Sheets("Register").Select
 
 Range("G8").Select
EndNOW:
End If
End Sub

Private Sub CheckBox10_Click() 'Penseta

If CheckBox10.value = False Then
 Sheets("Report").Range("J19").FormulaR1C1 = ""
 Sheets("Register").CheckBox10.value = False
 End If

If CheckBox10.value = True Then
    Q = Application.InputBox _
             (Prompt:="Introdu Asset Nr. Penseta", _
                    Title:="Asset Nr#", Type:=2) 'Default:="AM0024",
    
 If Q = 0 Then
 Sheets("Register").CheckBox10.value = False
 Sheets("Report").Range("J19").FormulaR1C1 = ""
 GoTo EndNOW
 End If
    
    Sheets("Report").Range("J19").FormulaR1C1 = "Penseta/" + Q
    Sheets("Register").Select
 
 Range("N8").Select
EndNOW:
End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.value = True Then
    Sheets("Report").Range("H18").FormulaR1C1 = "Multimetru"
    Sheets("Register").Select
 End If
 If CheckBox2.value = False Then
 Sheets("Report").Range("H18").FormulaR1C1 = ""
 End If
  Range("N8").Select
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.value = True Then
    Sheets("Report").Range("K18").FormulaR1C1 = "Micrometru"
    Sheets("Register").Select
 End If
 If CheckBox3.value = False Then
 Sheets("Report").Range("K18").FormulaR1C1 = ""
 End If
  Range("A8").Select
End Sub

Private Sub CheckBox4_Click()
If CheckBox4.value = True Then
    Sheets("Report").Range("J18").FormulaR1C1 = "Rigla"
    Sheets("Register").Select
 End If
 If CheckBox4.value = False Then
 Sheets("Report").Range("J18").FormulaR1C1 = ""
 End If
 Range("A8").Select
End Sub

Private Sub CheckBox5_Click()
If CheckBox5.value = True Then
    Sheets("Report").Range("I18").FormulaR1C1 = "Ruleta"
    Sheets("Register").Select
 End If
 If CheckBox5.value = False Then
 Sheets("Report").Range("I18").FormulaR1C1 = ""
 End If
 Range("A8").Select
End Sub

Private Sub CheckBox6_Click()
If CheckBox6.value = True Then
    Sheets("Report").Range("L18").FormulaR1C1 = "Durimetru"
    Sheets("Register").Select
 End If
 If CheckBox6.value = False Then
 Sheets("Report").Range("L18").FormulaR1C1 = ""
 End If
 Range("A8").Select
End Sub

Private Sub CheckBox7_Click()
If CheckBox7.value = True Then
    Sheets("Report").Range("G19").FormulaR1C1 = "Pini"
    Sheets("Register").Select
 End If
 If CheckBox7.value = False Then
 Sheets("Report").Range("G19").FormulaR1C1 = ""
 End If
 Range("A8").Select
End Sub

Private Sub CheckBox8_Click()
If CheckBox8.value = True Then
    Sheets("Report").Range("H19").FormulaR1C1 = "Lere"
    Sheets("Register").Select
 End If
 If CheckBox8.value = False Then
 Sheets("Report").Range("H19").FormulaR1C1 = ""
 End If
 Range("A8").Select
End Sub

Private Sub CheckBox9_Click()
If CheckBox9.value = True Then
    Sheets("Report").Range("I19").FormulaR1C1 = "Etalon"
    Sheets("Register").Select
 End If
 If CheckBox9.value = False Then
 Sheets("Report").Range("I19").FormulaR1C1 = ""
 End If
 Range("A8").Select
End Sub

Function myFileExists(ByVal strPath As String) As Boolean
'Function returns true if file exists, false otherwise
    If Dir(strPath) > "" Then
        myFileExists = True
    Else
        myFileExists = False
    End If
End Function

Private Sub CommandButton3_Click() 'AQL
    Call AQL.AQL 'AQL module
End Sub

Private Sub CommandButton6_Click() 'NOK Button

If Range("G8").value = "" Then
MsgBox "Lipseste P.O.-ul"
Range("G8").Select
Exit Sub
End If

If Range("H8").value = "" Then
MsgBox "Lipseste Cantitatea"
Range("H8").Select
Exit Sub
End If

ActiveCell.Select
    Call openMyFile 'Module2
    Workbooks("PARTURI NOK INCOMING.xlsm").Activate
    Dim Lastrow As Long   'Find...
    With ActiveSheet      '..Last...
        Lastrow = .Cells(.Rows.count, "B").End(xlUp).Select '..empty cell
    End With
    
 'Deletes entire line so that won't appear an Error like in this example: ='[Report-Register.xlsm]
    ActiveCell.Rows("2:2").EntireRow.Delete
    With ActiveSheet      '..Last...
        Lastrow = .Cells(.Rows.count, "B").End(xlUp).Select '..empty cell
    End With
    
  'Number
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.FormulaR1C1 = "=SUM(R[-1]C)+1"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
           
  'PN
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.Formula = "='[Report-Register.xlsm]Register'!$B$8"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
  'MPN
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Formula = "='[Report-Register.xlsm]Register'!$C$8"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
  'MPN PRODUCT
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Formula = "='[Report-Register.xlsm]Register'!$D$8"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
  
  'Supplier
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.Formula = "='[Report-Register.xlsm]Register'!$F$8"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        
  'Date
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.NumberFormat = "[$-409]d-mmm-yy;@"
    ActiveCell.FormulaR1C1 = "=TODAY()"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
   
  'PO
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.Formula = "='[Report-Register.xlsm]Register'!$G$8"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        
  'Quantity
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.Formula = "='[Report-Register.xlsm]Register'!$H$8"
    Selection.HorizontalAlignment = xlRight
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        
          
  'No
   ActiveCell.Offset(0, 2).Range("A1").Select
   ActiveCell.FormulaR1C1 = "No"
   
  'User names initials
   ActiveCell.Offset(0, 16).Range("A1").Select
   ActiveCell.Formula = "='[Report-Register.xlsm]Register'!$K$8"
   Selection.Copy
   Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        
  'WeekNumber
    ActiveCell.Offset(0, -15).Range("A1").Select
    Selection.HorizontalAlignment = xlCenter
    ActiveCell.Formula = "='[Report-Register.xlsm]Data'!$J$2"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    
ChooseLine.Show 'Call the UserForm
End Sub
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    Dim myrng As Range
     
    Set myrng = Range("A8:M8:G9:H10")
    Range("A8:M8").Hyperlinks.Delete  'Remove Hyperlinks within the range
     
    'Target Is an argument which refers To the range being changed
    If Not Intersect(Target, myrng) Is Nothing Then
        'this check Is To prevent the procedure from repeatedly calling itself
        If Not (Target.Text = UCase(Target.Text)) Then
            Target = UCase(Target.Text)
        End If
    End If
    
    With ActiveSheet.Range("A8:M8")
        .Font.Size = 10             'Make font size 10 in cells A8 to M8
        .Font.Name = "Arial"
        .Font.Strikethrough = False
        .Font.Superscript = False
        .Font.Subscript = False
        .Font.OutlineFont = False
        .Font.Shadow = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.ColorIndex = xlAutomatic
        .Font.Bold = False
        .Font.Italic = False
        .Rows.AutoFit
    End With
    Range("H8:M8").HorizontalAlignment = xlCenter  'Align to center cells H8 to M8
    Range("H8:M8").HorizontalAlignment = xlCenter  'Align to center cells H8 to M8
    Range("C8:D8").NumberFormat = "@"
'XX Rohs + Pass/Fail
    
    If Range("L8").Text = "Y" Then
    Sheet1.CheckBox1.value = True  'Rohs Agile Yes
    Sheet1.CheckBox2.value = False 'Rohs Agile No
    Else
    If Range("L8").Text = "N" Then
    Sheet1.CheckBox1.value = False
    Sheet1.CheckBox2.value = True
    End If
    End If

    If Range("M8").Text = "Y" Then
    Sheet1.CheckBox3.value = True  'Rohs Produs Yes
    Sheet1.CheckBox4.value = False 'Rohs Produs No
    Else
    If Range("M8").Text = "N" Then
    Sheet1.CheckBox3.value = False
    Sheet1.CheckBox4.value = True
    End If
    End If

    If Range("I8").Text = "X" Then
    Sheet1.CheckBox5.value = True  'Pass
    Sheet1.CheckBox6.value = False 'Fail
    
    Else
    If Range("I8").Text = "" Then
    Sheet1.CheckBox5.value = False
    Sheet1.CheckBox6.value = True
    End If
    End If
'XX end of Rohs + Pass/Fail XXX

End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

If Not Intersect(Target, Range("G11")) Is Nothing Then Range("H8").Select
If Not Intersect(Target, Range("H9")) Is Nothing Then Range("L8").Select
If Not Intersect(Target, Range("L9")) Is Nothing Then Range("G8").Select
If Not Intersect(Target, Range("M7")) Is Nothing Then Run "value" 'Data Module
'If Not Intersect(Target, Range("G9")) Is Nothing Then Range("H8").Select
If Not Intersect(Target, Range("I8")) Is Nothing Then Range("B8").Select
 
Set InRange2 = Intersect(Target, Range("B9"))
    Range("C8:D8").Interior.Pattern = xlNone
If IsEmpty(Range("B8").value) = True Then
  GoTo nothing_1
  Else
  If Not InRange2 Is Nothing Then
      Run "InsertTime" 'Data
      Run "Start" 'Data module
      Range("A8:M8").ClearComments
  End If
nothing_1:
End If
End Sub
