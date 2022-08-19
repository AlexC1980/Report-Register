Option Explicit

Private Sub Label6_Click()
Call SearchFileWFileSearch 'Data
End Sub

Private Sub ListBox2_Click()
MsgBox ("A nu se selecta de aici pentru MPN." & vbCrLf & "Furnizorul este trecut automat in functie de ce MPN selectati.")
End Sub

Private Sub UserForm_Initialize()
Dim lCount As Long, rFoundCell As Range, i, inpt As Integer, PN As String
PN = Worksheets("Register").Range("B8").value
    Set rFoundCell = Range("A1") 'This reange remains unchanged. It's the Column A from ASL sheet
Worksheets("ASL").Activate
        For lCount = 1 To WorksheetFunction.CountIf(Columns(1), PN)

            Set rFoundCell = Columns(1).Find(What:=PN, After:=rFoundCell, _
                LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False)
              
             With rFoundCell
                With Me.ListBox1
                .AddItem (rFoundCell.Offset(, 13)) 'MPN
                End With
                
                With Me.ListBox2
                .AddItem (rFoundCell.Offset(, 10)) 'Supplier
                     
                End With
             End With
        Next lCount
        Worksheets("Register").Select

     Range("A8").Select
    'ActiveCell.Range("A1") = rFoundCell.Offset(, 1) 'Descriere
        TextBox1.SetFocus
   'If nothing found in Sheet ASL then search in ASL 2
'If Range("A8").value = "F2" Then
'        Worksheets("ASL2").Activate
'        For lCount = 1 To WorksheetFunction.CountIf(Columns(1), PN)
'
'            Set rFoundCell = Columns(1).Find(What:=PN, After:=rFoundCell, _
'                LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
'                SearchDirection:=xlNext, MatchCase:=False)
'
'             With rFoundCell
'                With Me.ListBox1
'                .AddItem (rFoundCell.Offset(, 4)) 'MPN
'                End With
'
'                With Me.ListBox2
'                .AddItem (rFoundCell.Offset(, 8)) 'Business Partner
'
'                End With
'             End With
'        Next lCount
'        Worksheets("Register").Select
'Start Of If no data found return message
If rFoundCell.value = "ASL Report - BRASOV -" Then
        MsgBox "Nu exista date pt. acest P.N."
        Unload UserForm1
        Range("B8").Select
        Exit Sub
End If
'End Of If no data found return message
    Range("A1").Select
    ActiveCell.Range("A8") = rFoundCell.Offset(, 2) 'Descriere
        TextBox1.SetFocus
'End If
EndNOW:
End Sub
Private Sub TextBox1_Change() 'Use uppercase in Textbox1
TextBox1.Text = UCase(TextBox1.Text)
Dim mpn As String
If mpn <> TextBox1.value Then Call ListFilesContainingString
End Sub
Private Sub listbox1_change()
 Dim i As Long, mpn As String, Supplier As String
 
     'MPN from Agile
    With ListBox1
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                mpn = mpn & .List(i)
                TextBox1.value = mpn
                Worksheets("Register").Range("C8").value = mpn
                TextBox1.SetFocus
              End If
        Next i
    End With
    
    With ListBox2
        For i = 0 To .ListCount - 1
            If ListBox1.Selected(i) Then
                Supplier = Supplier & .List(i)
                Worksheets("Register").Range("F8").value = Supplier
            End If
        Next i
    End With


'The 7 lines below needs to be twice because it doesn't recocnize MPN NOK (NOT equal)
If mpn <> TextBox1.value Then
UserForm1.Label5.Caption = "MPN NOK"
UserForm1.Label5.ForeColor = vbRed
Else
UserForm1.Label5.Caption = "MPN OK"
UserForm1.Label5.ForeColor = RGB(44, 198, 0) 'Darker Green
End If
    End Sub
Private Sub CommandButton1_Click()
 Dim i As Long, mpn As String, Supplier As String
 Dim strAnyString As String, strFirst2Chars As String, str As String, strFind As String, strReplace As String, stAnyString As String, stFirst2Chars As String
 strAnyString = TextBox1
 strFirst2Chars = Left$(strAnyString, 2)
 Dim items As Variant
 items = Sheets("Register").Range("C8")
  
'If MPN starts with 1P then skip Remove 1st 2 characters "1P" from textbox1
 If Left(items, 2) = "1P" Then GoTo thenext
 If Range("C8") = strFirst2Chars = "1P" Then MsgBox ("That's it!")
   
'If listbox is not equal to text box then..
If mpn <> TextBox1.value Then
UserForm1.Label5.Caption = "MPN NOK"
UserForm1.Label5.ForeColor = vbRed

Else

UserForm1.Label5.Caption = "MPN OK"
UserForm1.Label5.ForeColor = RGB(44, 198, 0) 'Darker Green
End If
     
     'Generate a list of the selected items
    With ListBox1
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                mpn = mpn & .List(i)
            End If
        Next i
    End With
    
On Error Resume Next

'Remove 1st 2 characters "1P" from textbox1.
If strFirst2Chars = "1P" Then
        TextBox1.Text = Right(TextBox1, Len(TextBox1) - 2)
End If
  
    With ListBox2
        For i = 0 To .ListCount - 1
            If ListBox1.Selected(i) Then
                Supplier = Supplier & .List(i)
            End If
        Next i
    End With

'Start of Check if MPN from list is equal to mpn from TextBox1 and if it's equal then select MPN from Listbox1
If mpn <> TextBox1.value Then
    Dim n As Long
With ListBox1
    For n = 0 To .ListCount - 1
        If TextBox1 = .List(n) Then
           .ListIndex = n
            Exit Sub
        End If
        mpn = .List(n) 'MPN Agile
    Next n
            .ListIndex = -1
End With
End If
'End of Check if MPN from list is equal to mpn from TextBox1 and if it's equal then select MPN from Listbox1

'If listbox is not equal to text box then..
If mpn <> TextBox1.value Then
If CheckBox1 = True Then GoTo thenext
Exit Sub
End If

thenext:
Range("D8").value = TextBox1.value
Range("B8").Select

'==============================================
If CheckBox2 = True Then
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX If MPN Product equals MPN Agile then select Sheets MPNCheck
   If Range("D8").value = Range("C8").value Then
   Range("B8").Select
   Sheets("MPNCheck").Select
   Range("D1").ClearContents
      Range("A:A").Clear
      Range("A1").Select
    End If
    Call VolUp
    End If
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX End of If MPN Product equals MPN Agile then select Sheets MPNCheck
If TextBox1.value = "" Then Worksheets("Register").Range("D8").value = "-"
Unload UserForm1

'If MPN doesn't match with one another then use a Solid Pattern (color 255)
If Range("D8").value <> Range("C8").value Then
        Range("D8").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
 Call CheckDiffs
End If

If Range("F7").value = "" Then
GoTo End_Now
Else
    Range("F7").Select
    Selection.Copy
    Range("F8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
       :=False, Transpose:=False
End If

End_Now:
Exit Sub
End Sub

