Sub printsheet1()
'
' printsheet1 Macro
' Macro recorded 05.03.2010 by culeaa
' Keyboard Shortcut: Ctrl+q

'XXXXX Check if Register is opened XXXXX
' Register = Worksheets("Data").Range("B2") 'Path is in WorkBook "Report-Register", Sheet "Data", cell B2
'    Dim TestWorkbook As Workbook
'    Set TestWorkbook = Nothing
'    On Error Resume Next
'    Set TestWorkbook = Workbooks(Register)
'    On Error GoTo 0

'    If TestWorkbook Is Nothing Then
'        MsgBox "Verifica denumirea, extensia si luna curenta ca Foaie(Sheet) in registru.", vbOKOnly + vbCritical, "Registrul nu este deschis"
'     Exit Sub
'    End If
'XX End of Check if Register is opened XX


Application.DisplayAlerts = True


   Sheets("Register").Select
    Range("A8:M8").Select
    Selection.Interior.ColorIndex = xlNone 'this is for filled color false
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
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
    End With
    
    MyName = Worksheets("Register").Range("B8") & Format(Date, " dd-mm-yy") & Format(Time, "_hh.mm")
    Path = Worksheets("Data").Range("B3")
    Sheets(Array("Report", "Register")).Copy 'Create new Workbook and.....
     
'End of Record time for making a PN 'See module Data for Start Record time for making a PN
 ActiveWorkbook.Comments = "Time spent working on this P.N.: " & Range("A1") - TimeValue(Now)
 Range("A1").Clear

    
 'xxxxxxxxx  Format to Text Ranges in Sheet Report
 For Each cell In Range("C8:D9,E8,I8,M8,P8,D13,H12,D15,D17,D19,D21,G16,G17,G22:I22,O28")
 If cell.Errors.Item(xlNumberAsText).value = False Then cell.Formula = cell.Text
 Next
 'xxxxxxxxx
   Sheets("Register").Select
   Range("B1").Select
   ActiveCell.FormulaR1C1 = "" 'Delete week nr. from Register sheet
   Range("A8:M8").Select
   Sheets("Report").Select
  Application.DisplayAlerts = False
 
 'ActiveSheet.PrintOut 'Print sheet Report
   
   ActiveWorkbook.SaveAs FileName:=Path & MyName & ".xlsx" '.....save to Network
   ActiveWorkbook.Close 'Close new workbook
   
  
    Sheets("Register").Select
    Range("B8").Select
'   Range("A8:N8").Select
'    Range("N8").Activate
'    Selection.Copy
     
' GoTo "Register" and Paste the content copied from Sheet1 (Report)

'xxxxxxxxxxxxxxxxxxxReisterxxxxxxxxxxxxxxx
'On Error GoTo ErrHandler: 'If no current month then create it
'   Dim LValue As String
'   mm = Range("B2")
'   LValue = MonthName(mm, False)
'    MyPath = Worksheets("Data").Range("B2") 'Path is in WorkBook "Report-Register", Sheet "Data", cell B2
'    Workbooks(MyPath).Activate
'    Sheets(LValue).Select 'Detect the actual Month
    
'ErrHandler:
'    If Err.Number = 9 Then
'       Worksheets.Add(After:=Worksheets(Worksheets.count)).Name = LValue
'    Resume
'    End If
 
   'Application.CommandBars("Reviewing").Controls("&Update File").Execute  'UPDATE cells
    
'    If IsEmpty(ActiveCell) Then GoTo EXIT_SUB 'If the cell is empty go to exit_sub
'         Do ' Begin loop
'Find the last used row in column B
'    Dim Lastrow As Long
'    With ActiveSheet
'        Lastrow = .Cells(.Rows.count, "B").End(xlUp).Select
'    End With
    
'            ActiveCell.Offset(2, -1).Select ' Steps down one row to the next cell.

            ' Test contents of active cell; if empty, exit loop
            ' or Loop While Not IsEmpty(ActiveCell).

'         Loop Until IsEmpty(ActiveCell)
'EXIT_SUB:
'   If Not (IsEmpty((Range("A:N")))) Then
'    ActiveCell.Offset(0, 1).EntireRow.Range("B1").Select
'    ActiveSheet.Paste
'    ActiveCell.Offset(0, -1).Select
'      ActiveCell.FormulaR1C1 = "=TODAY()"
'    ActiveCell.Select
'    Selection.Copy
'    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
'        xlNone, SkipBlanks:=False, Transpose:=False
'    Selection.HorizontalAlignment = xlLeft

'   End If
   
    'Move 1 Cell(s) Down
'   ActiveCell.Offset(1, 0).Select
'     Do While Not IsEmpty(ActiveCell)
'    ActiveCell.Offset(1, 0).Select
'    Loop
'Application.CommandBars("Reviewing").Controls("&Update File").Execute  'UPDATE cells
'    ActiveWorkbook.Save
    
'End_Sub:
    End Sub
