Option Explicit
 
 'You must change these lines to your desired folder location!
Public Const fpath = "G:\incoming\"
Public Const fname = "PARTURI NOK INCOMING.xlsm"
 
Sub NOK(control As IRibbonControl)
     
    On Error Resume Next
    Workbooks(fname).Activate
    Sheets("PARTURI SUSPECTE INCOMING").Select
     
     'If Excel cannot activate the book, then it's not open, which will
     '  in turn create an error not of the value 0 (no error)
     
    If Err = 0 Then
        Exit Sub 'Exit macro if no error
    End If
     
    Err.Clear 'Clear erroneous errors
    Workbooks.Open fpath & fname
    Sheets("PARTURI SUSPECTE INCOMING").Select
End Sub

Sub MPNCheck(control As IRibbonControl)
    If ActiveWorkbook.Worksheets("MPNCheck") Is ActiveSheet Then
MsgBox ("Butonul MPNCheck nu este activ decat in foaia Register.")
Else
    If ActiveWorkbook.Worksheets("Register") Is ActiveSheet Then
    Sheets("MPNCheck").Select
Application.StatusBar = False
Range("D1").ClearContents

    Range("A:A").Clear
    Range("A1").Select
    Call VolUp
End If
End If
End Sub

Sub Save_Report_Register(control As IRibbonControl)
Dim MyName As String, Path As String
 With ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule 'Delete startup code

            .DeleteLines 1, .CountOfLines

 End With
Application.DisplayAlerts = False
 Sheets("ASL").Select
    ActiveWindow.SelectedSheets.Delete
 Sheets("Data").Select
 Range("B5").Select
 Selection.ClearContents
 Sheets("Register").Select
Application.DisplayAlerts = True
    MyName = Worksheets("Register").Range("B8")
    Path = Worksheets("Data").Range("B4")
    ActiveWorkbook.SaveAs FileName:=Path & MyName & ".xlsm"
    Range("A8").Select 'This is for making the workbook active because it becomes inactive, dunno why

End Sub
Sub calc_cells(control As IRibbonControl)
Worksheets("MPNCheck").Activate

'Looking for VBA to check if all data in a column is the same (text)
    Dim myRange As Range
    Dim myValue
    Dim allSame As Boolean
    
'   Set column to check
    Set myRange = Range("A:A")
    
'   Get first value from myRange
    myValue = myRange(1, 1).value
    
    allSame = (WorksheetFunction.CountA(myRange) = WorksheetFunction.CountIf(myRange, myValue))
    
'   Return whether or not they are all the same (TRUE/FALSE)
    If allSame = False Then
        MsgBox ("Unul sau mai multe MPN-uri gresite.")
        Exit Sub
    End If
'End of Looking for VBA to check if all data in a column is the same (text)

Call VolMute
'
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlUnique
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
  'Select column A to the last cell
    Dim rRange1 As Range, rRange2 As Range
     
    Set rRange1 = Range(Cells(1, "A"), Cells(Rows.count, "A").End(xlUp))
    Set rRange2 = Range(Cells(1, "A"), Cells(Rows.count, "A").End(xlUp))

     
    Application.GoTo Union(rRange1, rRange2)
    
'Start of Count selected cells
    Dim cell As Object
    Dim count As Integer
    count = 0
    For Each cell In Selection
        count = count + 1
    Next cell
'End of Count selected cells

  Dim Q As Variant
Q = Application.InputBox _
             (Prompt:="Introdu cantitatea de pe rolã.", _
                    Title:="Cantitate", Type:=1) 'Value 1 is for number, 2 is for text (that's InputBox Method)
        If Q = 0 Then
    End If
    Worksheets("Register").Activate
    Worksheets("Register").Range("H8") = count * Q
    
End Sub
Sub update(control As IRibbonControl)
If ActiveWorkbook.Worksheets("Register") Is ActiveSheet Then

On Error GoTo Handler:
Dim argh As Double
argh = Shell("\\BRR-FS03\groups\public\Culea Alex\Report-Register\tools\UPDATE Report-Register.bat", vbNormalFocus)
ActiveWorkbook.Close
Handler:
    MsgBox "Error " & Err.Number & ". " & Err.Description
    Exit Sub
 Else
 MsgBox ("Butonul Update nu este activ decat in foaia Register.")
End If
End Sub
'Start of If file NOK is opened don't open again.
Function FileInUse(sFileName As String) As Boolean
    On Error Resume Next
    Open sFileName For Binary Access Read Lock Read As #1
    Close #1
    FileInUse = IIf(Err.Number > 0, True, False)
    On Error GoTo 0
End Function
Sub NCRf(control As IRibbonControl)
If ActiveWorkbook.Worksheets("Register") Is ActiveSheet Then
    Dim sPath As String
    Dim sFileName As String

        'change as required
    sFileName = "NCR.docm"
    
    If FileInUse(sFileName) Then
        'read / write file in use
    Exit Sub
    End If
'End of If file NOK is opened don't open again.
Call NCRbutton
Else
MsgBox ("Butonul NCR File nu este activ decat in foaia Register.")
End If
End Sub
Sub About_R(control As IRibbonControl) 'Dialog Launcher
About.Show
End Sub
Sub Created(control As IRibbonControl)
About.Show
End Sub
