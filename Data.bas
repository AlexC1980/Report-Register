Public fso As New FileSystemObject
Sub ReadFile()
Dim txtstr As TextStream
Dim FileName As String
Dim File As File
FileName = "G:\Incoming\Pt. incoming\Report-Register\tools\Numar Raport NCR.txt"

If fso.FileExists(FileName) Then
  Set File = fso.GetFile(FileName)
  Set txtstr = File.OpenAsTextStream(ForReading, TristateUseDefault)
  'Worksheets("Data").Range("B4").value = txtstr.ReadAll
  txtstr.Close
End If
End Sub
Sub Start()

Application.ScreenUpdating = False
' Save/Record PN for later. In case Report isn't saved you r asked if u wanna scann again and if you answer no then PN from B7 (Recorded PN) will be copied to B8
    Sheets("Register").Select
    Range("B8").Select
    Selection.Copy
    Range("B7").Select
    ActiveSheet.Paste
   With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10040115
        .TintAndShade = 0
        .PatternTintAndShade = 0
   End With
   With Selection.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -6737101
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
   End With
'End of Save/Record PN

'XXXX Start of clear data XXXX
   
    Sheets("Register").Select 'Clear data from Data Sheet
    Range("G8").Select
    Selection.ClearContents
    Range("D8").Select
    Selection.ClearContents
    Range("C8").Select
    Selection.ClearContents
    Range("A8").Select
    Selection.ClearContents
    Range("G10").Select
    ActiveCell.FormulaR1C1 = "-" 'Clear "Date Code" Cell
    Range("N8").Select
    Selection.ClearContents
    Range("B8").Select

'Clear data from Report Sheet
    Sheets("Report").Select
    Range("D21").Select
    ActiveCell.Formula = "='[Report-Register.xlsm]Register'!$H$8"
    Range("G22").Select
    ActiveCell.FormulaR1C1 = "100%"
    Range("H22").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("I22").Select
    ActiveCell.FormulaR1C1 = "-"
    Range("C29:M53").Select
    Selection.ClearContents
    ActiveSheet.Shapes("TextBox 1").Select
    ActiveWindow.SmallScroll Down:=14
    Selection.Characters.Text = ""
    Application.GoTo Sheets("Report").Range("A1"), True
    Sheets("Register").Select

'Remove any color Pattern
    Range("A8:M8").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

'Uncheck boxes from Register Sheet
Sheets("Register").Select
Sheets("Register").CheckBox1.value = False
Sheets("Register").CheckBox2.value = False
Sheets("Register").CheckBox3.value = False
Sheets("Register").CheckBox4.value = False
Sheets("Register").CheckBox5.value = False
Sheets("Register").CheckBox6.value = False
Sheets("Register").CheckBox7.value = False
Sheets("Register").CheckBox8.value = False
Sheets("Register").CheckBox9.value = False
Sheets("Register").CheckBox10.value = False
'XXXX End of clear data XXXX

'XXXX Start of detect Project XXXX
Dim stringToFind As String
Dim Found As Range
Dim strAnyString As String
Dim strFirst3Chars As String
project = Range("B8")
strAnyString = project
strFirst3Chars = Left$(strAnyString, 3)

If strFirst3Chars = "ABI" Then
Range("E8") = "ABB JAGUAR"
End If

If strFirst3Chars = "EMR" Then
Range("E8") = "EMERSON SPECTRONIX"
End If

If strFirst3Chars = "BAR" Then
Range("E8") = "BARCO"
End If

If strFirst3Chars = "BEI" Then
Range("E8") = "Dedicated Computing"
End If

If strFirst3Chars = "CIN" Then
Range("E8") = "CINIONIC"
End If

If strFirst3Chars = "FLU" Then
Range("E8") = "FLUKE"
End If

If strFirst3Chars = "GEH" Then
Range("E8") = "GE"
End If

If strFirst3Chars = "MRL" Then
Range("E8") = "MAREL"
End If

If strFirst3Chars = "KCI" Then
Range("E8") = "KCI"
End If

If strFirst3Chars = "PHF" Then
Range("E8") = "PARKER"
End If

If strFirst3Chars = "PHI" Then
Range("E8") = "PHILIPS"
End If

If strFirst3Chars = "PRH" Then
Range("E8") = "PREH"
End If

If strFirst3Chars = "ROS" Then
Range("E8") = "ROSEMOUNT"
End If

If strFirst3Chars = "RAT" Then
Range("E8") = "RATIONAL"
End If

If strFirst3Chars = "RNS" Then
Range("E8") = "Rohde & Schwarz"
End If

If strFirst3Chars = "SET" Then
Range("E8") = "SETRA"
End If

If strFirst3Chars = "SMS" Then
Range("E8") = "SIEMENS"
End If

If strFirst3Chars = "WEI" Then
Range("E8") = "LOWENSTEIN MEDICAL"
End If
'XXXX End of detect Project XXXX

Worksheets("ASL").Activate
On Error GoTo ErrHandler:
      UserForm1.Show
      Worksheets("Register").Range("B8").Select
Exit Sub
ErrHandler:
End Sub
Sub value()
    Sheets("Data").Select
    Range("K2").Select
End Sub
Sub value2()
    Range("N2").Select
    Selection.Copy
    Sheets("Register").Select
    Range("N8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub
Sub InsertTime()
    Sheets("Data").Select
    Range("J1").Select
    ActiveCell.FormulaR1C1 = TimeValue(Now)
    Range("I1").Select
    Sheets("Register").Select
End Sub
Sub NoVBE()
    MsgBox ("Can't access VBA via ALT+F11.")
End Sub
Sub ListFilesContainingString()
'Identify Preh Project and skip finding email because there is a problem, look below "characters not accepted by Microsoft".
Dim strAnyString As String
Dim strFirst3Chars As String
project = Worksheets("Register").Range("B8")
strAnyString = project
strFirst3Chars = Left$(strAnyString, 3)
If strFirst3Chars = "PRH" Then
    GoTo End1
End If
' End Of Identify Preh Project

    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    emailPath = Worksheets("Data").Range("B7")
NextCode:
    getfolder = (emailPath)
    Set fldr = Nothing
   
    wrd = Range("B8")
      If wrd = "" Then
        MsgBox "???"
        Exit Sub
    End If
   
    strFile = Dir(getfolder & "\*" & wrd & "*") 'There is a problem when a PN has characters not accepted by Microsoft.
    fc = 0
    Do While Len(strFile) > 0
        fc = fc + 1
        strFile = Dir
   Loop
    If fc > 0 Then UserForm1.Label6.Caption = "Exista E-mail"
End1:
End Sub
Sub SearchFileWFileSearch()
Path = Worksheets("Data").Range("B7")
SearchPath = "search-ms:displayname=" & d & "%20E-mails&crumb=&crumb=name%3A~<" & Range("B8") & "%20OR%20System.Generic.String%3A" & Range("B8")
searchlocation = "&crumb=location:" & Path
Call Shell("explorer.exe """ & SearchPath & d & searchlocation & "", 1)
End Sub
