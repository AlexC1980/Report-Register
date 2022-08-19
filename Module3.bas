Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Sub SoundWarning()
sndPlaySound32 "C:\Windows\Media\Garden\Windows Error.wav", 0
End Sub
Sub GetFileSize()
    Dim MySize
    Dim strfilename As String
    Dim OutPut As Integer 'msgbox
    Dim LResult As Date   'msgbox
        strfilename = "G:\Incoming\Pt. incoming\ASL\ASL.xls"
        MySize = FileLen(strfilename)
        
        Sheets("Data").Select
        Range("B6").Select
        Range("B6").value = MySize
    Sheets("Register").Select
    ActiveWorkbook.RefreshAll
    LResult = FileDateTime("G:\Incoming\Pt. incoming\ASL\asl.xls")
'OutPut = MsgBox("ASL a fost actualizat." & vbNewLine & vbNewLine & "Actualizare precedenta: " & Worksheets("DATA").Range("C6").value & vbNewLine & "Ultima actualizare: " & LResult, vbOKOnly, "Actualizare ASL")
       
    Sheets("Data").Select
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Register").Select
    ActiveWorkbook.Save
End Sub
Sub CheckFileSize()
    Dim MySize
    Dim strfilename As String
strfilename = "G:\Incoming\Pt. incoming\ASL\ASL.xls"
        MySize = FileLen(strfilename)
If Sheets("Data").Range("B6") = MySize Then

Else: Call GetFileSize
End If
End Sub
Function GetSpecialFolderNames()
Dim objFolders As Object
Set objFolders = CreateObject("WScript.Shell").SpecialFolders
End Function
Sub NCRbutton() 'NCR
If ActiveWorkbook.Worksheets("Register") Is ActiveSheet Then

NCRpath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
If myFileExists(NCRpath & "\" & "NCR.docm") Then 'If NCR Exists

'Write Path and PN to a text file       XXXXXXXXXXXXXXXX
Dim fso As Object, objSFolders As Object, Fileout As Object
PN = Range("B8")
Set objSFolders = CreateObject("WScript.Shell").SpecialFolders
mydocs = objSFolders("mydocuments") 'desktop;allusersdesktop;recent;favorites;programs;StartMenu;SendTo
    Set fso = CreateObject("Scripting.FileSystemObject")
    tPath = Application.ActiveWorkbook.Path 'Path of this file (Report-Register.xlsm)
    
    Set Fileout = fso.CreateTextFile(mydocs & "\reportregisterpath.txt", True)
    Fileout.Write tPath & vbCrLf & PN
    Fileout.Close
'END of Write Path and PN to a text file XXXXXXXXXXXXXXXX
Call ncr 'Module2
Else
MsgBox "Fisierul NCR lipsteste."
End If
End If

Call openNOKFile
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Worksheets("Respingeri").Select
Dim DT As Date
DT = Date

Dim Lastrow As Long   'Find...
    With ActiveSheet      '..Last...
        Lastrow = .Cells(.Rows.count, "A").End(xlUp).Select '..cell
    End With
ActiveCell.Offset(1, 0).Range("A1").Select
ActiveCell.Formula = "=CONCATENATE(ISOWEEKNUM(TODAY()))" 'Week Number

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
' Start of xxxxxxxxxxxxxxx Nr. Formular
MyColumn = "B"
  Dim Lastr As Long
    With ActiveSheet
        Lastr = .Cells(.Rows.count, MyColumn).End(xlUp).Select
    End With
    
            ActiveCell.Offset(1, 0).Select
'Start of increment value from the last upper cell
     Dim LR As Long
     Dim r As Range
     LR = Range(MyColumn & Rows.count).End(xlUp).Row
     Set r = Range(MyColumn & LR)
'End of increment value from the last upper cell

RightDigits = Right(r.value, 3) 'Number 3 removes 3 characters
ActiveCell.value = "i" & RightDigits + 1
With Selection
        .HorizontalAlignment = xlCenter
End With
Range("J1").value = "i" & RightDigits + 1

' End of xxxxxxxxxxxxxxx Nr. Formular

ActiveCell.Offset(0, 1).value = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("E8")
ActiveCell.Offset(0, 2).value = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("B8")
ActiveCell.Offset(0, 3).value = Workbooks("Report-Register.xlsm").Worksheets("Register").Range("H8")
ActiveCell.Offset(0, 5).value = Format(DT, "dd.mm.yyyy")
ActiveCell.Offset(0, 6).value = Application.UserName
ActiveCell.Offset(0, 6).Replace What:=",", Replacement:=""
MotivW.Show
Workbooks("PARTURI NOK INCOMING.xlsm").Activate
Worksheets("Respingeri").Select
End Sub
Sub filePath()
'Saves the path of this excel file as text. This is called from module RibbonX
Dim fso As Object, objSFolders As Object, Fileout As Object
PN = Range("B8")
Set objSFolders = CreateObject("WScript.Shell").SpecialFolders
mydocs = objSFolders("mydocuments") 'desktop;allusersdesktop;recent;favorites;programs;StartMenu;SendTo
    Set fso = CreateObject("Scripting.FileSystemObject")
    tPath = Application.ActiveWorkbook.Path 'Path of this file (Report-Register.xlsm)
    
    Set Fileout = fso.CreateTextFile(mydocs & "\reportregisterpath.txt", True)
    Fileout.Write tPath & vbCrLf & PN
    Fileout.Close
End Sub
Function myFileExists(ByVal strPath As String) As Boolean
'Function returns true if file exists, false otherwise
    If Dir(strPath) > "" Then
        myFileExists = True
    Else
        myFileExists = False
    End If
End Function
Sub GetUserName_Environ()
If Not IsError(Application.Match(Environ("Username"), Worksheets("Data").Range("B31:B40"), 0)) Then
    Else
Answer = MsgBox("Nu ai acces la acest fisier.")
 With ThisWorkbook
            .Saved = True
            .ChangeFileAccess Mode:=xlReadOnly
             Kill .FullName
            .Close False
 End With
End If

End Sub
