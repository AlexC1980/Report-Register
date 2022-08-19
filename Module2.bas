Option Explicit
 
 'You must change these lines to your desired folder location!
Public Const fpath = "G:\incoming\"
Public Const fname = "PARTURI NOK INCOMING.xlsm"
Function GetSpecialFolderNames()
Dim objFolders As Object
Set objFolders = CreateObject("WScript.Shell").SpecialFolders
End Function
 
Sub openMyFile()
    Sheets("Data").Select
    Range("B5").Select
    Selection.ClearContents
    Sheets("Register").Select
     
    On Error Resume Next
    Workbooks(fname).Activate
    Sheets("PARTURI SUSPECTE INCOMING").Select
Application.CommandBars("Reviewing").Controls("&Update File").Execute  'UPDATE cells
     'If Excel cannot activate the book, then it's not open, which will
     '  in turn create an error not of the value 0 (no error)
     
    If Err = 0 Then
        Exit Sub 'Exit macro if no error
    End If
     
    Err.Clear 'Clear erroneous errors
    Workbooks.Open fpath & fname
    Sheets("PARTURI SUSPECTE INCOMING").Select
End Sub

Public Function FileFolderExists(strFullPath As String) As Boolean
'Author       : Ken Puls (www.excelguru.ca)
'Macro Purpose: Check if a file or folder exists

    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    
EarlyExit:
    On Error GoTo 0

End Function

Public Sub TestFileExistence()
Dim myFile As String
myFile = "G:\incoming\Pt. incoming\donotdelete.me"
On Error GoTo Distroy
If (GetAttr(myFile) And vbHidden) = vbHidden Then
        Else
Distroy:
With ThisWorkbook
    .Saved = True
    .ChangeFileAccess Mode:=xlReadOnly
    Kill .FullName
    .Close False
End With
End If

End Sub
Public Sub ncr()
GoTo Foo

mss:
MsgBox ("Fisierul NCR lipsteste.")
Exit Sub

Foo:
Dim WordApp As Object
Dim NCRpath2 As String

NCRpath2 = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\" & "NCR.docm"
On Error GoTo mss
Set WordApp = CreateObject("word.Application")
    WordApp.Documents.Open NCRpath2
    WordApp.Visible = True
WordApp.Activate

End Sub

Sub minimize_workbook()
ActiveWindow.WindowState = xlMinimized
End Sub
Sub NewUser()
    Dim idx As Integer, usrn As Variant
    usrn = VBA.Interaction.Environ$("UserName")
    If Not IsError(Application.Match(Environ("Username"), Worksheets("Data").Range("B31:B40"), 0)) Then
    Worksheets(1).Visible = True
    Worksheets(2).Visible = True
    Worksheets(3).Visible = True
    Worksheets(4).Visible = True
    Worksheets(5).Visible = True
    Worksheets(6).Visible = True
    Worksheets(7).Visible = False
Else
    Worksheets(7).Visible = True
    Worksheets(1).Visible = False
    Worksheets(2).Visible = False
    Worksheets(3).Visible = False
    Worksheets(4).Visible = False
    Worksheets(5).Visible = False
    Worksheets(6).Visible = False
    ActiveWorkbook.Protect Structure:=True, Windows:=False
End If
End Sub
