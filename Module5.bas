Option Explicit

Public Const fpath = "G:\incoming\"
Public Const fname = "PARTURI NOK INCOMING.xlsm"
Function GetSpecialFolderNames()
Dim objFolders As Object
Set objFolders = CreateObject("WScript.Shell").SpecialFolders
End Function
 
Sub openNOKFile()
    On Error Resume Next
    Workbooks(fname).Activate
    Sheets("PARTURI SUSPECTE INCOMING").Select
Application.CommandBars("Reviewing").Controls("&Update File").Execute  'UPDATE cells
    If Err = 0 Then
        Exit Sub 'Exit macro if no error
    End If
     
    Err.Clear 'Clear erroneous errors
    Workbooks.Open fpath & fname
End Sub
