Private Sub Workbook_Open()
'Call GetUserName_Environ 'If username not in Sheet Data then close (kill) workbook
Application.WindowState = xlMaximized

Application.OnKey "%w", "minimize_workbook"
    
Dim usrname As Variant
usrname = Environ("Username")
If usrname = "culeaa" Then
    Application.CommandBars("Ply").Enabled = True
Else
    Application.CommandBars("Ply").Enabled = False 'Disable right click on sheet tab in Excel
    Application.OnKey "%{F11}", "NoVBE"
End If
    
'Call Get_User_Name
'Call filePath
'Call copyNCR
'Call ReadFile 'Write file is in File NOK
sHostName = Environ$("computername")
Application.StatusBar = sHostName & " | " & VBA.Interaction.Environ$("UserName")
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.ThisWorkbook.Saved = True 'Tells Excel that the file has already been saved (this prevents..
    '...Excel from requesting that you save the file when you close it)
End Sub
Function GetSpecialFolderNames()
Dim objFolders As Object
Set objFolders = CreateObject("WScript.Shell").SpecialFolders
End Function
Sub copyNCR()
Dim oFSO As Object
Set oFSO = CreateObject("Scripting.FileSystemObject")
mydocs = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\"
Call oFSO.copyfile("G:\Incoming\Pt. incoming\NCR.docm", mydocs, True)
End Sub
