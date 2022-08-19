'Get windows user name

' Makes sure all variables are dimensioned in each subroutine.
     Option Explicit

     ' Access the GetUserNameA function in advapi32.dll and
     ' call the function GetUserName.
     Declare PtrSafe Function GetUsername Lib "advapi32.dll" Alias "GetUserNameA" _
     (ByVal lpBuffer As String, nSize As Long) As Long
Function GetFirstLetters(rng As Range) As String 'This is for getting the initials from The Full name of a username form Sheet "DATA"
'Update 20140325
    Dim arr
    Dim i As Long
    arr = VBA.Split(rng, " ")
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            GetFirstLetters = GetFirstLetters & Left(arr(i), 1)
        Next i
    Else
        GetFirstLetters = Left(arr, 1)
    End If
End Function
'Main routine to Dimension variables, retrieve user name and display answer.
     Sub Get_User_Name()
On Error Resume Next
     'Dimension variables
     Dim lpBuff As String * 25
     Dim ret As Long, UserName As String

     ' Get the user name minus any trailing spaces found in the name.
     ret = GetUsername(lpBuff, 25)
     UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
     Dim usern1 As Range, usern2 As Range, usern3 As Range, usern4 As Range, usern5 As Range, usern6 As Range, usern7 As Range, usern8 As Range, usern9 As Range, usern10 As Range
       
     Set usern1 = Sheets("DATA").Range("B31")
     Set usern2 = Sheets("DATA").Range("B32")
     Set usern3 = Sheets("DATA").Range("B33")
     Set usern4 = Sheets("DATA").Range("B34")
     Set usern5 = Sheets("DATA").Range("B35")
     Set usern6 = Sheets("DATA").Range("B36")
     Set usern7 = Sheets("DATA").Range("B37")
     Set usern8 = Sheets("DATA").Range("B38")
     Set usern9 = Sheets("DATA").Range("B39")
     Set usern10 = Sheets("DATA").Range("B40")

      If UserName = usern1 Then
      Worksheets("Report").Range("D11").FormulaR1C1 = usern1.Offset(0, 1).Range("A1")
      Sheets("Register").Range("K8").FormulaR1C1 = GetFirstLetters(usern1.Offset(0, 1).Range("A1"))
      Sheets("Register").Select
      Range("B8").Select
      ActiveWorkbook.Save
      End If
      
      If UserName = usern2 Then
      Worksheets("Report").Range("D11").FormulaR1C1 = usern2.Offset(0, 1).Range("A1")
      Sheets("Register").Range("K8").FormulaR1C1 = GetFirstLetters(usern2.Offset(0, 1).Range("A1"))
      Sheets("Register").Select
      Range("B8").Select
      ActiveWorkbook.Save
      End If
      
      If UserName = usern3 Then
      Worksheets("Report").Range("D11").FormulaR1C1 = usern3.Offset(0, 1).Range("A1")
      Sheets("Register").Range("K8").FormulaR1C1 = GetFirstLetters(usern3.Offset(0, 1).Range("A1"))
      Sheets("Register").Select
      Range("B8").Select
      ActiveWorkbook.Save
      End If
      
      If UserName = usern4 Then
      Worksheets("Report").Range("D11").FormulaR1C1 = usern4.Offset(0, 1).Range("A1")
      Sheets("Register").Range("K8").FormulaR1C1 = GetFirstLetters(usern4.Offset(0, 1).Range("A1"))
      Sheets("Register").Select
      Range("B8").Select
      ActiveWorkbook.Save
      End If
      
      If UserName = usern5 Then
      Worksheets("Report").Range("D11").FormulaR1C1 = usern5.Offset(0, 1).Range("A1")
      Sheets("Register").Range("K8").FormulaR1C1 = GetFirstLetters(usern5.Offset(0, 1).Range("A1"))
      Sheets("Register").Select
      Range("B8").Select
      ActiveWorkbook.Save
      End If
      
      If UserName = usern6 Then
      Worksheets("Report").Range("D11").FormulaR1C1 = usern6.Offset(0, 1).Range("A1")
      Sheets("Register").Range("K8").FormulaR1C1 = GetFirstLetters(usern6.Offset(0, 1).Range("A1"))
      Sheets("Register").Select
      Range("B8").Select
      ActiveWorkbook.Save
      End If
      
      If UserName = usern7 Then
      Worksheets("Report").Range("D11").FormulaR1C1 = usern7.Offset(0, 1).Range("A1")
      Sheets("Register").Range("K8").FormulaR1C1 = GetFirstLetters(usern7.Offset(0, 1).Range("A1"))
      Sheets("Register").Select
      Range("B8").Select
      ActiveWorkbook.Save
      End If
      
      If UserName = usern8 Then
      Worksheets("Report").Range("D11").FormulaR1C1 = usern8.Offset(0, 1).Range("A1")
      Sheets("Register").Range("K8").FormulaR1C1 = GetFirstLetters(usern8.Offset(0, 1).Range("A1"))
      Sheets("Register").Select
      Range("B8").Select
      ActiveWorkbook.Save
      End If
      
      If UserName = usern9 Then
      Worksheets("Report").Range("D11").FormulaR1C1 = usern9.Offset(0, 1).Range("A1")
      Sheets("Register").Range("K8").FormulaR1C1 = GetFirstLetters(usern9.Offset(0, 1).Range("A1"))
      Sheets("Register").Select
      Range("B8").Select
      ActiveWorkbook.Save
      End If
      
      If UserName = usern10 Then
      Worksheets("Report").Range("D11").FormulaR1C1 = usern10.Offset(0, 1).Range("A1")
      Sheets("Register").Range("K8").FormulaR1C1 = GetFirstLetters(usern10.Offset(0, 1).Range("A1"))
      Sheets("Register").Select
      Range("B8").Select
      ActiveWorkbook.Save
      End If
End Sub
