Option Explicit
'Option Compare Text

Const redCIndex As Long = 3
Const blackCIndex As Long = 0

Sub CheckDiffs() 'Check differences in Cell C8 and D8
    Dim CSensitivity As Long
    Dim oneCell As Range

            CSensitivity = 0
   
    With ThisWorkbook.Sheets("Register").Range("C8"): Rem adjust
        For Each oneCell In Range(.Cells(1, 1), .Cells(.Rows.count, 1).End(xlUp))
            oneCell.Font.Color = blackCIndex
            oneCell.Offset(0, 1).Font.ColorIndex = blackCIndex
            Call highlightDifference(oneCell.Offset(0, 1), oneCell, CSensitivity)
        Next oneCell
    End With
End Sub
Sub highlightDifference(refCell As Range, testCell As Range, Optional CaseSensitivity As Long)
    Rem default caseSenstivity = 0 for case insensitive, set CaseSensitivity = 1
    
    Dim refString As String, testString As String
    Dim i As Long, startPoint As Long, newPoint As Long
    

    CaseSensitivity = Sgn(CaseSensitivity) ^ 2
    
    
    With testCell.Font
        .ColorIndex = redCIndex
        .FontStyle = "Bold"
    End With
     
    refString = refCell.Text
    testString = testCell.Text
    startPoint = 1
    For i = 1 To Len(refString)
        newPoint = InStr(startPoint, testString, Mid(refString, i, 1), CaseSensitivity)
        If newPoint <> 0 Then
            With testCell.Characters(newPoint, 1).Font
                .ColorIndex = blackCIndex
                .FontStyle = "Regular"
            End With
            startPoint = newPoint + 1
        End If
    Next i

End Sub
