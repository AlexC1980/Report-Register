Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Not Intersect(Target, Range("L3")) Is Nothing Then Run "value2"
End Sub
