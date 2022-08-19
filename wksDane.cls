Private Sub Worksheet_SelectionChange(ByVal Target As Range)

Dim xRange As Range
Set xRange = Cells(Rows.count, 4).End(xlUp)
Set xRange = Range(Range("A:A"), xRange)
NumRows = Application.CountIf(xRange, Range("A1"))

Application.StatusBar = "Cantitate: " & NumRows * Range("D1") & "  |  " & Worksheets("Register").Range("B8")

End Sub
Private Sub Worksheet_Change(ByVal Target As Range)

Dim rngKolory As Range
Dim rngDoPokolorowania As Range
Dim LicznikKolorow As Integer
Dim Licznik As Integer
Dim rngKolumna As Range
Dim rngDaneWypelnione As Range

' cells with colors to choose from
Set rngKolory = wksDane.Range("rngDaneStartB").Resize(7, 1) '7 is the number of color duplicates
' cells with data to be "colored"
Set rngDoPokolorowania = wksDane.Range(Range("rngDaneStart"), Cells(65535, Range("rngDaneStart").Column).End(xlUp))

' column with data
Set rngKolumna = Columns("A")

With wksDane
    Set rngDaneWypelnione = .Range(.Range("rngDaneStart"), .Range("rngDaneStart").Offset(10000).End(xlUp))
End With

If Not Intersect(Target, rngKolumna) Is Nothing Then

Application.ScreenUpdating = False '

' Let's clear the whole data area (set background color to default)

rngDaneWypelnione.Resize(rngDaneWypelnione.count + 1).Interior.ColorIndex = xlNone

LicznikKolorow = 1 ' color counter reset

With rngDoPokolorowania
   ' first cell
   If Application.WorksheetFunction.CountIf(rngDoPokolorowania, .Cells(1).value) > 1 Then
      .Cells(1).Interior.ColorIndex = rngKolory.Cells(LicznikKolorow).Interior.ColorIndex
      LicznikKolorow = LicznikKolorow + 1
      If LicznikKolorow > rngKolory.count Then LicznikKolorow = 1
   End If
    
    'more than one cell
    If rngDaneWypelnione.count > 1 Then
        ' for following cells
        For Licznik = 2 To .count
            If Application.WorksheetFunction.CountIf(rngDoPokolorowania, _
                                                    .Cells(Licznik).value) > 1 Then
                If Application.WorksheetFunction.CountIf(Range("rngDaneStart").Resize(Licznik - 1), .Cells(Licznik).value) > 0 Then
                    .Cells(Licznik).Interior.ColorIndex = _
                    rngDaneWypelnione.Find(What:=.Cells(Licznik).value, After:=.Cells(Licznik), SearchDirection:=xlPrevious, LookAt:=xlWhole).Interior.ColorIndex
                Else
                    .Cells(Licznik).Interior.ColorIndex = rngKolory.Cells(LicznikKolorow).Interior.ColorIndex
                    LicznikKolorow = LicznikKolorow + 1
                If LicznikKolorow > rngKolory.count Then LicznikKolorow = 1
                End If
            End If
       Next Licznik
    End If
End With
Application.ScreenUpdating = True
End If
If ActiveCell = Range("A1") Then ActiveCell.Select Else If ActiveCell.Offset(-1, 0).Interior.ColorIndex = xlNone Then Call SoundWarning
End Sub
