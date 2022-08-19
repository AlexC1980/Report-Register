Private Sub CheckBox1_Click()
If CheckBox1 = True Then
   CheckBox2 = False
   Worksheets("Register").Range("L8").value = "Y"
End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2 = True Then
   CheckBox1 = False
   Worksheets("Register").Range("L8").value = "N"
End If
End Sub

Private Sub CheckBox3_Click()
If CheckBox3 = True Then
   CheckBox4 = False
   Worksheets("Register").Range("M8").value = "Y"
End If
End Sub

Private Sub CheckBox4_Click()
If CheckBox4 = True Then
   CheckBox3 = False
   Worksheets("Register").Range("M8").value = "N"
End If
End Sub

Private Sub CheckBox5_Click() 'Pass
If CheckBox5 = True Then
   CheckBox6 = False
   Worksheets("Register").Range("I8").value = "X"
   Worksheets("Register").Range("J8").value = ""
End If
End Sub

Private Sub CheckBox6_Click() 'Fail
If CheckBox6 = True Then
   CheckBox5 = False
   Worksheets("Register").Range("I8").value = ""
   Worksheets("Register").Range("J8").value = "X"
End If
End Sub
