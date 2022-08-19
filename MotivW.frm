'This is called by NCRbutton from Module 3
Private Sub CommandButton1_Click()
ActiveCell.Offset(0, 4).value = TextBox1
Unload MotivW
ActiveWorkbook.Save
End Sub

Private Sub TextBox1_Change()
If Me.TextBox1 = vbNullString Then MsgBox ("Te rog scrie un motiv.")
TextBox1.Text = UCase(TextBox1.Text)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        MsgBox "Nu poti inchide fereastra apasand pe X!"
    End If
End Sub
