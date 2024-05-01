' Cancel Click Event // Source-Code 'Data Labs' 
' The option to cancel will be added to 'Data Labs' Source-Code

Private Sub cmdCancel_Click()
    ' Code to handle the click event of the "Cancel" button
    Unload Me
    ThisWorkbook.Close savechanges:=False
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' The code to disable the close button on the user form
    If CloseMode = 0 Then Cancel = True
End Sub


