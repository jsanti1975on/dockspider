Private Sub btnCalculate_Click()
    Dim userInput As Double
    
    ' Retrieve user input from the TextBox
    userInput = CDbl(txtInput.Value)
    
    ' Send the user input to cell J11
    Sheets("SplashParkCalculator").Range("J11").Value = userInput
End Sub

Private Sub btnRetrieve_Click()
    ' Retrieve the value from cell J5
    Dim retrievedValue As Double
    retrievedValue = Sheets("SplashParkCalculator").Range("J5").Value
    
    ' Display the retrieved value in the label
    lblDisplay.Caption = "Retrieved Value: " & Format(retrievedValue, "0.00")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Make Excel application visible again when the user form is closed
    Application.Visible = True
    
    ' Unload the user form
    Unload Me
End Sub
Private Sub btnCalculate_Click()
    Dim userInput As Double
    
    ' Retrieve user input from the TextBox
    userInput = CDbl(txtInput.Value)
    
    ' Send the user input to cell J11
    Sheets("SplashParkCalculator").Range("J11").Value = userInput
End Sub

Private Sub btnRetrieve_Click()
    ' Retrieve the value from cell J5
    Dim retrievedValue As Double
    retrievedValue = Sheets("SplashParkCalculator").Range("J5").Value
    
    ' Display the retrieved value in the label
    lblDisplay.Caption = "Retrieved Value: " & Format(retrievedValue, "0.00")
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Make Excel application visible again when the user form is closed
    Application.Visible = True
    
    ' Unload the user form
    Unload Me
End Sub
