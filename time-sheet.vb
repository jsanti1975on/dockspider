Private Sub UserForm_Initialize()
    ' Add items to ComboBox1
    ComboBox1.AddItem "Option 1"
    ComboBox1.AddItem "Option 2"
    ComboBox1.AddItem "Option 3"
    
    ' Add items to ComboBox2
    ComboBox2.AddItem "Option 1"
    ComboBox2.AddItem "Option 2"
    ComboBox2.AddItem "Option 3"
End Sub
Private Sub SubmitButton_Click()
    ' Populate the selected values from ComboBoxes and TextBoxes into Sheet1
    With ThisWorkbook.Sheets("Sheet1")
        .Range("B2").Value = ComboBox1.Value
        .Range("B3").Value = ComboBox2.Value
        .Range("C2").Value = TextBox1.Value
        .Range("C3").Value = TextBox2.Value
    End With
    
    ' Clear the form after submission
    ComboBox1.Value = ""
    ComboBox2.Value = ""
    TextBox1.Value = ""
    TextBox2.Value = ""
End Sub
Private Sub UserForm_Initialize()
    ' Add items to ComboBox1
    ComboBox1.AddItem "Option 1"
    ComboBox1.AddItem "Option 2"
    ComboBox1.AddItem "Option 3"
    
    ' Add items to ComboBox2
    ComboBox2.AddItem "Option 1"
    ComboBox2.AddItem "Option 2"
    ComboBox2.AddItem "Option 3"
End Sub

Private Sub SubmitButton_Click()
    ' Populate the selected values from ComboBoxes and TextBoxes into Sheet1
    With ThisWorkbook.Sheets("Sheet1")
        .Range("B2").Value = ComboBox1.Value
        .Range("B3").Value = ComboBox2.Value
        .Range("C2").Value = TextBox1.Value
        .Range("C3").Value = TextBox2.Value
    End With
    
    ' Clear the form after submission
    ComboBox1.Value = ""
    ComboBox2.Value = ""
    TextBox1.Value = ""
    TextBox2.Value = ""
End Sub
