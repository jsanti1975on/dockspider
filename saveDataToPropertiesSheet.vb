' Declare a variable to hold the reference to the user form
Dim EntryLogForm As UserForm

' Procedure to show the user form
Sub ShowTenantEntryLogForm()
    ' Create a new instance of the user form
    Set EntryLogForm = New UserForm_EntryLog
    
    ' Show the user form
    EntryLogForm.Show
End Sub

' Procedure to handle Save button click event
Sub SaveEntryLogData()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Set the worksheet reference (change "Properties" to your sheet name)
    Set ws = ThisWorkbook.Worksheets("Properties")
    
    ' Find the last used row in the Properties sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Copy data from user form controls to the Properties sheet
    With ws
        .Cells(lastRow, 1).Value = EntryLogForm.TextBox_Name.Value
        .Cells(lastRow, 2).Value = EntryLogForm.TextBox_DateArrival.Value
        .Cells(lastRow, 3).Value = EntryLogForm.TextBox_DateDeparture.Value
        .Cells(lastRow, 4).Value = EntryLogForm.TextBox_SlipNumber.Value
        .Cells(lastRow, 5).Value = EntryLogForm.TextBox_Phone1.Value
        .Cells(lastRow, 6).Value = EntryLogForm.TextBox_Phone2.Value
        .Cells(lastRow, 7).Value = EntryLogForm.TextBox_Email.Value
        .Cells(lastRow, 8).Value = EntryLogForm.TextBox_Model.Value
        .Cells(lastRow, 9).Value = EntryLogForm.TextBox_LOA.Value
        .Cells(lastRow, 10).Value = EntryLogForm.TextBox_Registration.Value
        .Cells(lastRow, 11).Value = EntryLogForm.TextBox_Insurance.Value
    End With
    
    ' Close the user form after saving
    EntryLogForm.Hide
    Set EntryLogForm = Nothing
    
    ' Inform user that data has been saved
    MsgBox "Entry log data has been saved.", vbInformation
End Sub

' Procedure to handle Reset button click event
Sub ResetEntryLogForm()
    ' Clear all text boxes on the user form
    With EntryLogForm
        .TextBox_Name.Value = ""
        .TextBox_DateArrival.Value = ""
        .TextBox_DateDeparture.Value = ""
        .TextBox_SlipNumber.Value = ""
        .TextBox_Phone1.Value = ""
        .TextBox_Phone2.Value = ""
        .TextBox_Email.Value = ""
        .TextBox_Model.Value = ""
        .TextBox_LOA.Value = ""
        .TextBox_Registration.Value = ""
        .TextBox_Insurance.Value = ""
    End With
End Sub
