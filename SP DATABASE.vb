Option Explicit

' Click Event for the reset button
Private Sub cmdReset_Click()
    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Do you want to reset the form?", vbYesNo + vbInformation, "Confirmation")
    If msgValue = vbNo Then Exit Sub
    Call Reset
End Sub

' Click Event for the save button
Private Sub cmdSave_Click()
    Dim msgValue As VbMsgBoxResult
    msgValue = MsgBox("Do you want to save this data?", vbYesNo + vbInformation, "Confirmation")
    If msgValue = vbNo Then Exit Sub
    Call Submit
    Call Reset
End Sub

Option Explicit

' Reset the form fields
Sub Reset()
    With frmDatabase
        .txtSaveNumber.Value = ""
        .txtDataCashCounted.Value = ""
        .txtFolderSaveLocation.Value = ""
        .txtAmountCounted.Value = ""
        .txtBagNumber.Value = ""
        .optOpenYes.Value = False
        .optOpenNo.Value = False
        .optDepositedYes.Value = False
        .optDepositedNo.Value = False
    End With
End Sub

' Submit the form data to the sheet
Sub Submit()
    Dim sh As Worksheet
    Dim iRow As Long
    
    Set sh = ThisWorkbook.Sheets("Database")
    iRow = WorksheetFunction.CountA(sh.Range("A:A")) + 1
    
    With sh
        .Cells(iRow, 1).Value = Format(Now(), "DD-MM-YYYY HH:MM:SS") ' Save # (Timestamp)
        .Cells(iRow, 2).Value = frmDatabase.txtDataCashCounted.Value ' Data Cash Counted
        .Cells(iRow, 3).Value = frmDatabase.txtFolderSaveLocation.Value ' Folder Save Location
        .Cells(iRow, 4).Value = frmDatabase.txtAmountCounted.Value ' Amount Counted
        .Cells(iRow, 5).Value = frmDatabase.txtBagNumber.Value ' Bag Number
        ' Option Button for Open
        If frmDatabase.optOpenYes.Value = True Then
            .Cells(iRow, 6).Value = "Yes"
        ElseIf frmDatabase.optOpenNo.Value = True Then
            .Cells(iRow, 6).Value = "No"
        End If
        .Cells(iRow, 7).Value = Format(Now(), "DD-MM-YYYY HH:MM:SS") ' Time Stamp
        ' Option Button for Deposited
        If frmDatabase.optDepositedYes.Value = True Then
            .Cells(iRow, 8).Value = "Yes"
        ElseIf frmDatabase.optDepositedNo.Value = True Then
            .Cells(iRow, 8).Value = "No"
        End If
    End With
End Sub

' Show form by assignment to an object
Sub Show_Form()
    frmDatabase.Show
End Sub
