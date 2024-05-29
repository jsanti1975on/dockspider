Option Explicit

' Click Event for the Aircraft Database
Private Sub cmdReset_Click()

    Dim msgValue As VbMsgBoxResult

    msgValue = MsgBox("Do you want to reset the form?", vbYesNo + vbInformation, "Confirmation")

    If msgValue = vbNo Then Exit Sub

    Call Reset

End Sub

Private Sub cmdSave_Click()

    Dim msgValue As VbMsgBoxResult

    msgValue = MsgBox("Do you want to save this data?", vbYesNo + vbInformation, "Confirmation")

    If msgValue = vbNo Then Exit Sub

    Call Submit
    Call Reset

End Sub

' Reset the form fields
Sub Reset()
    With frmAircraft
        .txtSaveNumber.Value = ""
        .txtArrival.Value = ""
        .txtDeparture.Value = ""
        .txtName.Value = ""
        .txtPhone.Value = ""
        .txtAddress.Value = ""
        .txtEmail.Value = ""
        .txtMake.Value = ""
        .txtNNumber.Value = ""
        .txtColors.Value = ""
        .optApron.Value = False
        .optDock.Value = False
        .optBeach.Value = False
    End With
End Sub

' Submit the form data to the Aircraft sheet
Sub Submit()

    Dim sh As Worksheet
    Dim iRow As Long

    Set sh = ThisWorkbook.Sheets("Aircraft")
    
    iRow = WorksheetFunction.CountA(sh.Range("A:A")) + 1

    With sh
        .Cells(iRow, 1).Value = frmAircraft.txtSaveNumber.Value
        .Cells(iRow, 2).Value = frmAircraft.txtArrival.Value
        .Cells(iRow, 3).Value = frmAircraft.txtDeparture.Value
        .Cells(iRow, 4).Value = frmAircraft.txtName.Value
        .Cells(iRow, 5).Value = frmAircraft.txtPhone.Value
        .Cells(iRow, 6).Value = frmAircraft.txtAddress.Value
        .Cells(iRow, 7).Value = frmAircraft.txtEmail.Value
        .Cells(iRow, 8).Value = frmAircraft.txtMake.Value
        .Cells(iRow, 9).Value = frmAircraft.txtNNumber.Value
        .Cells(iRow, 10).Value = frmAircraft.txtColors.Value
        If frmAircraft.optApron.Value = True Then
            .Cells(iRow, 11).Value = "Apron"
        ElseIf frmAircraft.optDock.Value = True Then
            .Cells(iRow, 11).Value = "Dock"
        ElseIf frmAircraft.optBeach.Value = True Then
            .Cells(iRow, 11).Value = "Beach"
        End If
    End With

End Sub

' Show the form
Sub Show_Form()

    frmAircraft.Show

End Sub
