Option Explicit

Sub Reset()
    With frmTenant
        ' The save number will be a time stamp
        .txtTimeStamp.Value = "'"
        .txtArrival.Value = "'"
        .txtReturnCount.Value = "'"
        .txtName.Value = "'"
        .txtPrimaryPhone.Value = "'"
        .txtAddress.Value = "'"
        .txtEmail.Value = "'"
        .txtMake.Value = "'"
        .txtFLNumber.Value = "'"
        .txtVesselColor.Value = "'"
        .txtPreviousSlips.Value = "'"
        .txtEmergencyContact.Value = "'"
        .txtModel.Value = "'"
        .optApron.Value = False
        .optDock.Value = False
        .optBeach.Value = False
    End With
End Sub

Sub Submit()
    ' Submit the form data to the Tenant sheet
    ' Add the time stamp line
    Dim sh As Worksheet
    Dim iRow As Long
    
    Set sh = ThisWorkbook.Sheets("Tenant")
    
    iRow = WorksheetFunction.CountA(sh.Range("A:A")) + 1
    
    With sh
        .Cells(iRow, 1).Value = [Text(Now(),"DD-MM-YYYY HH:MM:SS")]
        .Cells(iRow, 2).Value = frmTenant.txtArrival.Value
        .Cells(iRow, 3).Value = frmTenant.txtReturnCount.Value
        .Cells(iRow, 4).Value = frmTenant.txtName.Value
        .Cells(iRow, 5).Value = frmTenant.txtPrimaryPhone.Value
        .Cells(iRow, 6).Value = frmTenant.txtAddress.Value
        .Cells(iRow, 7).Value = frmTenant.txtEmail.Value
        .Cells(iRow, 8).Value = frmTenant.txtMake.Value
        .Cells(iRow, 9).Value = frmTenant.txtFLNumber.Value
        .Cells(iRow, 10).Value = frmTenant.txtVesselColor.Value
        .Cells(iRow, 11).Value = frmTenant.txtPreviousSlips.Value
        .Cells(iRow, 12).Value = frmTenant.txtEmergencyContact.Value
        .Cells(iRow, 13).Value = frmTenant.txtModel.Value
        
        ' This If Else statement may need testing
        If frmTenant.optApron.Value = True Then
            .Cells(iRow, 14).Value = "Apron"
        ElseIf frmTenant.optDock.Value = True Then
            .Cells(iRow, 14).Value = "Dock"
        ElseIf frmTenant.optBeach.Value = True Then
            .Cells(iRow, 14).Value = "Beach"
        End If
    End With
End Sub

' Show form by assignment to an object
Sub Show_Form()
    frmTenant.Show
End Sub

' Click Events
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
