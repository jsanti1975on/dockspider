' This VBA code dynamically creates a user form (DynamicUserForm) with three label-text box pairs for data entry and two command buttons (Save and Reset).
' The loop (For i = 1 To 3) creates three sets of labels and text boxes vertically aligned on the form.
' Each label (lbl) and text box (txtBox) is created and positioned relative to the previous control using the Top property.
' Command buttons (btnSave and btnReset) are added below the label-text box pairs.
' The VBA.UserForms.Add(frm.Name).Show line displays the user form.
' After the user closes the form, the dynamically created form (frm) is removed from the VBA project using ThisWorkbook.VBProject.VBComponents.Remove frm.
' Usage 
' Open the VBA editor (Alt + F11).
' Insert a new module (Insert > Module).
' Copy and paste the above code into the module.
' Run the CreateUserFormWithCode subroutine to create and display the dynamic user form.
' This approach is particularly useful when you need to generate user interfaces programmatically at runtime.



Sub CreateUserFormWithCode()
    Dim frm As Object
    Dim lbl As Object
    Dim txtBox As Object
    Dim btnSave As Object
    Dim btnReset As Object
    Dim yPos As Integer
    Dim ctrlHeight As Integer
    Dim ctrlWidth As Integer
    Dim spacing As Integer
    
    ' Initialize form parameters
    yPos = 20 ' Starting Y-position for controls
    ctrlHeight = 20 ' Height of controls
    ctrlWidth = 120 ' Width of controls
    spacing = 10 ' Spacing between controls
    
    ' Create a new user form
    Set frm = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = vbext_ct_MSForm (user form)
    frm.Name = "DynamicUserForm"
    frm.Properties("Caption") = "Dynamic User Form"
    
    ' Add labels and text boxes for data entry
    For i = 1 To 3
        ' Create label
        Set lbl = frm.Designer.Controls.Add("Forms.Label.1")
        With lbl
            .Caption = "Field " & i & ":"
            .Left = 20
            .Top = yPos
            .Width = ctrlWidth
            .Height = ctrlHeight
        End With
        
        ' Create text box
        Set txtBox = frm.Designer.Controls.Add("Forms.TextBox.1")
        With txtBox
            .Name = "TextBox" & i
            .Left = 150
            .Top = yPos
            .Width = ctrlWidth
            .Height = ctrlHeight
        End With
        
        ' Increment Y-position for the next control
        yPos = yPos + ctrlHeight + spacing
    Next i
    
    ' Add Save button
    Set btnSave = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With btnSave
        .Caption = "Save"
        .Name = "SaveButton"
        .Left = 20
        .Top = yPos
        .Width = ctrlWidth
        .Height = ctrlHeight
    End With
    
    ' Add Reset button
    Set btnReset = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With btnReset
        .Caption = "Reset"
        .Name = "ResetButton"
        .Left = 150
        .Top = yPos
        .Width = ctrlWidth
        .Height = ctrlHeight
    End With
    
    ' Show the user form
    VBA.UserForms.Add(frm.Name).Show
    
    ' Clean up: remove the dynamically created form from the VBA project
    ThisWorkbook.VBProject.VBComponents.Remove frm
End Sub
