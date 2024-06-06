Sub CompareColumnsFlexible()
    Dim ws As Worksheet
    Dim lastRowD As Long
    Dim lastRowE As Long
    Dim i As Long, j As Long
    Dim matchFound As Boolean
    Dim cellD As Range
    Dim cellE As Range
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet's name if different

    ' Find the last row with data in columns D and E
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Loop through each cell in column D
    For i = 2 To lastRowD ' Assuming there is a header row, start from the second row
        Set cellD = ws.Cells(i, 4) ' Column D
        matchFound = False
        
        ' Check if the value in column D is found in column E
        For j = 2 To lastRowE
            Set cellE = ws.Cells(j, 5) ' Column E
            If cellD.Value = cellE.Value Then
                matchFound = True
                Exit For
            End If
        Next j
        
        ' Highlight cell in column D if no match is found
        If Not matchFound Then
            cellD.Interior.Color = RGB(255, 0, 0) ' Red
        Else
            cellD.Interior.ColorIndex = xlNone
        End If
    Next i
    
    ' Loop through each cell in column E to highlight unmatched values
    For j = 2 To lastRowE
        Set cellE = ws.Cells(j, 5) ' Column E
        matchFound = False
        
        ' Check if the value in column E is found in column D
        For i = 2 To lastRowD
            Set cellD = ws.Cells(i, 4) ' Column D
            If cellE.Value = cellD.Value Then
                matchFound = True
                Exit For
            End If
        Next i
        
        ' Highlight cell in column E if no match is found
        If Not matchFound Then
            cellE.Interior.Color = RGB(255, 0, 0) ' Red
        Else
            cellE.Interior.ColorIndex = xlNone
        End If
    Next j

    ' Notify the user that the comparison is complete
    MsgBox "Comparison complete. Differences are highlighted in red.", vbInformation
End Sub
