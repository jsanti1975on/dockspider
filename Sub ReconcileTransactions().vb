Sub ReconcileTransactions()
    Dim ws As Worksheet
    Dim lastRowPOS As Long
    Dim lastRowCC As Long
    Dim i As Long, j As Long
    Dim matchFound As Boolean
    Dim posAmount As Range
    Dim posMethod As Range
    Dim ccAmount As Range
    Dim ccMethod As Range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Find the last row with data in both POS and Credit Card columns
    lastRowPOS = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowCC = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    ' Loop through each transaction in POS data
    For i = 2 To lastRowPOS ' Assuming there is a header row, start from the second row
        Set posAmount = ws.Cells(i, 1) ' Column A for POS amount
        Set posMethod = ws.Cells(i, 2) ' Column B for POS method of payment
        matchFound = False
        
        ' Check if the transaction in POS data is found in Credit Card data
        For j = 2 To lastRowCC
            Set ccAmount = ws.Cells(j, 4) ' Column D for Credit Card amount
            Set ccMethod = ws.Cells(j, 5) ' Column E for Credit Card method of payment
            If posAmount.Value = ccAmount.Value And posMethod.Value = ccMethod.Value Then
                matchFound = True
                Exit For
            End If
        Next j
        
        ' Highlight cell in Column C if no match is found
        If Not matchFound Then
            ws.Cells(i, 3).Value = "Mismatch"
            ws.Cells(i, 3).Interior.Color = RGB(255, 0, 0) ' Red
        Else
            ws.Cells(i, 3).Value = ""
            ws.Cells(i, 3).Interior.ColorIndex = xlNone
        End If
    Next i
    
    ' Loop through each transaction in Credit Card data to highlight unmatched values
    For j = 2 To lastRowCC
        Set ccAmount = ws.Cells(j, 4) ' Column D for Credit Card amount
        Set ccMethod = ws.Cells(j, 5) ' Column E for Credit Card method of payment
        matchFound = False
        
        ' Check if the transaction in Credit Card data is found in POS data
        For i = 2 To lastRowPOS
            Set posAmount = ws.Cells(i, 1) ' Column A for POS amount
            Set posMethod = ws.Cells(i, 2) ' Column B for POS method of payment
            If ccAmount.Value = posAmount.Value And ccMethod.Value = posMethod.Value Then
                matchFound = True
                Exit For
            End If
        Next i
        
        ' Highlight cell in Column C if no match is found
        If Not matchFound Then
            ws.Cells(j, 3).Value = "Mismatch"
            ws.Cells(j, 3).Interior.Color = RGB(255, 0, 0) ' Red
        Else
            ws.Cells(j, 3).Value = ""
            ws.Cells(j, 3).Interior.ColorIndex = xlNone
        End If
    Next j

    ' Notify the user that the reconciliation is complete
    MsgBox "Reconciliation complete. Discrepancies are highlighted in red.", vbInformation
End Sub
