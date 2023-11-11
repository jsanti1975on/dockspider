Sub CompareCashAndCreditTransactions()
    Dim wsData As Worksheet
    Dim wsSummary As Worksheet
    Dim LastRow As Long
    Dim PaymentType As String
    Dim TotalAmount As Double
    Dim TransactionDate As Date
    Dim i As Long
    
    ' Set references to the worksheets
    Set wsData = ThisWorkbook.Sheets("2012")
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    ' Clear previous data in the summary sheet
    wsSummary.Cells.Clear
    
    ' Initialize row number in the summary sheet
    i = 2
    
    ' Loop through each row in the data sheet
    LastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    For i = 2 To LastRow ' Assuming the data starts from row 2
    
        PaymentType = wsData.Cells(i, "J").Value ' Assuming payment type is in column J
        TotalAmount = wsData.Cells(i, "I").Value ' Assuming sales amount is in column I
        TransactionDate = wsData.Cells(i, "A").Value ' Assuming transaction date is in column A
        
        ' Check if the transaction is in January 2012
        If Year(TransactionDate) = 2012 And Month(TransactionDate) = 1 Then
            ' Add the data to the respective cells in the summary sheet
            If PaymentType = "CASH" Then
                wsSummary.Cells(2, "A").Value = wsSummary.Cells(2, "A").Value + TotalAmount
            ElseIf PaymentType = "MASTERCARD" Or PaymentType = "VISA" Or PaymentType = "AMEX" Then
                wsSummary.Cells(2, "B").Value = wsSummary.Cells(2, "B").Value + TotalAmount
            End If
        End If
    Next i
End Sub
