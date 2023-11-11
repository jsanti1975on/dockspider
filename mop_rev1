Sub SumPaymentTypesWithInvoiceAndDate()
    Dim wsData As Worksheet
    Dim wsScript As Worksheet
    Dim LastRow As Long
    Dim PaymentType As String
    Dim TotalAmount As Double
    Dim InvoiceNumber As String
    Dim TransactionDate As Date
    Dim i As Long
    
    ' Set references to the worksheets
    Set wsData = ThisWorkbook.Sheets("2012")
    Set wsScript = ThisWorkbook.Sheets("2012_script")
    
    ' Clear previous data in the script sheet
    wsScript.Cells.Clear
    
    ' Initialize row number in the script sheet
    i = 2
    
    ' Initialize a collection to track unique invoice numbers
    Dim InvoiceNumbers As Collection
    Set InvoiceNumbers = New Collection
    
    ' Loop through each row in the data sheet
    LastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    For i = 2 To LastRow ' Assuming the data starts from row 2
    
        PaymentType = wsData.Cells(i, "J").Value ' Assuming payment type is in column J
        TotalAmount = wsData.Cells(i, "I").Value ' Assuming sales amount is in column I
        InvoiceNumber = wsData.Cells(i, "D").Value ' Assuming invoice number is in column D
        TransactionDate = wsData.Cells(i, "A").Value ' Assuming transaction date is in column A
	 
        ' Check if the payment type is one of the desired types or if it's "CHECK"
        If PaymentType = "CASH" Or PaymentType = "MASTERCARD" Or PaymentType = "VISA" Or PaymentType = "CHECK" Or PaymentType = "AMEX" Then
            ' Check if the invoice number is not in the collection (i.e., it's unique)
            If Not Contains(InvoiceNumbers, InvoiceNumber) Then
                ' Add the invoice number to the collection
                InvoiceNumbers.Add InvoiceNumber
                ' Add the data to the respective cells in the script sheet
                wsScript.Cells(i, "A").Value = TransactionDate
                wsScript.Cells(i, "B").Value = PaymentType
                wsScript.Cells(i, "C").Value = InvoiceNumber
                wsScript.Cells(i, "D").Value = wsScript.Cells(i, "D").Value + TotalAmount
            End If
        End If
    Next i
End Sub

Function Contains(col As Collection, item As Variant) As Boolean
    On Error Resume Next
    Contains = Not IsEmpty(col(item))
    On Error GoTo 0
End Function
