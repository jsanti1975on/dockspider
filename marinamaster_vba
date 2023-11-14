Sub UpdateTenantList()
    Dim spreadsheet As Worksheet
    Dim bsListSheet As Worksheet
    Dim tenantListSheet As Worksheet
    Dim bsData As Variant
    Dim tenantData() As Variant
    Dim maxSlipNumber As Integer
    Dim i As Integer, j As Integer
    Dim slip As Integer
    Dim found As Boolean
    
    ' Set references to the sheets
    Set spreadsheet = ThisWorkbook
    Set bsListSheet = spreadsheet.Sheets("bs_list")
    Set tenantListSheet = spreadsheet.Sheets("tenant_list")
    
    ' Clear the existing data in the tenant_list sheet
    tenantListSheet.Cells.Clear
    
    ' Retrieve data from bs_list sheet
    bsData = bsListSheet.Range("A2:B" & bsListSheet.Cells(bsListSheet.Rows.Count, "A").End(xlUp).Row).Value
    
    ' Initialize variables
    maxSlipNumber = 0
    
    ' Process data
    ReDim tenantData(1 To 1, 1 To 2)
    For i = 1 To UBound(bsData, 1)
        Dim name As String
        Dim slipNumber As Integer
        name = bsData(i, 1)
        slipNumber = bsData(i, 2)
        
        If name <> "" And slipNumber <> 0 Then
            ' Keep track of the maximum slip number
            If slipNumber > maxSlipNumber Then
                maxSlipNumber = slipNumber
            End If
            
            ' Add data to tenantData
            ReDim Preserve tenantData(1 To UBound(tenantData) + 1, 1 To 2)
            tenantData(UBound(tenantData), 1) = name
            tenantData(UBound(tenantData), 2) = slipNumber
        End If
    Next i
    
    ' Ensure there are rows in tenant_list for each slip number
    For slip = 1 To maxSlipNumber
        found = False
        For j = 1 To UBound(tenantData, 1)
            If tenantData(j, 2) = slip Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            ReDim Preserve tenantData(1 To UBound(tenantData) + 1, 1 To 2)
            tenantData(UBound(tenantData), 1) = ""
            tenantData(UBound(tenantData), 2) = slip
        End If
    Next slip
    
    ' Sort the tenantData by slipNumber
    QuickSort tenantData, 1, UBound(tenantData, 1), 2
    
    ' Write the sorted tenantData to the tenant_list sheet
    tenantListSheet.Range("A1:B" & UBound(tenantData, 1)).Value = tenantData
End Sub

Sub QuickSort(arr As Variant, low As Long, high As Long, keyIndex As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant
    Dim temp As Variant
    
    i = low
    j = high
    pivot = arr((low + high) \ 2, keyIndex)
    
    Do While i <= j
        Do While arr(i, keyIndex) < pivot
            i = i + 1
        Loop
        
        Do While pivot < arr(j, keyIndex)
            j = j - 1
        Loop
        
        If i <= j Then
            For k = LBound(arr, 2) To UBound(arr, 2)
                temp = arr(i, k)
                arr(i, k) = arr(j, k)
                arr(j, k) = temp
            Next k
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If low < j Then QuickSort arr, low, j, keyIndex
    If i < high Then QuickSort arr, i, high, keyIndex
End Sub
