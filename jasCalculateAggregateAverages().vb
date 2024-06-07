Sub jasCalculateAggregateAverages()
    Dim inputSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim inputData As Variant
    Dim outputData As Variant
    Dim skuData As Object
    Dim i As Long, j As Long
    Dim sku As String
    Dim qty As Double, cost As Double, totalFormulaResult As Double
    Dim avgCost As Double
    Dim row As Long
    Dim cValue As Double, eValue As Double, fValue As Double, hValue As Double

    ' Set references to sheets
    Set inputSheet = ThisWorkbook.Sheets("input")
    Set outputSheet = ThisWorkbook.Sheets("output")

    ' Get data from the 'input' sheet
    inputData = inputSheet.Range("A2:H100").Value

    ' Create a dictionary to store aggregated data based on SKU
    Set skuData = CreateObject("Scripting.Dictionary")

    ' Loop through the input data
    For i = 1 To UBound(inputData, 1) ' Start from 1 to skip the header row
        sku = inputData(i, 2) ' Assuming SKU is in column B
        qty = inputData(i, 7) ' Assuming QTY is in column G
        cost = inputData(i, 8) ' Assuming Cost is in column H
        totalFormulaResult = inputData(i, 6) ' Assuming the formula result is in column F

        If Not skuData.Exists(sku) Then
            skuData.Add sku, Array(qty, cost, totalFormulaResult, 1)
        Else
            skuData(sku)(0) = skuData(sku)(0) + qty
            skuData(sku)(1) = skuData(sku)(1) + cost
            skuData(sku)(2) = skuData(sku)(2) + totalFormulaResult
            skuData(sku)(3) = skuData(sku)(3) + 1
        End If
    Next i

    ' Prepare output data
    ReDim outputData(1 To skuData.Count + 1, 1 To 4)
    outputData(1, 1) = "SKU"
    outputData(1, 2) = "QTY"
    outputData(1, 3) = "Average Cost"
    outputData(1, 4) = "Total Formula Result"

    j = 2
    For Each sku In skuData.Keys
        avgCost = skuData(sku)(1) / skuData(sku)(3)
        outputData(j, 1) = sku
        outputData(j, 2) = skuData(sku)(0)
        outputData(j, 3) = avgCost
        outputData(j, 4) = skuData(sku)(2)
        j = j + 1
    Next sku

    ' Write aggregated data to the 'output' sheet
    outputSheet.Range(outputSheet.Cells(1, 1), outputSheet.Cells(UBound(outputData, 1), UBound(outputData, 2))).Value = outputData

    ' Delay for 4 seconds (4000 milliseconds)
    Application.Wait (Now + TimeValue("0:00:04"))

    ' Check for discrepancies in columns E, F, and H
    For row = 2 To 24
        cValue = inputSheet.Cells(row, 3).Value ' Value in column C
        eValue = inputSheet.Cells(row, 5).Value ' Value in column E
        fValue = inputSheet.Cells(row, 6).Value ' Value in column F
        hValue = inputSheet.Cells(row, 8).Value ' Value in column H

        ' Check for discrepancy
        If ((cValue * hValue) / cValue) <> fValue Then
            ' Discrepancy found, fill cell in column I with red
            inputSheet.Cells(row, 9).Interior.Color = RGB(255, 0, 0) ' Red fill color
        Else
            ' No discrepancy, fill cell in column I with green
            inputSheet.Cells(row, 9).Interior.Color = RGB(0, 255, 0) ' Green fill color
        End If
    Next row
End Sub
