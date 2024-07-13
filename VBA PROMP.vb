Sub CheckFuelData()
    Dim ws As Worksheet
    Dim startRecApron As Double, startAvApron As Double, startRecDock As Double, startAvDock As Double
    Dim finishRecApron As Double, finishAvApron As Double, finishRecDock As Double, finishAvDock As Double
    Dim errorSheets As String
    Dim isNumericStart As Boolean, isNumericFinish As Boolean
    
    errorSheets = "Sheets with errors:" & vbCrLf
    
    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize numeric checks
        isNumericStart = True
        isNumericFinish = True
        
        ' Read start meter values
        If IsNumeric(ws.Range("K17").Value) Then
            startRecApron = ws.Range("K17").Value
        Else
            isNumericStart = False
        End If
        If IsNumeric(ws.Range("K18").Value) Then
            startAvApron = ws.Range("K18").Value
        Else
            isNumericStart = False
        End If
        If IsNumeric(ws.Range("K19").Value) Then
            startRecDock = ws.Range("K19").Value
        Else
            isNumericStart = False
        End If
        If IsNumeric(ws.Range("K20").Value) Then
            startAvDock = ws.Range("K20").Value
        Else
            isNumericStart = False
        End If
        
        ' Read finish meter values
        If IsNumeric(ws.Range("K23").Value) Then
            finishRecApron = ws.Range("K23").Value
        Else
            isNumericFinish = False
        End If
        If IsNumeric(ws.Range("K24").Value) Then
            finishAvApron = ws.Range("K24").Value
        Else
            isNumericFinish = False
        End If
        If IsNumeric(ws.Range("K25").Value) Then
            finishRecDock = ws.Range("K25").Value
        Else
            isNumericFinish = False
        End If
        If IsNumeric(ws.Range("K26").Value) Then
            finishAvDock = ws.Range("K26").Value
        Else
            isNumericFinish = False
        End If
        
        ' Check for inconsistencies if both start and finish values are numeric
        If isNumericStart And isNumericFinish Then
            If finishRecApron <= startRecApron Or _
               finishAvApron <= startAvApron Or _
               finishRecDock <= startRecDock Or _
               finishAvDock <= startAvDock Then
                errorSheets = errorSheets & ws.Name & vbCrLf
            End If
        Else
            ' Add sheet to error list if non-numeric data found
            errorSheets = errorSheets & ws.Name & " (Non-numeric data)" & vbCrLf
        End If
    Next ws
    
    ' Display sheets with errors
    If errorSheets = "Sheets with errors:" & vbCrLf Then
        MsgBox "No errors found in the data."
    Else
        MsgBox errorSheets, vbCritical, "Data Inconsistencies Found"
    End If
End Sub
' Promp 2 
Sub CheckTankMeterReads()
    Dim ws As Worksheet
    Dim startValues As Dictionary
    Dim finishValues As Dictionary
    Dim startCell As Range, finishCell As Range
    Dim dayNumber As Integer
    Dim errors As String
    Dim i As Integer
    
    ' Initialize dictionaries to store start and finish values
    Set startValues = New Dictionary
    Set finishValues = New Dictionary

    ' Initialize error string
    errors = ""

    ' Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets
        ' Extract day number from sheet name assuming sheet names end with the day number
        dayNumber = CInt(Right(ws.Name, 2))
        
        ' Read start values
        Set startCell = ws.Range("K17:K20")
        For i = 1 To 4
            If Not startValues.Exists(dayNumber) Then
                startValues(dayNumber) = ws.Cells(startCell.Row + i - 1, startCell.Column).Value
            Else
                startValues(dayNumber) = startValues(dayNumber) + "," + ws.Cells(startCell.Row + i - 1, startCell.Column).Value
            End If
        Next i

        ' Read finish values
        Set finishCell = ws.Range("K23:K26")
        For i = 1 To 4
            If Not finishValues.Exists(dayNumber) Then
                finishValues(dayNumber) = ws.Cells(finishCell.Row + i - 1, finishCell.Column).Value
            Else
                finishValues(dayNumber) = finishValues(dayNumber) + "," + ws.Cells(finishCell.Row + i - 1, finishCell.Column).Value
            End If
        Next i
    Next ws

    ' Check if start values are less than finish values for each day
    For Each dayNumber In startValues.Keys
        Dim startArray() As String
        Dim finishArray() As String
        
        startArray = Split(startValues(dayNumber), ",")
        finishArray = Split(finishValues(dayNumber), ",")
        
        For i = 0 To UBound(startArray)
            If Val(startArray(i)) >= Val(finishArray(i)) Then
                errors = errors & "Error on Day " & dayNumber & " for pump " & (i + 1) & ": Start = " & startArray(i) & ", Finish = " & finishArray(i) & vbCrLf
            End If
        Next i
    Next dayNumber

    ' Display errors if any
    If errors = "" Then
        MsgBox "All values are correct.", vbInformation
    Else
        MsgBox errors, vbExclamation, "Errors Found"
    End If
End Sub
