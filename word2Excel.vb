REM Code modified from "Dr. Todd Grande's" version 
Sub AircraftRev1()
    Dim wordApp As Object
    Dim wDoc As Object
    Dim r As Long ' Use Long instead of Integer for row counter
    Dim i As Integer
    
    ' Create an instance of Word Application
    Set wordApp = CreateObject("word.application")
    
    ' Open the Word application
    Set wDoc = wordApp.Documents.Open(ThisWorkbook.Path & "\" & Sheets("Sheet1").Range("E1").Value & ".docx")
    
    ' Make Word application visible (for testing purposes)
    wordApp.Visible = True
    
    ' Initialize row counter for the "Overnight01" sheet
    r = 1
    
    ' Loop through the content controls in the Word document and paste the text into "Overnight01" sheet
    For i = 1 To 10
        Sheets("Overnight01").Cells(r, i).Value = wDoc.ContentControls(i).Range.Text
    Next i
    
    ' Close the Word document and quit Word application
    wDoc.Close
    wordApp.Quit
    
    ' Release object reference
    Set wDoc = Nothing
    Set wordApp = Nothing
           
End Sub
