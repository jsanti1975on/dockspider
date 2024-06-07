Sub Pause(duration_ms As Double)
    Dim start_time As Double
    start_time = Timer
    Do
        DoEvents
    Loop Until (Timer - start_time) * 1000 >= duration_ms
End Sub

 Sub AnimateSmileyFace()
    Dim mouthOpen As Shape
    Dim mouthClosed As Shape
    Dim i As Integer
    
    ' Set the shapes
    Set mouthOpen = ActiveSheet.Shapes("MouthOpen")
    Set mouthClosed = ActiveSheet.Shapes("MouthClosed")
    
    ' Initial visibility
    mouthOpen.Visible = msoTrue
    mouthClosed.Visible = msoFalse
    
    ' Loop to animate the mouth opening and closing
    For i = 1 To 10
        ' Toggle visibility
        mouthOpen.Visible = Not mouthOpen.Visible
        mouthClosed.Visible = Not mouthClosed.Visible
        
        ' Pause for a short duration (500 milliseconds)
        Pause 500
    Next i
End Sub

Sub AnimateTextBubble()
    Dim textBubble As Shape
    Dim i As Integer
    
    ' Ensure the text bubble shape exists
    On Error Resume Next
    Set textBubble = ActiveSheet.Shapes("TextBubble")
    On Error GoTo 0
    
    If textBubble Is Nothing Then
        MsgBox "Shape 'TextBubble' not found on the active sheet."
        Exit Sub
    End If
    
    ' Loop to animate the text bubble appearing and disappearing
    For i = 1 To 5 ' Adjust the number of times as needed
        ' Show the text bubble
        textBubble.Visible = msoTrue
        
        ' Animate the smiley face while the text bubble is visible
        AnimateSmileyFace
        
        ' Hide the text bubble
        textBubble.Visible = msoFalse
        
        ' Pause for a short duration (1 second)
        Pause 1000
    Next i
End Sub

