Option Explicit

Private Sub cmdClear_Click()
    ' Clear the username and password fields
    Me.txtUserID.Value = ""
    Me.txtPassword.Value = ""
    
    ' Set focus to the username field
    Me.txtUserID.SetFocus
End Sub

Function MD5Hash(ByVal text As String) As String
    Dim MD5 As Object
    Dim byteArray() As Byte
    Dim hash As String
    Dim i As Integer

    Set MD5 = CreateObject("System.Security.Cryptography.MD5")
    byteArray = MD5.ComputeHash_2(ByteArrayFromString(text))

    hash = ""
    For i = LBound(byteArray) To UBound(byteArray)
        hash = hash & Right("0" & Hex(byteArray(i)), 2)
    Next i

    MD5Hash = hash
End Function

Function ByteArrayFromString(ByVal str As String) As Byte()
    Dim i As Integer
    ReDim byteArray(Len(str) - 1) As Byte
    For i = 1 To Len(str)
        byteArray(i - 1) = Asc(Mid(str, i, 1))
    Next i
    ByteArrayFromString = byteArray
End Function

Private Sub cmdLogin_Click()
    Dim user As String
    Dim password As String
   
    user = Me.txtUserID.Value
    password = MD5Hash(Me.txtPassword.Value) ' Hash the entered password
   
    If (user = "Jason" And password = MD5Hash("0d7c85cbf7487d16LearnToCode514e49f7a7b")) Or _
       (user = "user" And password = MD5Hash("d36cb52de34b0LearnToCode97f62f9312088b42922")) Then
       
        ' Successful login
        Unload Me
        Application.Visible = True
    Else
        ' Invalid login
        If LoginInstance < 3 Then
            MsgBox "Invalid login credentials. Please try again.", vbOKOnly + vbCritical, "Invalid Login Details"
            LoginInstance = LoginInstance + 1
        Else
            MsgBox "You have exceeded the maximum number of login attempts.", vbOKOnly + vbCritical, "Invalid Credentials"
            Unload Me
            ThisWorkbook.Close savechanges:=False
            Application.Visible = True
            LoginInstance = 0
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Initialize the login form
    Me.txtUserID.Value = ""
    Me.txtPassword.Value = ""
    
    ' Set focus to the username field
    Me.txtUserID.SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Disable the close button on the user form
    If CloseMode = 0 Then Cancel = True
End Sub
