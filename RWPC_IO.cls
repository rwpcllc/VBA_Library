"Function getUserInput(prompt As String, title As String)
    On Error GoTo FAILURE
    Dim userInput As String
    userInput = InputBox(prompt, title)
    If (userInput = """" Or userInput = Null) Then
        Err.Raise 1, , ""Invalid user input blank or null""
    End If
    getUserInput = userInput
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""getUserInput():"" & Err.Description
End Function
"
