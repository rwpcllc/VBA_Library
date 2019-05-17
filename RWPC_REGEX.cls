'Description: takes in a source string and pattern and applies regex against the source
'then returns the first match
Function RegexFirstMatch(source As String, pattern As String)
    On Error GoTo FAILURE
    Dim regex As New RegExp
    Dim matches As Object
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = pattern
    End With
    
    If regex.test(source) Then
        Set matches = regex.Execute(source)
    End If
    
    If (matches.Count <> 0) Then
        RegexFirstMatch = matches.Item(0)
    Else
        RegexFirstMatch = ""
    End If
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "RegexMatch():" & Err.Description

End Function

'Description: takes in a source string and pattern and applies regex against the source
'then returns a boolean
'Args: none
'Return: boolean
Function RegexMatchBool(source As String, pattern As String)
    On Error GoTo FAILURE
    Dim regex As New RegExp
    Dim matches As Object
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = pattern
    End With
    RegexMatchBool = regex.test(source)
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "RegexMatchBool():" & Err.Description

End Function

