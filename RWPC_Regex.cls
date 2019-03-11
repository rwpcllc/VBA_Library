VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RWPC_Regex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Description: takes in a source string and pattern and applies regex against the source
'then returns the first match
Function RegexMatch(source As String, pattern As String)
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
        RegexMatch = matches.Item(0)
    Else
        RegexMatch = ""
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
