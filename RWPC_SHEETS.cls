"'*********************************************************************************************
'Description: Checks if provided sheet exists in workbook)
'Arg: sheetName (String)
'*********************************************************************************************
Function spreadsheetExists(sheetName As String) As Boolean
    On Error GoTo FAILURE
    spreadsheetExists = Evaluate(""ISREF('"" & sheetName & ""'!A1)"")
    Exit Function

FAILURE:
    Err.Raise Err.Number, Err.source, ""WorksheetExists():"" & Err.Description
End Function

'*********************************************************************************************
'Description: Checks if provided sheet name is empty (devoid of any content in any cells)
'Arg: sheetName (String)
'*********************************************************************************************
Function spreadsheetIsEmpty(sheetName As String)
    On Error GoTo FAILURE
    Sheets(sheetName).Select
    spreadsheetIsEmpty = (WorksheetFunction.CountA(ActiveSheet.UsedRange) = 0)
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""spreadsheetIsEmtpy("" & sheetName & ""):"" & Err.Description
    
End Function
"
