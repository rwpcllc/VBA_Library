VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RWPC_Sheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************************************
'Description: Checks if provided sheet exists in workbook)
'Arg: sheetName (String)
'*********************************************************************************************
Function spreadsheetExists(sheetName As String) As Boolean
    On Error GoTo FAILURE
    spreadsheetExists = Evaluate("ISREF('" & sheetName & "'!A1)")
    Exit Function

FAILURE:
    Err.Raise Err.Number, Err.source, "WorksheetExists():" & Err.Description
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
    Err.Raise Err.Number, Err.source, "spreadsheetIsEmtpy(" & sheetName & "):" & Err.Description
    
End Function

