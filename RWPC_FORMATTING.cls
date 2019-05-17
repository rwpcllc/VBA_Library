Public Function copyColumnFormat(sourceCol As String, targetCol As String, Optional sourceSheet As String, Optional targetSheet As String)
    On Error GoTo FAILURE
    If (sourceSheet = vbNullString Or sourceSheet = "") Then
        sourceSheet = ActiveSheet.Name
    End If
    
    If (targetSheet = vbNullString Or targetSheet = "") Then
        targetSheet = ActiveSheet.Name
    End If
    
    Sheets(sourceSheet).Select
    Range(sourceCol & ":" & sourceCol).Select
    Selection.Copy
    Sheets(targetSheet).Select
    Columns(targetCol & ":" & targetCol).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Exit Function
FAILURE:
    Err.Raise 1, , "copyColumnFormat():" & Err.DescriptionQ
End Function

