"'***********************************************************************************************
'Description: This function converts a supplied letter into the corresponding column number
'Args: (string) colLetter -- letter of the column
'Return: (long) column number
'***********************************************************************************************
Function convertColLetterToNum(colLetter As Variant)
    On Error GoTo FAILURE
    If (Len(colLetter) > 3 Or colLetter > ""XFD"") Then
        Err.Raise 1, , ""provided letter exceeds Excel limit of XFD""
    End If
    convertColLetterToNum = Range(colLetter & 1).column
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""convertColLetterToNum()"" + Err.Description
End Function

'***********************************************************************************************
'Description: This function converts a supplied long into its corresponding column letter
'Args: (long) colNum -- number of the column
'Return: (string) column letter
'***********************************************************************************************
Function convertColNumToLetter(colNum As Variant)
    On Error GoTo FAILURE
    If (Not IsNumeric(colNum)) Then
        Err.Raise 1, , ""Invalid input, number not provided as argument""
    End If
    Dim num As Long: num = CLng(colNum)
    
    If (num > 16384 Or num < 1) Then
       Err.Raise 1, , ""Column number: "" & colNum & "" exceeds Excel range of 1 to 16384""
    End If

    Dim n As Long
    Dim C As Byte
    Dim s As String

    n = num
    Do
        C = ((n - 1) Mod 26)
        s = Chr(C + 65) & s
        n = (n - C) \ 26
    Loop While n > 0
    convertColNumToLetter = s
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""convertColNumToLetter():"" & Err.Description
End Function

'***********************************************************************************************
'Description: Searches for column name in given header row number
'Args: (long) rowNum -- row with headers, (String) colName -- name to search for
'Return: (long) column number of
'***********************************************************************************************
Function findColByName(headerRow As Long, colName As String, Optional returnString As Boolean = False)
    On Error GoTo FAILURE
    Dim numCols As Long: numCols = findLastColumn(headerRow)
    Dim nameFound As Boolean: nameFound = False
    For i = 1 To numCols Step 1
        If (Cells(headerRow, i) = colName) Then
            findColByName = i
            nameFound = True
            Exit For
        End If
    Next i
    
    If (Not nameFound) Then
        Err.Raise 1, , ""Could not find column named: "" & colName
    End If
    
    If (returnString) Then
        findColByName = convertColNumToLetter(findColByName)
    End If
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""findColByName("" & headerRow & "","" & colName & ""):"" & Err.Description
End Function

'***********************************************************************************************
'Description: Find the last column of a given row and returns its value
'Args: (long) rowNum -- number of the row to search in
'Return: (long) column number
'***********************************************************************************************
Function findLastColumn(rowNum As Long, Optional returnString As Boolean = False)
    On Error GoTo FAILURE
    Range(""XFD"" & rowNum).Select
    If (Selection = """") Then
        Selection.End(xlToLeft).Select
    End If
    
    If (ActiveCell.value = """") Then
        Err.Raise 1, , ""No values found in given row""
    End If
    If (returnString) Then
        findLastColumn = convertColNumToLetter(ActiveCell.column)
    Else
        findLastColumn = ActiveCell.column
    End If
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""findLastColumn():"" & Err.Description
End Function

Function findLastRow(col As Variant)
    On Error GoTo FAILURE
    Dim column As String, numCol As Long
    
    If (IsNumeric(col)) Then
        If (col > 16384 Or col < 1) Then
            Err.Raise 1, , ""Column number: "" & col & "" exceeds Excel column range 1 to 16384""
        End If
              
        column = convertColNumToLetter(col)
    Else
        If (col > ""XFD"" Or Len(col) > 3) Then
            Err.Raise 1, , ""Column Letter: "" & col & "" exceeds Excel column range A to XFD""
        End If
        column = col
    End If
        
    Range(column & 1048576).Select
    If (ActiveCell.value = """") Then
        Selection.End(xlUp).Select
    End If
    
    If (ActiveCell.value = """") Then
        Err.Raise 1, ""Src"", ""Cannot find last row, column is empty""
    End If
    
    findLastRow = ActiveCell.Row
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""findLastRow("" & col & ""):"" & Err.Description
End Function

'***********************************************************************************************
'Description: moves a columnm behind another
'Args: (String) sourceCol -- column to move, (String)targetCol --- target to move behind
'Return: none
'***********************************************************************************************
Function moveColumnBehind(sourceCol As String, targetCol As String)
    On Error GoTo FAILURE
    If (sourceCol > ""XFD"" Or targetCol > ""XFD"") Then
        Err.Raise 1, , ""Target or Source columns exceed Excel Column limit range""
    End If
    
    If (sourceCol = targetCol) Then
        Err.Raise 1, ""Source column equals target column""
    End If
    
    Columns(sourceCol & "":"" & sourceCol).Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns(targetCol & "":"" & targetCol).Select
    Selection.Insert Shift:=xlToRight
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""moveColumnBehindColumn("" & sourceCol & "","" & targetCol & ""):"" & Err.Description
End Function


'***********************************************************************************************
'Description: Deletes a column by the given letter
'Args: (string) colLetter -- letter of column to delete
'Return: none
'***********************************************************************************************
Function removeColByLetter(colLetter As String)
    On Error GoTo FAILURE
    Columns(colLetter & "":"" & colLetter).Select
    Selection.Delete Shift:=xlToLeft
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""removeColByLetter("" & colLetter & ""):"" & Err.Description
End Function


'***********************************************************************************************
'Description: Deletes a column by the given name
'Args: (String) colName -- name of column to delete
'Return: none
'***********************************************************************************************
Function removeColByName(headerRow As Long, colName As String)
    On Error GoTo FAILURE
    Dim colNum As Long: colNum = findColByName(headerRow, colName)
    Dim colLetter As String: colLetter = convertColNumToLetter(colNum)
    removeColByLetter (colLetter)
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""removeColByName("" & headerRow & "","" & colName & ""): "" & Err.Description
    
End Function

'***********************************************************************************************
'Description: Deletes a column by the given number
'Args: (long) colNum -- number of column to delete
'Return: none
'***********************************************************************************************
Function removeColByNum(colNum As Long)
    On Error GoTo FAILURE
    Dim colLetter As String: colLetter = convertColNumToLetter(colNum)
    removeColByLetter (colLetter)
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""removeColByNum("" & colNum & ""):"" & Err.Description
End Function

'***********************************************************************************************
'Description: Deletes given row
'Args: (long) rowNum -- number of the row to delete
'Return: none
'***********************************************************************************************
Function removeRow(rowNum As Long)
    On Error GoTo FAILURE
    Rows(rowNum & "":"" & rowNum).Select
    Selection.Delete Shift:=xlUp
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""removeRow():"" & Err.Description
End Function


'***********************************************************************************************
'Description: Resizes all columns
'Args: none
'Return: none
'***********************************************************************************************
Function resizeAll()
    On Error GoTo FAILURE
    Cells.Select
    Cells.EntireColumn.AutoFit
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, ""resizeAll():"" & Err.Description
End Function

"
