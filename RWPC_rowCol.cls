VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RWPC_rowCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************************************
'Description: This function converts a supplied long into its corresponding column letter
'Args: (long) colNum -- number of the column
'Return: (string) column letter
'***********************************************************************************************
Function convertColNumToLetter(colNum As Long)
    On Error GoTo FAILURE
    If (colNum > 16384 Or colNum < 1) Then
       Err.Raise 1, , "Column number: " & colNum & " exceeds Excel range of 1 to 16384"
    End If

    Dim n As Long
    Dim C As Byte
    Dim s As String

    n = colNum
    Do
        C = ((n - 1) Mod 26)
        s = Chr(C + 65) & s
        n = (n - C) \ 26
    Loop While n > 0
    convertColNumToLetter = s
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "convertColNumToLetter():" & Err.Description
End Function

'***********************************************************************************************
'Description: Searches for column name in given header row number
'Args: (long) rowNum -- row with headers, (String) colName -- name to search for
'Return: (long) column number of
'***********************************************************************************************
Function findColByName(headerRow As Long, colName As String)
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
        findColByName = -1
    End If
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "findColByName(" & headerRow & "," & colName & "):" & Err.Description
End Function

'***********************************************************************************************
'Description: Find the last column of a given row and returns its value
'Args: (long) rowNum -- number of the row to search in
'Return: (long) column number
'***********************************************************************************************
Function findLastColumn(rowNum As Long)
    On Error GoTo FAILURE
    Range("XFD" & rowNum).Select
    If (Selection = "") Then
        Selection.End(xlToLeft).Select
    End If
    
    If (ActiveCell.Value = "") Then
        Err.Raise 1, , "No values found in given row"
    End If
    findLastColumn = ActiveCell.column
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "findLastColumn():" & Err.Description
End Function

'***********************************************************************************************
'Description: This function obtains the row count for the provided column
'Args: (String) col -- letter of the column
'Return: (long)row count
'***********************************************************************************************
Function findLastRowByColLetter(col As String)
    On Error GoTo FAILURE:
    Dim lRow As Long, lCol As Long
    
    If (col > "XFD" Or Len(col) > 3) Then
       Err.Raise 1, , "Column letter: " & col & " exceeds Excel column range A to XFD"
    End If
    
    Range(col & 1048576).Select
    If (ActiveCell.Value = "") Then
        Selection.End(xlUp).Select
    End If
    
    If (ActiveCell.Value = "") Then
        Err.Raise 1, "Src", "Cannot find last row, column is empty"
    End If
    findLastRowByColLetter = ActiveCell.row
    
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "findLastRowByColLetter(" & col & "):" & Err.Description
End Function

'***********************************************************************************************
'Description: This function obtains the row count for the provided column, because of VBA
'non-support of overloaded functions, this is a pass through function that converts the number
'supplied to a long and then calls findLastRowByColLetter
'Args: (long) colNum -- number of the column
'Return: (long)row count
'***********************************************************************************************
Function findLastRowByColNum(colNum As Long, rowCol As RWPC_rowCol)
    On Error GoTo FAILURE
    If (colNum > 16384 Or colNum < 1) Then
       Err.Raise 1, , "Column number: " & colNum & " exceeds Excel range of 1 to 16384"
    End If
    
    findLastRowByColNum = findLastRowByColLetter(rowCol.convertColNumToLetter(colNum))
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "findLastRowByColNum():" & Err.Description
End Function

'***********************************************************************************************
'Description: moves a columnm behind another
'Args: (String) sourceCol -- column to move, (String)targetCol --- target to move behind
'Return: none
'***********************************************************************************************
Function moveColumnBehindColumn(sourceCol As String, targetCol As String)
    On Error GoTo FAILURE
    If (sourceCol > "XFD" Or targetCol > "XFD") Then
        Err.Raise 1, , "Target or Source columns exceed Excel Column limit range"
    End If
    
    If (sourceCol = targetCol) Then
        Err.Raise 1, "Source column equals target column"
    End If
    
    Columns(sourceCol & ":" & sourceCol).Select
    Application.CutCopyMode = False
    Selection.Cut
    Columns(targetCol & ":" & targetCol).Select
    Selection.Insert Shift:=xlToRight
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "moveColumnBehindColumn(" & sourceCol & "," & targetCol & "):" & Err.Description
End Function

'***********************************************************************************************
'Description: Deletes a column by the given letter
'Args: (string) colLetter -- letter of column to delete
'Return: none
'***********************************************************************************************
Function removeColByLetter(colLetter As String)
    On Error GoTo FAILURE
    Columns(colLetter & ":" & colLetter).Select
    Selection.Delete Shift:=xlToLeft
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "removeColByLetter(" & colLetter & "):" & Err.Description
End Function

'***********************************************************************************************
'Description: Deletes a column by the given name
'Args: (String) colName -- name of column to delete
'Return: none
'***********************************************************************************************
Function removeColByName(headerRow As Long, colName As String, rowCol As RWPC_rowCol)
    On Error GoTo FAILURE
    Dim colNum As Long: colNum = rowCol.findColByName(headerRow, colName)
    Dim colLetter As String: colLetter = rowCol.convertColNumToLetter(colNum)
    removeColByLetter (colLetter)
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "removeColByName(" & headerRow & "," & colName & "): " & Err.Description
    
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
    Err.Raise Err.Number, Err.source, "removeColByNum(" & colNum & "):" & Err.Description
End Function

'***********************************************************************************************
'Description: Deletes given row
'Args: (long) rowNum -- number of the row to delete
'Return: none
'***********************************************************************************************
Function removeRow(rowNum As Long)
    On Error GoTo FAILURE
    Rows(rowNum & ":" & rowNum).Select
    Selection.Delete Shift:=xlUp
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "removeRow():" & Err.Description
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
    Err.Raise Err.Number, Err.source, "resizeAll():" & Err.Description
End Function

