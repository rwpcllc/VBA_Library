'*******************************************************************
'Description remove duplicates based on columnms 1 and 2 of a given header row, this needs work
'Args: (long) headerRow -- row with headers, (Variant()) array of criteria columns
'Return: none
'*******************************************************************
Function removeDuplicatesWithHeader(headerRow As Long, colToFindRowCount As String, rowCol As RWPC_ROWCOL, criteriaCols As Variant)
    On Error GoTo FAILURE
    Dim lastRow As Long: lastRow = rowCol.findLastRow(colToFindRowCount)
    Dim lastCol As String: lastCol = rowCol.convertColNumToLetter(rowCol.findLastColumn(headerRow))
       
    ActiveSheet.Range("$A$" & headerRow & ":$" & lastCol & "$" & lastRow).RemoveDuplicates Columns:=(criteriaCols), _
        Header:=xlYes
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "removeDuplicates():" & Err.Description
End Function

'*******************************************************************
'Description Summarizes Data By Provided Columns into a HashMap
'Args: (variant) criteraCols -- array of column(s) to summarize by, (variant) sumCol -- column to summaraize values on
'Return: (Dictionary) map of summarized values
'*******************************************************************
Function summarizeDataByColumns(criteriaCols As Variant, sumCol As Variant, Optional headers As Boolean = True, Optional delimeter As String = "~")
    On Error GoTo FAILURE
    Dim rowCol As RWPC_ROWCOL, startRow As Long, hash As Dictionary
    Set rowCol = New RWPC_ROWCOL
    Set hash = New Dictionary
        
    'If Letter was provided convert to number
    If (VarType(sumCol) = 8) Then
        Dim sumColumn As String: sumColumn = sumCol
        sumCol = rowCol.convertColLetterToNum(sumColumn)
    End If
    
    startRow = IIf(headers, 2, 1)
    lastRow = rowCol.findLastRow(sumCol)
    
    For i = startRow To lastRow Step 1
        Dim key As String: key = ""
        For j = 0 To UBound(criteriaCols) Step 1
            If (j = 0) Then
                key = key & Cells(i, criteriaCols(j))
            Else
                key = key & delimeter & Cells(i, criteriaCols(j))
            End If
        Next j
        
        Dim value As Variant
        value = Cells(i, sumCol)
        
        If (hash.Exists(key)) Then
            hash(key) = hash(key) + value
        Else
            hash.Add key, value
        End If
    Next i
    
    Set summarizeDataByColumns = hash
    
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "summarizeDataByColumns():" & Err.Description
End Function


