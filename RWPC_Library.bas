Attribute VB_Name = "RWPC_Library"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"

'Prompt user for file path (treasury file)
    Dim filePath As String
    filePath = PromptForFileLocation()
    If (filePath = "Null") Then
        MsgBox ("No file chosen! Aborting")
        Exit Sub
    End If
    
'Import treasury file as text for each column
    Application.CutCopyMode = False
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & filePath, Destination:= _
        Range("$A$1"))
        .Name = "all_tas_betc"
        .FieldNames = True
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    MsgBox (ActiveSheet.QueryTables(1).ResultRange.Columns.Count)
    
End Sub
'*********************************************************************************************
'Description: Function prompts user to pick a file on their computer/mapped/networked drive
'  The function then returns the file path, or null if nothing was selected.
'*********************************************************************************************
Function PromptForFileLocation()
Dim picker As Object
Set picker = Application.FileDialog(msoFileDialogFilePicker)
picker.AllowMultiSelect = False
picker.Filters.Add "ALL", "*.*", 1
picker.Show

If (picker.SelectedItems.Count > 0) Then
    PromptForFileLocation = picker.SelectedItems.Item(1)
Else
    PromptForFileLocation = "null"
End If

End Function

'***********************************************************************************************
'Description: This function obtains the row count for the provided column
'Args: (String) col -- letter of the column
'Return: (long)row count
'***********************************************************************************************
Function findLastRowByColLetter(col As String)
    Dim lRow As Long, lCol As Long
    
    If (col > "XFD" Or Len(col) > 3) Then
       MsgBox ("Column letter: " & col & " exceeds Excel column range A to XFD")
       findLastRowByColLetter = "null"
       Exit Function
    End If
    
    lRow = Cells.Find(What:="*", _
        After:=Range(col & "1"), _
        LookAt:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).row

    findLastRowByColLetter = lRow
    
End Function
'***********************************************************************************************
'Description: This function obtains the row count for the provided column, because of VBA
'non-support of overloaded functions, this is a pass through function that converts the number
'supplied to a long and then calls findLastRowByColLetter
'Args: (long) colNum -- number of the column
'Return: (long)row count
'***********************************************************************************************
Function findLastRowByColNum(colNum As Long)
    If (colNum > 16384 Or colNum < 1) Then
       MsgBox "Column number: " & colNum & " exceeds Excel range of 1 to 16384", vbOKOnly
       findLastRowByColNum = "null"
       Exit Function
    End If
    
    findLastRowByColNum = findLastRowByColLetter(convertColNumToLetter(colNum))

End Function

'***********************************************************************************************
'Description: This function converts a supplied long into its corresponding column letter
'Args: (long) colNum -- number of the column
'Return: (string) column letter
'***********************************************************************************************
Function convertColNumToLetter(colNum As Long)
    If (colNum > 16384 Or colNum < 1) Then
       MsgBox "Column number: " & colNum & " exceeds Excel range of 1 to 16384", vbOKOnly
       convertColNumToLetter = "null"
       Exit Function
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

End Function
'***********************************************************************************************
'Description: Imports CSV file specified in filepath var, and uses the colTypes array to determine
' column types when importing {General, text date, etc.}
'Args: (String) filePath, (Variant Arr) colTypes
'Return: none
'***********************************************************************************************
Function ImportCSV(filePath As String, colTypes As Variant)
Application.CutCopyMode = False
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & filePath, Destination:= _
        Range("$A$1"))
        .FieldNames = True
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = colTypes
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Function

