VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RWPC_FileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'*********************************************************************************************
'Description: Function prompts user to pick a file on their computer/mapped/networked drive
'  The function then returns the file path, or null if nothing was selected.
'*********************************************************************************************
Function PromptForFileLocation()
    'On Error GoTo FAILURE
    Dim picker As Object
    Set picker = Application.FileDialog(msoFileDialogFilePicker)
    picker.AllowMultiSelect = False
    picker.Filters.Add "ALL", "*.*", 1
    picker.Show
    
    'If (picker.SelectedItems.Count > 0) Then
    '    PromptForFileLocation = picker.SelectedItems.Item(1)
    'Else
    '    PromptForFileLocation = "null"
    'End If
FAILURE:
    Err.Raise Err.Number, Err.source, "PromptForfileLocation()" & Err.Description

End Function


'***********************************************************************************************
'Description: Imports CSV file specified in filepath var, and uses the colTypes array to determine
' column types when importing {General, text date, etc.}
'Args: (String) filePath, (Variant Arr) colTypes
'Return: none
'***********************************************************************************************
Function ImportCSV(filePath As String, colTypes As Variant)
    On Error GoTo FAILURE
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
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "ImportCSV(" & filePath & "):" & Err.Description
End Function
'*******************************************************************
'Description Saves file at given filepath
'Args: (String) filePath
'Return: none
'*******************************************************************
Function SaveFile(filePath As String)
    On Error GoTo FAILURE
    ActiveWorkbook.SaveAs filename:= _
        filePath, FileFormat _
        :=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "SaveFile(" & filePath & "}: " & Err.Description
End Function
'*****************************************************************************
'Description: Prompts user with filepicker to get folder to save a file in
'Args: none
'Return: (String) path to folder
'*****************************************************************************
Function PromptForSaveLocation()
    On Error GoTo FAILURE
    Dim picker As Object
    Set picker = Application.FileDialog(msoFileDialogFolderPicker)
    picker.AllowMultiSelect = False
    picker.Show
    
    If (picker.SelectedItems.Count > 0) Then
        PromptForSaveLocation = picker.SelectedItems.Item(1)
    Else
        Err.Raise 1, "Source", "User did not select a save location."
    End If
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "PromptForSaveLocation():" & Err.Description
End Function

