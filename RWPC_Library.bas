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

Function PromptForFileLocation()
Dim picker As Object
Set picker = Application.FileDialog(msoFileDialogFilePicker)
picker.AllowMultiSelect = False
picker.Filters.Add "ALL", "*.*", 1
picker.Show

If (picker.SelectedItems.Count > 0) Then
    PromptForFileLocation = picker.SelectedItems.Item(1)
Else
    PromptForFileLocation = "Null"
End If

End Function
