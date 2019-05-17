Attribute VB_Name = "BIO"
Dim importTypes() As Variant
'some comment'
Sub AID_Main_Account_Maintenance()
Dim currSheet As String
currSheet = Application.ActiveSheet.Name

'load config
LoadConfig

'if the active sheet is not the config sheet && totally blank use that sheet, otherwise create a new sheet

'get filepath
Dim filePath As String
filePath = RWPC_Library.PromptForFileLocation

'import file
If (Not spreadsheetIsEmpty(currSheet)) Then
    MsgBox "The current sheet is not blank, adding a new sheet to complete the import", vbInformation
    Sheets.Add
    currSheet = Application.ActiveSheet.Name
End If
 
Sheets(currSheet).Select

RWPC_Library.ImportCSV filePath, importTypes

End Sub

Function LoadConfig()
Sheets("Config").Select
Dim friendlyTypes As Variant
friendlyTypes = Array("General", _
    "Text", _
    "Date: DMY", _
    "Date: DYM", _
    "Date: MDY", _
    "Date: MYD", _
    "Date: YDM", _
    "Date: YMD", _
    "Skip Column")

importTypes = Array()

Dim rowCount As Long
rowCount = RWPC_Library.findLastRowByColLetter("A")

For i = 2 To rowCount
    ReDim Preserve importTypes(UBound(importTypes) + 1)
    importTypes(UBound(importTypes)) = Application.Match(Cells(i, 2), friendlyTypes, False)
Next i

End Function
