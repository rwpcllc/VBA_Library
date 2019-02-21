Attribute VB_Name = "BIO"
Dim importTypes() As Variant

Sub AID_Main_Account_Maintenance()
Dim currSheet As String
currSheet = Application.ActiveSheet.Name

'load config
LoadConfig

'get filepath
Dim filePath As String
filePath = RWPC_Library.PromptForFileLocation

'import file
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