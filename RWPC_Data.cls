'exddint a;lksdjf;klasdjfd;lkjds'
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RWPC_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************
'Description remove duplicates based on columnms 1 and 2 of a given header row, this needs work
'Args: (long) headerRow -- row with headers, (Variant()) array of criteria columns
'Return: none
'*******************************************************************
Function removeDuplicatesWithHeader(headerRow As Long, colToFindRowCount As String, rowCol As RWPC_rowCol, criteriaCols As Variant)
    On Error GoTo FAILURE
    Dim lastRow As Long: lastRow = rowCol.findLastRowByColLetter(colToFindRowCount)
    Dim lastCol As String: lastCol = rowCol.convertColNumToLetter(rowCol.findLastColumn(headerRow))
       
    ActiveSheet.Range("$A$" & headerRow & ":$" & lastCol & "$" & lastRow).removeDuplicates Columns:=(criteriaCols), _
        Header:=xlYes
    Exit Function
FAILURE:
    Err.Raise Err.Number, Err.source, "removeDuplicates():" & Err.Description
End Function


