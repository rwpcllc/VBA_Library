VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RWPC_Library"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public FileIO As RWPC_FileIO
Public Sheets As RWPC_Sheets
Public rowCol As RWPC_rowCol
Public regex As RWPC_Regex
Public data As RWPC_Data

Public Function initFullLib()
    Set FileIO = New RWPC_FileIO
    Set Sheets = New RWPC_Sheets
    Set rowCol = New RWPC_rowCol
    Set regex = New RWPC_Regex
    Set data = New RWPC_Data
End Function
