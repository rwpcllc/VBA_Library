"
Sub writeLibraryOut()
Dim doExport As Boolean: doExport = False
Dim fileName As String
Dim filePath As String: filePath = ""C:\Users\Patrick Camargo\Desktop\VBA_Library\""
Dim exportFile As String
Dim source As VBIDE.CodeModule

For Each element In ActiveWorkbook.VBProject.VBComponents
    doExport = True
    Select Case element.Type
        Case vbext_ct_ClassModule
            fileName = element.Name & "".cls""
        Case vbext_ct_MSForm
            fileName = element.Name & "".frm""
        Case vbext_ct_StdModule
            fileName = element.Name & "".bas""
        Case Else
            doExport = False
    End Select
    If (doExport) Then
        exportFile = filePath & fileName
        Set source = ActiveWorkbook.VBProject.VBComponents(element.Name).CodeModule
        Open exportFile For Output As #1
        Write #1, source.Lines(1, source.CountOfLines)
        Close #1
    End If
Next


End Sub"
