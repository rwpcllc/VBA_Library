Attribute VB_Name = "UnitTests"
Private RWPC_LIB As RWPC_Library
Private TESTS As Variant

Public Sub runallTests()
    Dim rowCtr As Integer: rowCtr = 2
    setup
    loadTests
    
    For i = 0 To (UBound(TESTS)) Step 1
        Cells(rowCtr, 1).Value = TESTS(i)
        Cells(rowCtr, 2).Value = Application.Run("TEST_" & TESTS(i))
        rowCtr = rowCtr + 1
    Next i
    
    wrapup
    
End Sub

Private Function loadTests()
    TESTS = Array("RegexMatch", "RegexMatchBool", "removeDuplicatesWithHeader", "RegexMatch", "RegexMatchBool", "convertColNumToLetter", "findColByName", _
        "findLastColumn", "findLastRowByColLetter", "findLastRowByColNum", "moveColumnBehindColumn", "removeColByLetter", "removeColByName", "removeColByNum", _
        "removeRow", "resizeAll", "spreadsheetExists", "spreadsheetIsEmpty")
    
End Function

'DATA Library Tests
Private Function TEST_removeDuplicatesWithHeader()
    On Error GoTo FAILURE
    TEST_removeDuplicatesWithHeader = "PASSED"
    Sheets.Add
    Cells(1, 1) = "A"
    Cells(1, 2) = "R"
    Cells(2, 1) = "B"
    Cells(2, 2) = "R"
    Cells(3, 1) = "B"
    Cells(3, 2) = "R"
    Cells(4, 1) = "C"
    Cells(4, 2) = "R"
    RWPC_LIB.data.removeDuplicatesWithHeader 1, "A", RWPC_LIB.rowCol, Array(1, 2)
    If (Not Cells(3, 1) = "C") Then
        TEST_removeDuplicatesWithHeader = "FAILED"
    End If
    DelTestSheet
    
    Exit Function
FAILURE:
    TEST_removeDuplicatesWithHeader = "FAILED"
    DelTestSheet
    
End Function





'****************************************************************************************************************************
'REGEX Library Tests
'****************************************************************************************************************************
Private Function TEST_RegexMatch()
    On Error GoTo FAILURE
    TEST_RegexMatch = "PASSED"
    Dim result As String: result = RWPC_LIB.regex.RegexMatch("Sample text 03-05-2019", "\d{2}-\d{2}-\d{4}")
    If (Not result = "03-05-2019") Then
        TEST_RegexMatch = "FAILED"
    End If
    Exit Function
FAILURE:
    TEST_RegexMatch = "FAILED"
End Function

Private Function TEST_RegexMatchBool()
    On Error GoTo FAILURE
    TEST_RegexMatchBool = "PASSED"
    If (Not RWPC_LIB.regex.RegexMatchBool("Sample text 2019-03-02.", "\d{4}-\d{2}-\d{2}")) Then
        TEST_RegexMatchBool = "FAILED"
    End If
    Exit Function
FAILURE:
    TEST_RegexMatchBool = "FAILED"
End Function







'****************************************************************************************************************************
'rowCol Library Tests
'****************************************************************************************************************************
Function TEST_convertColNumToLetter()
    On Error GoTo EXPECT_FAILURE1
    TEST_convertColNumToLetter = "PASSED"
    RWPC_LIB.rowCol.convertColNumToLetter 20000
    TEST_convertColNumToLetter = "FAILED"
    Exit Function
EXPECT_FAILURE1:
    Resume NEXT_TEST1
NEXT_TEST1:
    On Error GoTo FAILURE
    convertColNumToLetter = "PASSED"
    Dim result As String: result = RWPC_LIB.rowCol.convertColNumToLetter(27)
    If (Not result = "AA") Then
        convertColNumToLetter = "FAILED"
    End If
    Exit Function
FAILURE:
    TEST_convertColNumToLetter = "FAILED"
End Function

Function TEST_findColByName()
    On Error GoTo FAILURE
    TEST_findColByName = "PASSED"
    Sheets.Add
    Cells(1, 1) = "Col1"
    Cells(1, 2) = "Col2"
    Dim result As Long: result = RWPC_LIB.rowCol.findColByName(1, "Dog")
    If (Not result = -1) Then
        TEST_findColByName = "FAILED"
    End If
    
    result = RWPC_LIB.rowCol.findColByName(1, "Col2")
    If (Not result = 2) Then
        TEST_findColByName = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    TEST_findColByName = "FAILED"
    DelTestSheet
End Function

Function TEST_findLastColumn()
    On Error GoTo EXPECT_FAILURE1
    TEST_findLastColumn = "PASSED"
    Sheets.Add
    RWPC_LIB.rowCol.findLastColumn 1
    TEST_findLastColumn = "FAILED"
    DelTestSheet
    Exit Function
EXPECT_FAILURE1:
    Resume NEXT_TEST1
NEXT_TEST1:
    On Error GoTo FAILURE
    Cells(1, 1) = "COL 1"
    Cells(1, 2) = "Col 2"
    Cells(1, 3) = "Col 3"
    Dim result As Long: result = RWPC_LIB.rowCol.findLastColumn(1)
    If (Not result = 3) Then
        TEST_findLastColumn = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    TEST_findLastColumn = "FAILED"
    DelTestSheet
End Function

Function TEST_findLastRowByColLetter()
    'Expect an error if sheet is empty
    On Error GoTo EXPECT_FAILURE1
    TEST_findLastRowByColLetter = "PASSED"
    Sheets.Add
    RWPC_LIB.rowCol.findLastRowByColLetter ("A")
    TEST_findLastRowByColLetter = "FAILED"
    DelTestSheet
    Exit Function
EXPECT_FAILURE1:
    Resume NEXT_TEST1
NEXT_TEST1:
    'Expect an error if invalid column (from valid column range in Excel) is supplied
    On Error GoTo EXPECT_FAILURE2
    TEST_findLastRowByColLetter = "PASSED"
    RWPC_LIB.rowCol.findLastRowByColLetter "XXXXXZ"
    TEST_findLastRowByColLetter = "FAILED"
    DelTestSheet
    Exit Function
EXPECT_FAILURE2:
    Resume NEXT_TEST2
NEXT_TEST2:
    On Error GoTo FAILURE
    TEST_findLastRowByColLetter = "PASSED"
    Cells(1, 1) = "Data1"
    Cells(2, 1) = "Data2"
    Cells(3, 1) = "Data3"
    Dim result As Long: result = RWPC_LIB.rowCol.findLastRowByColLetter("A")
    If (Not result = 3) Then
        TEST_findLastRowByColLetter = "FAILED"
    End If
    
    'Test that if data fills all the way to the bottom, the bottom column is returned
    Cells(1048576, 1) = "Data1048576"
    result = RWPC_LIB.rowCol.findLastRowByColLetter("A")
    If (Not result = 1048576) Then
        TEST_findLastRowByColLetter = "FAILED"
    End If
    
    DelTestSheet
    Exit Function
FAILURE:
    TEST_findLastRowByColLetter = "FAILED"
    DelTestSheet
End Function

Function TEST_findLastRowByColNum()
    'Exepct error if empty sheet is searched for last row
    On Error GoTo EXPECT_ERROR1
    TEST_findLastRowByColNum = "PASSED"
    Sheets.Add
    RWPC_LIB.rowCol.findLastRowByColNum 1, RWPC_LIB.rowCol
    TEST_findLastRowByColNum = "FAILED"
    DelTestSheet
    Exit Function
EXPECT_ERROR1:
    Resume NEXT_TEST1
NEXT_TEST1:
    'Expect error if invalid column number is provided
    On Error GoTo EXPECT_ERROR2
    RWPC_LIB.rowCol.findLastRowByColNum 20000, RWPC_LIB.rowCol
    TEST_findLastRowByColNum = "FAILED"
    DelTestSheet
    Exit Function
EXPECT_ERROR2:
    Resume NEXT_TEST2
NEXT_TEST2:
    On Error GoTo FAILURE
    TEST_findLastRowByColNum = "PASSED"
    Cells(1, 2) = "A"
    Cells(2, 2) = "B"
    Cells(3, 2) = "C"
    Cells(4, 2) = "D"
    Dim result As Long: result = RWPC_LIB.rowCol.findLastRowByColNum(2, RWPC_LIB.rowCol)
    If (Not result = 4) Then
        TEST_findLastRowByColNum = "FAILED"
    End If
    
    Cells(1048576, 2) = "ZZ"
    result = RWPC_LIB.rowCol.findLastRowByColNum(2, RWPC_LIB.rowCol)
    If (Not result = 1048576) Then
        TEST_findLastRowByColNum = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    TEST_findLastRowByColNum = "FAILED"
    DelTestSheet
End Function

Function TEST_moveColumnBehindColumn()
    On Error GoTo EXPECT_FAILURE1
    TEST_moveColumnBehindColumn = "PASSED"
    Sheets.Add
    Cells(1, 1) = "Col A"
    Cells(1, 2) = "Col B"
    Cells(1, 3) = "Col C"
    Cells(1, 4) = "Col D"
    'Provide invalid columns
    RWPC_LIB.rowCol.moveColumnBehindColumn "XXXX", "XXXXZ"
    TEST_moveColumnBehindColumn = "FAILED"
    Exit Function
EXPECT_FAILURE1:
    Resume NEXT_TEST1
NEXT_TEST1:
    On Error GoTo EXPECT_FAILURE2
    RWPC_LIB.rowCol.moveColumnBehindColumn "A", "A"
    TEST_moveColumnBehindColumn = "FAILED"
    DelTestSheet
    Exit Function
EXPECT_FAILURE2:
    Resume NEXT_TEST2
NEXT_TEST2:
    On Error GoTo FAILURE
    RWPC_LIB.rowCol.moveColumnBehindColumn "C", "B"
    If (Not Cells(1, 2) = "Col C") Then
        TEST_moveColumnBehindColumn = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    TEST_moveColumnBehindColumn = "FAILED"
    DelTestSheet
End Function

Function TEST_removeColByLetter()
    On Error GoTo EXPECT_FAILURE1
    TEST_removeColByLetter = "PASSED"
    Sheets.Add
    Cells(1, 1) = "Col A"
    Cells(1, 2) = "Col B"
    Cells(1, 3) = "Col C"
    'Expect error raised when removing invalid column
    RWPC_LIB.rowCol.removeColByLetter "XXXX"
    TEST_removeColByLetter = "FAILED"
    Exit Function
EXPECT_FAILURE1:
    Resume NEXT_TEST1
NEXT_TEST1:
    On Error GoTo FAILURE
    RWPC_LIB.rowCol.removeColByLetter "B"
    If (Not Cells(1, 2) = "Col C") Then
        TEST_removeColByLetter = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    TEST_removeColByLetter = "FAILED"
    DelTestSheet
End Function

Function TEST_removeColByName()
    On Error GoTo EXPECT_FAILURE1
    TEST_removeColByName = "PASSED"
    'REMOVE BAD COL
    Sheets.Add
    Cells(1, 1) = "Col1"
    Cells(1, 2) = "Col2"
    Cells(1, 3) = "Col3"
    
    RWPC_LIB.rowCol.removeColByName 1, "Dog", RWPC_LIB.rowCol
    TEST_removeColByName = "FAILED"
    DelTestSheet
    Exit Function
EXPECT_FAILURE1:
    Resume NEXT_TEST1
NEXT_TEST1:
    On Error GoTo FAILURE
    RWPC_LIB.rowCol.removeColByName 1, "Col2", RWPC_LIB.rowCol
    If (Not Cells(1, 2) = "Col3") Then
        TEST_removeColByName = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    MsgBox (Err.Description)
    TEST_removeColByName = "FAILED"
    DelTestSheet
End Function

Function TEST_removeColByNum()
    On Error GoTo EXPECT_FAILURE1
    TEST_removeColByNum = "PASSED"
    Sheets.Add
    'Test removing bad col, expect error thrown
    RWPC_LIB.rowCol.removeColByNum 20000
    TEST_removeColByNum = "FAILED"
    DelTestSheet
    Exit Function
EXPECT_FAILURE1:
    Resume NEXT_TEST1
NEXT_TEST1:
    On Error GoTo FAILURE
    Cells(1, 1) = "Col A"
    Cells(1, 2) = "Col B"
    Cells(1, 3) = "Col C"
    RWPC_LIB.rowCol.removeColByNum 2
    If (Not Cells(1, 2) = "Col C") Then
        TEST_removeColByNum = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    TEST_removeColByNum = "FAILED"
    DelTestSheet
End Function

Function TEST_removeRow()
    On Error GoTo EXPECT_FAILURE1
    TEST_removeRow = "PASSED"
    Sheets.Add
    Cells(1, 1) = "Row 1"
    Cells(2, 1) = "Row 2"
    Cells(3, 1) = "Row 3"
    'Test removing bad row, expect an error
    RWPC_LIB.rowCol.removeRow 2000000
    TEST_removeRow = "FAILED"
    DelTestSheet
    Exit Function
EXPECT_FAILURE1:
    Resume NEXT_TEST1
NEXT_TEST1:
    On Error GoTo FAILURE
    RWPC_LIB.rowCol.removeRow 2
    If (Not Cells(2, 1) = "Row 3") Then
        TEST_removeRow = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    TEST_removeRow = "FAILED"
    DelTestSheet
End Function

Function TEST_resizeAll()
    On Error GoTo FAILURE
    TEST_resizeAll = "PASSED"
    Sheets.Add
    Cells(1, 1) = "AB"
    RWPC_LIB.rowCol.resizeAll
    Range("A1").Select
    If (ActiveCell.Width = 2.33) Then
        TEST_resizeAll = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    TEST_resizeAll = "FAILED"
    DelTestSheet
End Function







'****************************************************************************************************************************
'sheets Library Tests
'****************************************************************************************************************************
Function TEST_spreadsheetExists()
    On Error GoTo FAILURE
    TEST_spreadsheetExists = "PASSED"
    Sheets.Add
    ActiveSheet.name = "DOG"
    Dim result As Boolean: result = RWPC_LIB.Sheets.spreadsheetExists("DOG")
    If (Not result) Then
        TEST_spreadsheetExists = "FAILED"
    End If
    
    result = RWPC_LIB.Sheets.spreadsheetExists("CAT")
    If (result) Then
        TEST_spreadsheetExists = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    TEST_spreadsheetExists = "FAILED"
    DelTestSheet
End Function

Function TEST_spreadsheetIsEmpty()
    On Error GoTo FAILURE
    TEST_spreadsheetIsEmpty = "PASSED"
    Sheets.Add
    ActiveSheet.name = "DOG"
    Dim result As Boolean: result = RWPC_LIB.Sheets.spreadsheetIsEmpty("DOG")
    If (Not result) Then
        TEST_spreadsheetIsEmpty = "FAILED"
    End If
    Cells(1, 1) = "Col A"
    
    result = RWPC_LIB.Sheets.spreadsheetIsEmpty("DOG")
    If (result) Then
        TEST_spreadsheetIsEmpty = "FAILED"
    End If
    DelTestSheet
    Exit Function
FAILURE:
    TEST_spreadsheetIsEmpty = "FAILED"
    DelTestSheet
End Function





'******************************************************************************************************************************
'Setup functions for testing library
'******************************************************************************************************************************
Private Function setup()
    Set RWPC_LIB = New RWPC_Library
    RWPC_LIB.initFullLib
    If (Evaluate("ISREF('Unit Testing Results'!A1)")) Then
        Application.DisplayAlerts = False
        Sheets("Unit Testing Results").Delete
        Application.DisplayAlerts = True
    End If
    
    Sheets.Add
    Application.ActiveSheet.name = "Unit Testing Results"
    Cells(1, 1).Value = "Function"
    Cells(1, 2).Value = "Result"
End Function


Private Function DelTestSheet()
    Application.DisplayAlerts = False
    If (Not ActiveSheet.name = "Unit Testing Results") Then
        ActiveWindow.SelectedSheets.Delete
    End If
    Application.DisplayAlerts = True
    Sheets("Unit Testing Results").Select
End Function

Private Function wrapup()
    Cells.Select
    Cells.EntireColumn.AutoFit
End Function
