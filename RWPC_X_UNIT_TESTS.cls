"Private formatUtil As RWPC_FORMATTING
Private dataUtil As RWPC_DATA
Private rowColUtil As RWPC_ROWCOL
Private regexUtil As RWPC_REGEX
Private sheetsUtil As RWPC_SHEETS

Private TEST_CTR As Long

Public Function runTests()
    Set rowColUtil = New RWPC_ROWCOL
    If (Evaluate(""ISREF('UnitTestResults'!A1)"")) Then
        Application.DisplayAlerts = False
        Sheets(""UnitTestResults"").Delete
        Application.DisplayAlerts = True
    End If
        
    Sheets.Add
    ActiveSheet.Name = ""UnitTestResults""
    Cells(1, 1) = ""Test""
    Cells(1, 2) = ""Result""
    TEST_CTR = 1
    
    formattingTests
    dataTests
    rowcolTests
    regexTests
    sheetsTests
    
    rowColUtil.resizeAll
    Cells(1, 1).Select
End Function




'******************************************************************************************
'*FORMATTING TESTS
'* - testCopyColumnFormat
'******************************************************************************************
Private Function formattingTests()
    Set formatUtil = New RWPC_FORMATTING
    Sheets(""UnitTestResults"").Select
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test FORMATTING.copyColumnFormat""
    Cells(TEST_CTR, 2) = testCopyColumnFormat
    
End Function


Private Function testCopyColumnFormat()
    On Error GoTo FAILURE
    testCopyColumnFormat = ""PASSED""
    Sheets.Add
    Cells(1, 1) = ""1""
    Range(""A1"").Select
    Selection.NumberFormat = ""$#,##0.00""
    
    Cells(1, 2) = 3
    
    formatUtil.copyColumnFormat ""A"", ""B""
    
    Range(""B1"").Select
    
    If (Not Selection.NumberFormat = ""$#,##0.00"") Then
        testCopyColumnFormat = ""FAILED""
    End If
    delSheet
    Exit Function
FAILURE:
    delSheet
    Err.Raise Err.Number, Err.source, ""Failed Unit Test testCopyColumnFormat():"" & Err.Description
End Function



'******************************************************************************************
'*DATA TESTS
'* - testRemoveDuplicatesWithHeader
'* - testSummarizeDataByColumns

'******************************************************************************************
Private Function dataTests()
    Set dataUtil = New RWPC_DATA
    Sheets(""UnitTestResults"").Select
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test DATA.removeDuplicatesWithHeader""
    Cells(TEST_CTR, 2) = testRemoveDuplicatesWithHeader
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test DATA.summarizeDataByColumns""
    Cells(TEST_CTR, 2) = testSummarizeDataByColumns
    
End Function

Private Function testRemoveDuplicatesWithHeader()
    On Error GoTo FAILURE
    testRemoveDuplicatesWithHeader = ""PASSED""
    Sheets.Add
    Cells(1, 1) = ""ColA""
    Cells(1, 2) = ""ColB""
    Cells(2, 1) = ""A""
    Cells(2, 2) = ""B""
    Cells(3, 1) = ""A""
    Cells(3, 2) = ""B""
    Cells(4, 1) = ""B""
    Cells(4, 2) = ""A""
    
    Dim rowCol As RWPC_ROWCOL
    Set rowCol = New RWPC_ROWCOL
    dataUtil.removeDuplicatesWithHeader 1, ""A"", rowCol, Array(1, 2)
    
    If (Not Cells(3, 1) = ""B"" Or Not Cells(3, 2) = ""A"") Then
        testRemoveDuplicatesWithHeader = ""FAILED""
        Err.Raise 1, , ""Failed testRemoveDuplicatesWithHeader""
    End If
    delSheet
    Exit Function
FAILURE:
    delSheet
    Err.Raise Err.Number, Err.source, ""testRemoveDuplicatesWithHeader(): "" & Err.Description
End Function

Private Function testSummarizeDataByColumns()
    testSummarizeDataByColumns = ""PASSED""
    Sheets.Add
    Cells(1, 1) = ""Group""
    Cells(1, 2) = ""SubGroup""
    Cells(1, 3) = ""Revenue""
    Cells(2, 1) = ""Visa""
    Cells(2, 2) = ""DPS""
    Cells(2, 3) = 500
    Cells(3, 1) = ""Visa""
    Cells(3, 2) = ""DPS""
    Cells(3, 3) = 600
    Cells(4, 1) = ""Visa""
    Cells(4, 2) = ""CyberSource""
    Cells(4, 3) = 200
    Dim results As Variant
    Set results = dataUtil.summarizeDataByColumns(Array(""A"", ""B""), ""C"", True)
    delSheet
    Sheets.Add
    Dim rowCtr As Long: rowCtr = 1
    Dim colCtr As Long: colCtr = 1
    
    For Each r In results.Keys
        Dim valsArr As Variant
        valsArr = Split(r, ""~"")
        colCtr = 1
        For Each v In valsArr
            Cells(rowCtr, colCtr) = v
            colCtr = colCtr + 1
        Next
        Cells(rowCtr, colCtr) = results(r)
        rowCtr = rowCtr + 1
    Next
    
    If (Not Cells(2, 1) = ""Visa"" Or Not Cells(2, 2) = ""CyberSource"" Or Not Cells(2, 3) = 200) Then
        testSummarizeDataByColumns = ""FAILED""
        Err.Raise 1, , ""Failed testSummarizeDataByColumns""
    End If
    delSheet
    Exit Function
FAILURE:
    delSheet
    Err.Raise Err.Number, Err.source, ""testSummarizeDataByColumns(): "" & Err.Description
End Function


'******************************************************************************************
'*ROWCOL TESTS
'* - convertColLetterToNum
'* - convertColNumToLetter
'* - findColByName
'* - findLastColumn
'* - findLastRow
'* - moveColumnBehind
'* - removeColByLetter
'* - removeColByName
'* - removeColByNum
'* - removeRow
'* - resizeAll
'******************************************************************************************
Private Function rowcolTests()
    Set rowColUtil = New RWPC_ROWCOL
    Sheets(""UnitTestResults"").Select
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.convertColLetterToNum""
    Cells(TEST_CTR, 2) = testConvertColLetterToNum
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.convertColNumToLetter""
    Cells(TEST_CTR, 2) = testConvertColNumToLetter
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.findColByName""
    Cells(TEST_CTR, 2) = testFindColByName
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.findLastColumn""
    Cells(TEST_CTR, 2) = testFindLastColumn
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.findLastRow""
    Cells(TEST_CTR, 2) = testFindLastRow
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.moveColumnBehind""
    Cells(TEST_CTR, 2) = testMoveColumnBehind
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.removeColByLetter""
    Cells(TEST_CTR, 2) = testRemoveColByLetter
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.removeColByName""
    Cells(TEST_CTR, 2) = testRemoveColByName
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.removeColByNum""
    Cells(TEST_CTR, 2) = testRemoveColByNum
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.removeRow""
    Cells(TEST_CTR, 2) = testRemoveRow
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test ROWCOL.resizeAll""
    Cells(TEST_CTR, 2) = testResizeAll
    
    
End Function

Private Function testConvertColLetterToNum()
    On Error GoTo EXPECT_FAILURE1
    testConvertColLetterToNum = ""PASSED""
    Dim rowCol As Long
    rowCol = rowColUtil.convertColLetterToNum(""ZZZZ"")
    testConvertColLetterToNum = ""PASSED""
    Exit Function
EXPECT_FAILURE1:
    Resume RESUME_FAILURE
RESUME_FAILURE:
    On Error GoTo FAILURE
    testConvertColLetterToNum = ""PASSED""
    colNum = rowColUtil.convertColLetterToNum(""AA"")
    If (Not colNum = 27) Then
        testConvertColLetterToNum = ""FAILED""
    End If
    Exit Function
FAILURE:
    testConvertColLetterToNum = ""FAILED""
End Function

Private Function testConvertColNumToLetter()
    On Error GoTo EXPECT_FAILURE1
    testConvertColNumToLetter = ""PASSED""
    Dim rowCol As String
    rowCol = rowColUtil.convertColNumToLetter(99999)
    testConvertColNumToLetter = ""FAILED""
    Exit Function
EXPECT_FAILURE1:
    Resume RESUME_FAILURE
RESUME_FAILURE:
    On Error GoTo FAILURE
    testConvertColNumToLetter = ""PASSED""
    colNum = rowColUtil.convertColNumToLetter(27)
    If (Not colNum = ""AA"") Then
        testConvertColNumToLetter = ""FAILED""
    End If
    Exit Function
FAILURE:
    testConvertColNumToLetter = ""FAILED""
End Function

Private Function testFindColByName()
    On Error GoTo EXPECT_FAILURE1
    testFindColByName = ""PASSED""
    Sheets.Add
    Cells(1, 1) = ""HELLO""
    Cells(1, 2) = ""WORLD""
    Cells(1, 3) = ""TEST""
    Dim colNum As Long
    colNum = rowColUtil.findColByName(1, ""SAMPLE"")
    testFindColByName = ""FAILED""
    Exit Function
EXPECT_FAILURE1:
    Resume RESUME_FAILURE
RESUME_FAILURE:
    On Error GoTo FAILURE
    testFindColByName = ""PASSED""
    colNum = rowColUtil.findColByName(1, ""WORLD"")
    If (Not colNum = 2) Then
        testFindColByName = ""FAILED""
    End If
    
    Dim colLetter As String
    colLetter = rowColUtil.findColByName(1, ""WORLD"", True)
    If (Not colLetter = ""B"") Then
        testFindColByName = ""FAILED""
    End If
    delSheet
    Exit Function
FAILURE:
    delSheet
    testFindColByName = ""FAILED""
End Function

Private Function testFindLastColumn()
    On Error GoTo EXPECT_FAILURE1
    testFindLastColumn = ""PASSED""
    Sheets.Add
    Dim colNum As Long
    colNum = rowColUtil.findLastColumn(1)
EXPECT_FAILURE1:
    Resume RESUME_FAILURE1
RESUME_FAILURE1:
    On Error GoTo FAILURE
    Cells(1, 1) = ""HELLO""
    Cells(1, 2) = ""WORLD""
    Cells(1, 3) = ""TEST""
    colNum = rowColUtil.findLastColumn(1)
    If (Not colNum = 3) Then
        testFindLastColumn = ""FAILED""
    End If
    
    Dim colLetter As String
    colLetter = rowColUtil.findLastColumn(1, True)
    If (Not colLetter = ""C"") Then
        testFindLastColumn = ""FAILED""
    End If
    delSheet
    Exit Function
FAILURE:
    testFindLastColumn = ""FAILED""
    delSheet
End Function

Private Function testFindLastRow()
    On Error GoTo EXPECT_FAILURE1
    testFindLastRow = ""PASSED""
    Sheets.Add
    Dim rowNum As Long
    rowNum = rowColUtil.findLastRow(""A"")
    testFindLastRow = ""FAILED""
    Exit Function
EXPECT_FAILURE1:
    Resume RESUME_FAILURE1
RESUME_FAILURE1:
    On Error GoTo FAILURE
    Cells(1, 1) = ""HELLO""
    Cells(2, 1) = ""WORLD""
    Cells(3, 1) = ""TEST""
    rowNum = rowColUtil.findLastRow(""A"")
    If (Not rowNum = 3) Then
        testFindLastRow = ""FAILED""
    End If
    rowNum = rowColUtil.findLastRow(1)
    If (Not rowNum = 3) Then
        testFindLastRow = ""FAILED""
    End If
    delSheet
    Exit Function
FAILURE:
    testFindLastRow = ""FAILED""
    delSheet
End Function

Private Function testMoveColumnBehind()
    On Error GoTo EXPECT_FAILURE1
    testMoveColumnBehind = ""PASSED""
    Sheets.Add
    Cells(1, 1) = ""DOLLARS""
    Cells(2, 1) = 500
    Cells(1, 2) = ""CREDIT""
    Cells(2, 2) = 700
    rowColUtil.moveColumnBehind ""B"", ""B""
    testMoveColumnBehind = ""FAILED""
    Exit Function
EXPECT_FAILURE1:
    Resume RESUME_FAILURE1
RESUME_FAILURE1:
    On Error GoTo EXPECT_FAILURE2
    rowColUtil.moveColumnBehind ""XXXX"", ""XXXX""
    testMoveColumnBehind = ""FAILED""
    Exit Function
EXPECT_FAILURE2:
    Resume RESUME_FAILURE2
RESUME_FAILURE2:
    On Error GoTo FAILURE
    rowColUtil.moveColumnBehind ""B"", ""A""
    If (Not Cells(2, 1) = 700) Then
        testMoveColumnBehind = ""FAILED""
    End If
    delSheet
    Exit Function
FAILURE:
    testMoveColumnBehind = ""FAILED""
    delSheet
End Function

Private Function testRemoveColByLetter()
    On Error GoTo FAILURE
    testRemoveColByLetter = ""PASSED""
    Sheets.Add
    Cells(1, 1) = ""Hello""
    Cells(1, 2) = ""World""
    Cells(1, 3) = ""Test""
    rowColUtil.removeColByLetter (""B"")
    If (Not Cells(1, 2) = ""Test"") Then
        testRemoveColByLetter = ""FAILED""
    End If
    delSheet
    Exit Function
FAILURE:
    delSheet
    testRemoveColByLetter = ""FAILED""
End Function

Private Function testRemoveColByName()
    On Error GoTo EXPECT_FAILURE1
    testRemoveColByName = ""PASSED""
    Sheets.Add
    Cells(1, 1) = ""HELLO""
    Cells(1, 2) = ""WORLD""
    Cells(1, 3) = ""TEST""
    rowColUtil.removeColByName 1, ""PUPPY""
    testRemoveColByName = ""FAILED""
    Exit Function
EXPECT_FAILURE1:
    Resume RESUME_FAILURE1
RESUME_FAILURE1:
    On Error GoTo FAILURE
    rowColUtil.removeColByName 1, ""WORLD""
    If (Not Cells(1, 2) = ""TEST"") Then
        testRemoveColByName = ""FAILED""
    End If
    delSheet
    Exit Function
FAILURE:
    delSheet
    testRemoveColByName = ""FAILED""
End Function

Private Function testRemoveColByNum()
    On Error GoTo FAILURE
    testRemoveColByNum = ""PASSED""
    Sheets.Add
    Cells(1, 1) = ""HELLO""
    Cells(1, 2) = ""WORLD""
    Cells(1, 3) = ""TEST""
    rowColUtil.removeColByNum (2)
    If (Not Cells(1, 2) = ""TEST"") Then
        testRemoveColByNum = ""FAILED""
    End If
    delSheet
    Exit Function
FAILURE:
    delSheet
    testRemoveColByNum = ""FAILED""
End Function

Private Function testRemoveRow()
    On Error GoTo FAILURE
    testRemoveRow = ""PASSED""
    Sheets.Add
    Cells(1, 1) = ""HELLO""
    Cells(2, 1) = ""WORLD""
    Cells(3, 1) = ""TEST""
    rowColUtil.removeRow 2
    If (Not Cells(2, 1) = ""TEST"") Then
        testRemoveRow = ""FAILED""
    End If
    delSheet
    Exit Function
FAILURE:
    testRemoveRow = ""FAILED""
    delSheet
End Function


'*****************************************************
'regex tests
'*****************************************************
Private Function regexTests()
    Set regexUtil = New RWPC_REGEX
    Sheets(""UnitTestResults"").Select
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test REGEX.regexFirstMatch""
    Cells(TEST_CTR, 2) = testRegexFirstMatch
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test REGEX.regexMatchBool""
    Cells(TEST_CTR, 2) = testRegexMatchBool
    
End Function

Private Function testRegexFirstMatch()
    testRegexFirstMatch = ""PASSED""
    On Error GoTo FAILURE
    Dim match As String: match = regexUtil.RegexFirstMatch(""Version 19.04 \n Version 19.05"", ""\d{2}\.\d{2}"")
    If (Not match = ""19.04"") Then
        testRegexFirstMatch = ""FAILED""
    End If
    Exit Function
FAILURE:
    testRegexFirstMatch = ""FAILED""
End Function

Private Function testRegexMatchBool()
    testRegexMatchBool = ""PASSED""
    On Error GoTo FAILURE
    If (Not regexUtil.RegexMatchBool(""version 19.04"", ""\d{2}\.\d{2}"")) Then
        testRegexMatchBool = ""FAILED""
    End If
    Exit Function
FAILURE:
    testRegexMatchBool = ""FAILED""
End Function

Private Function sheetsTests()
    Set sheetsUtil = New RWPC_SHEETS
    Sheets(""UnitTestResults"").Select
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test SHEETS.SpreadsheetExists""
    Cells(TEST_CTR, 2) = testSpreadsheetExists
    incrementTestCount
    Cells(TEST_CTR, 1) = ""Test SHEETS.spreadsheetIsEmpty""
    Cells(TEST_CTR, 2) = testSpreadsheetIsEmpty
    
End Function

Private Function testSpreadsheetExists()
    testSpreadsheetExists = ""PASSED""
    On Error GoTo FAILURE
    If (Not sheetsUtil.spreadsheetExists(""UnitTestResults"")) Then
        testSpreadsheetExists = ""FAILED""
    End If
    Exit Function
FAILURE:
    testSpreadsheetExists = ""FAILED""
End Function

Private Function testSpreadsheetIsEmpty()
    testSpreadsheetIsEmpty = ""PASSED""
    On Error GoTo FAILURE
    Sheets.Add
    If (Not sheetsUtil.spreadsheetIsEmpty(ActiveSheet.Name)) Then
        testSpreadsheetIsEmpty = ""FAILED""
    End If
    delSheet
    Exit Function
FAILURE:
    delSheet
    testSpreadsheetIsEmpty = ""FAILED""
End Function


Private Function testResizeAll()
    On Error GoTo FAILURE
    testResizeAll = ""PASSED""
    Sheets.Add
    Cells(1, 1) = ""HELLO""
    rowColUtil.resizeAll
    If (Not Cells(1, 1).ColumnWidth = 5.13) Then
        testResizeAll = Cells(1, 1).ColumnWidth
    End If
    delSheet
    Exit Function
FAILURE:
    delSheet
    testResizeAll = ""FAILED""
End Function
Private Function incrementTestCount()
    TEST_CTR = TEST_CTR + 1
End Function
Private Function delSheet()
    If Not ActiveSheet.Name = ""UnitTestResults"" Then
        Application.DisplayAlerts = False
        ActiveSheet.Delete
    End If
    Sheets(""UnitTestResults"").Select
    Application.DisplayAlerts = True
End Function"
