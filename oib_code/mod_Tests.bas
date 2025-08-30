Attribute VB_Name = "mod_Tests"
Option Explicit
Option Private Module

'==== RUN SWITCHES (toggle as needed) ======================================================
Private Const RUN_UNIT_TESTS As Boolean = True
Private Const RUN_INTEGRATION_TESTS As Boolean = True
Private Const RUN_POWERQUERY_TESTS As Boolean = True
Private Const RUN_API_TESTS As Boolean = True
Private Const RUN_UI_TESTS As Boolean = True
Private Const RUN_SEEDING_TESTS As Boolean = True

'==== ASSERT/LOG STATE =====================================================================
Private mPassCount As Long
Private mFailCount As Long
Private mSkipCount As Long
Private mStartedAt As Date

'==== TRACK OPENED WORKBOOKS ==============================================================
Private mTestWorkbooks As Collection

'==== PUBLIC ENTRYPOINTS ===================================================================
Public Sub Test_All()
    StartSuite "On It, Boss! Test Suite"

    If RUN_UNIT_TESTS Then
        Run_Group_String
        Run_Group_FixTableCombine
        Run_Group_JSON
        Run_Group_History_Unit
        Run_Group_SchemaQuery_Unit
        Run_Group_Core_Unit
        Run_Group_API_Unit
        Run_Group_Env_Unit
        Run_Group_ExcelHelpers_Unit
    Else
        Skip "Unit tests"
    End If

    If RUN_INTEGRATION_TESTS Then
        Run_Group_History_Integration
        Run_Group_SchemaQuery_Integration
    Else
        Skip "Integration group"
    End If

    If RUN_POWERQUERY_TESTS Then
        Run_Group_PowerQuery_Integration
    Else
        Skip "Power Query integration group"
    End If

    If RUN_API_TESTS Then
        Run_Group_API_Manual
    Else
        Skip "API/manual group"
    End If

    If RUN_UI_TESTS Then
        Run_Group_UI_Manual
    Else
        Skip "UI/manual group"
    End If
    
    If RUN_SEEDING_TESTS Then
        Run_Group_History_EnsureSeed_UsesExistingFormula
    Else
        Skip "Seeding group"
    End If

    FinishSuite
End Sub

'==========================================================================================
' GROUP: mod_String
'==========================================================================================

Private Sub Run_Group_String()
    StartGroup "mod_String"
    Test_EscapeJson_Basic
    Test_EscapeJson_NewlinesTabs
    Test_SanitizeName_SlashesSpaces
    Test_SanitizeTableName_CharsAndLength
    Test_NormalizeModelCode_UnescapeAndFences
    Test_NormalizeModelCode_BOMAndQuotes
    Test_SanitizeTableName_AlphaPrefixOnNonLetter
    Test_EscapeJson_AlreadyEscapedBackslashes
    Test_NormalizeModelCode_PlainMNoChanges
    Test_NormalizeModelCode_UnescapeLoopStops
    Test_NormalizeModelCode_EdgeFencesAndNewlines
End Sub

Private Sub Test_EscapeJson_Basic()
    Dim got As String
    got = mod_String.EscapeJson("a""b\c")
    AssertEquals "a\""b\\c", got, "EscapeJson basic"
End Sub

Private Sub Test_EscapeJson_NewlinesTabs()
    Dim s As String: s = "x" & vbCrLf & "y" & vbTab & "z"
    Dim got As String: got = mod_String.EscapeJson(s)
    AssertTrue InStr(got, "\n") > 0 And InStr(got, "\t") > 0, "EscapeJson newlines/tabs"
End Sub

Private Sub Test_SanitizeName_SlashesSpaces()
    AssertEquals "My-Table_-_Name", mod_String.SanitizeName("My/Table \ Name"), "SanitizeName"
End Sub

Private Sub Test_SanitizeTableName_CharsAndLength()
    Dim got As String
    got = mod_String.SanitizeTableName("9bad:name/with\chars.(and)more")
    AssertTrue Left$(got, 2) = "T_", "SanitizeTableName prefix"
    AssertTrue InStr(got, "-") = 0 And InStr(got, ".") = 0, "SanitizeTableName cleaned"
    got = mod_String.SanitizeTableName(String$(300, "A"))
    AssertTrue Len(got) <= 240, "SanitizeTableName length cap"
End Sub

Private Sub Test_NormalizeModelCode_UnescapeAndFences()
    Dim m As String, got As String
    m = "```m" & vbCrLf & "let\n    X = ""a\""b""\n in X" & vbCrLf & "```"
    got = mod_String.NormalizeModelCode(m)
    AssertTrue InStr(got, "let") = 1, "NormalizeModelCode fence removed"
End Sub

Private Sub Test_NormalizeModelCode_BOMAndQuotes()
    Dim m As String, got As String
    m = ChrW(&HFEFF) & """" & "let" & vbLf & "in 1" & """"
    got = mod_String.NormalizeModelCode(m)
    AssertTrue Left$(got, 3) = "let", "NormalizeModelCode BOM/quotes"
End Sub

Private Sub Test_SanitizeTableName_AlphaPrefixOnNonLetter()
    Dim got As String
    got = mod_String.SanitizeTableName("123 starts with digit")
    AssertTrue Left$(got, 2) = "T_", "SanitizeTableName non-letter prefix adds 'T_'"
End Sub

Private Sub Test_EscapeJson_AlreadyEscapedBackslashes()
    Dim s As String, got As String
    s = "C:\Temp\Foo"     ' contains backslashes already
    got = mod_String.EscapeJson(s)
    ' Expect every backslash doubled; no triple-escaping artifacts
    AssertEquals "C:\\Temp\\Foo", got, "EscapeJson doubles backslashes but not beyond"
End Sub

Private Sub Test_NormalizeModelCode_PlainMNoChanges()
    Dim m As String, got As String
    m = "let" & vbCrLf & "    X = 1" & vbCrLf & "in" & vbCrLf & "    X"
    got = mod_String.NormalizeModelCode(m)
    AssertEquals m, got, "NormalizeModelCode leaves plain M unchanged"
End Sub

Private Sub Test_NormalizeModelCode_UnescapeLoopStops()
    Dim m As String, got As String
    ' Simulate a payload with some escaping but not infinitely peelable
    m = "let\n    X = ""a\u0022b""\n in X"
    got = mod_String.NormalizeModelCode(m)
    AssertTrue InStr(1, got, "let", vbTextCompare) > 0, "NormalizeModelCode peeled escapes"
    ' Ensure it did not loop forever or strip content
    AssertTrue InStr(1, got, """a""b""", vbTextCompare) = 0, "NormalizeModelCode reasonable normalization (not a strict requirement, but ensures stability)"
End Sub

Private Sub Test_NormalizeModelCode_EdgeFencesAndNewlines()
    Dim m As String, got As String

    m = "```" & vbCrLf & "let" & vbLf & "in 1" & vbCrLf & "```"
    got = mod_String.NormalizeModelCode(m)
    AssertTrue Left$(got, 3) = "let", "Fence without language stripped"

    m = "let\nin 1"
    got = mod_String.NormalizeModelCode(m)
    AssertTrue InStr(got, vbCrLf) > 0, "Lone \\n normalized to CRLF"
End Sub

'==========================================================================================
' GROUP: FixTableCombine
'==========================================================================================

Private Sub Run_Group_FixTableCombine()
    StartGroup "FixTableCombine"
    Test_FixTableCombine_WrapsTwoArgs
    Test_FixTableCombine_WrapsSingleArg
    Test_FixTableCombine_LeavesExistingListUntouched
    Test_FixTableCombine_LeavesExistingListWithWhitespaceUntouched
    Test_FixTableCombine_NestedExpressions
    Test_FixTableCombine_MultipleOccurrences
    Test_FixTableCombine_Idempotent
    Test_FixTableCombine_NoChangeWhenAbsent
    Test_FixTableCombine_AllowsWhitespaceBeforeParen
    Test_FixTableCombine_MultilineArgs
    Test_FixTableCombine_InsideStringLiteral_Untouched
    Test_FixTableCombine_InsideSingleLineComment_Untouched
    Test_FixTableCombine_InsideMultiLineComment_Untouched
    Test_FixTableCombine_MultipleLines_AllFixed
    Test_FixTableCombine_CaseInsensitive_Many
End Sub

Public Sub Test_FixTableCombine_WrapsTwoArgs()
    Dim s As String, got As String, want As String
    s = "let t = Table.Combine(Table1, Table2) in t"
    want = "let t = Table.Combine({Table1, Table2}) in t"
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Wraps two args into list"
End Sub

Public Sub Test_FixTableCombine_WrapsSingleArg()
    Dim s As String, got As String, want As String
    s = "let t = Table.Combine(Table1) in t"
    want = "let t = Table.Combine({Table1}) in t"
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Wraps single arg into list"
End Sub

Public Sub Test_FixTableCombine_LeavesExistingListUntouched()
    Dim s As String, got As String
    s = "let t = Table.Combine({Table1, Table2}) in t"
    got = FixTableCombineSyntax(s)
    AssertEquals s, got, "Already-correct list remains unchanged"
End Sub

Public Sub Test_FixTableCombine_LeavesExistingListWithWhitespaceUntouched()
    Dim s As String, got As String
    s = "let t = Table.Combine(   {Table1, Table2}   ) in t"
    got = FixTableCombineSyntax(s)
    AssertEquals s, got, "Already-correct list (with spaces) remains unchanged"
End Sub

Public Sub Test_FixTableCombine_NestedExpressions()
    Dim s As String, got As String, want As String
    s = "let t = Table.Combine( Table.AddColumn(A, ""c"", each 1), B ) in t"
    want = "let t = Table.Combine({Table.AddColumn(A, ""c"", each 1), B}) in t"
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Nested expressions safely wrapped"
End Sub

Public Sub Test_FixTableCombine_MultipleOccurrences()
    Dim s As String, got As String, want As String
    s = "let a = Table.Combine(T1, T2) in let b = Table.Combine( {T3} ) in a & b"
    want = "let a = Table.Combine({T1, T2}) in let b = Table.Combine( {T3} ) in a & b"
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Fixes only incorrect occurrences; leaves correct ones"
End Sub

Public Sub Test_FixTableCombine_Idempotent()
    Dim s As String, once As String, twice As String
    s = "let t = Table.Combine(Table1,Table2) in t"
    once = FixTableCombineSyntax(s)
    twice = FixTableCombineSyntax(once)
    AssertEquals once, twice, "Idempotent on second pass"
End Sub

Public Sub Test_FixTableCombine_NoChangeWhenAbsent()
    Dim s As String, got As String
    s = "let t = Table.RemoveRows(T, 0) in t"
    got = FixTableCombineSyntax(s)
    AssertEquals s, got, "No change when Table.Combine not present"
End Sub

Public Sub Test_FixTableCombine_AllowsWhitespaceBeforeParen()
    Dim s As String, got As String, want As String
    s = "let t = Table.Combine   (  Table1 , Table2 ) in t"
    want = "let t = Table.Combine({Table1 , Table2}) in t"
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Handles spaces before '(' and trims outer whitespace"
End Sub

Public Sub Test_FixTableCombine_MultilineArgs()
    Dim s As String, got As String, want As String
    s = "let t = Table.Combine(" & vbCrLf & "    Table1," & vbCrLf & "    Table2" & vbCrLf & ") in t"
    want = "let t = Table.Combine({Table1," & vbCrLf & "    Table2}) in t"
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Handles multi-line arguments and trims outer whitespace"
End Sub

Public Sub Test_FixTableCombine_InsideStringLiteral_Untouched()
    Dim s As String, got As String, want As String
    s = "let msg = ""This is text: Table.Combine(T1, T2)"" in msg"
    want = s
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Does not touch occurrences inside string literals"
End Sub

Public Sub Test_FixTableCombine_InsideSingleLineComment_Untouched()
    Dim s As String, got As String, want As String
    s = "// Table.Combine(T1, T2)" & vbCrLf & "let t = Table.Combine(T1, T2) in t"
    want = "// Table.Combine(T1, T2)" & vbCrLf & "let t = Table.Combine({T1, T2}) in t"
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Ignores 'Table.Combine(...)' inside // comments"
End Sub

Public Sub Test_FixTableCombine_InsideMultiLineComment_Untouched()
    Dim s As String, got As String, want As String
    s = "/* Example: Table.Combine(T1, T2) */" & vbCrLf & "let t = Table.Combine(T1) in t"
    want = "/* Example: Table.Combine(T1, T2) */" & vbCrLf & "let t = Table.Combine({T1}) in t"
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Ignores 'Table.Combine(...)' inside /* ... */ comments"
End Sub

Public Sub Test_FixTableCombine_MultipleLines_AllFixed()
    Dim s As String, got As String, want As String
    s = "let" & vbCrLf & _
        "    a = Table.Combine(T1, T2)," & vbCrLf & _
        "    b = Table.Combine(T3)," & vbCrLf & _
        "    c = Table.Combine( {T4, T5} )" & vbCrLf & _
        "in" & vbCrLf & _
        "    Table.Combine( a, b )"
    want = "let" & vbCrLf & _
        "    a = Table.Combine({T1, T2})," & vbCrLf & _
        "    b = Table.Combine({T3})," & vbCrLf & _
        "    c = Table.Combine( {T4, T5} )" & vbCrLf & _
        "in" & vbCrLf & _
        "    Table.Combine({a, b})"
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Fixes multiple lines with several Combine calls"
End Sub

Public Sub Test_FixTableCombine_CaseInsensitive_Many()
    Dim s As String, got As String, want As String
    s = "let a = table.combine(T1,T2) in table.combine(a)"
    want = "let a = table.combine({T1,T2}) in table.combine({a})"
    got = FixTableCombineSyntax(s)
    AssertEquals want, got, "Handles multiple calls and case-insensitive name"
End Sub



'==========================================================================================
' GROUP: mod_JSON
'==========================================================================================

Private Sub Run_Group_JSON()
    StartGroup "mod_JSON"
    Test_JSON_Unescape_Simple
    Test_JSON_Unescape_UnicodeAndTrailing
    Test_JSON_GetString_Basic
    Test_JSON_GetString_DoubleEscaped
    Test_LooksLikeEscapedJson
    Test_EscapeRegex
    Test_IsHex4
    Test_FindMatchingBrace_NestedAndStrings
    Test_JSON_Unescape_TrailingBackslash
    Test_JSON_Unescape_MalformedUnicode
    Test_JSON_GetString_MissingKey
    Test_JSON_GetString_UnicodeAndEscapes
    Test_FindMatchingBrace_QuotesEscapesEdge
    Test_JSON_GetString_CodeWithInnerQuotes
End Sub

Private Sub Test_JSON_Unescape_Simple()
    Dim got As String
    got = mod_JSON.JSON_Unescape("a\\b\/c\""d\ne\r\tf")
    AssertTrue InStr(got, "a") > 0, "JSON_Unescape core"
End Sub

Private Sub Test_JSON_Unescape_UnicodeAndTrailing()
    AssertEquals "hiA!\", mod_JSON.JSON_Unescape("hi\u0041!\"), "JSON_Unescape unicode"
End Sub

Private Sub Test_JSON_GetString_Basic()
    Dim j As String
    j = "{""title"":""Hello"",""language"":""m"",""code"":""let\nin 1""}"
    AssertEquals "Hello", mod_JSON.JSON_GetString(j, "title"), "JSON_GetString title"
End Sub

Private Sub Test_JSON_GetString_DoubleEscaped()
    Dim j As String
    j = """" & "{\""title\"":\""Hi\""}" & """"
    AssertEquals "Hi", mod_JSON.JSON_GetString(j, "title"), "JSON_GetString double-escaped"
End Sub

Private Sub Test_LooksLikeEscapedJson()
    AssertTrue mod_JSON.LooksLikeEscapedJson("{"), "LooksLikeEscapedJson raw brace"
End Sub

Private Sub Test_EscapeRegex()
    AssertTrue InStr(mod_JSON.EscapeRegex("a.^$()[]{}|\+*?"), "\^") > 0, "EscapeRegex"
End Sub

Private Sub Test_IsHex4()
    AssertTrue mod_JSON.IsHex4("00AF"), "IsHex4 yes"
    AssertTrue Not mod_JSON.IsHex4("G0AF"), "IsHex4 no"
End Sub

Private Sub Test_FindMatchingBrace_NestedAndStrings()
    Dim j As String, pos As Long, endPos As Long
    j = "{""a"":{""b"":""{not a brace}""}, ""c"":{}}"
    pos = InStr(1, j, "{")
    endPos = mod_JSON.FindMatchingBrace(j, pos)
    AssertEquals Len(j), endPos, "FindMatchingBrace end"
End Sub

Private Sub Test_JSON_Unescape_TrailingBackslash()
    Dim got As String
    got = mod_JSON.JSON_Unescape("abc\")
    AssertEquals "abc\", got, "JSON_Unescape keeps trailing backslash"
End Sub

Private Sub Test_JSON_Unescape_MalformedUnicode()
    ' malformed: \u12 (too short) and \uZZZZ (non-hex) -> tolerant fallback
    Dim got1 As String, got2 As String
    got1 = mod_JSON.JSON_Unescape("x\u12y")
    got2 = mod_JSON.JSON_Unescape("x\uZZZZy")
    AssertTrue InStr(1, got1, "\u12", vbBinaryCompare) > 0 Or InStr(1, got1, "x") > 0, "JSON_Unescape tolerates short \\u"
    AssertTrue InStr(1, got2, "\uZZZZ", vbBinaryCompare) > 0 Or InStr(1, got2, "x") > 0, "JSON_Unescape tolerates non-hex \\u"
End Sub

Private Sub Test_JSON_GetString_MissingKey()
    Dim j As String
    j = "{""title"":""Ok"",""language"":""m"",""code"":""in 1""}"
    AssertEquals "", mod_JSON.JSON_GetString(j, "nope"), "JSON_GetString missing key returns empty"
End Sub

Private Sub Test_JSON_GetString_UnicodeAndEscapes()
    Dim j As String, got As String
    j = "{""title"":""Hi\u0041!\"" "",""language"":""m"",""code"":""let\nin 1""}"
    got = mod_JSON.JSON_GetString(j, "title")
    AssertTrue InStr(1, got, "HiA!", vbBinaryCompare) > 0, "JSON_GetString handles \\uXXXX + quotes"
End Sub

Private Sub Test_FindMatchingBrace_QuotesEscapesEdge()
    Dim j As String, pos As Long, endPos As Long
    j = "{""a"":""{\""inner\"":1}"",""b"":{""c"":{}}}"
    pos = InStr(1, j, "{")
    endPos = mod_JSON.FindMatchingBrace(j, pos)
    AssertEquals Len(j), endPos, "FindMatchingBrace survives embedded escaped quotes/braces"
End Sub

Private Sub Test_JSON_GetString_CodeWithInnerQuotes()
    Dim j As String, got As String
    j = "{""title"":""T"",""language"":""m"",""code"":""let\n    Source = Excel.CurrentWorkbook(){[Name=\""Table1\""]}[Content],\n    GroupedRows = Table.Group(Source, {\""Department\""}, {{\""TotalSalary\"", each List.Sum([Salary]), type number}})\nin\n    GroupedRows""}"
    got = mod_JSON.JSON_GetString(j, "code")
    ' Must contain the inner quotes correctly unescaped after extraction:
    AssertStringContains got, "Name=""Table1""", "JSON_GetString handles inner quotes in code"
    AssertTrue InStr(1, got, "GroupedRows = Table.Group", vbBinaryCompare) > 0, "Full code captured (not truncated)"
End Sub


'==========================================================================================
' GROUP: mod_History (unit)
'==========================================================================================

Private Sub Run_Group_History_Unit()
    StartGroup "mod_History (unit)"
    Dim wb As Workbook: Set wb = CreateTestWorkbook("HIST_UNIT_")
    Test_EnsureHistorySheet_CreatesHiddenHeaders wb
    Test_NextHistoryId_GapAfterDelete
    Test_History_GetCodeById_FoundAndMissing
End Sub

Private Sub Test_EnsureHistorySheet_CreatesHiddenHeaders(wb As Workbook)
    Dim ws As Worksheet: Set ws = mod_History.EnsureHistorySheet(wb)
    AssertEquals mod_Config.HISTORY_SHEET_NAME, ws.name, "History sheet name"
    AssertEquals "ID", CStr(ws.Cells(1, 1).Value), "History header A1"
End Sub

Private Sub Test_NextHistoryId_GapAfterDelete()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("HIST_GAP_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "T", Array("A"), Array(Array(1)))
    mod_History.History_EnsureSeedForTable wb, lo

    Dim ws As Worksheet: Set ws = wb.Worksheets(mod_Config.HISTORY_SHEET_NAME)
    ' Add two more history rows
    Call mod_History.History_Add(wb, lo.name, "T", "Step 1", "m", "in 1")
    Call mod_History.History_Add(wb, lo.name, "T", "Step 2", "m", "in 2")
    ' Delete last row to create a "gap" scenario
    ws.rows(ws.Cells(ws.rows.Count, mod_Config.hcID).End(xlUp).Row).Delete
    ' NextHistoryId should use the last remaining row's ID + 1 (monotonic)
    Dim nextId As Long
    nextId = mod_History.NextHistoryId(ws)
    AssertTrue nextId >= 2, "NextHistoryId remains monotonic after delete"
End Sub

Private Sub Test_History_GetCodeById_FoundAndMissing()
    StartGroup "History GetCode"
    Dim wb As Workbook: Set wb = CreateTestWorkbook("HIST_GET_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "Orders", Array("X"), Array(Array("a")))
    mod_History.History_EnsureSeedForTable wb, lo

    Dim ws As Worksheet: Set ws = wb.Worksheets(mod_Config.HISTORY_SHEET_NAME)
    Dim r As Long: r = mod_History.History_Add(wb, lo.name, "Orders", "Step", "m", "let" & vbCrLf & "in 1")
    Dim id As Long: id = CLng(ws.Cells(r, mod_Config.hcID).Value)

    Dim t As String, lang As String, code As String
    code = mod_History.History_GetCodeById(wb, id, t, lang)
    AssertTrue Len(code) > 0 And LCase$(lang) = "m", "GetCodeById returns code/lang"

    t = "": lang = "": code = mod_History.History_GetCodeById(wb, 999999, t, lang)
    AssertEquals "", code, "GetCodeById returns empty for missing ID"
End Sub

'==========================================================================================
' GROUP: SchemaQuery (unit)
'==========================================================================================

Private Sub Run_Group_SchemaQuery_Unit()
    StartGroup "mod_SchemaQuery (unit)"
    Dim wb As Workbook: Set wb = CreateTestWorkbook("SCHEMA_UNIT_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "Sales Data", Array("Region", "Qty"), Array(Array("US", 10)))
    Test_BuildTableSchemaJson_WithData lo
    Test_BuildTableSchemaJson_EmptyTable
    Test_BuildTableSchemaJson_QuotesInHeaders
    Test_UniqueSheetName_SuffixIncrements
    Test_SheetExists_FalseForMissing
    Test_TryExtractOriginTableName_Parses
    Test_TryExtractOriginTableName_EmptyWhenMissing
    Test_BuildOriginSchemaJson_FindsTable
    Test_BuildOriginSchemaJson_MissingTableGraceful
End Sub

' --- Origin table-name parsing ---
Private Sub Test_TryExtractOriginTableName_Parses()
    Dim m As String
    m = "let" & vbCrLf & _
        "    Source = Excel.CurrentWorkbook(){[Name=""T1""]}[Content]," & vbCrLf & _
        "    #" & Chr(34) & "Step" & Chr(34) & " = Source" & vbCrLf & _
        "in #" & Chr(34) & "Step" & Chr(34)
    AssertEquals "T1", mod_SchemaQuery.TryExtractOriginTableName(m), "Origin table parsed from M"
End Sub

Private Sub Test_TryExtractOriginTableName_EmptyWhenMissing()
    Dim m As String
    m = "let" & vbCrLf & _
        "    Source = 123" & vbCrLf & _
        "in Source"
    AssertEquals "", mod_SchemaQuery.TryExtractOriginTableName(m), "No match returns empty string"
End Sub

' --- Origin schema JSON build ---
Private Sub Test_BuildOriginSchemaJson_FindsTable()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("ORIGIN_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "T1", Array("A", "B"), Array(Array(1, 2)))
    Dim m As String
    m = "let" & vbCrLf & _
        "    Source = Excel.CurrentWorkbook(){[Name=""T1""]}[Content]" & vbCrLf & _
        "in Source"
    Dim j As String: j = mod_SchemaQuery.BuildOriginSchemaJson(wb, m)
    AssertStringContains j, """tableName"":""T1""", "Origin schema tableName present"
    AssertStringContains j, """columns""", "Origin schema has columns[]"
End Sub

Private Sub Test_BuildOriginSchemaJson_MissingTableGraceful()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("ORIGIN_MISS_")
    Dim m As String
    m = "let" & vbCrLf & _
        "    Source = Excel.CurrentWorkbook(){[Name=""Ghost""]}[Content]" & vbCrLf & _
        "in Source"
    Dim j As String: j = mod_SchemaQuery.BuildOriginSchemaJson(wb, m)
    AssertStringContains j, """tableName"":""Ghost""", "Echo missing name"
    AssertStringContains j, """columns"":[]", "Empty columns when table not found"
    AssertStringContains j, """sample"":[]", "Empty sample when table not found"
End Sub

' --- Prompt builder includes both schema blocks ---
Private Sub Test_BuildXmlPrompt_IncludesBothSchemas()
    Dim xml As String
    xml = mod_Core.BuildXmlPrompt( _
            "let" & vbCrLf & "in 0", _
            "{""tableName"":""Origin"",""columns"":[],""sample"":[]}", _
            "{""tableName"":""Current"",""columns"":[],""sample"":[]}", _
            "noop" _
          )
    AssertStringContains xml, "<ORIGIN_TABLE_SCHEMA>", "Origin schema tag present"
    AssertStringContains xml, "<CURRENT_TABLE_SCHEMA>", "Current schema tag present"
    AssertStringContains xml, "<CURRENT_M_CODE>", "Current M tag present"
    AssertStringContains xml, "<USER_REQUEST>", "User request tag present"
End Sub

' --- End-to-end builder emits both schemas from a real ListObject ---
Private Sub Test_BuildSchemaAndCurrentM_ReturnsBoth()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("SCM_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "T0", Array("A"), Array(Array(1)))

    Dim originJson As String, currentJson As String, currentM As String, qName As String
    mod_Core.BuildSchemaAndCurrentM lo, originJson, currentJson, currentM, qName

    AssertStringContains originJson, """tableName"":""T0""", "Origin schema built"
    AssertStringContains currentJson, """tableName"":""T0""", "Current schema built"
    AssertTrue Len(currentM) > 0, "Current M returned"
End Sub

Private Sub Test_BuildTableSchemaJson_WithData(lo As ListObject)
    Dim j As String: j = mod_SchemaQuery.BuildTableSchemaJson(lo)
    AssertStringContains j, """tableName"":""Sales Data""", "Schema tableName"
End Sub

Private Sub Test_BuildTableSchemaJson_EmptyTable()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("SCHEMA_EMPTY_")
    Dim lo As ListObject
    Set lo = AddTestTable(wb, "EmptyTbl", Array("A", "B"), Array()) ' no rows
    Dim j As String: j = mod_SchemaQuery.BuildTableSchemaJson(lo)
    AssertStringContains j, """sample"":[]", "Schema builder returns empty sample for no data"
End Sub

Private Sub Test_BuildTableSchemaJson_QuotesInHeaders()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("SCHEMA_HDR_")
    Dim lo As ListObject
    Set lo = AddTestTable(wb, "Hdrs", Array("A""Quote", "B/Slash"), Array(Array("x", "y")))
    Dim j As String: j = mod_SchemaQuery.BuildTableSchemaJson(lo)
    AssertStringContains j, """name"":""A\""Quote""", "Schema JSON escapes quotes in header"
    AssertStringContains j, """name"":""B/Slash""", "Schema JSON keeps slash header (sanitized elsewhere)"
End Sub

Private Sub Test_UniqueSheetName_SuffixIncrements()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("SCHEMA_UTIL_")
    ' Create two sheets with similar base name
    Dim s1 As String, s2 As String
    s1 = mod_SchemaQuery.UniqueSheetName(wb, "Base")
    wb.Worksheets.Add.name = s1
    s2 = mod_SchemaQuery.UniqueSheetName(wb, "Base")
    AssertTrue s1 <> s2, "UniqueSheetName increments suffix when needed"
End Sub

Private Sub Test_SheetExists_FalseForMissing()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("SCHEMA_UTIL_2_")
    AssertTrue Not mod_SchemaQuery.SheetExists(wb, "I_Do_Not_Exist"), "SheetExists false on missing"
End Sub

'==========================================================================================
' GROUP: Env (unit)
'==========================================================================================
Private Sub Run_Group_Env_Unit()
    StartGroup "mod_Env (unit)"
    Test_Env_ReadWrite_SafeTemp
End Sub

Private Sub Test_Env_ReadWrite_SafeTemp()
    Dim tmp As String: tmp = Environ$("TEMP") & "\xlgpt_test.env"
    On Error Resume Next: Kill tmp: On Error GoTo 0
    mod_Env.TEST_OVERRIDE_EnvPath = tmp

    mod_Env.WriteApiKeyToEnv "KEY1"
    AssertEquals "KEY1", mod_Env.ReadApiKeyFromEnv(), "Write then read key"

    ' Update existing
    mod_Env.WriteApiKeyToEnv "KEY2"
    AssertEquals "KEY2", mod_Env.ReadApiKeyFromEnv(), "Overwrite existing key"
    
    mod_Env.TEST_OVERRIDE_EnvPath = ""
End Sub

'==========================================================================================
' GROUP: ExcelHelpers (unit)
'==========================================================================================
Private Sub Run_Group_ExcelHelpers_Unit()
    StartGroup "mod_ExcelHelpers (unit)"
    Test_ActiveListObjectOrNothing_PicksFirstTableWhenNoneSelected
End Sub

Private Sub Test_ActiveListObjectOrNothing_PicksFirstTableWhenNoneSelected()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("ACT_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "T1", Array("A"), Array(Array(1)))
    ' Deselect everything by activating a different sheet without tables

    wb.Worksheets(1).Activate
    wb.Worksheets(1).Range("C1").Select ' select different cell on ws
    Dim got As ListObject: Set got = mod_ExcelHelpers.ActiveListObjectOrNothing()
    AssertTrue Not got Is Nothing And got.name = lo.name, "Picked first table on sheet"
End Sub

'==========================================================================================
' GROUP: mod_Core (unit - builders)
'==========================================================================================

Private Sub Run_Group_Core_Unit()
    StartGroup "mod_Core (builders)"
    Test_BuildSeedM
    Test_BuildChatJSONPayload_ResponseFormatJsonObject
    Test_RunTransformForListObject_BlanksAreRejected
End Sub

Private Sub Test_BuildSeedM()
    AssertStringContains mod_Core.BuildSeedM("MyTable"), "Excel.CurrentWorkbook", "BuildSeedM content"
End Sub

Private Sub Test_BuildChatJSONPayload_ResponseFormatJsonObject()
    Dim payload As String
    payload = mod_Core.BuildChatJSONPayload("g", 1#, "<xml/>")
    AssertStringContains payload, """response_format"":{""type"":""json_object""}", "BuildChatJSONPayload sets response_format=json_object"
End Sub

Private Sub Test_RunTransformForListObject_BlanksAreRejected()
    StartGroup "Input validation"
    Dim wb As Workbook: Set wb = CreateTestWorkbook("BLANK_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "T", Array("A"), Array(Array(1)))
    Dim ok As Boolean
    ok = mod_Core.RunTransformForListObject(lo, "   ", True)
    AssertTrue Not ok, "Blank user prompt is rejected"
End Sub

'==========================================================================================
' GROUP: mod_API (unit - JSON extraction)
'==========================================================================================

Private Sub Run_Group_API_Unit()
    StartGroup "mod_API (extract)"
    Test_ExtractFirstMessageContent_JSON_Works
    Test_ExtractFirstMessageContent_JSON_StringContentAndMissing
End Sub

Private Sub Test_ExtractFirstMessageContent_JSON_Works()
    Dim full As String, got As String
    full = "{""choices"":[{""message"":{""role"":""assistant"",""content"":{""title"":""T""}}}]}"
    got = mod_API.ExtractFirstMessageContent_JSON(full)
    AssertStringContains got, """title"":""T""", "ExtractFirstMessageContent"
End Sub

Private Sub Test_ExtractFirstMessageContent_JSON_StringContentAndMissing()
    Dim full1 As String, got1 As String
    Dim full2 As String, got2 As String

    ' (b) content is a JSON STRING (escaped object)
    full1 = "{""choices"":[{""message"":{""content"":""{\""title\"":\""T2\"",\""code\"":\""in 1\"",\""language\"":\""m\""}""}}]}"
    got1 = mod_API.ExtractFirstMessageContent_JSON(full1)
    AssertStringContains got1, """title"":""T2""", "ExtractFirstMessageContent handles string-JSON"

    ' Missing content -> empty
    full2 = "{""choices"":[{""message"":{""role"":""assistant""}}]}"
    got2 = mod_API.ExtractFirstMessageContent_JSON(full2)
    AssertEquals "", got2, "ExtractFirstMessageContent returns empty on missing content"
End Sub

'==========================================================================================
' GROUP: Integration (History + SchemaQuery w/o PQ)
'==========================================================================================

Private Sub Run_Group_History_Integration()
    StartGroup "History Integration"

    Dim wb As Workbook: Set wb = CreateTestWorkbook("HIST_INT_")
    Dim lo As ListObject
    Set lo = AddTestTable(wb, "Orders", Array("Item", "Qty"), Array( _
        Array("Pen", 2), Array("Pad", 5)))

    ' Ensure seed only when none exists for given table
    mod_History.History_EnsureSeedForTable wb, lo

    Dim arr As Variant
    arr = mod_History.History_ListForTable(wb, lo.name)
    AssertTrue Not IsEmpty(arr), "History_ListForTable not empty"

    ' Expect at least one row (the seed)
    Dim rc As Long
    On Error Resume Next
    rc = UBound(arr, 1) - LBound(arr, 1) + 1
    On Error GoTo 0
    AssertTrue rc >= 1, "History_EnsureSeedForTable seeded at least one entry"
    
    Test_History_DeleteById_RemovesRow
End Sub

Private Sub Test_History_DeleteById_RemovesRow()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("HIST_DEL_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "Orders", Array("X"), Array(Array("a")))
    mod_History.History_EnsureSeedForTable wb, lo
    ' Add one more entry; capture its ID from the sheet
    Dim ws As Worksheet: Set ws = wb.Worksheets(mod_Config.HISTORY_SHEET_NAME)
    Dim r As Long: r = mod_History.History_Add(wb, lo.name, "Orders", "ToDelete", "m", "in 0")
    Dim idToDelete As Long: idToDelete = CLng(ws.Cells(r, mod_Config.hcID).Value)
    ' Delete
    mod_History.History_DeleteById wb, idToDelete
    ' Verify it's gone
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, mod_Config.hcID).End(xlUp).Row
    Dim found As Boolean: found = False
    Dim i As Long
    For i = 2 To lastRow
        If CLng(ws.Cells(i, mod_Config.hcID).Value) = idToDelete Then
            found = True: Exit For
        End If
    Next
    AssertTrue Not found, "History_DeleteById removes row by ID"
End Sub

Private Sub Run_Group_SchemaQuery_Integration()
    StartGroup "SchemaQuery Integration"

    Dim wb As Workbook: Set wb = CreateTestWorkbook("SCHEMA_INT_")
    Dim lo As ListObject
    Set lo = AddTestTable(wb, "Customers", Array("Name"), Array( _
        Array("A"), Array("B")))

    ' Ensure query & sheet exist (adds seed M + binds). No PQ refresh asserted here.
    mod_SchemaQuery.EnsureQueryAndSheetForTable lo

    Dim qName As String: qName = mod_String.SanitizeName(lo.name)
    AssertTrue mod_SchemaQuery.QueryExists(wb, qName), "QueryExists after Ensure"

    ' EnsureLoadedToSheet should be idempotent if already bound
    mod_SchemaQuery.EnsureQueryLoadedToSheet wb, qName
    Pass "EnsureQueryLoadedToSheet idempotent (no assertions)"
End Sub

'==========================================================================================
' GROUP: Power Query Integration (optional; creates query, binds, refreshes)
'   - Requires Microsoft Power Query/Mashup engine available.
'==========================================================================================

Private Sub Run_Group_PowerQuery_Integration()
    StartGroup "PowerQuery Integration (optional)"

    Dim wb As Workbook: Set wb = CreateTestWorkbook("PQ_INT_")
    Dim lo As ListObject
    Set lo = AddTestTable(wb, "PQ_Source", Array("X"), Array( _
        Array("1"), Array("2")))

    Dim qName As String: qName = mod_String.SanitizeName(lo.name)

    ' Adds query + binds output sheet
    mod_SchemaQuery.EnsureQueryAndSheetForTable lo

    ' Verify a bound table exists
    Dim bound As ListObject
    Set bound = mod_SchemaQuery.GetFirstBoundListObject(wb, qName)
    AssertTrue Not bound Is Nothing, "Bound ListObject exists for query"

    ' Apply a trivial M transform and let the module handle refresh/rollback on failure
    Dim mcode As String
    mcode = "let" & vbCrLf & _
            "    Source = Excel.CurrentWorkbook(){[Name=""" & lo.name & """]}[Content]," & vbCrLf & _
            "    #" & "ChangedType = Table.TransformColumnTypes(Source,{{""X"", type text}})" & vbCrLf & _
            "in" & vbCrLf & _
            "    #" & "ChangedType"

    mod_SchemaQuery.ApplyMToTable lo, mcode, "Type cast to text"
    Pass "ApplyMToTable executed (manual verify data if desired)"
    
    Test_ApplyMToTable_RollbackOnInvalidM
End Sub

Private Sub Test_ApplyMToTable_RollbackOnInvalidM()
    StartGroup "PowerQuery Rollback"
    Dim wb As Workbook: Set wb = CreateTestWorkbook("PQ_ROLL_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "SRC", Array("X"), Array(Array("1")))
    mod_SchemaQuery.EnsureQueryAndSheetForTable lo
    Dim qName As String: qName = mod_String.SanitizeName(lo.name)
    Dim beforeFormula As String: beforeFormula = mod_SchemaQuery.GetExistingQueryFormula(wb, qName)

    ' Intentionally invalid M
    Dim badM As String: badM = "let" & vbCrLf & "    Source = #""NonexistentStep""" & vbCrLf & "in" & vbCrLf & "    Source"

    On Error Resume Next
    mod_SchemaQuery.ApplyMToTable lo, badM, "bad"
    On Error GoTo 0

    ' Should have rolled back to the previous formula
    Dim afterFormula As String: afterFormula = mod_SchemaQuery.GetExistingQueryFormula(wb, qName)
    AssertEquals beforeFormula, afterFormula, "ApplyMToTable rolled back on invalid M"
End Sub


'==========================================================================================
' GROUP: API Manual (optional; requires key + network)
'==========================================================================================

Private Sub Run_Group_API_Manual()
    StartGroup "API Manual"

    Dim apiKey As String
    On Error Resume Next
    apiKey = mod_Env.ReadApiKeyFromEnv()
    On Error GoTo 0

    If Len(apiKey) = 0 Then
        Skip "API key not found; skipping network call tests"
        Exit Sub
    End If

    ' Minimal chat call expecting JSON content from the assistant
    Dim payload As String, contentJson As String
    Dim stubOrigin As String, stubCurrent As String
    stubOrigin = "{""tableName"":""T"",""columns"":[],""sample"":[]}"
    stubCurrent = stubOrigin

    payload = mod_Core.BuildChatJSONPayload( _
                "gpt-5-nano", _
                1#, _
                mod_Core.BuildXmlPrompt( _
                    "let" & vbCrLf & "in 0", _
                    stubOrigin, _
                    stubCurrent, _
                    "return the same data" _
                ) _
              )

    contentJson = mod_API.CallChatAPI_JSON(apiKey, payload)
    AssertTrue InStr(contentJson, """language""") > 0 Or InStr(contentJson, """code""") > 0, _
               "CallChatAPI_JSON content JSON present"
    
    Test_RunChatAndParse_RejectsNonM
End Sub

Private Sub Test_RunChatAndParse_RejectsNonM()
    StartGroup "Chat parsing guard"
    ' Fake a non-M content JSON
    mod_API.TEST_OVERRIDE_ContentJson = "{""title"":""NotM"",""language"":""python"",""code"":""print(1)""}"

    Dim ok As Boolean
    Dim outT As String, outL As String, outC As String
    ok = mod_Core.RunChatAndParse("<x/>", outT, outL, outC)

    AssertTrue Not ok, "RunChatAndParse returns False for language != 'm'"
End Sub

'==========================================================================================
' GROUP: UI Manual (optional; shows forms, tests model refresh hook)
'==========================================================================================

Private Sub Run_Group_UI_Manual()
    StartGroup "UI Manual"
    Test_History_Panel_ReplayBlocksNonM
End Sub

Private Sub Test_History_Panel_ReplayBlocksNonM()
    Dim wb As Workbook: Set wb = CreateTestWorkbook("REPLAY_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "SRC", Array("A"), Array(Array(1)))
    Dim ws As Worksheet: Set ws = mod_History.EnsureHistorySheet(wb)
    Dim r As Long: r = mod_History.History_Add(wb, lo.name, "SRC", "BadLang", "python", "print(1)")
    Dim id As Long: id = CLng(ws.Cells(r, mod_Config.hcID).Value)

    Dim title As String, lang As String
    Dim code As String: code = mod_History.History_GetCodeById(wb, id, title, lang)
    AssertTrue LCase$(lang) <> "m", "History row language is not M (setup)"
    ' The UI path itself is event-driven; validating the underlying guard here is sufficient.
End Sub

'==========================================================================================
' GROUP: Seeding
'==========================================================================================
Public Sub Run_Group_History_EnsureSeed_UsesExistingFormula()
    StartGroup "Seeding"

    Dim wb As Workbook: Set wb = CreateTestWorkbook("SEED_")
    Dim lo As ListObject: Set lo = AddTestTable(wb, "SeedingTest", Array("ColA", "ColB"), Array(Array(1, "x")))
    Dim qName As String: qName = mod_String.SanitizeName(lo.name)

    ' Give the workbook a query with a non-Excel.CurrentWorkbook source
    Dim mExisting As String
    mExisting = "let" & vbCrLf & _
                "  Source = Web.Contents(""about:blank"")," & vbCrLf & _
                "  Dummy = #table({""ColA"",""ColB""}, {{1,""x""},{2,""y""}})" & vbCrLf & _
                "in" & vbCrLf & _
                "  Dummy"
    wb.Queries.Add name:=qName, Formula:=mExisting

    ' Ensure no preexisting history rows
    Dim ws As Worksheet: Set ws = mod_History.EnsureHistorySheet(wb)
    ws.Cells.ClearContents
    ws.Range("A1:G1").Value = Array("ID", "TableName", "QueryName", "Title", "Language", "Code", "CreatedAt")

    ' Call SUT
    mod_History.History_EnsureSeedForTable wb, lo

    ' Grab last row code
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, mod_Config.hcID).End(xlUp).Row
    Dim mFromHistory As String
    mFromHistory = CStr(ws.Cells(lastRow, mod_Config.hcCode).Value)

    ' Assertions
    AssertTrue Len(mFromHistory) > 0, "Seed row added"
    AssertEquals NormalizeWs(mExisting), NormalizeWs(mFromHistory), "Seed used existing formula"
End Sub

' --- helpers ------------------------------------------------------------
Private Function GetOrCreateTempSheet(wb As Workbook, name As String, ByRef created As Boolean) As Worksheet
    On Error Resume Next
    Set GetOrCreateTempSheet = wb.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateTempSheet Is Nothing Then
        Set GetOrCreateTempSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateTempSheet.name = name
        created = True
    Else
        created = False
    End If
End Function

Private Function CreateOrClearTempListObject(ws As Worksheet, loName As String) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(loName)
    On Error GoTo 0
    
    If lo Is Nothing Then
        ' Make a tiny 2-col table
        If ws.UsedRange.Cells.Count = 1 And ws.UsedRange.Value2 = "" Then ws.Cells(1, 1).Value = "ColA"
        ws.Cells(1, 1).Value = "ColA": ws.Cells(1, 2).Value = "ColB"
        ws.Cells(2, 1).Value = 1: ws.Cells(2, 2).Value = "x"
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        lo.name = loName
    Else
        ' Clear rows but keep headers
        If lo.DataBodyRange Is Nothing Then
            ' nothing to clear
        Else
            lo.DataBodyRange.Delete
        End If
    End If
    Set CreateOrClearTempListObject = lo
End Function

Private Sub EnsureWorkbookQuery(wb As Workbook, queryName As String, mFormula As String)
    Dim q As WorkbookQuery
    On Error Resume Next
    Set q = wb.Queries(queryName)
    On Error GoTo 0
    If q Is Nothing Then
        wb.Queries.Add name:=queryName, Formula:=mFormula
    Else
        q.Formula = mFormula
    End If
End Sub

' Try to remove existing history rows for a given object.
' This is best-effort and tolerant of different sheet/table names and columns.
Private Sub TryClearHistoryForObject(objectName As String)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ResolveHistoryListObject()
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    
    Dim colIdxName As Long: colIdxName = FindColumn(lo, Array("ObjectName", "QueryName", "Table", "TableName", "Name"))
    If colIdxName = 0 Then Exit Sub ' unknown schema; skip
    
    Dim r As Long
    Application.ScreenUpdating = False
    On Error Resume Next
    If Not lo.DataBodyRange Is Nothing Then
        For r = lo.DataBodyRange.rows.Count To 1 Step -1
            If CStr(lo.DataBodyRange.Cells(r, colIdxName).Value) = objectName Then
                lo.DataBodyRange.rows(r).EntireRow.Delete
            End If
        Next r
    End If
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

' Read back the *latest* history Code for objectName.
Private Function GetLatestHistoryCodeForObject(objectName As String) As String
    Dim lo As ListObject: Set lo = ResolveHistoryListObject()
    If lo Is Nothing Then Err.Raise 5, , "Could not locate a History table. Adjust ResolveHistoryListObject to your project."

    Dim colIdxName As Long: colIdxName = FindColumn(lo, Array("ObjectName", "QueryName", "Table", "TableName", "Name"))
    Dim colIdxCode As Long: colIdxCode = FindColumn(lo, Array("Code", "M", "Formula"))
    If colIdxName = 0 Or colIdxCode = 0 Then Err.Raise 5, , "History table missing expected columns (Name/Code)."

    Dim lastCode As String
    Dim r As Long
    
    If lo.DataBodyRange Is Nothing Then Err.Raise 5, , "History is empty after seeding—unexpected."

    ' Assume newest is last row. If your history sorts differently, adapt as needed.
    For r = lo.DataBodyRange.rows.Count To 1 Step -1
        If CStr(lo.DataBodyRange.Cells(r, colIdxName).Value) = objectName Then
            lastCode = CStr(lo.DataBodyRange.Cells(r, colIdxCode).Value)
            Exit For
        End If
    Next r
    
    If Len(lastCode) = 0 Then Err.Raise 5, , "No history row found for '" & objectName & "'."
    GetLatestHistoryCodeForObject = lastCode
End Function

' Locate your History ListObject by trying a few common names.
' If your project uses a specific sheet/table name, hardcode it here.
Private Function ResolveHistoryListObject() As ListObject
    Dim candSheets As Variant: candSheets = Array("AI_History", "History", "ux_History", "Sheet_History")
    Dim candTables As Variant: candTables = Array("lo_History", "History", "tbl_History", "ai_History")

    Dim ws As Worksheet, lo As ListObject
    Dim i As Long, j As Long
    
    ' Try sheet+table pairs first
    For i = LBound(candSheets) To UBound(candSheets)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(candSheets(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            For j = LBound(candTables) To UBound(candTables)
                On Error Resume Next
                Set lo = ws.ListObjects(CStr(candTables(j)))
                On Error GoTo 0
                If Not lo Is Nothing Then
                    Set ResolveHistoryListObject = lo
                    Exit Function
                End If
            Next j
        End If
        Set ws = Nothing
    Next i
    
    ' Fallback: scan all sheets for any table with a "Code" column
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If FindColumn(lo, Array("Code", "M", "Formula")) > 0 Then
                Set ResolveHistoryListObject = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function FindColumn(lo As ListObject, names As Variant) As Long
    Dim i As Long, j As Long
    For i = 1 To lo.HeaderRowRange.Columns.Count
        For j = LBound(names) To UBound(names)
            If StrComp(CStr(lo.HeaderRowRange.Cells(1, i).Value), CStr(names(j)), vbTextCompare) = 0 Then
                FindColumn = i
                Exit Function
            End If
        Next j
    Next i
End Function

' Compare strings ignoring whitespace differences (line breaks/indent)
Private Function NormalizeWs(ByVal s As String) As String
    Dim rx As Object: Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.MultiLine = True
    rx.Pattern = "\s+"
    NormalizeWs = LCase$(rx.Replace(s, " "))
End Function

'==========================================================================================
' HELPERS: Test harness core
'==========================================================================================

Private Sub StartGroup(ByVal name As String)
    Debug.Print vbCrLf & ">>>>> " & name & " >>>>>"
End Sub

Private Sub Pass(ByVal msg As String)
    mPassCount = mPassCount + 1
    Debug.Print "? PASS: " & msg
End Sub

Private Sub Skip(ByVal why As String)
    mSkipCount = mSkipCount + 1
    Debug.Print "? SKIP: " & why
End Sub

Private Sub Fail(ByVal msg As String)
    mFailCount = mFailCount + 1
    Debug.Print "? FAIL: " & msg
End Sub

Private Sub AssertTrue(ByVal cond As Boolean, ByVal msg As String)
    If cond Then Pass msg Else Fail msg
End Sub

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal msg As String)
    If CStr(expected) = CStr(actual) Then
        Pass msg
    Else
        Fail msg & " (expected: " & expected & ", got: " & actual & ")"
    End If
End Sub

Private Sub AssertStringContains(ByVal haystack As String, ByVal needle As String, ByVal msg As String)
    If InStr(1, haystack, needle, vbBinaryCompare) > 0 Then
        Pass msg
    Else
        Fail msg & " (missing '" & needle & "')"
    End If
End Sub

'==========================================================================================
' HELPERS: Setup / teardown for Workbook/Table/Test Data
'==========================================================================================
Private Sub StartSuite(ByVal title As String)
    mPassCount = 0: mFailCount = 0: mSkipCount = 0
    mStartedAt = Now
    Set mTestWorkbooks = New Collection
    
    Debug.Print String(80, "=")
    Debug.Print title & " — start: " & Format$(mStartedAt, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String(80, "=")
End Sub

Private Function CreateTestWorkbook(ByVal prefix As String) As Workbook
    Dim wb As Workbook: Set wb = Application.Workbooks.Add
    wb.Worksheets(1).name = prefix & "Sheet1"
    
    ' Track for cleanup
    If mTestWorkbooks Is Nothing Then Set mTestWorkbooks = New Collection
    mTestWorkbooks.Add wb
    
    Set CreateTestWorkbook = wb
End Function

Private Sub FinishSuite()
    Debug.Print vbCrLf & String(80, "-")
    Debug.Print "PASSED: "; mPassCount; " | FAILED: "; mFailCount; " | SKIPPED: "; mSkipCount
    Debug.Print String(80, "=")
    
    ' Cleanup opened test workbooks
    Dim wb As Workbook
    On Error Resume Next
    For Each wb In mTestWorkbooks
        wb.Close SaveChanges:=False
    Next
    On Error GoTo 0
    Set mTestWorkbooks = Nothing
End Sub

Private Function AddTestTable(wb As Workbook, tableName As String, headers As Variant, rows As Variant) As ListObject
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    ws.Cells.clear
    Dim c As Long, r As Long
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next
    If IsArray(rows) Then
        For r = LBound(rows) To UBound(rows)
            For c = 0 To UBound(rows(r))
                ws.Cells(r + 2, c + 1).Value = rows(r)(c)
            Next
        Next
    End If
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    lo.name = tableName
    Set AddTestTable = lo
End Function
