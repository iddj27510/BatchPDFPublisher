Attribute VB_Name = "modTests"
Option Explicit
' Module    : modTests
' Purpose   : Assertion-based test suite. Safety net for refactoring.
'             Run from Immediate Window: RunAllTests
' Reference : Appendix A.101, I.6
' Version   : 9.19.9 (Inventor 2027)
'
' SIZE NOTE: This module exceeds the 800 LOC guideline (section 1.10)
' and is documented as a justified exception.
'
' Justification:
'   1. Single responsibility: assertion-based unit tests.
'   2. Content is test cases, not production logic. Splitting into
'      per-area modules would trigger over-modularization (section 1.1)
'      and force shared helpers (AssertEquals, TEST_WriteFile,
'      TEST_ReadFile) into a separate test-utils module.
'   3. Analogous to a single *Tests.cs file in xUnit / NUnit.
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   No functional changes.
'
' v9.19.6:
'   Added Test_STATE_Init_ClearsRegCache (2 assertions) — verifies
'   the REG_ClearCache call at the top of STATE_Init. Without it,
'   registry values cached in one session would survive a reset.
'   Added Test_PRJ_PutClipboardFallback_EnsuresSandbox (3 checks) —
'   verifies that the fallback clipboard writer creates the sandbox
'   itself when it has been wiped. Regression guard for the
'   modClipboard v9.19.6 PRJ_EnsureSandbox addition.
'
' v9.19.6:
'   Test_STATE_ValidateConfig updated: STATE_ValidateConfig is now
'   a Sub (was Function always returning True). The call site drops
'   the return value capture.
'   Added Test_STS_Classify_ValidateStatus (4 new assertions):
'   Missing / Stale / Mismatch / OK statuses added to TStatus in
'   v9.19.6 and consumed by SC_ValidateDrawings.
'
' Test inventory (43 tests):
'   1.  Test_SN_SanitizePath
'   2.  Test_PDF_SafeReplace_Logical
'   3.  Test_RP_StatusTag
'   4.  Test_CACHE_RoundTrip
'   5.  Test_L_FallbackChain
'   6.  Test_STS_Classify
'   7.  Test_PRJ_FormatElapsed
'   8.  Test_PRJ_FormatFileSize
'   9.  Test_PRJ_FormatDate
'   10. Test_RP_MatchesFilter
'   11. Test_STATE_AssertGlobals
'   12. Test_ERR_Classify
'   13. Test_PRJ_PathCombine
'   14. Test_RP_BuildSummaryText
'   15. Test_STATE_ValidateConfig
'   16. Test_PDF_ValidateHeader
'   17. Test_PDF_BuildTempPath
'   18. Test_CACHE_CaseInsensitive
'   19. Test_PRJ_FormatElapsed_Edge
'   20. Test_PRJ_FormatDate_Edge
'   21. Test_SN_SanitizePath_UNC
'   22. Test_ERR_IsRecoverable_All
'   23. Test_ERR_DescribeClass_All
'   24. Test_L_Placeholders
'   25. Test_U_UnicodeDecoder
'   26. Test_LANG_NormalizeCode
'   27. Test_RP_CountFiles
'   28. Test_RP_PickSummaryColor
'   29. Test_RP_PickStatusBold
'   30. Test_RP_StatusColor
'   31. Test_CACHE_EvictIfNeeded
'   32. Test_PRJ_GetRelPath
'   33. Test_RETRY_DeleteFile_MoveFile
'   34. Test_QA_AnalyzePDF_TooSmall
'   35. Test_RP_FindStat
'   36. Test_RP_CountErrorRows
'   37. Test_FormatFileSize_Boundary
'   38. Test_STS_Classify_Edge
'   39. Test_LANG_SupportedCodes
'   40. Test_FormatElapsed_Boundary
'   41. Test_STS_Classify_ValidateStatus
'   42. Test_STATE_Init_ClearsRegCache
'   43. Test_PRJ_PutClipboardFallback_EnsuresSandbox
'
' Notes     : No line-continuation characters in this module.
'             No decorative separator comment lines.

#Const PRODUCTION_BUILD = False

#If Not PRODUCTION_BUILD Then

Public Sub RunAllTests()
    Dim nFail As Long
    Dim nRun  As Long
    Dim nTest As Long

    Debug.Print ""
    Debug.Print "========================================"
    Debug.Print "Batch PDF Publisher - Test Suite v" & PRJ_VERSION
    Debug.Print "========================================"

    nFail = 0
    nRun = 0

    nTest = Test_SN_SanitizePath()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PDF_SafeReplace_Logical()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_RP_StatusTag()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_CACHE_RoundTrip()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_L_FallbackChain()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_STS_Classify()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PRJ_FormatElapsed()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PRJ_FormatFileSize()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PRJ_FormatDate()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_RP_MatchesFilter()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_STATE_AssertGlobals()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_ERR_Classify()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PRJ_PathCombine()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_RP_BuildSummaryText()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_STATE_ValidateConfig()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PDF_ValidateHeader()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PDF_BuildTempPath()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_CACHE_CaseInsensitive()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PRJ_FormatElapsed_Edge()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PRJ_FormatDate_Edge()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_SN_SanitizePath_UNC()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_ERR_IsRecoverable_All()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_ERR_DescribeClass_All()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_L_Placeholders()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_U_UnicodeDecoder()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_LANG_NormalizeCode()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_RP_CountFiles()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_RP_PickSummaryColor()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_RP_PickStatusBold()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_RP_StatusColor()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_CACHE_EvictIfNeeded()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PRJ_GetRelPath()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_RETRY_DeleteFile_MoveFile()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_QA_AnalyzePDF_TooSmall()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_RP_FindStat()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_RP_CountErrorRows()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_FormatFileSize_Boundary()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_STS_Classify_Edge()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_LANG_SupportedCodes()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_FormatElapsed_Boundary()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_STS_Classify_ValidateStatus()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_STATE_Init_ClearsRegCache()
    nFail = nFail + nTest: nRun = nRun + 1

    nTest = Test_PRJ_PutClipboardFallback_EnsuresSandbox()
    nFail = nFail + nTest: nRun = nRun + 1

    Debug.Print "----------------------------------------"
    Debug.Print "TOTAL: " & (nRun - nFail) & " / " & nRun & " passed, " & nFail & " failed"
    Debug.Print "========================================"
    Debug.Print ""
End Sub

#End If

Private Function Test_SN_SanitizePath() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(SN_SanitizePath("  C:\Foo  "), "C:\Foo", "trim")
    nFail = nFail + AssertEquals(SN_SanitizePath("""C:\Foo"""), "C:\Foo", "quotes")
    nFail = nFail + AssertEquals(SN_SanitizePath("C:\Foo\\Bar"), "C:\Foo\Bar", "double-slash")
    nFail = nFail + AssertEquals(SN_SanitizePath("\\server\share\..\..\Windows"), "", "UNC escape")
    nFail = nFail + AssertEquals(SN_SanitizePath(String(300, "x")), "", "too long")
    nFail = nFail + AssertEquals(SN_SanitizePath("C:\Foo\..\Bar"), "C:\Bar", "parent resolve")
    nFail = nFail + AssertEquals(SN_SanitizePath("""\\server\share\..\..\Windows"""), "", "quoted UNC escape")
    nFail = nFail + AssertEquals(SN_SanitizePath("""\\server\share\sub"""), "\\server\share\sub", "quoted UNC valid")
    nFail = nFail + AssertEquals(SN_SanitizePath(""), "", "empty")
    nFail = nFail + AssertEquals(SN_SanitizePath("   "), "", "whitespace-only")
    nFail = nFail + AssertEquals(SN_SanitizePath("C:\Foo<Bar"), "C:\FooBar", "illegal char strip")
    nFail = nFail + AssertEquals(SN_SanitizePath("C:\Foo\..\..\Bar"), "C:\Bar", "root clamp")

    If nFail = 0 Then Debug.Print "PASS [SN_SanitizePath]"
    Test_SN_SanitizePath = nFail
    Exit Function

EH:
    Debug.Print "FAIL [SN_SanitizePath] - unexpected error: " & Err.Description
    Test_SN_SanitizePath = 1
End Function

Private Function Test_PDF_SafeReplace_Logical() As Long
    Dim sTestDir As String
    Dim sPDF     As String
    Dim sTemp    As String
    Dim sBak     As String
    Dim sErr     As String
    Dim sContent As String

    On Error GoTo EH

    PRJ_EnsureSandbox
    sTestDir = PRJ_SandboxPath() & "\_test_safereplace"

    On Error Resume Next
    If PRJ_FSO().FolderExists(sTestDir) Then PRJ_FSO().DeleteFolder sTestDir, True
    On Error GoTo EH
    PRJ_FSO().CreateFolder sTestDir

    sPDF = sTestDir & "\target.pdf"
    sTemp = sTestDir & "\source.tmp"

    TEST_WriteFile sPDF, "OLD_CONTENT"
    TEST_WriteFile sTemp, "NEW_CONTENT"

    sErr = PDF_SafeReplace(sTemp, sPDF)
    If Len(sErr) > 0 Then
        Debug.Print "FAIL [PDF_SafeReplace] - SafeReplace error: " & sErr
        Test_PDF_SafeReplace_Logical = 1
        GoTo Cleanup
    End If

    sContent = TEST_ReadFile(sPDF)
    If sContent <> "NEW_CONTENT" Then
        Debug.Print "FAIL [PDF_SafeReplace] - content mismatch: [" & sContent & "]"
        Test_PDF_SafeReplace_Logical = 1
        GoTo Cleanup
    End If

    PDF_CommitReplace sPDF

    sBak = sPDF & ".bak"
    If PRJ_FSO().FileExists(sBak) Then
        Debug.Print "FAIL [PDF_SafeReplace] - .bak still exists after commit"
        Test_PDF_SafeReplace_Logical = 1
        GoTo Cleanup
    End If

    Debug.Print "PASS [PDF_SafeReplace]"
    Test_PDF_SafeReplace_Logical = 0

Cleanup:
    On Error Resume Next
    If PRJ_FSO().FolderExists(sTestDir) Then PRJ_FSO().DeleteFolder sTestDir, True
    On Error GoTo 0
    Exit Function

EH:
    Debug.Print "FAIL [PDF_SafeReplace] - " & Err.Description
    Test_PDF_SafeReplace_Logical = 1
    Resume Cleanup
End Function

Private Function Test_RP_StatusTag() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(RP_StatusTag("Error: something"), "[Error]", "error")
    nFail = nFail + AssertEquals(RP_StatusTag("Skipped (Exists)"), "[Skipped]", "skipped")
    nFail = nFail + AssertEquals(RP_StatusTag("Overwritten | 5/5 pages"), "[Overwritten]", "overwritten")
    nFail = nFail + AssertEquals(RP_StatusTag("HasDrawing"), "[Has Drawing]", "hasdrawing")
    nFail = nFail + AssertEquals(RP_StatusTag("NoDrawing"), "[No Drawing]", "nodrawing")
    nFail = nFail + AssertEquals(RP_StatusTag("Created | 5/5 pages"), "[Created]", "created")
    nFail = nFail + AssertEquals(RP_StatusTag("UnknownStatus"), "[Created]", "unknown-fallback")

    If nFail = 0 Then Debug.Print "PASS [RP_StatusTag]"
    Test_RP_StatusTag = nFail
    Exit Function

EH:
    Debug.Print "FAIL [RP_StatusTag] - unexpected error: " & Err.Description
    Test_RP_StatusTag = 1
End Function

Private Function Test_CACHE_RoundTrip() As Long
    Dim sTestFile As String
    Dim n         As Long

    On Error GoTo EH

    PRJ_EnsureSandbox
    CACHE_Init

    sTestFile = PRJ_SandboxPath() & "\_test_cache.pdf"
    TEST_WriteFile sTestFile, "%PDF-1.4 test"

    CACHE_InvalidatePDF sTestFile
    n = CACHE_GetPageCount(sTestFile)
    If n <> -1 Then
        Debug.Print "FAIL [CACHE] - expected -1 miss, got " & n
        Test_CACHE_RoundTrip = 1
        GoTo Cleanup
    End If

    CACHE_PutPageCount sTestFile, 42
    n = CACHE_GetPageCount(sTestFile)
    If n <> 42 Then
        Debug.Print "FAIL [CACHE] - expected 42 hit, got " & n
        Test_CACHE_RoundTrip = 1
        GoTo Cleanup
    End If

    CACHE_InvalidatePDF sTestFile
    n = CACHE_GetPageCount(sTestFile)
    If n <> -1 Then
        Debug.Print "FAIL [CACHE] - expected -1 after invalidate, got " & n
        Test_CACHE_RoundTrip = 1
        GoTo Cleanup
    End If

    Debug.Print "PASS [CACHE]"
    Test_CACHE_RoundTrip = 0

Cleanup:
    On Error Resume Next
    If PRJ_FSO().FileExists(sTestFile) Then PRJ_FSO().DeleteFile sTestFile, True
    On Error GoTo 0
    Exit Function

EH:
    Debug.Print "FAIL [CACHE] - " & Err.Description
    Test_CACHE_RoundTrip = 1
    Resume Cleanup
End Function

Private Function Test_L_FallbackChain() As Long
    Dim s          As String
    Dim sTemplate  As String

    On Error GoTo EH

    s = L("title")
    If Len(s) = 0 Then
        Debug.Print "FAIL [L()] - title is empty"
        Test_L_FallbackChain = 1
        Exit Function
    End If
    If Left$(s, 2) = "??" Then
        Debug.Print "FAIL [L()] - title missing translation: " & s
        Test_L_FallbackChain = 1
        Exit Function
    End If

    s = L("nonexistent.test.key.zzz")
    If s <> "??nonexistent.test.key.zzz??" Then
        Debug.Print "FAIL [L()] - expected ??key?? marker, got: " & s
        Test_L_FallbackChain = 1
        Exit Function
    End If

    sTemplate = LN("scan.sum.drawings", 5)
    If InStr(sTemplate, "{0}") > 0 Then
        Debug.Print "FAIL [LN()] - placeholder not replaced: " & sTemplate
        Test_L_FallbackChain = 1
        Exit Function
    End If

    Debug.Print "PASS [L()]"
    Test_L_FallbackChain = 0
    Exit Function

EH:
    Debug.Print "FAIL [L()] - unexpected error: " & Err.Description
    Test_L_FallbackChain = 1
End Function

Private Function Test_STS_Classify() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(CStr(STS_Classify("Error: something")), CStr(ST_ERROR), "error")
    nFail = nFail + AssertEquals(CStr(STS_Classify("Skipped (Exists)")), CStr(ST_SKIPPED), "skipped")
    nFail = nFail + AssertEquals(CStr(STS_Classify("Skipped (Empty)")), CStr(ST_SKIPPED_EMPTY), "skipped-empty")
    nFail = nFail + AssertEquals(CStr(STS_Classify("Overwritten | 5/5 pages")), CStr(ST_OVERWRITTEN), "overwritten")
    nFail = nFail + AssertEquals(CStr(STS_Classify("Created | 5/5 pages")), CStr(ST_CREATED), "created")
    nFail = nFail + AssertEquals(CStr(STS_Classify("HasDrawing")), CStr(ST_HAS_DRAWING), "hasdrawing")
    nFail = nFail + AssertEquals(CStr(STS_Classify("NoDrawing")), CStr(ST_NO_DRAWING), "nodrawing")
    nFail = nFail + AssertEquals(CStr(STS_Classify("UnknownStatus")), CStr(ST_UNKNOWN), "unknown")

    If nFail = 0 Then Debug.Print "PASS [STS_Classify]"
    Test_STS_Classify = nFail
    Exit Function

EH:
    Debug.Print "FAIL [STS_Classify] - unexpected error: " & Err.Description
    Test_STS_Classify = 1
End Function

Private Function Test_PRJ_FormatElapsed() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    If InStr(PRJ_FormatElapsed(0), "0.00 s") = 0 Then
        Debug.Print "FAIL [FormatElapsed] - 0s: " & PRJ_FormatElapsed(0)
        nFail = nFail + 1
    End If

    If InStr(PRJ_FormatElapsed(1.5), "1.50 s") = 0 Then
        Debug.Print "FAIL [FormatElapsed] - 1.5s: " & PRJ_FormatElapsed(1.5)
        nFail = nFail + 1
    End If

    Dim sMin As String
    sMin = PRJ_FormatElapsed(65)
    If InStr(sMin, "1 min") = 0 Or InStr(sMin, "5.00 s") = 0 Then
        Debug.Print "FAIL [FormatElapsed] - 65s: " & sMin
        nFail = nFail + 1
    End If

    Dim sHour As String
    sHour = PRJ_FormatElapsed(3665)
    If InStr(sHour, "1 h") = 0 Or InStr(sHour, "1 min") = 0 Or InStr(sHour, "5.00 s") = 0 Then
        Debug.Print "FAIL [FormatElapsed] - 3665s: " & sHour
        nFail = nFail + 1
    End If

    Dim sNeg As String
    sNeg = PRJ_FormatElapsed(-1)
    If Len(sNeg) = 0 Then
        Debug.Print "FAIL [FormatElapsed] - negative returned empty"
        nFail = nFail + 1
    End If

    If nFail = 0 Then Debug.Print "PASS [FormatElapsed]"
    Test_PRJ_FormatElapsed = nFail
    Exit Function

EH:
    Debug.Print "FAIL [FormatElapsed] - unexpected error: " & Err.Description
    Test_PRJ_FormatElapsed = 1
End Function

Private Function Test_PRJ_FormatFileSize() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(PRJ_FormatFileSize(-1), "unknown", "negative")
    nFail = nFail + AssertEquals(PRJ_FormatFileSize(0), "0 B", "zero")
    nFail = nFail + AssertEquals(PRJ_FormatFileSize(512), "512 B", "bytes")
    nFail = nFail + AssertEquals(PRJ_FormatFileSize(1024), "1 KB", "kb")
    nFail = nFail + AssertEquals(PRJ_FormatFileSize(1048576), "1.0 MB", "mb")
    nFail = nFail + AssertEquals(PRJ_FormatFileSize(1073741824), "1.0 GB", "gb")

    If nFail = 0 Then Debug.Print "PASS [FormatFileSize]"
    Test_PRJ_FormatFileSize = nFail
    Exit Function

EH:
    Debug.Print "FAIL [FormatFileSize] - unexpected error: " & Err.Description
    Test_PRJ_FormatFileSize = 1
End Function

Private Function Test_PRJ_FormatDate() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(PRJ_FormatDate(0), "unknown", "zero")
    nFail = nFail + AssertEquals(PRJ_FormatDate(DateSerial(1970, 1, 1)), "unknown", "pre-1980")

    Dim sValid As String
    sValid = PRJ_FormatDate(DateSerial(2024, 6, 15))
    If Len(sValid) = 0 Or sValid = "unknown" Then
        Debug.Print "FAIL [FormatDate] - valid date: " & sValid
        nFail = nFail + 1
    End If

    If nFail = 0 Then Debug.Print "PASS [FormatDate]"
    Test_PRJ_FormatDate = nFail
    Exit Function

EH:
    Debug.Print "FAIL [FormatDate] - unexpected error: " & Err.Description
    Test_PRJ_FormatDate = 1
End Function

Private Function Test_RP_MatchesFilter() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    Dim sSaved As String
    sSaved = g_sActiveFilter

    g_sActiveFilter = "All"
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("Created | 5/5 pages")), "True", "all-created")
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("Error: x")), "True", "all-error")

    g_sActiveFilter = "Err"
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("Error: x")), "True", "err-match")
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("Created | 5/5 pages")), "False", "err-nomatch")

    g_sActiveFilter = "Skip"
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("Skipped (Exists)")), "True", "skip-match")
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("Skipped (Empty)")), "True", "skip-empty-match")
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("Created | 5/5 pages")), "False", "skip-nomatch")

    g_sActiveFilter = "Ok"
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("Created | 5/5 pages")), "True", "ok-created")
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("Overwritten | 5/5 pages")), "True", "ok-overwritten")
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("Error: x")), "False", "ok-error")

    g_sActiveFilter = "NoDWG"
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("NoDrawing")), "True", "nodwg-match")
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("HasDrawing")), "False", "nodwg-nomatch")

    g_sActiveFilter = "NoIdw"
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("NoDrawing")), "True", "noidw-match")
    nFail = nFail + AssertEquals(CStr(RP_MatchesFilter("HasDrawing")), "False", "noidw-nomatch")

    g_sActiveFilter = sSaved

    If nFail = 0 Then Debug.Print "PASS [RP_MatchesFilter]"
    Test_RP_MatchesFilter = nFail
    Exit Function

EH:
    On Error Resume Next
    g_sActiveFilter = sSaved
    On Error GoTo 0
    Debug.Print "FAIL [RP_MatchesFilter] - unexpected error: " & Err.Description
    Test_RP_MatchesFilter = 1
End Function

Private Function Test_STATE_AssertGlobals() As Long
    On Error GoTo EH

    STATE_Init

    Dim sReport As String
    Dim bOK As Boolean
    bOK = STATE_AssertGlobals(sReport)

    If Not bOK Then
        Debug.Print "FAIL [STATE_AssertGlobals]"
        Debug.Print sReport
        Test_STATE_AssertGlobals = 1
        Exit Function
    End If

    Debug.Print "PASS [STATE_AssertGlobals]"
    Test_STATE_AssertGlobals = 0
    Exit Function

EH:
    Debug.Print "FAIL [STATE_AssertGlobals] - unexpected error: " & Err.Description
    Test_STATE_AssertGlobals = 1
End Function

Private Function Test_ERR_Classify() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(CStr(ERR_Classify(18)), CStr(EC_USER_CANCEL), "user-cancel")
    nFail = nFail + AssertEquals(CStr(ERR_Classify(53)), CStr(EC_NETWORK), "network-53")
    nFail = nFail + AssertEquals(CStr(ERR_Classify(76)), CStr(EC_NETWORK), "network-76")
    nFail = nFail + AssertEquals(CStr(ERR_Classify(70)), CStr(EC_PERMISSION), "permission-70")
    nFail = nFail + AssertEquals(CStr(ERR_Classify(75)), CStr(EC_PERMISSION), "permission-75")
    nFail = nFail + AssertEquals(CStr(ERR_Classify(13)), CStr(EC_DATA), "data")
    nFail = nFail + AssertEquals(CStr(ERR_Classify(55)), CStr(EC_FILE_LOCK), "file-lock")
    nFail = nFail + AssertEquals(CStr(ERR_Classify(429)), CStr(EC_COM), "com")
    nFail = nFail + AssertEquals(CStr(ERR_Classify(999)), CStr(EC_UNKNOWN), "unknown")

    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_NETWORK)), "True", "recoverable-network")
    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_FILE_LOCK)), "True", "recoverable-lock")
    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_USER_CANCEL)), "True", "recoverable-cancel")
    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_PERMISSION)), "False", "not-recoverable-perm")
    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_COM)), "False", "not-recoverable-com")

    If nFail = 0 Then Debug.Print "PASS [ERR_Classify]"
    Test_ERR_Classify = nFail
    Exit Function

EH:
    Debug.Print "FAIL [ERR_Classify] - unexpected error: " & Err.Description
    Test_ERR_Classify = 1
End Function

Private Function Test_PRJ_PathCombine() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(PRJ_PathCombine("C:\Foo", "bar.txt"), "C:\Foo\bar.txt", "basic")
    nFail = nFail + AssertEquals(PRJ_PathCombine("", "bar.txt"), "bar.txt", "empty-base")
    nFail = nFail + AssertEquals(PRJ_PathCombine("C:\Foo", ""), "C:\Foo", "empty-child")
    nFail = nFail + AssertEquals(PRJ_PathCombine("", ""), "", "both-empty")

    If nFail = 0 Then Debug.Print "PASS [PRJ_PathCombine]"
    Test_PRJ_PathCombine = nFail
    Exit Function

EH:
    Debug.Print "FAIL [PRJ_PathCombine] - unexpected error: " & Err.Description
    Test_PRJ_PathCombine = 1
End Function

Private Function Test_RP_BuildSummaryText() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    Dim t As TReportCounts
    t.nOK = 5
    t.nOver = 2
    t.nSkip = 1
    t.nEmpty = 0
    t.nErr = 0

    Dim s As String
    s = RP_BuildSummaryText(t, 1.5)

    If Len(s) = 0 Then
        Debug.Print "FAIL [RP_BuildSummaryText] - empty result"
        nFail = nFail + 1
    End If
    If InStr(s, "5") = 0 Then
        Debug.Print "FAIL [RP_BuildSummaryText] - missing nOK value 5: " & s
        nFail = nFail + 1
    End If
    If InStr(s, "2") = 0 Then
        Debug.Print "FAIL [RP_BuildSummaryText] - missing nOver value 2: " & s
        nFail = nFail + 1
    End If
    If InStr(s, "??") > 0 Then
        Debug.Print "FAIL [RP_BuildSummaryText] - unresolved key marker: " & s
        nFail = nFail + 1
    End If

    If nFail = 0 Then Debug.Print "PASS [RP_BuildSummaryText]"
    Test_RP_BuildSummaryText = nFail
    Exit Function

EH:
    Debug.Print "FAIL [RP_BuildSummaryText] - unexpected error: " & Err.Description
    Test_RP_BuildSummaryText = 1
End Function

Private Function Test_STATE_ValidateConfig() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    Dim sOrigTheme As String, sOrigSub As String
    Dim sOrigForce As String, sOrigLevel As String, sOrigLang As String

    sOrigTheme = REG_Get(REG_KEY_THEME, "I")
    sOrigSub = REG_Get(REG_KEY_SUB, "True")
    sOrigForce = REG_Get(REG_KEY_FORCE, "False")
    sOrigLevel = REG_Get(REG_KEY_LOGLEVEL, "1")
    sOrigLang = REG_Get(REG_KEY_LANG, "")

    REG_Set REG_KEY_THEME, "X"
    REG_Set REG_KEY_SUB, "maybe"
    REG_Set REG_KEY_FORCE, "yes"
    REG_Set REG_KEY_LOGLEVEL, "99"
    REG_Set REG_KEY_LANG, "zz"

    Dim sReport As String
    STATE_ValidateConfig sReport

    nFail = nFail + AssertEquals(REG_Get(REG_KEY_THEME, ""), "I", "theme-reset")
    nFail = nFail + AssertEquals(LCase$(REG_Get(REG_KEY_SUB, "")), "true", "sub-reset")
    nFail = nFail + AssertEquals(LCase$(REG_Get(REG_KEY_FORCE, "")), "false", "force-reset")
    nFail = nFail + AssertEquals(REG_Get(REG_KEY_LOGLEVEL, ""), "1", "loglevel-reset")
    nFail = nFail + AssertEquals(REG_Get(REG_KEY_LANG, "x"), "", "lang-reset")

    REG_Set REG_KEY_THEME, "IS"
    STATE_ValidateConfig sReport
    nFail = nFail + AssertEquals(REG_Get(REG_KEY_THEME, ""), "I", "theme-2char-rejected")

    REG_Set REG_KEY_THEME, "ISL"
    STATE_ValidateConfig sReport
    nFail = nFail + AssertEquals(REG_Get(REG_KEY_THEME, ""), "I", "theme-3char-rejected")

    REG_Set REG_KEY_THEME, sOrigTheme
    REG_Set REG_KEY_SUB, sOrigSub
    REG_Set REG_KEY_FORCE, sOrigForce
    REG_Set REG_KEY_LOGLEVEL, sOrigLevel
    REG_Set REG_KEY_LANG, sOrigLang

    If nFail = 0 Then Debug.Print "PASS [STATE_ValidateConfig]"
    Test_STATE_ValidateConfig = nFail
    Exit Function

EH:
    On Error Resume Next
    REG_Set REG_KEY_THEME, sOrigTheme
    REG_Set REG_KEY_SUB, sOrigSub
    REG_Set REG_KEY_FORCE, sOrigForce
    REG_Set REG_KEY_LOGLEVEL, sOrigLevel
    REG_Set REG_KEY_LANG, sOrigLang
    On Error GoTo 0
    Debug.Print "FAIL [STATE_ValidateConfig] - unexpected error: " & Err.Description
    Test_STATE_ValidateConfig = 1
End Function

Private Function Test_PDF_ValidateHeader() As Long
    Dim sTestDir   As String
    Dim sValidPDF  As String
    Dim sBadPDF    As String
    Dim sMissing   As String
    Dim nFail      As Long

    On Error GoTo EH

    nFail = 0
    PRJ_EnsureSandbox
    sTestDir = PRJ_SandboxPath() & "\_test_pdfheader"

    On Error Resume Next
    If PRJ_FSO().FolderExists(sTestDir) Then PRJ_FSO().DeleteFolder sTestDir, True
    On Error GoTo EH
    PRJ_FSO().CreateFolder sTestDir

    sValidPDF = sTestDir & "\valid.pdf"
    sBadPDF = sTestDir & "\invalid.pdf"
    sMissing = sTestDir & "\missing.pdf"

    TEST_WriteFile sValidPDF, "%PDF-1.4 test"
    TEST_WriteFile sBadPDF, "NOTPDF"

    nFail = nFail + AssertEquals(CStr(PDF_ValidateHeader(sValidPDF)), "True", "valid-pdf")
    nFail = nFail + AssertEquals(CStr(PDF_ValidateHeader(sBadPDF)), "False", "invalid")
    nFail = nFail + AssertEquals(CStr(PDF_ValidateHeader(sMissing)), "False", "missing")
    nFail = nFail + AssertEquals(CStr(PDF_ValidateHeader("")), "False", "empty")

    On Error Resume Next
    If PRJ_FSO().FolderExists(sTestDir) Then PRJ_FSO().DeleteFolder sTestDir, True
    On Error GoTo 0

    If nFail = 0 Then Debug.Print "PASS [PDF_ValidateHeader]"
    Test_PDF_ValidateHeader = nFail
    Exit Function

EH:
    Debug.Print "FAIL [PDF_ValidateHeader] - " & Err.Description
    Test_PDF_ValidateHeader = 1
End Function

Private Function Test_PDF_BuildTempPath() As Long
    Dim sSaved      As Long
    Dim sPath       As String
    Dim sPath2      As String
    Dim sSandbox    As String
    Dim nFail       As Long

    On Error GoTo EH

    nFail = 0
    sSaved = g_nTempCounter

    sSandbox = PRJ_SandboxPath()

    sPath = PDF_BuildTempPath("C:\Foo\Bar.pdf")

    If Len(sPath) = 0 Then
        Debug.Print "FAIL [PDF_BuildTempPath] - empty result"
        Test_PDF_BuildTempPath = 1
        GoTo Cleanup
    End If

    If Left$(sPath, Len(sSandbox)) <> sSandbox Then
        Debug.Print "FAIL [PDF_BuildTempPath] - not in sandbox: " & sPath
        nFail = nFail + 1
    End If

    If InStr(sPath, "_tmp_") = 0 Then
        Debug.Print "FAIL [PDF_BuildTempPath] - missing _tmp_ prefix: " & sPath
        nFail = nFail + 1
    End If

    If Right$(sPath, 4) <> ".pdf" Then
        Debug.Print "FAIL [PDF_BuildTempPath] - missing .pdf extension: " & sPath
        nFail = nFail + 1
    End If

    sPath2 = PDF_BuildTempPath("C:\Foo\Bar.pdf")
    If sPath = sPath2 Then
        Debug.Print "FAIL [PDF_BuildTempPath] - duplicate path: " & sPath
        nFail = nFail + 1
    End If

Cleanup:
    g_nTempCounter = sSaved

    If nFail = 0 Then Debug.Print "PASS [PDF_BuildTempPath]"
    Test_PDF_BuildTempPath = nFail
    Exit Function

EH:
    g_nTempCounter = sSaved
    Debug.Print "FAIL [PDF_BuildTempPath] - " & Err.Description
    Test_PDF_BuildTempPath = 1
End Function

Private Function Test_CACHE_CaseInsensitive() As Long
    Dim sTestFile As String
    Dim sUpper    As String
    Dim sLower    As String
    Dim n         As Long

    On Error GoTo EH

    PRJ_EnsureSandbox
    CACHE_Init

    sTestFile = PRJ_SandboxPath() & "\_test_cache_case.pdf"
    TEST_WriteFile sTestFile, "%PDF-1.4 case test"

    sUpper = UCase$(sTestFile)
    sLower = LCase$(sTestFile)

    CACHE_InvalidatePDF sTestFile
    CACHE_PutPageCount sTestFile, 77

    n = CACHE_GetPageCount(sUpper)
    If n <> 77 Then
        Debug.Print "FAIL [CACHE-CaseInsensitive] - uppercase lookup failed: " & n
        Test_CACHE_CaseInsensitive = 1
        GoTo Cleanup
    End If

    n = CACHE_GetPageCount(sLower)
    If n <> 77 Then
        Debug.Print "FAIL [CACHE-CaseInsensitive] - lowercase lookup failed: " & n
        Test_CACHE_CaseInsensitive = 1
        GoTo Cleanup
    End If

    CACHE_InvalidatePDF sTestFile
    Debug.Print "PASS [CACHE-CaseInsensitive]"
    Test_CACHE_CaseInsensitive = 0

Cleanup:
    On Error Resume Next
    If PRJ_FSO().FileExists(sTestFile) Then PRJ_FSO().DeleteFile sTestFile, True
    On Error GoTo 0
    Exit Function

EH:
    Debug.Print "FAIL [CACHE-CaseInsensitive] - " & Err.Description
    Test_CACHE_CaseInsensitive = 1
    Resume Cleanup
End Function

Private Function Test_PRJ_FormatElapsed_Edge() As Long
    On Error GoTo EH

    Dim nFail As Long
    Dim s     As String

    nFail = 0

    s = PRJ_FormatElapsed(0)
    If InStr(s, "0.00 s") = 0 Then
        Debug.Print "FAIL [FormatElapsed_Edge] - 0s: " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatElapsed(0.5)
    If InStr(s, "0.50 s") = 0 Then
        Debug.Print "FAIL [FormatElapsed_Edge] - 0.5s: " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatElapsed(60)
    If InStr(s, "1 min") = 0 Then
        Debug.Print "FAIL [FormatElapsed_Edge] - 60s: " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatElapsed(3600)
    If InStr(s, "1 h") = 0 Then
        Debug.Print "FAIL [FormatElapsed_Edge] - 3600s: " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatElapsed(3661.5)
    If InStr(s, "1 h") = 0 Or InStr(s, "1 min") = 0 Or InStr(s, "1.50 s") = 0 Then
        Debug.Print "FAIL [FormatElapsed_Edge] - 3661.5s: " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatElapsed(-1)
    If InStr(s, "23 h") = 0 Or InStr(s, "59 min") = 0 Then
        Debug.Print "FAIL [FormatElapsed_Edge] - negative wrap: " & s
        nFail = nFail + 1
    End If

    If nFail = 0 Then Debug.Print "PASS [FormatElapsed_Edge]"
    Test_PRJ_FormatElapsed_Edge = nFail
    Exit Function

EH:
    Debug.Print "FAIL [FormatElapsed_Edge] - unexpected error: " & Err.Description
    Test_PRJ_FormatElapsed_Edge = 1
End Function

Private Function Test_PRJ_FormatDate_Edge() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(PRJ_FormatDate(0), "unknown", "zero-date")
    nFail = nFail + AssertEquals(PRJ_FormatDate(DateSerial(1970, 1, 1)), "unknown", "pre-1980")
    nFail = nFail + AssertEquals(PRJ_FormatDate(DateSerial(1979, 12, 31)), "unknown", "1979-dec")

    Dim sValid As String
    sValid = PRJ_FormatDate(DateSerial(2024, 6, 15))
    If sValid = "unknown" Or Len(sValid) = 0 Then
        Debug.Print "FAIL [FormatDate_Edge] - valid 2024: " & sValid
        nFail = nFail + 1
    End If

    sValid = PRJ_FormatDate(DateSerial(1980, 1, 1))
    If sValid = "unknown" Or Len(sValid) = 0 Then
        Debug.Print "FAIL [FormatDate_Edge] - 1980 boundary: " & sValid
        nFail = nFail + 1
    End If

    If nFail = 0 Then Debug.Print "PASS [FormatDate_Edge]"
    Test_PRJ_FormatDate_Edge = nFail
    Exit Function

EH:
    Debug.Print "FAIL [FormatDate_Edge] - unexpected error: " & Err.Description
    Test_PRJ_FormatDate_Edge = 1
End Function

Private Function Test_SN_SanitizePath_UNC() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(SN_SanitizePath("\\server\share"), "\\server\share", "plain-share")
    nFail = nFail + AssertEquals(SN_SanitizePath("\\server\share\file.pdf"), "\\server\share\file.pdf", "share-with-file")
    nFail = nFail + AssertEquals(SN_SanitizePath("\\server\share\sub\..\other"), "\\server\share\other", "parent-within-share")
    nFail = nFail + AssertEquals(SN_SanitizePath("\\server\share\..\..\Windows"), "", "parent-escape")
    nFail = nFail + AssertEquals(SN_SanitizePath("\\server\share\sub\..\..\..\Windows"), "", "deep-escape")
    nFail = nFail + AssertEquals(SN_SanitizePath("""\\server\share\sub"""), "\\server\share\sub", "quoted-valid")
    nFail = nFail + AssertEquals(SN_SanitizePath("""\\server\share\..\..\Windows"""), "", "quoted-escape")
    nFail = nFail + AssertEquals(SN_SanitizePath("\\server"), "", "no-share")
    nFail = nFail + AssertEquals(SN_SanitizePath("\\"), "", "root-only")

    If nFail = 0 Then Debug.Print "PASS [SN_SanitizePath_UNC]"
    Test_SN_SanitizePath_UNC = nFail
    Exit Function

EH:
    Debug.Print "FAIL [SN_SanitizePath_UNC] - unexpected error: " & Err.Description
    Test_SN_SanitizePath_UNC = 1
End Function

Private Function Test_ERR_IsRecoverable_All() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_NETWORK)), "True", "network")
    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_FILE_LOCK)), "True", "file-lock")
    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_USER_CANCEL)), "True", "user-cancel")

    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_PERMISSION)), "False", "permission")
    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_DATA)), "False", "data")
    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_COM)), "False", "com")
    nFail = nFail + AssertEquals(CStr(ERR_IsRecoverable(EC_UNKNOWN)), "False", "unknown")

    If nFail = 0 Then Debug.Print "PASS [ERR_IsRecoverable_All]"
    Test_ERR_IsRecoverable_All = nFail
    Exit Function

EH:
    Debug.Print "FAIL [ERR_IsRecoverable_All] - unexpected error: " & Err.Description
    Test_ERR_IsRecoverable_All = 1
End Function

Private Function Test_ERR_DescribeClass_All() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(ERR_DescribeClass(EC_NETWORK), "network", "network")
    nFail = nFail + AssertEquals(ERR_DescribeClass(EC_PERMISSION), "permission", "permission")
    nFail = nFail + AssertEquals(ERR_DescribeClass(EC_DATA), "data", "data")
    nFail = nFail + AssertEquals(ERR_DescribeClass(EC_FILE_LOCK), "file-lock", "file-lock")
    nFail = nFail + AssertEquals(ERR_DescribeClass(EC_COM), "com", "com")
    nFail = nFail + AssertEquals(ERR_DescribeClass(EC_USER_CANCEL), "user-cancel", "user-cancel")
    nFail = nFail + AssertEquals(ERR_DescribeClass(EC_UNKNOWN), "unknown", "unknown")

    If nFail = 0 Then Debug.Print "PASS [ERR_DescribeClass_All]"
    Test_ERR_DescribeClass_All = nFail
    Exit Function

EH:
    Debug.Print "FAIL [ERR_DescribeClass_All] - unexpected error: " & Err.Description
    Test_ERR_DescribeClass_All = 1
End Function

Private Function Test_L_Placeholders() As Long
    On Error GoTo EH

    Dim nFail As Long
    Dim s     As String

    nFail = 0

    s = L2("scan.sum.drawings", "5")
    If InStr(s, "{0}") > 0 Then
        Debug.Print "FAIL [L_Placeholders] - L2 left {0}: " & s
        nFail = nFail + 1
    End If
    If InStr(s, "5") = 0 Then
        Debug.Print "FAIL [L_Placeholders] - L2 missing value 5: " & s
        nFail = nFail + 1
    End If
    If InStr(s, "??") > 0 Then
        Debug.Print "FAIL [L_Placeholders] - L2 unresolved key: " & s
        nFail = nFail + 1
    End If

    s = L3("pdf.err.pagematch", "10", "20")
    If InStr(s, "{0}") > 0 Or InStr(s, "{1}") > 0 Then
        Debug.Print "FAIL [L_Placeholders] - L3 left placeholder: " & s
        nFail = nFail + 1
    End If
    If InStr(s, "10") = 0 Then
        Debug.Print "FAIL [L_Placeholders] - L3 missing value 10: " & s
        nFail = nFail + 1
    End If
    If InStr(s, "20") = 0 Then
        Debug.Print "FAIL [L_Placeholders] - L3 missing value 20: " & s
        nFail = nFail + 1
    End If
    If InStr(s, "??") > 0 Then
        Debug.Print "FAIL [L_Placeholders] - L3 unresolved key: " & s
        nFail = nFail + 1
    End If

    s = LN("scan.sum.drawings", "7")
    If InStr(s, "7") = 0 Then
        Debug.Print "FAIL [L_Placeholders] - LN missing value 7: " & s
        nFail = nFail + 1
    End If

    s = LN("pdf.err.pagematch", "10", "20")
    If InStr(s, "10") = 0 Or InStr(s, "20") = 0 Then
        Debug.Print "FAIL [L_Placeholders] - LN missing values: " & s
        nFail = nFail + 1
    End If

    If nFail = 0 Then Debug.Print "PASS [L_Placeholders]"
    Test_L_Placeholders = nFail
    Exit Function

EH:
    Debug.Print "FAIL [L_Placeholders] - unexpected error: " & Err.Description
    Test_L_Placeholders = 1
End Function

Private Function Test_U_UnicodeDecoder() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(U("Hello"), "Hello", "plain")
    nFail = nFail + AssertEquals(U("\u0041"), "A", "single-A")
    nFail = nFail + AssertEquals(U("\u00e9"), ChrW$(&HE9), "single-eacute")
    nFail = nFail + AssertEquals(U("A\u0042C"), "ABC", "mixed")
    nFail = nFail + AssertEquals(U("\u0041\u0042\u0043"), "ABC", "multi")

    Dim sMalformed As String
    sMalformed = U("\u00")
    If Len(sMalformed) <> 4 Then
        Debug.Print "FAIL [U_UnicodeDecoder] - malformed length: " & Len(sMalformed)
        nFail = nFail + 1
    End If

    nFail = nFail + AssertEquals(U(""), "", "empty")

    If nFail = 0 Then Debug.Print "PASS [U_UnicodeDecoder]"
    Test_U_UnicodeDecoder = nFail
    Exit Function

EH:
    Debug.Print "FAIL [U_UnicodeDecoder] - unexpected error: " & Err.Description
    Test_U_UnicodeDecoder = 1
End Function

Private Function Test_LANG_NormalizeCode() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(LANG_NormalizeCode("zh-cn"), "zh", "zh-cn")
    nFail = nFail + AssertEquals(LANG_NormalizeCode("zh-tw"), "zht", "zh-tw")
    nFail = nFail + AssertEquals(LANG_NormalizeCode("ZH-CN"), "zh", "zh-cn-upper")
    nFail = nFail + AssertEquals(LANG_NormalizeCode("ZH-TW"), "zht", "zh-tw-upper")
    nFail = nFail + AssertEquals(LANG_NormalizeCode("en"), "en", "en")
    nFail = nFail + AssertEquals(LANG_NormalizeCode("EN"), "en", "en-upper")
    nFail = nFail + AssertEquals(LANG_NormalizeCode("english"), "en", "english-truncated")
    nFail = nFail + AssertEquals(LANG_NormalizeCode("german"), "ge", "german-truncated")
    nFail = nFail + AssertEquals(LANG_NormalizeCode(""), "", "empty")
    nFail = nFail + AssertEquals(LANG_NormalizeCode("   "), "", "whitespace")
    nFail = nFail + AssertEquals(LANG_NormalizeCode("  en  "), "en", "whitespace-padded")

    If nFail = 0 Then Debug.Print "PASS [LANG_NormalizeCode]"
    Test_LANG_NormalizeCode = nFail
    Exit Function

EH:
    Debug.Print "FAIL [LANG_NormalizeCode] - unexpected error: " & Err.Description
    Test_LANG_NormalizeCode = 1
End Function

Private Function Test_RP_CountFiles() As Long
    Dim nFail  As Long
    Dim t      As TReportCounts
    Dim saved  As Object

    On Error GoTo EH

    nFail = 0

    Set saved = g_colFiles
    Set g_colFiles = CreateObject("System.Collections.ArrayList")

    g_colFiles.Add Array("C:\a.idw", "Created | 1/1 pages")
    g_colFiles.Add Array("C:\b.idw", "Overwritten | 2/2 pages")
    g_colFiles.Add Array("C:\c.idw", "Skipped (Exists)")
    g_colFiles.Add Array("C:\d.idw", "Skipped (Empty)")
    g_colFiles.Add Array("C:\e.idw", "Error: something")
    g_colFiles.Add Array("C:\f.ipt", "HasDrawing")
    g_colFiles.Add Array("C:\g.ipt", "NoDrawing")

    RP_CountFiles t

    nFail = nFail + AssertEquals(CStr(t.nOK), "1", "nOK")
    nFail = nFail + AssertEquals(CStr(t.nOver), "1", "nOver")
    nFail = nFail + AssertEquals(CStr(t.nSkip), "2", "nSkip")
    nFail = nFail + AssertEquals(CStr(t.nEmpty), "1", "nEmpty")
    nFail = nFail + AssertEquals(CStr(t.nErr), "1", "nErr")

    Set g_colFiles = saved

    If nFail = 0 Then Debug.Print "PASS [RP_CountFiles]"
    Test_RP_CountFiles = nFail
    Exit Function

EH:
    On Error Resume Next
    Set g_colFiles = saved
    On Error GoTo 0
    Debug.Print "FAIL [RP_CountFiles] - unexpected error: " & Err.Description
    Test_RP_CountFiles = 1
End Function

Private Function Test_RP_PickSummaryColor() As Long
    On Error GoTo EH

    Dim nFail As Long
    Dim t     As TReportCounts

    nFail = 0

    t.nOK = 0: t.nOver = 0: t.nSkip = 0: t.nEmpty = 0: t.nErr = 0
    nFail = nFail + AssertEquals(CStr(RP_PickSummaryColor(t)), CStr(CLR_TEXT_SEC), "all-zero")

    t.nOK = 3
    nFail = nFail + AssertEquals(CStr(RP_PickSummaryColor(t)), CStr(CLR_OK), "ok-only")

    t.nSkip = 1
    nFail = nFail + AssertEquals(CStr(RP_PickSummaryColor(t)), CStr(CLR_WARN), "with-skip")

    t.nErr = 1
    nFail = nFail + AssertEquals(CStr(RP_PickSummaryColor(t)), CStr(CLR_ERROR), "with-error")

    If nFail = 0 Then Debug.Print "PASS [RP_PickSummaryColor]"
    Test_RP_PickSummaryColor = nFail
    Exit Function

EH:
    Debug.Print "FAIL [RP_PickSummaryColor] - unexpected error: " & Err.Description
    Test_RP_PickSummaryColor = 1
End Function

Private Function Test_RP_PickStatusBold() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(CStr(RP_PickStatusBold("Error: x")), "True", "error-bold")
    nFail = nFail + AssertEquals(CStr(RP_PickStatusBold("Created | 1/1 pages")), "False", "created-not-bold")
    nFail = nFail + AssertEquals(CStr(RP_PickStatusBold("Skipped (Exists)")), "False", "skipped-not-bold")
    nFail = nFail + AssertEquals(CStr(RP_PickStatusBold("Overwritten | 1/1 pages")), "False", "overwritten-not-bold")
    nFail = nFail + AssertEquals(CStr(RP_PickStatusBold("HasDrawing")), "False", "hasdrawing-not-bold")
    nFail = nFail + AssertEquals(CStr(RP_PickStatusBold("NoDrawing")), "False", "nodrawing-not-bold")

    If nFail = 0 Then Debug.Print "PASS [RP_PickStatusBold]"
    Test_RP_PickStatusBold = nFail
    Exit Function

EH:
    Debug.Print "FAIL [RP_PickStatusBold] - unexpected error: " & Err.Description
    Test_RP_PickStatusBold = 1
End Function

Private Function Test_RP_StatusColor() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(CStr(RP_StatusColor("Error: x")), CStr(CLR_ERROR), "error")
    nFail = nFail + AssertEquals(CStr(RP_StatusColor("Skipped (Exists)")), CStr(CLR_WARN), "skipped")
    nFail = nFail + AssertEquals(CStr(RP_StatusColor("Skipped (Empty)")), CStr(CLR_WARN), "skipped-empty")
    nFail = nFail + AssertEquals(CStr(RP_StatusColor("HasDrawing")), CStr(CLR_TEXT_SEC), "hasdrawing")
    nFail = nFail + AssertEquals(CStr(RP_StatusColor("NoDrawing")), CStr(CLR_TEXT_SEC), "nodrawing")
    nFail = nFail + AssertEquals(CStr(RP_StatusColor("Created | 1/1 pages")), CStr(CLR_OK), "created")
    nFail = nFail + AssertEquals(CStr(RP_StatusColor("Overwritten | 1/1 pages")), CStr(CLR_OK), "overwritten")

    If nFail = 0 Then Debug.Print "PASS [RP_StatusColor]"
    Test_RP_StatusColor = nFail
    Exit Function

EH:
    Debug.Print "FAIL [RP_StatusColor] - unexpected error: " & Err.Description
    Test_RP_StatusColor = 1
End Function

Private Function Test_CACHE_EvictIfNeeded() As Long
    Dim saved     As Object
    Dim sTestFile As String
    Dim i         As Long
    Dim nFill     As Long
    Dim nBefore   As Long
    Dim nAfter    As Long
    Dim nExpected As Long

    On Error GoTo EH

    PRJ_EnsureSandbox
    CACHE_Init

    Set saved = g_dPageCache
    Set g_dPageCache = CreateObject("Scripting.Dictionary")
    g_dPageCache.CompareMode = vbTextCompare

    nFill = CACHE_MAX_ENTRIES + 10
    For i = 1 To nFill
        g_dPageCache("fake_key_" & CStr(i)) = "20260101000000|1"
    Next i

    nBefore = g_dPageCache.Count
    If nBefore <> nFill Then
        Debug.Print "FAIL [CACHE_EvictIfNeeded] - fill count " & nBefore & " expected " & nFill
        Test_CACHE_EvictIfNeeded = 1
        GoTo Cleanup
    End If

    sTestFile = PRJ_SandboxPath() & "\_test_evict.pdf"
    TEST_WriteFile sTestFile, "%PDF-1.4 evict"

    CACHE_PutPageCount sTestFile, 1

    nAfter = g_dPageCache.Count
    nExpected = nFill + 1 - CACHE_EVICT_BATCH

    If nAfter <> nExpected Then
        Debug.Print "FAIL [CACHE_EvictIfNeeded] - after=" & nAfter & " expected=" & nExpected
        Test_CACHE_EvictIfNeeded = 1
        GoTo Cleanup
    End If

    Debug.Print "PASS [CACHE_EvictIfNeeded]"
    Test_CACHE_EvictIfNeeded = 0

Cleanup:
    Set g_dPageCache = saved
    On Error Resume Next
    If PRJ_FSO().FileExists(sTestFile) Then PRJ_FSO().DeleteFile sTestFile, True
    On Error GoTo 0
    Exit Function

EH:
    On Error Resume Next
    Set g_dPageCache = saved
    On Error GoTo 0
    Debug.Print "FAIL [CACHE_EvictIfNeeded] - " & Err.Description
    Test_CACHE_EvictIfNeeded = 1
End Function

Private Function Test_PRJ_GetRelPath() As Long
    Dim nFail       As Long
    Dim savedFolder As String

    On Error GoTo EH

    nFail = 0
    savedFolder = g_sLastFolder

    g_sLastFolder = ""
    nFail = nFail + AssertEquals(PRJ_GetRelPath("C:\Foo\Bar.pdf"), "C:\Foo\Bar.pdf", "no-folder")

    g_sLastFolder = "C:\Foo"
    nFail = nFail + AssertEquals(PRJ_GetRelPath("C:\Foo\Bar.pdf"), "Bar.pdf", "direct-child")

    g_sLastFolder = "C:\Foo"
    nFail = nFail + AssertEquals(PRJ_GetRelPath("C:\Foo\Sub\Bar.pdf"), "Sub\Bar.pdf", "nested")

    g_sLastFolder = "C:\Foo"
    nFail = nFail + AssertEquals(PRJ_GetRelPath("c:\foo\sub\bar.pdf"), "sub\bar.pdf", "case-insensitive")

    g_sLastFolder = "C:\Foo\"
    nFail = nFail + AssertEquals(PRJ_GetRelPath("C:\Foo\Bar.pdf"), "Bar.pdf", "trailing-slash")

    g_sLastFolder = "C:\Other"
    nFail = nFail + AssertEquals(PRJ_GetRelPath("C:\Foo\Bar.pdf"), "C:\Foo\Bar.pdf", "outside")

    g_sLastFolder = savedFolder

    If nFail = 0 Then Debug.Print "PASS [PRJ_GetRelPath]"
    Test_PRJ_GetRelPath = nFail
    Exit Function

EH:
    g_sLastFolder = savedFolder
    Debug.Print "FAIL [PRJ_GetRelPath] - " & Err.Description
    Test_PRJ_GetRelPath = 1
End Function

Private Function Test_RETRY_DeleteFile_MoveFile() As Long
    Dim sTestDir As String
    Dim sSrc     As String
    Dim sDst     As String
    Dim bOK      As Boolean

    On Error GoTo EH

    PRJ_EnsureSandbox
    sTestDir = PRJ_SandboxPath() & "\_test_retry"

    On Error Resume Next
    If PRJ_FSO().FolderExists(sTestDir) Then PRJ_FSO().DeleteFolder sTestDir, True
    On Error GoTo EH
    PRJ_FSO().CreateFolder sTestDir

    sSrc = sTestDir & "\del_src.txt"
    TEST_WriteFile sSrc, "content"

    bOK = RETRY_DeleteFile(sSrc)
    If Not bOK Then
        Debug.Print "FAIL [RETRY] - DeleteFile returned False on existing file"
        Test_RETRY_DeleteFile_MoveFile = 1
        GoTo Cleanup
    End If
    If PRJ_FSO().FileExists(sSrc) Then
        Debug.Print "FAIL [RETRY] - file still exists after DeleteFile"
        Test_RETRY_DeleteFile_MoveFile = 1
        GoTo Cleanup
    End If

    sSrc = sTestDir & "\move_src.txt"
    sDst = sTestDir & "\move_dst.txt"
    TEST_WriteFile sSrc, "move content"

    bOK = RETRY_MoveFile(sSrc, sDst)
    If Not bOK Then
        Debug.Print "FAIL [RETRY] - MoveFile returned False on existing file"
        Test_RETRY_DeleteFile_MoveFile = 1
        GoTo Cleanup
    End If
    If PRJ_FSO().FileExists(sSrc) Then
        Debug.Print "FAIL [RETRY] - source still exists after MoveFile"
        Test_RETRY_DeleteFile_MoveFile = 1
        GoTo Cleanup
    End If
    If Not PRJ_FSO().FileExists(sDst) Then
        Debug.Print "FAIL [RETRY] - destination missing after MoveFile"
        Test_RETRY_DeleteFile_MoveFile = 1
        GoTo Cleanup
    End If

    Debug.Print "PASS [RETRY_DeleteFile_MoveFile]"
    Test_RETRY_DeleteFile_MoveFile = 0

Cleanup:
    On Error Resume Next
    If PRJ_FSO().FolderExists(sTestDir) Then PRJ_FSO().DeleteFolder sTestDir, True
    On Error GoTo 0
    Exit Function

EH:
    Debug.Print "FAIL [RETRY] - " & Err.Description
    Test_RETRY_DeleteFile_MoveFile = 1
    Resume Cleanup
End Function

Private Function Test_QA_AnalyzePDF_TooSmall() As Long
    Dim sTestDir    As String
    Dim sSmall      As String
    Dim sBig        As String
    Dim sBigContent As String
    Dim sResult     As String
    Dim savedFlag   As Boolean
    Dim nFail       As Long
    Dim i           As Long

    savedFlag = g_bAnalyzePDF

    On Error GoTo EH

    nFail = 0
    PRJ_EnsureSandbox
    sTestDir = PRJ_SandboxPath() & "\_test_qa"

    On Error Resume Next
    If PRJ_FSO().FolderExists(sTestDir) Then PRJ_FSO().DeleteFolder sTestDir, True
    On Error GoTo EH
    PRJ_FSO().CreateFolder sTestDir

    sSmall = sTestDir & "\small.pdf"
    sBig = sTestDir & "\big.pdf"

    TEST_WriteFile sSmall, "%PDF-1.4 " & String(90, "x")

    sBigContent = "%PDF-1.4 "
    For i = 1 To 1200
        sBigContent = sBigContent & "0123456789"
    Next i
    TEST_WriteFile sBig, sBigContent

    g_bAnalyzePDF = True

    sResult = QA_AnalyzePDF(sSmall)
    If InStr(sResult, QA_STAT_PDF_TOO_SMALL) = 0 Then
        Debug.Print "FAIL [QA_AnalyzePDF_TooSmall] - no warning for small: [" & sResult & "]"
        nFail = nFail + 1
    End If

    sResult = QA_AnalyzePDF(sBig)
    If Len(sResult) > 0 Then
        Debug.Print "FAIL [QA_AnalyzePDF_TooSmall] - unexpected warning for big: [" & sResult & "]"
        nFail = nFail + 1
    End If

    g_bAnalyzePDF = False
    sResult = QA_AnalyzePDF(sSmall)
    If Len(sResult) > 0 Then
        Debug.Print "FAIL [QA_AnalyzePDF_TooSmall] - warning when flag off: [" & sResult & "]"
        nFail = nFail + 1
    End If

    g_bAnalyzePDF = savedFlag

    On Error Resume Next
    If PRJ_FSO().FolderExists(sTestDir) Then PRJ_FSO().DeleteFolder sTestDir, True
    On Error GoTo 0

    If nFail = 0 Then Debug.Print "PASS [QA_AnalyzePDF_TooSmall]"
    Test_QA_AnalyzePDF_TooSmall = nFail
    Exit Function

EH:
    g_bAnalyzePDF = savedFlag
    Debug.Print "FAIL [QA_AnalyzePDF_TooSmall] - " & Err.Description
    Test_QA_AnalyzePDF_TooSmall = 1
End Function

Private Function Test_RP_FindStat() As Long
    Dim nFail  As Long
    Dim saved  As Object
    Dim sFound As String

    On Error GoTo EH

    nFail = 0

    Set saved = g_colFiles
    Set g_colFiles = CreateObject("System.Collections.ArrayList")

    g_colFiles.Add Array("C:\a.idw", "Created | 1/1 pages")
    g_colFiles.Add Array("C:\b.idw", "Error: something")
    g_colFiles.Add Array("C:\c.idw", "Overwritten | 2/2 pages")

    sFound = RP_FindStat("C:\a.idw")
    nFail = nFail + AssertEquals(sFound, "Created | 1/1 pages", "find-a")

    sFound = RP_FindStat("C:\b.idw")
    nFail = nFail + AssertEquals(sFound, "Error: something", "find-b")

    sFound = RP_FindStat("C:\c.idw")
    nFail = nFail + AssertEquals(sFound, "Overwritten | 2/2 pages", "find-c")

    sFound = RP_FindStat("C:\missing.idw")
    nFail = nFail + AssertEquals(sFound, "", "missing")

    Set g_colFiles = saved

    If nFail = 0 Then Debug.Print "PASS [RP_FindStat]"
    Test_RP_FindStat = nFail
    Exit Function

EH:
    On Error Resume Next
    Set g_colFiles = saved
    On Error GoTo 0
    Debug.Print "FAIL [RP_FindStat] - " & Err.Description
    Test_RP_FindStat = 1
End Function

Private Function Test_RP_CountErrorRows() As Long
    Dim nFail As Long
    Dim saved As Object
    Dim n     As Long

    On Error GoTo EH

    nFail = 0

    Set saved = g_colFiles
    Set g_colFiles = CreateObject("System.Collections.ArrayList")

    n = RP_CountErrorRows()
    nFail = nFail + AssertEquals(CStr(n), "0", "empty")

    g_colFiles.Add Array("C:\a.idw", "Error: x")
    g_colFiles.Add Array("C:\b.idw", "Created | 1/1 pages")
    g_colFiles.Add Array("C:\c.idw", "Error: y")
    g_colFiles.Add Array("C:\d.idw", "Skipped (Exists)")
    g_colFiles.Add Array("C:\e.idw", "Error: z")

    n = RP_CountErrorRows()
    nFail = nFail + AssertEquals(CStr(n), "3", "three-errors")

    Set g_colFiles = saved

    If nFail = 0 Then Debug.Print "PASS [RP_CountErrorRows]"
    Test_RP_CountErrorRows = nFail
    Exit Function

EH:
    On Error Resume Next
    Set g_colFiles = saved
    On Error GoTo 0
    Debug.Print "FAIL [RP_CountErrorRows] - " & Err.Description
    Test_RP_CountErrorRows = 1
End Function

Private Function Test_FormatFileSize_Boundary() As Long
    On Error GoTo EH

    Dim nFail As Long
    Dim s     As String

    nFail = 0

    nFail = nFail + AssertEquals(PRJ_FormatFileSize(0), "0 B", "zero")
    nFail = nFail + AssertEquals(PRJ_FormatFileSize(1), "1 B", "one")

    s = PRJ_FormatFileSize(1023)
    If Right$(s, 2) = "KB" Or Right$(s, 2) = "MB" Or Right$(s, 2) = "GB" Then
        Debug.Print "FAIL [FormatFileSize_Boundary] - 1023 in KB+: " & s
        nFail = nFail + 1
    End If
    If Right$(s, 1) <> "B" Then
        Debug.Print "FAIL [FormatFileSize_Boundary] - 1023 missing B: " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatFileSize(1024)
    If Right$(s, 2) <> "KB" Then
        Debug.Print "FAIL [FormatFileSize_Boundary] - 1024 not KB: " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatFileSize(1048575)
    If Right$(s, 2) <> "KB" Then
        Debug.Print "FAIL [FormatFileSize_Boundary] - 1048575 not KB: " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatFileSize(1048576)
    If Right$(s, 2) <> "MB" Then
        Debug.Print "FAIL [FormatFileSize_Boundary] - 1048576 not MB: " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatFileSize(1073741823)
    If Right$(s, 2) <> "MB" Then
        Debug.Print "FAIL [FormatFileSize_Boundary] - 1073741823 not MB: " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatFileSize(1073741824)
    If Right$(s, 2) <> "GB" Then
        Debug.Print "FAIL [FormatFileSize_Boundary] - 1073741824 not GB: " & s
        nFail = nFail + 1
    End If

    If nFail = 0 Then Debug.Print "PASS [FormatFileSize_Boundary]"
    Test_FormatFileSize_Boundary = nFail
    Exit Function

EH:
    Debug.Print "FAIL [FormatFileSize_Boundary] - " & Err.Description
    Test_FormatFileSize_Boundary = 1
End Function

Private Function Test_STS_Classify_Edge() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(CStr(STS_Classify("")), CStr(ST_UNKNOWN), "empty")
    nFail = nFail + AssertEquals(CStr(STS_Classify("   ")), CStr(ST_UNKNOWN), "whitespace")

    nFail = nFail + AssertEquals(CStr(STS_Classify("error: x")), CStr(ST_UNKNOWN), "lowercase-error")
    nFail = nFail + AssertEquals(CStr(STS_Classify("created")), CStr(ST_UNKNOWN), "lowercase-created")

    nFail = nFail + AssertEquals(CStr(STS_Classify("Creat")), CStr(ST_UNKNOWN), "partial-created")

    nFail = nFail + AssertEquals(CStr(STS_Classify("Created")), CStr(ST_CREATED), "exact-created")
    nFail = nFail + AssertEquals(CStr(STS_Classify("Error")), CStr(ST_ERROR), "exact-error")

    nFail = nFail + AssertEquals(CStr(STS_Classify("CreatedXYZ")), CStr(ST_CREATED), "created-extended")
    nFail = nFail + AssertEquals(CStr(STS_Classify("ErrorXYZ")), CStr(ST_ERROR), "error-extended")

    If nFail = 0 Then Debug.Print "PASS [STS_Classify_Edge]"
    Test_STS_Classify_Edge = nFail
    Exit Function

EH:
    Debug.Print "FAIL [STS_Classify_Edge] - unexpected error: " & Err.Description
    Test_STS_Classify_Edge = 1
End Function

Private Function Test_LANG_SupportedCodes() As Long
    Dim nFail     As Long
    Dim vCodes    As Variant
    Dim aExpected As Variant
    Dim i         As Long
    Dim j         As Long
    Dim bFound    As Boolean

    On Error GoTo EH

    nFail = 0

    vCodes = LANG_SupportedCodes()
    aExpected = Array("", "en", "de", "fr", "it", "es", "pt", _
                      "ru", "pl", "cs", "ja", "ko", "zh", "zht")

    For i = LBound(aExpected) To UBound(aExpected)
        bFound = False
        For j = LBound(vCodes) To UBound(vCodes)
            If CStr(vCodes(j)) = CStr(aExpected(i)) Then
                bFound = True
                Exit For
            End If
        Next j
        If Not bFound Then
            Debug.Print "FAIL [LANG_SupportedCodes] - missing code: '" & CStr(aExpected(i)) & "'"
            nFail = nFail + 1
        End If
    Next i

    Dim nActual As Long
    Dim nExpected As Long
    nActual = UBound(vCodes) - LBound(vCodes) + 1
    nExpected = UBound(aExpected) - LBound(aExpected) + 1

    If nActual <> nExpected Then
        Debug.Print "FAIL [LANG_SupportedCodes] - count " & nActual & " expected " & nExpected
        nFail = nFail + 1
    End If

    If nFail = 0 Then Debug.Print "PASS [LANG_SupportedCodes]"
    Test_LANG_SupportedCodes = nFail
    Exit Function

EH:
    Debug.Print "FAIL [LANG_SupportedCodes] - unexpected error: " & Err.Description
    Test_LANG_SupportedCodes = 1
End Function

Private Function Test_FormatElapsed_Boundary() As Long
    On Error GoTo EH

    Dim nFail As Long
    Dim s     As String

    nFail = 0

    s = PRJ_FormatElapsed(59.99)
    If InStr(s, "min") > 0 Then
        Debug.Print "FAIL [FormatElapsed_Boundary] - 59.99 shows min: " & s
        nFail = nFail + 1
    End If
    If InStr(s, "59.99 s") = 0 Then
        Debug.Print "FAIL [FormatElapsed_Boundary] - 59.99 not '59.99 s': " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatElapsed(60)
    If InStr(s, "1 min") = 0 Then
        Debug.Print "FAIL [FormatElapsed_Boundary] - 60 missing '1 min': " & s
        nFail = nFail + 1
    End If
    If InStr(s, " h ") > 0 Then
        Debug.Print "FAIL [FormatElapsed_Boundary] - 60 shows 'h': " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatElapsed(3599)
    If InStr(s, " h ") > 0 Then
        Debug.Print "FAIL [FormatElapsed_Boundary] - 3599 shows 'h': " & s
        nFail = nFail + 1
    End If
    If InStr(s, "59 min") = 0 Then
        Debug.Print "FAIL [FormatElapsed_Boundary] - 3599 missing '59 min': " & s
        nFail = nFail + 1
    End If

    s = PRJ_FormatElapsed(3600)
    If InStr(s, "1 h") = 0 Then
        Debug.Print "FAIL [FormatElapsed_Boundary] - 3600 missing '1 h': " & s
        nFail = nFail + 1
    End If

    If nFail = 0 Then Debug.Print "PASS [FormatElapsed_Boundary]"
    Test_FormatElapsed_Boundary = nFail
    Exit Function

EH:
    Debug.Print "FAIL [FormatElapsed_Boundary] - unexpected error: " & Err.Description
    Test_FormatElapsed_Boundary = 1
End Function

Private Function Test_STS_Classify_ValidateStatus() As Long
    On Error GoTo EH

    Dim nFail As Long
    nFail = 0

    nFail = nFail + AssertEquals(CStr(STS_Classify("Missing")), CStr(ST_MISSING), "missing")
    nFail = nFail + AssertEquals(CStr(STS_Classify("Stale")), CStr(ST_STALE), "stale")
    nFail = nFail + AssertEquals(CStr(STS_Classify("Mismatch:3/5")), CStr(ST_MISMATCH), "mismatch-prefix")
    nFail = nFail + AssertEquals(CStr(STS_Classify("Mismatch")), CStr(ST_MISMATCH), "mismatch-bare")
    nFail = nFail + AssertEquals(CStr(STS_Classify("OK")), CStr(ST_OK), "ok")

    nFail = nFail + AssertEquals(CStr(STS_Classify("ok")), CStr(ST_UNKNOWN), "lowercase-ok")
    nFail = nFail + AssertEquals(CStr(STS_Classify("MissingFile")), CStr(ST_UNKNOWN), "missing-extended-rejected")

    If nFail = 0 Then Debug.Print "PASS [STS_Classify_ValidateStatus]"
    Test_STS_Classify_ValidateStatus = nFail
    Exit Function

EH:
    Debug.Print "FAIL [STS_Classify_ValidateStatus] - unexpected error: " & Err.Description
    Test_STS_Classify_ValidateStatus = 1
End Function

Private Function Test_STATE_Init_ClearsRegCache() As Long
    Dim nFail As Long
    Dim sSaved As String
    Dim sCached As String

    nFail = 0
    sSaved = "I"

    On Error GoTo EH

    sSaved = REG_Get(REG_KEY_THEME, "I")
    If Len(sSaved) = 0 Then sSaved = "I"

    sCached = REG_Get(REG_KEY_THEME, "I")

    SaveSetting REG_SECTION, REG_PREFS, REG_KEY_THEME, "X"

    nFail = nFail + AssertEquals(REG_Get(REG_KEY_THEME, ""), sCached, "cache-still-stale")

    STATE_Init

    nFail = nFail + AssertEquals(REG_Get(REG_KEY_THEME, ""), "X", "post-init-refresh")

    GoTo Cleanup

EH:
    Debug.Print "FAIL [STATE_Init_ClearsRegCache] - " & Err.Description
    nFail = nFail + 1

Cleanup:
    On Error Resume Next
    SaveSetting REG_SECTION, REG_PREFS, REG_KEY_THEME, sSaved
    REG_ClearCache
    On Error GoTo 0

    If nFail = 0 Then Debug.Print "PASS [STATE_Init_ClearsRegCache]"
    Test_STATE_Init_ClearsRegCache = nFail
End Function

Private Function Test_PRJ_PutClipboardFallback_EnsuresSandbox() As Long
    Dim nFail As Long
    Dim sSandbox As String
    Dim sPayload As String
    Dim bOK As Boolean

    nFail = 0

    On Error GoTo EH

    sSandbox = PRJ_SandboxPath()

    ' v9.19.6: close log handles before wiping the sandbox.
    ' The main log and audit log are opened with VBA Open For Append
    ' and hold a write lock on files inside the sandbox. DeleteFolder
    ' fails (Error 75) with those locks in place, and the test would
    ' see a folder that "was not deleted" - a false FAIL.
    PRJ_LogFlush

    On Error Resume Next
    Err.Clear
    If PRJ_FSO().FolderExists(sSandbox) Then
        PRJ_FSO().DeleteFolder sSandbox, True
        If Err.Number <> 0 Then
            Debug.Print "  WARN [PutClipboardFallback] - DeleteFolder err " & _
                        Err.Number & ": " & Err.Description
            Err.Clear
        End If
    End If
    On Error GoTo EH

    If PRJ_FSO().FolderExists(sSandbox) Then
        Debug.Print "FAIL [PutClipboardFallback] - sandbox still present after wipe"
        nFail = nFail + 1
    End If

    sPayload = "test-content-" & CStr(Int(Timer * 1000))

    bOK = PRJ_PutClipboardFallback(sPayload)

    If Not bOK Then
        Debug.Print "FAIL [PutClipboardFallback] - returned False"
        nFail = nFail + 1
    End If

    If Not PRJ_FSO().FolderExists(sSandbox) Then
        Debug.Print "FAIL [PutClipboardFallback] - sandbox not recreated"
        nFail = nFail + 1
    End If

    GoTo Cleanup

EH:
    Debug.Print "FAIL [PRJ_PutClipboardFallback_EnsuresSandbox] - " & Err.Description
    nFail = nFail + 1

Cleanup:
    On Error Resume Next
    PRJ_EnsureSandbox
    On Error GoTo 0

    If nFail = 0 Then Debug.Print "PASS [PRJ_PutClipboardFallback_EnsuresSandbox]"
    Test_PRJ_PutClipboardFallback_EnsuresSandbox = nFail
End Function

' END: modTests v9.19.6

Private Function AssertEquals(ByVal a As Variant, ByVal b As Variant, ByVal sTag As String) As Long
    If CStr(a) = CStr(b) Then
        AssertEquals = 0
    Else
        Debug.Print "  FAIL [" & sTag & "]: expected [" & CStr(b) & "], got [" & CStr(a) & "]"
        AssertEquals = 1
    End If
End Function

Private Sub TEST_WriteFile(ByVal sPath As String, ByVal sContent As String)
    Dim ff As Integer
    ff = FreeFile
    Open sPath For Output As #ff
    Print #ff, sContent
    Close #ff
End Sub

Private Function TEST_ReadFile(ByVal sPath As String) As String
    Dim ff     As Integer
    Dim sLine  As String
    Dim sAll   As String

    sAll = ""
    ff = FreeFile
    Open sPath For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, sLine
        sAll = sAll & sLine
    Loop
    Close #ff
    TEST_ReadFile = sAll
End Function

' END: modTests v9.19.9

