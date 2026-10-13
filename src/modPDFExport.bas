Attribute VB_Name = "modPDFExport"
Option Explicit
' Module    : modPDFExport
' Purpose   : IDW -> PDF export pipeline orchestrator
' Reference : Appendix A.13, A.31
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + Requires fix):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Fixed Requires section: STAT_* constants and PDF_GUID live in
'   modConstants, not modState. The prior wording "(modState)" was
'   incorrect — verified against the actual declarations.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in procedure signatures and in a few multi-line
'   expressions). No functional changes.
'
' v9.18.0 (§9.10 API versioning):
'   Added PDF_FindAddIn. Primary lookup by PDF_GUID; fallback to
'   enumeration by DisplayName containing "PDF". Handles Autodesk
'   GUID changes in future Inventor versions.
' v9.18.0 (§4.11, §4.15 Localized error strings):
'   Hardcoded EN replaced with L() keys:
'     "Could not open IDW"     -> L("pdf.err.open")
'     "Prepare failed: {0}"    -> L2("pdf.err.prepare", ...)
'     "header validation"      -> L("pdf.err.header")
'     "PDF has N pages"        -> L3("pdf.err.pagematch", ...)
'     "Unknown error"          -> L("pdf.err.unknown")
' v9.18.0 (§3.1 Procedure length):
'   Split PRJ_CreatePDF into thin wrapper + PDF_RunPipeline.
' v9.18.0: Err.Clear before each Inventor COM call.
'
' Requires  : STAT_ERROR, STAT_SKIPPED_EMPTY, STAT_OVERWRITTEN,
'             STAT_CREATED, PDF_GUID (modConstants).
'             TPdfTiming (modPDFProgress).
'             TM_Now, TM_Elapsed (modTime).
'             QA_AnalyzeIDW, QA_AnalyzePDF (modQA).
'             PDF_BuildTempPath, PDF_ExportToTemp, PDF_ValidateHeader,
'             PDF_SafeReplace, PDF_CommitReplace, PDF_RollbackReplace
'             (modPDFFileIO).
'             PDF_Elapsed, PDF_LogTiming, PRJ_UpdatePhase
'             (modPDFProgress).
'             CACHE_InvalidatePDF (modCache).
'             PRJ_GetIDWSheetCount, PRJ_GetPDFPageCount (modInspect).
'             PRJ_EnsureSandbox (modSandbox).
'             PRJ_FSO (modCOM).
'             PRJ_LogWarn, PRJ_LogError, PRJ_LogDebug (modLog).
'             L, L2, L3, LN (modLang).
'
' Emits     : Status strings in one of the canonical forms:
'               STAT_CREATED       & " · N/N pages"
'               STAT_OVERWRITTEN   & " · N/N pages"
'               "Skipped (Exists) · N/N pages"
'               "Skipped (Exists)"
'               STAT_SKIPPED_EMPTY                       (no views)
'               STAT_ERROR & ": <reason>"                (failure)
'             Callers must parse via STS_Classify (modStatus).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and multi-line expressions are
'             excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public Function PRJ_CreatePDF(sIDW As String, bForce As Boolean) As String
    Dim sPDF       As String
    Dim nIDWSheets As Long
    Dim bExisted   As Boolean
    Dim sEarly     As String

    sEarly = PDF_EarlyExitReason(sIDW, sPDF, nIDWSheets, bForce, bExisted)
    If Len(sEarly) > 0 Then
        PRJ_CreatePDF = sEarly
        Exit Function
    End If

    PRJ_CreatePDF = PDF_RunPipeline(sIDW, sPDF, nIDWSheets, bExisted)
End Function

Private Function PDF_RunPipeline(ByVal sIDW As String, _
                                  ByVal sPDF As String, _
                                  ByVal nIDWSheets As Long, _
                                  ByVal bExisted As Boolean) As String
    Dim oDoc          As DrawingDocument
    Dim bWeOpened     As Boolean
    Dim sTemp         As String
    Dim sErr          As String
    Dim sQA_IDW       As String
    Dim sQA_PDF       As String
    Dim bPageMismatch As Boolean
    Dim dT1           As Double
    Dim tT            As TPdfTiming

    On Error GoTo ErrH
    tT.dTotal = TM_Now()
    PRJ_EnsureSandbox

    PRJ_UpdatePhase L("report.msg.opening")
    dT1 = TM_Now()
    Set oDoc = PDF_OpenIDW(sIDW, bWeOpened)
    tT.dOpen = PDF_Elapsed(dT1, TM_Now())
    If oDoc Is Nothing Then
        PDF_RunPipeline = STAT_ERROR & ": " & L("pdf.err.open")
        Exit Function
    End If

    PRJ_UpdatePhase L("report.msg.updating")
    dT1 = TM_Now()
    sErr = PDF_PrepareDocument(oDoc, nIDWSheets)
    tT.dUpdate = PDF_Elapsed(dT1, TM_Now())
    If Len(sErr) > 0 Then GoTo Cleanup

    sQA_IDW = QA_AnalyzeIDW(oDoc)
    If Len(sQA_IDW) > 0 Then
        PRJ_LogWarn "PDF_RunPipeline.QA_IDW", "[" & sIDW & "] " & sQA_IDW
    End If

    sTemp = PDF_BuildTempPath(sPDF)

    PRJ_UpdatePhase L("report.msg.exporting")
    dT1 = TM_Now()
    sErr = PDF_ExportToTemp(oDoc, sTemp)
    tT.dExport = PDF_Elapsed(dT1, TM_Now())
    If Len(sErr) > 0 Then GoTo Cleanup

    dT1 = TM_Now()
    If Not PDF_ValidateHeader(sTemp) Then
        If Not PRJ_FSO().FileExists(sTemp) Then
            sErr = LN("msg.pdf.temp_missing", sTemp)
        Else
            sErr = L("pdf.err.header")
        End If
        tT.dValidate = PDF_Elapsed(dT1, TM_Now())
        GoTo Cleanup
    End If
    tT.dValidate = PDF_Elapsed(dT1, TM_Now())

    If bWeOpened Then oDoc.Close False
    Set oDoc = Nothing
    bWeOpened = False

    dT1 = TM_Now()
    sErr = PDF_FinalizeExport(sTemp, sPDF, nIDWSheets, bPageMismatch)
    tT.dReplace = PDF_Elapsed(dT1, TM_Now())

    If bPageMismatch Then
        PDF_RunPipeline = sErr
        PDF_LogTiming sIDW, tT
        Exit Function
    End If
    If Len(sErr) > 0 Then GoTo Cleanup

    sQA_PDF = QA_AnalyzePDF(sPDF)
    If Len(sQA_PDF) > 0 Then
        PRJ_LogWarn "PDF_RunPipeline.QA_PDF", "[" & sPDF & "] " & sQA_PDF
    End If

    PDF_RunPipeline = PDF_BuildSuccessStatus(nIDWSheets, bExisted)
    PDF_LogTiming sIDW, tT
    Exit Function

Cleanup:
    PDF_CleanupAfterExport oDoc, bWeOpened, sTemp
    PDF_RunPipeline = PDF_FinalStatus(sErr)
    Exit Function

ErrH:
    sErr = Err.Description
    If Len(sErr) = 0 Then sErr = L("pdf.err.unknown") & " (" & Err.Number & ")"
    PRJ_LogError "PDF_RunPipeline", "[" & sIDW & "] " & sErr
    Resume Cleanup
End Function

Private Function PDF_EarlyExitReason(ByVal sIDW As String, _
                                      ByRef sPDF As String, _
                                      ByRef nIDWSheets As Long, _
                                      ByVal bForce As Boolean, _
                                      ByRef bExisted As Boolean) As String
    Dim nDot As Long
    PDF_EarlyExitReason = ""

    nDot = InStrRev(sIDW, ".")
    If nDot = 0 Then
        PDF_EarlyExitReason = STAT_ERROR & ": " & L2("pdf.err.noext", sIDW)
        Exit Function
    End If

    sPDF = Left$(sIDW, nDot) & "pdf"
    bExisted = PRJ_FSO().FileExists(sPDF)
    nIDWSheets = PRJ_GetIDWSheetCount(sIDW)

    If bExisted And Not bForce Then
        Dim sSkip As String
        sSkip = PDF_ShouldSkip(sIDW, sPDF, nIDWSheets)
        If Len(sSkip) > 0 Then PDF_EarlyExitReason = sSkip
    End If
End Function

Private Function PDF_ShouldSkip(ByVal sIDW As String, _
                                 ByVal sPDF As String, _
                                 ByVal nIDWSheets As Long) As String
    PDF_ShouldSkip = ""

    Dim nExPages  As Long
    Dim bIdwNewer As Boolean
    Dim bMismatch As Boolean

    nExPages = PRJ_GetPDFPageCount(sPDF)
    bIdwNewer = (PRJ_GetFileDate(sIDW) > PRJ_GetFileDate(sPDF))
    bMismatch = (nExPages > 0 And nIDWSheets > 0 And nExPages <> nIDWSheets)

    If Not bIdwNewer And Not bMismatch Then
        If nExPages > 0 And nIDWSheets > 0 Then
            PDF_ShouldSkip = L3("pdf.skipped.exists.pages", CStr(nExPages), CStr(nIDWSheets))
        Else
            PDF_ShouldSkip = L("pdf.skipped.exists")
        End If
    End If
End Function

Private Function PDF_OpenIDW(ByVal sIDW As String, _
                              ByRef bWeOpened As Boolean) As DrawingDocument
    bWeOpened = False

    On Error Resume Next
    Err.Clear
    Set PDF_OpenIDW = ThisApplication.Documents.ItemByName(sIDW)
    If Err.Number <> 0 Then
        Err.Clear
        Set PDF_OpenIDW = Nothing
    End If
    On Error GoTo 0

    If Not PDF_OpenIDW Is Nothing Then Exit Function

    On Error GoTo EH
    Dim oOpenOpts As NameValueMap
    Set oOpenOpts = ThisApplication.TransientObjects.CreateNameValueMap
    oOpenOpts.Add "SkipAllUnresolvedFiles", True

    On Error Resume Next
    Err.Clear
    oOpenOpts.Add "DeferUpdates", True
    If Err.Number <> 0 Then
        PRJ_LogDebug "PDF_OpenIDW", "DeferUpdates not supported: " & Err.Description
        Err.Clear
    End If
    On Error GoTo EH

    Set PDF_OpenIDW = ThisApplication.Documents.OpenWithOptions(sIDW, oOpenOpts, False)
    bWeOpened = True
    Exit Function

EH:
    bWeOpened = False
    Set PDF_OpenIDW = Nothing
End Function

Private Function PDF_PrepareDocument(ByVal oDoc As DrawingDocument, _
                                      ByRef nIDWSheets As Long) As String
    PDF_PrepareDocument = ""
    On Error GoTo EH

    Dim bDirty As Boolean
    bDirty = False
    On Error Resume Next
    Err.Clear
    bDirty = oDoc.Dirty
    If Err.Number <> 0 Then
        bDirty = False
        Err.Clear
    End If
    On Error GoTo EH

    If bDirty Then
        PRJ_LogWarn "PDF_PrepareDocument", _
            "Document has unsaved changes - not saved: " & oDoc.FullFileName
    End If

    oDoc.Update

    On Error Resume Next
    Err.Clear
    If nIDWSheets <= 0 Then nIDWSheets = oDoc.Sheets.Count
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo EH

    If PDF_CountViews(oDoc) = 0 Then
        PDF_PrepareDocument = STAT_SKIPPED_EMPTY
    End If
    Exit Function

EH:
    PRJ_LogError "PDF_PrepareDocument", Err.Description
    PDF_PrepareDocument = L2("pdf.err.prepare", Err.Description)
End Function

Private Function PDF_CountViews(oDoc As DrawingDocument) As Long
    PDF_CountViews = 0
    On Error Resume Next
    Dim oSheet As Object
    For Each oSheet In oDoc.Sheets
        PDF_CountViews = PDF_CountViews + oSheet.DrawingViews.Count
    Next
    On Error GoTo 0
End Function

Private Function PDF_FinalizeExport(ByVal sTemp As String, _
                                     ByVal sPDF As String, _
                                     ByVal nIDWSheets As Long, _
                                     ByRef bPageMismatch As Boolean) As String
    Dim sErr As String
    Dim nPDFPages As Long

    bPageMismatch = False
    PDF_FinalizeExport = ""

    sErr = PDF_SafeReplace(sTemp, sPDF)
    If Len(sErr) > 0 Then
        PDF_FinalizeExport = sErr
        Exit Function
    End If

    CACHE_InvalidatePDF sPDF

    nPDFPages = PRJ_GetPDFPageCount(sPDF)

    If nPDFPages > 0 And nIDWSheets > 0 And nPDFPages <> nIDWSheets Then
        PDF_RollbackReplace sPDF
        CACHE_InvalidatePDF sPDF
        bPageMismatch = True
        PDF_FinalizeExport = STAT_ERROR & ": " & _
                              L3("pdf.err.pagematch", CStr(nPDFPages), CStr(nIDWSheets))
        Exit Function
    End If

    PDF_CommitReplace sPDF
End Function

Private Sub PDF_CleanupAfterExport(ByRef oDoc As DrawingDocument, _
                                    ByVal bWeOpened As Boolean, _
                                    ByVal sTemp As String)
    On Error Resume Next
    If bWeOpened And Not oDoc Is Nothing Then oDoc.Close False
    Set oDoc = Nothing
    If Len(sTemp) > 0 Then
        If PRJ_FSO().FileExists(sTemp) Then PRJ_FSO().DeleteFile sTemp, True
    End If
    On Error GoTo 0
End Sub

Private Function PDF_FinalStatus(ByVal sErr As String) As String
    If sErr = STAT_SKIPPED_EMPTY Then
        PDF_FinalStatus = sErr
    Else
        PDF_FinalStatus = STAT_ERROR & ": " & sErr
    End If
End Function

Private Function PDF_BuildSuccessStatus(ByVal nIDWSheets As Long, _
                                         ByVal bExisted As Boolean) As String
    Dim sep As String
    sep = " " & Chr(183) & " "
    If bExisted Then
        PDF_BuildSuccessStatus = STAT_OVERWRITTEN & sep & nIDWSheets & "/" & nIDWSheets & " pages"
    Else
        PDF_BuildSuccessStatus = STAT_CREATED & sep & nIDWSheets & "/" & nIDWSheets & " pages"
    End If
End Function

' END: modPDFExport v9.19.9

