Attribute VB_Name = "modReportFilter"
Option Explicit
' Module    : modReportFilter
' Purpose   : Report classification, filtering, counting, formatting,
'             retry pass, row path resolution, summary composition,
'             and status visual mapping (color + bold).
' Reference : Appendix A.13, A.31
' Version   : 9.19.11 (Inventor 2027)
'
' v9.19.11 (summary timing prefix):
'   RP_BuildSummaryText now prefixes the elapsed-time segment with
'   the localized "Time" label ("Time: 9.93 s") instead of the
'   bare "9.93 s". Every segment in the summary line therefore
'   has a label; the orphan duration is gone. Uses the existing
'   report.time key, no new lang keys.
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Fixed stale inline comment: replaced '?' (encoding artifact
'   for an arrow) with '->' in the RP_GetSelectedPath column
'   layout note.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in RP_BuildSummaryText signature and in the
'   multi-argument LN() call inside it). No functional changes.
'
' v9.19.6 (RP_GetSelectedPath fix):
'   RP_GetSelectedPath now reads column 4 (hidden full path)
'   instead of column 1 (PDF size). The prior revision silently
'   returned "1.2 MB" for every row, which made lstReport_DblClick
'   pass an unopenable path to PRJ_OpenPathOrSelect and fail
'   silently. This blocked double-click-to-open-PDF on every row.
'   Fallback to column 1 (legacy position) retained for
'   compatibility if lstReport's column layout ever changes.
'
' v9.19.3 (Dead code removal):
'   Removed RP_BuildSelDetailLocal. It was used by the old lblSel
'   details panel, which was removed in v9.19.0 when row details
'   moved entirely into lstReport columns. No callers remained.
'
' v9.19.2 (Summary keys):
'   RP_BuildSummaryText consumes report.summary.full and
'   report.skip.empty.
'
' v9.18.0 (§3.10 Dead code):
'   Removed RP_FormatSummary.
' v9.18.0 (§1.11 UI/Logic separation):
'   Added RP_PickSummaryColor, RP_BuildSummaryText, RP_PickStatusBold.
' v9.18.0 (§1.3 UI/Logic separation):
'   RP_RetryPass updates g_colLog. RP_GetSelectedPath added.
'
' Requires  : g_colFiles, g_colLog, g_sActiveFilter (modState).
'             STS_Classify, TStatus, ST_* (modStatus).
'             PRJ_CreatePDF (modPDFExport).
'             PRJ_FSO (modCOM).
'             PRJ_GetFileSize, PRJ_GetFileDate, PRJ_FormatFileSize,
'             PRJ_FormatDate, PRJ_FormatElapsed (modFileInfo).
'             L, L2, L3, LN (modLang).
'             CLR_ERROR, CLR_WARN, CLR_OK, CLR_TEXT_SEC (modState).
'             PRJ_LogError (modLog).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and multi-argument call
'             continuation are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public Type TReportCounts
    nOK    As Long
    nOver  As Long
    nSkip  As Long
    nEmpty As Long
    nErr   As Long
End Type

Public Function RP_StatusTag(ByVal sRes As String) As String
    Select Case STS_Classify(sRes)
        Case ST_ERROR:          RP_StatusTag = "[Error]"
        Case ST_SKIPPED:        RP_StatusTag = "[Skipped]"
        Case ST_SKIPPED_EMPTY:  RP_StatusTag = "[Skipped]"
        Case ST_OVERWRITTEN:    RP_StatusTag = "[Overwritten]"
        Case ST_HAS_DRAWING:    RP_StatusTag = "[Has Drawing]"
        Case ST_NO_DRAWING:     RP_StatusTag = "[No Drawing]"
        Case ST_CREATED:        RP_StatusTag = "[Created]"
        Case Else:              RP_StatusTag = "[Created]"
    End Select
End Function

Public Function RP_MatchesFilter(ByVal sRes As String) As Boolean
    Dim t As TStatus
    Dim bIsErr   As Boolean
    Dim bIsSkip  As Boolean
    Dim bIsAudit As Boolean
    Dim bIsNoDWG As Boolean

    t = STS_Classify(sRes)
    bIsErr = (t = ST_ERROR)
    bIsSkip = (t = ST_SKIPPED Or t = ST_SKIPPED_EMPTY)
    bIsNoDWG = (t = ST_NO_DRAWING)
    bIsAudit = (t = ST_HAS_DRAWING Or t = ST_NO_DRAWING)

    Select Case g_sActiveFilter
        Case "All":            RP_MatchesFilter = True
        Case "Err":            RP_MatchesFilter = bIsErr
        Case "Skip":           RP_MatchesFilter = bIsSkip
        Case "Ok":             RP_MatchesFilter = Not bIsErr And Not bIsSkip And Not bIsAudit
        Case "NoIdw", "NoDWG": RP_MatchesFilter = bIsNoDWG
        Case Else:             RP_MatchesFilter = True
    End Select
End Function

Public Function RP_StatusColor(ByVal sRes As String) As Long
    Select Case STS_Classify(sRes)
        Case ST_ERROR:          RP_StatusColor = CLR_ERROR
        Case ST_SKIPPED:        RP_StatusColor = CLR_WARN
        Case ST_SKIPPED_EMPTY:  RP_StatusColor = CLR_WARN
        Case ST_HAS_DRAWING:    RP_StatusColor = CLR_TEXT_SEC
        Case ST_NO_DRAWING:     RP_StatusColor = CLR_TEXT_SEC
        Case Else:              RP_StatusColor = CLR_OK
    End Select
End Function

Public Function RP_PickStatusBold(ByVal sRes As String) As Boolean
    RP_PickStatusBold = (STS_Classify(sRes) = ST_ERROR)
End Function

Public Sub RP_CountFiles(ByRef t As TReportCounts)
    Dim i As Long
    Dim a As Variant
    Dim s As TStatus

    If g_colFiles Is Nothing Then Exit Sub

    For i = 0 To g_colFiles.Count - 1
        a = g_colFiles(i)
        s = STS_Classify(CStr(a(1)))
        Select Case s
            Case ST_CREATED:       t.nOK = t.nOK + 1
            Case ST_OVERWRITTEN:   t.nOver = t.nOver + 1
            Case ST_SKIPPED:       t.nSkip = t.nSkip + 1
            Case ST_SKIPPED_EMPTY: t.nSkip = t.nSkip + 1
                                   t.nEmpty = t.nEmpty + 1
            Case ST_ERROR:         t.nErr = t.nErr + 1
            Case ST_HAS_DRAWING:
            Case ST_NO_DRAWING:
            Case Else:             t.nErr = t.nErr + 1
        End Select
    Next i
End Sub

Public Function RP_PickSummaryColor(ByRef t As TReportCounts) As Long
    If t.nErr > 0 Then
        RP_PickSummaryColor = CLR_ERROR
    ElseIf t.nSkip > 0 Then
        RP_PickSummaryColor = CLR_WARN
    ElseIf t.nOK + t.nOver > 0 Then
        RP_PickSummaryColor = CLR_OK
    Else
        RP_PickSummaryColor = CLR_TEXT_SEC
    End If
End Function

Public Function RP_BuildSummaryText(ByRef t As TReportCounts, _
                                     ByVal dElapsedSec As Double) As String
    Dim sSkip As String
    If t.nEmpty > 0 Then
        sSkip = L3("report.skip.empty", CStr(t.nSkip), CStr(t.nEmpty))
    Else
        sSkip = CStr(t.nSkip)
    End If

    ' v9.19.11: prefix the timing with the localized "Time" label
    ' so the summary reads consistently with the other segments
    ' (every segment has a label, no orphan duration).
    Dim sTime As String
    sTime = L("report.time") & ": " & PRJ_FormatElapsed(dElapsedSec)

    RP_BuildSummaryText = LN("report.summary.full", _
        CStr(t.nOK), CStr(t.nOver), sSkip, CStr(t.nErr), sTime)
End Function

Public Function RP_FindStat(ByVal sPath As String) As String
    Dim i As Long
    Dim a As Variant
    RP_FindStat = ""
    If g_colFiles Is Nothing Then Exit Function
    For i = 0 To g_colFiles.Count - 1
        a = g_colFiles(i)
        If CStr(a(0)) = sPath Then
            RP_FindStat = CStr(a(1))
            Exit Function
        End If
    Next i
End Function

Public Function RP_CountErrorRows() As Long
    Dim i    As Long
    Dim aChk As Variant
    RP_CountErrorRows = 0
    If g_colFiles Is Nothing Then Exit Function
    For i = 0 To g_colFiles.Count - 1
        aChk = g_colFiles(i)
        If STS_Classify(CStr(aChk(1))) = ST_ERROR Then
            RP_CountErrorRows = RP_CountErrorRows + 1
        End If
    Next i
End Function

Public Sub RP_RetryPass(ByRef nOK As Long, ByRef nStill As Long)
    Dim i     As Long
    Dim a     As Variant
    Dim sStat As String
    If g_colFiles Is Nothing Then Exit Sub
    For i = 0 To g_colFiles.Count - 1
        a = g_colFiles(i)
        If STS_Classify(CStr(a(1))) = ST_ERROR Then
            sStat = PRJ_CreatePDF(CStr(a(0)), True)
            g_colFiles(i) = Array(a(0), sStat)
            If Not g_colLog Is Nothing Then
                If g_colLog.Exists(CStr(a(0))) Then
                    g_colLog(CStr(a(0))) = sStat
                Else
                    g_colLog.Add CStr(a(0)), sStat
                End If
            End If
            If STS_Classify(sStat) = ST_ERROR Then
                nStill = nStill + 1
            Else
                nOK = nOK + 1
            End If
            DoEvents
        End If
    Next i
End Sub

' RP_GetSelectedPath   resolve the full path of the currently
'   selected row in lstReport.
'
'   Column layout (from frmReport UserForm_Activate):
'     0 = status marker + status text ("[+] Created | 5/5 pages")
'     1 = PDF size ("1.2 MB")   -> was read by mistake pre-v9.19.4
'     2 = date ("15.06.2024 09:00")
'     3 = relative path
'     4 = full path (hidden, width 0)
'
'   v9.19.6: fixed to read column 4.
'   Prior revision read column 1 and passed the PDF size string
'   ("1.2 MB") to PRJ_OpenPathOrSelect. That produced an
'   unopenable path and the failure was swallowed because
'   SHELL_SelectInExplorer returns True when Shell dispatches
'   without error (even if the file does not exist).
'   Result: double-click on any row silently did nothing.
'
'   Fallback to column 1 remains for compatibility if lstReport's
'   column count ever drops below 5 in a future revision.
Public Function RP_GetSelectedPath(ByVal lb As Object) As String
    RP_GetSelectedPath = ""
    If lb Is Nothing Then Exit Function
    If lb.ListIndex < 0 Then Exit Function

    Dim sPath As String
    sPath = ""

    If lb.ColumnCount >= 5 Then
        sPath = CStr(lb.List(lb.ListIndex, 4))
    ElseIf lb.ColumnCount >= 2 Then
        sPath = CStr(lb.List(lb.ListIndex, 1))
    End If

    ' Defensive fallback: parses path from the visible row text.
    ' Never fires with the current lstReport (ColumnCount = 5),
    ' but kept in case ColumnCount changes in a future revision.
    If Len(sPath) = 0 Then
        Dim sLine As String
        sLine = CStr(lb.Value)
        If Len(sLine) > 0 Then
            Dim nPos As Long
            nPos = InStrRev(sLine, " · ")
            If nPos > 0 Then
                Dim sCand As String
                sCand = Trim$(Mid$(sLine, nPos + 3))
                If InStr(1, sCand, "\") > 0 Then sPath = sCand
            End If
        End If
    End If

    RP_GetSelectedPath = sPath
End Function

' END: modReportFilter v9.19.11

