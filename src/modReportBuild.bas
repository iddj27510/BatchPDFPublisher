Attribute VB_Name = "modReportBuild"
Option Explicit
' Module    : modReportBuild
' Purpose   : Report text construction (pure, no IO, no UI, no COM)
' Reference : Appendix A.13, A.31
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + header completion):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Added missing Notes and Requires sections for consistency with
'   the rest of the project (this module previously had neither).
'   No functional changes.
'
' v9.19.2 (§1.17 Resource Separation, §8.4 Date/time locale):
'   RP_BuildHeader no longer calls L("report.date.format") — uses
'   PRJ_FormatNow() (Windows Short Date + Short Time).
'   RP_BuildAuditBody rewritten to use the same unified summary
'   keys as modSummary.PRJ_GetAuditSummary (summary.hdr.models_full
'   / summary.hdr.with_idw / summary.hdr.no_idw). The previous
'   four keys (scan.audit.sum.parts/assemblies/with/without) were
'   removed from modLang in v9.19.1 as dead, but this module still
'   referenced them — audit reports showed ??key?? in every language.
'
' v9.16.2 (Faza 7.1b, M.23): g_sAbortReason in report header.
' v9.15.5 (Faza 6D.2, M.8c): TReportCounts; RP_BuildPdfBody uses Type.
' v9.15.3 (Faza 6C, H.5): STS_Classify.
' v9.14.6 (Faza 2): RP_BuildHeader localized.
'
' Requires  : g_colFiles, g_sLastFolder, g_bSubfolders,
'             g_bForceOverwrite, g_bAuditMode, g_sAbortReason,
'             g_dTimeStart, g_dTimeEnd, SECONDS_PER_DAY,
'             PRJ_VERSION (modState / modConstants).
'             RP_CountFiles, TReportCounts (modReportFilter).
'             STS_Classify, TStatus, ST_* (modStatus).
'             PRJ_GetPDFSize, PRJ_GetRelPath, PRJ_FormatElapsed,
'             PRJ_FormatNow (modFileInfo).
'             L, L2, LN (modLang).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and multi-argument call
'             continuation are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public Function PRJ_GetReportText() As String
    Dim s As String
    s = RP_BuildHeader()
    If Len(g_sAbortReason) > 0 Then
        s = s & g_sAbortReason & vbCrLf & vbCrLf
    End If
    s = s & RP_BuildPdfBody()
    If g_bAuditMode Then s = s & vbCrLf & RP_BuildAuditBody()
    PRJ_GetReportText = s
End Function

Private Function RP_BuildHeader() As String
    Dim s As String
    s = L2("report.hdr.title", PRJ_VERSION) & vbCrLf
    If g_bAuditMode Then
        s = s & L2("report.hdr.mode", "PDF + " & L("report.audit.title")) & vbCrLf
    Else
        s = s & L2("report.hdr.mode", "PDF") & vbCrLf
    End If
    s = s & L2("report.hdr.folder", g_sLastFolder) & vbCrLf
    s = s & L("report.hdr.subfolders") & ": " & RP_OnOff(g_bSubfolders)
    If Not g_bAuditMode Then
        s = s & "  " & L("report.hdr.force") & ": " & RP_OnOff(g_bForceOverwrite)
    End If
    s = s & vbCrLf
    ' v9.19.2: PRJ_FormatNow uses Windows Short Date + Short Time.
    ' Do not go back to a per-language template.
    s = s & L2("report.hdr.generated", PRJ_FormatNow()) & vbCrLf
    s = s & vbCrLf
    RP_BuildHeader = s
End Function

Private Function RP_OnOff(bState As Boolean) As String
    If bState Then
        RP_OnOff = L("report.state.on")
    Else
        RP_OnOff = L("report.state.off")
    End If
End Function

Private Function RP_BuildPdfBody() As String
    Dim t As TReportCounts
    RP_CountFiles t

    Dim dRaw As Double
    dRaw = g_dTimeEnd - g_dTimeStart
    If dRaw < 0 Then dRaw = dRaw + SECONDS_PER_DAY

    Dim sElapsed As String
    sElapsed = PRJ_FormatElapsed(dRaw)

    Dim s As String
    s = L("report.hdr.pdf.section") & vbCrLf
    s = s & L("report.created") & ": " & t.nOK & "  "
    s = s & L("report.overwritten") & ": " & t.nOver & "  "
    s = s & L("report.skipped") & ": " & t.nSkip
    If t.nEmpty > 0 Then
        s = s & " (" & t.nEmpty & " " & L("report.empty") & ")"
    End If
    s = s & "  " & L("report.errors") & ": " & t.nErr
    s = s & "  " & L("report.time") & ": " & sElapsed & vbCrLf
    s = s & vbCrLf
    s = s & L("report.hdr.pdf.detail") & vbCrLf
    s = s & RP_BuildPdfDetail()
    RP_BuildPdfBody = s
End Function

Private Function RP_BuildPdfDetail() As String
    RP_BuildPdfDetail = ""
    If g_colFiles Is Nothing Then Exit Function
    Dim nCount As Long
    nCount = g_colFiles.Count
    If nCount = 0 Then Exit Function
    Dim aLines() As String
    ReDim aLines(0 To nCount - 1)
    Dim nIdx As Long
    nIdx = 0
    Dim sep As String
    sep = " " & Chr(183) & " "
    Dim v As Variant, a As Variant
    Dim sRes As String, sSize As String
    Dim sLine As String, sRel As String
    Dim t As TStatus
    For Each v In g_colFiles
        a = v
        sRes = CStr(a(1))
        t = STS_Classify(sRes)
        If t <> ST_HAS_DRAWING And t <> ST_NO_DRAWING Then
            sSize = PRJ_GetPDFSize(CStr(a(0)))
            sRel = PRJ_GetRelPath(CStr(a(0)))
            sLine = PRJ_LocalizeStatus(sRes) & sep & sSize & sep & sRel
            aLines(nIdx) = sLine
            nIdx = nIdx + 1
        End If
    Next
    If nIdx = 0 Then Exit Function
    ReDim Preserve aLines(0 To nIdx - 1)
    RP_BuildPdfDetail = Join(aLines, vbCrLf) & vbCrLf
End Function

' RP_BuildAuditBody   unified with modSummary.PRJ_GetAuditSummary.
'   Uses the same three keys so the report file and the on-screen
'   lblSum2 show identical numbers in identical phrasing:
'     summary.hdr.models_full   "{0} models ({1} assemblies, {2} parts)"
'     summary.hdr.with_idw      "{0} with drawing" / "{0} covered"
'     summary.hdr.no_idw        "{0} without drawing" / "{0} missing"
'
' v9.19.2: was using scan.audit.sum.parts / .assemblies / .with /
'   .without — those four keys no longer exist in modLang.
Private Function RP_BuildAuditBody() As String
    Dim nHas As Long, nNo As Long
    Dim nParts As Long, nAssy As Long
    RP_CountAuditResults nHas, nNo, nParts, nAssy

    Dim nTotal As Long
    nTotal = nParts + nAssy

    Dim sep As String
    sep = " " & Chr(183) & " "

    Dim sSumLine As String
    sSumLine = LN("summary.hdr.models_full", _
                  CStr(nTotal), CStr(nAssy), CStr(nParts))
    sSumLine = sSumLine & sep & L2("summary.hdr.with_idw", CStr(nHas))
    sSumLine = sSumLine & sep & L2("summary.hdr.no_idw", CStr(nNo))

    Dim s As String
    s = L("report.hdr.audit.section") & vbCrLf
    s = s & sSumLine & vbCrLf
    s = s & vbCrLf
    s = s & L("report.hdr.audit.detail") & vbCrLf
    s = s & RP_BuildAuditDetail()
    RP_BuildAuditBody = s
End Function

Private Sub RP_CountAuditResults(ByRef nHas As Long, ByRef nNo As Long, _
                                  ByRef nParts As Long, ByRef nAssy As Long)
    If g_colFiles Is Nothing Then Exit Sub
    Dim v As Variant, a As Variant
    Dim sRes As String, sPath As String
    For Each v In g_colFiles
        a = v
        sRes = CStr(a(1))
        sPath = CStr(a(0))
        RP_ProcessAuditRow sRes, sPath, nHas, nNo, nParts, nAssy
    Next
End Sub

Private Sub RP_ProcessAuditRow(ByVal sRes As String, ByVal sPath As String, _
                                ByRef nHas As Long, ByRef nNo As Long, _
                                ByRef nParts As Long, ByRef nAssy As Long)
    Dim t As TStatus
    t = STS_Classify(sRes)
    If t = ST_HAS_DRAWING Then nHas = nHas + 1
    If t = ST_NO_DRAWING Then nNo = nNo + 1
    If t <> ST_HAS_DRAWING And t <> ST_NO_DRAWING Then Exit Sub
    Dim sExt As String
    sExt = LCase(Right(sPath, 4))
    If sExt = ".ipt" Then nParts = nParts + 1
    If sExt = ".iam" Then nAssy = nAssy + 1
End Sub

Private Function RP_BuildAuditDetail() As String
    RP_BuildAuditDetail = ""
    If g_colFiles Is Nothing Then Exit Function
    Dim nCount As Long
    nCount = g_colFiles.Count
    If nCount = 0 Then Exit Function
    Dim sep As String
    sep = " " & Chr(183) & " "
    Dim aLines() As String
    ReDim aLines(0 To nCount - 1)
    Dim nIdx As Long
    nIdx = 0
    Dim v As Variant, a As Variant, sRes As String
    For Each v In g_colFiles
        a = v
        sRes = CStr(a(1))
        If STS_Classify(sRes) = ST_NO_DRAWING Then
            aLines(nIdx) = PRJ_LocalizeStatus(sRes) & sep & PRJ_GetRelPath(CStr(a(0)))
            nIdx = nIdx + 1
        End If
    Next
    If nIdx = 0 Then Exit Function
    ReDim Preserve aLines(0 To nIdx - 1)
    RP_BuildAuditDetail = Join(aLines, vbCrLf) & vbCrLf
End Function

Public Function PRJ_LocalizeStatus(sRes As String) As String
    Dim sOut As String
    Select Case STS_Classify(sRes)
        Case ST_HAS_DRAWING
            sOut = L("report.tag.hasdrawing") & Mid(sRes, 11)
        Case ST_NO_DRAWING
            sOut = L("report.tag.nodrawing") & Mid(sRes, 10)
        Case ST_ERROR
            sOut = L("report.tag.error") & Mid(sRes, 6)
        Case ST_SKIPPED
            sOut = L("report.tag.skipped") & Mid(sRes, 8)
        Case ST_SKIPPED_EMPTY
            sOut = L("report.tag.skipped") & Mid(sRes, 8)
        Case ST_OVERWRITTEN
            sOut = L("report.tag.overwritten") & Mid(sRes, 12)
        Case ST_CREATED
            sOut = L("report.tag.created") & Mid(sRes, 8)
        Case Else
            sOut = sRes
    End Select
    sOut = Replace(sOut, " pages", " " & L("report.pages"))
    PRJ_LocalizeStatus = sOut
End Function

' END: modReportBuild v9.19.9

