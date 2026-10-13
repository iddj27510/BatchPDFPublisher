Attribute VB_Name = "modPDFProgress"
Option Explicit
' Module    : modPDFProgress
' Purpose   : Progress dialog updater and per-file timing instrumentation
' Reference : Appendix A.13, A.31, A.102
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + header completion):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Added missing Requires and Notes sections for consistency with
'   the rest of the project (this module previously had neither).
'   No functional changes.
'
' v9.19.3 (Full path in lblFile):
'   PRJ_UpdateProgress now writes the full file path into both
'   lblFile.Caption and lblFile.ControlTipText. Callers (modBatch)
'   already pass the full path since v9.19.3. The tooltip is
'   redundant when the caption fits, but keeps the path readable
'   when it is truncated by the label width.
'
' v9.19.0 (Progress color by %):
'   lblCount, lblPct and lblFill are colored by completion percent.
' v9.18.0: Err.Clear in PROG_CacheControls.
'
' Requires  : g_oProgress, g_nTotal, g_nDone, g_sCurrentRelPath
'             (modState).
'             CLR_* palette variables (modState / modPalette).
'             TM_Now, TM_Elapsed (modTime).
'             STS_Classify, TStatus, ST_* (modStatus).
'             L, LANG_ActiveCode (modLang).
'             PRJ_LogWarn, PRJ_LogError, PRJ_LogDebug (modLog).
'
' Emits     : TPdfTiming is consumed by modPDFExport.
'
' Notes     : No line-continuation characters in this module.
'             No decorative separator comment lines.

Public Type TPdfTiming
    dOpen     As Double
    dUpdate   As Double
    dExport   As Double
    dValidate As Double
    dReplace  As Double
    dTotal    As Double
End Type

Private m_dLastDoEvents As Double
Private m_oLblFile  As Object
Private m_oLblStat  As Object
Private m_oLblCount As Object
Private m_oLblPct   As Object
Private m_oLblBar   As Object
Private m_oLblFill  As Object
Private Const PDF_DOEVENTS_SEC  As Double = 0.05
Private Const PDF_SLOW_FILE_SEC As Double = 10#

Private m_sProgressProcessing As String
Private m_sProgressLang       As String

Public Sub PROG_CacheControls()
    Set m_oLblFile = Nothing
    Set m_oLblStat = Nothing
    Set m_oLblCount = Nothing
    Set m_oLblPct = Nothing
    Set m_oLblBar = Nothing
    Set m_oLblFill = Nothing
    If g_oProgress Is Nothing Then Exit Sub
    On Error Resume Next
    Err.Clear
    Set m_oLblFile = g_oProgress.Controls("lblFile")
    Set m_oLblStat = g_oProgress.Controls("lblStat")
    Set m_oLblCount = g_oProgress.Controls("lblCount")
    Set m_oLblPct = g_oProgress.Controls("lblPct")
    Set m_oLblBar = g_oProgress.Controls("lblBar")
    Set m_oLblFill = g_oProgress.Controls("lblFill")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Public Sub PROG_ReleaseControls()
    Set m_oLblFile = Nothing
    Set m_oLblStat = Nothing
    Set m_oLblCount = Nothing
    Set m_oLblPct = Nothing
    Set m_oLblBar = Nothing
    Set m_oLblFill = Nothing
End Sub

Private Function PROG_IsCached() As Boolean
    PROG_IsCached = False
    If m_oLblFile Is Nothing Then Exit Function
    If m_oLblStat Is Nothing Then Exit Function
    If m_oLblCount Is Nothing Then Exit Function
    If m_oLblPct Is Nothing Then Exit Function
    If m_oLblBar Is Nothing Then Exit Function
    If m_oLblFill Is Nothing Then Exit Function
    PROG_IsCached = True
End Function

Private Sub PDF_ThrottledDoEvents()
    Dim dNow As Double
    dNow = TM_Now()
    If dNow < m_dLastDoEvents Then
        m_dLastDoEvents = 0#
    End If
    If dNow - m_dLastDoEvents >= PDF_DOEVENTS_SEC Then
        m_dLastDoEvents = dNow
        DoEvents
    End If
End Sub

Public Function PDF_Elapsed(ByVal dStart As Double, ByVal dEnd As Double) As Double
    PDF_Elapsed = TM_Elapsed(dStart, dEnd)
End Function

Public Sub PDF_LogTiming(ByVal sIDW As String, ByRef tTiming As TPdfTiming)
    Dim dTotal As Double
    Dim sMsg As String

    dTotal = PDF_Elapsed(tTiming.dTotal, TM_Now())
    sMsg = "open=" & Format(tTiming.dOpen, "0.000") & "s"
    sMsg = sMsg & " update=" & Format(tTiming.dUpdate, "0.000") & "s"
    sMsg = sMsg & " export=" & Format(tTiming.dExport, "0.000") & "s"
    sMsg = sMsg & " validate=" & Format(tTiming.dValidate, "0.000") & "s"
    sMsg = sMsg & " replace=" & Format(tTiming.dReplace, "0.000") & "s"
    sMsg = sMsg & " total=" & Format(dTotal, "0.000") & "s"
    sMsg = sMsg & " [" & sIDW & "]"
    If dTotal >= PDF_SLOW_FILE_SEC Then
        PRJ_LogWarn "PRJ_CreatePDF.timing", sMsg
    Else
        PRJ_LogDebug "PRJ_CreatePDF.timing", sMsg
    End If
End Sub

Private Function PROG_PickPercentColor(ByVal p As Double) As Long
    If p >= 1# Then
        PROG_PickPercentColor = CLR_ACCENT
    ElseIf p >= 0.75 Then
        PROG_PickPercentColor = CLR_OK
    ElseIf p >= 0.5 Then
        PROG_PickPercentColor = CLR_WARN
    ElseIf p >= 0.25 Then
        PROG_PickPercentColor = CLR_TEMP_WARN
    Else
        PROG_PickPercentColor = CLR_ERROR
    End If
End Function

Public Sub PRJ_UpdateProgress(sFile As String, sStat As String)
    On Error GoTo EH
    If g_oProgress Is Nothing Then Exit Sub
    Dim oFile  As Object
    Dim oStat  As Object
    Dim oCount As Object
    Dim oPct   As Object
    Dim oBar   As Object
    Dim oFill  As Object
    If PROG_IsCached() Then
        Set oFile = m_oLblFile
        Set oStat = m_oLblStat
        Set oCount = m_oLblCount
        Set oPct = m_oLblPct
        Set oBar = m_oLblBar
        Set oFill = m_oLblFill
    Else
        PROG_CacheControls
        If Not PROG_IsCached() Then Exit Sub
        Set oFile = m_oLblFile
        Set oStat = m_oLblStat
        Set oCount = m_oLblCount
        Set oPct = m_oLblPct
        Set oBar = m_oLblBar
        Set oFill = m_oLblFill
    End If

    ' v9.19.3: full path in caption + tooltip.
    PROG_SetLabel oFile, sFile
    If Not oFile Is Nothing Then
        If oFile.ControlTipText <> sFile Then oFile.ControlTipText = sFile
    End If

    PROG_SetLabel oStat, sStat
    PROG_SetForeColor oStat, PROG_PickStatusColor(sStat)

    Dim p As Double
    p = IIf(g_nTotal > 0, g_nDone / g_nTotal, 0)

    Dim lPctClr As Long
    lPctClr = PROG_PickPercentColor(p)

    PROG_SetLabel oCount, g_nDone & " / " & g_nTotal
    PROG_SetForeColor oCount, lPctClr

    PROG_SetLabel oPct, Format(p, "0%")
    PROG_SetForeColor oPct, lPctClr
    On Error Resume Next
    If oPct Is Nothing Then
    ElseIf p >= 1# Then
        If Not oPct.Font.Bold Then oPct.Font.Bold = True
    Else
        If oPct.Font.Bold Then oPct.Font.Bold = False
    End If
    On Error GoTo EH

    PROG_SetBarWidth oBar, oFill, p
    PROG_SetForeColorFill oFill, lPctClr

    PDF_ThrottledDoEvents
    Exit Sub
EH:
    PRJ_LogError "PRJ_UpdateProgress", Err.Description
End Sub

Public Sub PRJ_UpdatePhase(ByVal sPhase As String)
    On Error Resume Next
    If Len(g_sCurrentRelPath) = 0 Then Exit Sub
    PRJ_UpdateProgress g_sCurrentRelPath, sPhase
    On Error GoTo 0
End Sub

Private Function PROG_GetProgressProcessingText() As String
    Dim sCurLang As String
    sCurLang = LANG_ActiveCode()
    If m_sProgressLang <> sCurLang Then
        m_sProgressProcessing = L("progress.processing")
        m_sProgressLang = sCurLang
    End If
    PROG_GetProgressProcessingText = m_sProgressProcessing
End Function

Private Function PROG_PickStatusColor(ByVal sStat As String) As Long
    Select Case STS_Classify(sStat)
        Case ST_ERROR:          PROG_PickStatusColor = CLR_ERROR
        Case ST_NO_DRAWING:     PROG_PickStatusColor = CLR_ERROR
        Case ST_MISSING:        PROG_PickStatusColor = CLR_ERROR
        Case ST_MISMATCH:       PROG_PickStatusColor = CLR_ERROR
        Case ST_STALE:          PROG_PickStatusColor = CLR_WARN
        Case ST_HAS_DRAWING:    PROG_PickStatusColor = CLR_OK
        Case ST_OK:             PROG_PickStatusColor = CLR_OK
        Case ST_SKIPPED:        PROG_PickStatusColor = CLR_WARN
        Case ST_SKIPPED_EMPTY:  PROG_PickStatusColor = CLR_WARN
        Case ST_UNKNOWN:
            If InStr(sStat, PROG_GetProgressProcessingText()) > 0 Then
                PROG_PickStatusColor = CLR_PROCESS
            ElseIf InStr(sStat, "Processing") > 0 Then
                PROG_PickStatusColor = CLR_PROCESS
            ElseIf InStr(sStat, L("report.msg.opening")) > 0 Then
                PROG_PickStatusColor = CLR_PROCESS
            ElseIf InStr(sStat, L("report.msg.updating")) > 0 Then
                PROG_PickStatusColor = CLR_PROCESS
            ElseIf InStr(sStat, L("report.msg.exporting")) > 0 Then
                PROG_PickStatusColor = CLR_PROCESS
            Else
                PROG_PickStatusColor = CLR_OK
            End If
        Case Else:              PROG_PickStatusColor = CLR_OK
    End Select
End Function

Private Sub PROG_SetLabel(ByVal oLbl As Object, ByVal sText As String)
    If oLbl Is Nothing Then Exit Sub
    If oLbl.Caption <> sText Then oLbl.Caption = sText
End Sub

Private Sub PROG_SetForeColor(ByVal oLbl As Object, ByVal lColor As Long)
    If oLbl Is Nothing Then Exit Sub
    If oLbl.ForeColor <> lColor Then oLbl.ForeColor = lColor
End Sub

Private Sub PROG_SetForeColorFill(ByVal oFill As Object, ByVal lColor As Long)
    If oFill Is Nothing Then Exit Sub
    If oFill.BackColor <> lColor Then oFill.BackColor = lColor
End Sub

Private Sub PROG_SetBarWidth(ByVal oBar As Object, ByVal oFill As Object, ByVal p As Double)
    If oBar Is Nothing Then Exit Sub
    If oFill Is Nothing Then Exit Sub
    Dim nNewWidth As Single
    nNewWidth = oBar.Width * p
    If oFill.Width <> nNewWidth Then oFill.Width = nNewWidth
End Sub

' END: modPDFProgress v9.19.9
