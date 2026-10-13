Attribute VB_Name = "modUIHelpers"
Option Explicit
' Module    : modUIHelpers
' Purpose   : Tooltips, badges, responsive sleep, text pair alignment,
'             keyboard accelerators, caption localization.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.9 (badge redesign — action-oriented):
'   SC_BuildBadgeText rewritten to describe what the batch will
'   do, not what is wrong:
'     Old: 27 drawings · 3 missing
'     New: 27 drawings · 3 to create · 24 current
'   Force ON moves all current files into the "to recreate" count
'   and appends (force). Errors stay visible in both modes.
'   The "all current" shortcut fires only when Force is OFF and
'   nothing is pending.
'   PRJ_ApplyTooltips trimmed from 15 to 12 tooltips. Removed
'   three restating the visible caption (lblBrowse, lblCopy,
'   txtPath). Remaining 12 each add information beyond the
'   caption or explain a non-obvious consequence.
'
'   SC_BuildBadgeTooltip keeps the diagnostic vocabulary
'   (missing / stale / mismatch) and appends (force) when Force
'   is on, so the tooltip mirrors the visible badge state.
'
'   Signature change: SC_BuildBadgeText now takes the whole
'   TBadgeCounts by reference, not five separate Longs.
'   SC_ApplyDrawingBadge updated to the new signature.
'
'   3 new lang keys consumed: scan.sum.to_create,
'   scan.sum.to_recreate, scan.sum.force_suffix. Reuses
'   existing scan.sum.ok and scan.sum.all.current.
'
'   Scan options on frmScan are single labels with an Active/
'   Inactive visual state, not CheckBoxes. PRJ_LocalizeCaptions,
'   PRJ_SetupAccelerators and PRJ_ApplyTooltips updated from
'   chkSubfolders / chkForce / chkAudit to lblOptSub /
'   lblOptForce / lblOptAudit. SC_CenterAll unchanged — the
'   toggle labels are positioned by frmScan.LayoutToggleLabels,
'   not by PRJ_AlignTxt.
'
' v9.19.7 (Audit gate + tooltip audit):
'   SC_UpdateSum2 now gates on g_bAuditMode. Prior version ran
'   the audit summary unconditionally, so "N models (X assemblies,
'   Y parts) ..." appeared even with the Audit checkbox off.
'   PRJ_ApplyTooltips trimmed from 24 tooltips to 17.
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   UI_AlignTxtPair removed (dead code, no callers).
'   SC_UpdateSum2 checks modValidate.SC_GetCachedAudit first.
'
' v9.19.3:
'   PRJ_SetAccelerator logs silent failures via PRJ_LogDebug.
'   SC_UpdateTitle removed (no-op, no callers).
' v9.19.0: PRJ_LocalizeCaptions added frmScan checkboxes and lblTitle.
' v9.18.16: SC_UpdateSum2 ForeColor before Caption.
' v9.18.15: SC_SetLabelText ForeColor before Caption.
' v9.18.14: SC_ShowSum2 no-op, SC_HideSum2 no Visible toggle.
'
' Notes     : Type declarations come first (VBA requirement).
'             No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public Type TBadgeCounts
    nTot      As Long
    nReady    As Long
    nMissing  As Long
    nStale    As Long
    nMismatch As Long
    nErr      As Long
End Type

Public g_tLastBadge As TBadgeCounts
Public g_sLastBadgeText As String

Public Sub PRJ_SetTip(oFrm As Object, sBase As String, sTip As String)
    On Error Resume Next
    Err.Clear
    oFrm.Controls(sBase).ControlTipText = sTip
    If Err.Number <> 0 Then
        PRJ_LogDebug "PRJ_SetTip", "Control not found: " & sBase
        Err.Clear
    End If
    On Error GoTo 0

    On Error Resume Next
    Err.Clear
    oFrm.Controls(sBase & "Txt").ControlTipText = sTip
    If Err.Number <> 0 Then
        PRJ_LogDebug "PRJ_SetTip", "Control not found: " & sBase & "Txt"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' PRJ_ApplyTooltips   set ControlTipText on controls whose caption
'   alone does not convey enough information. Rule: a tooltip must
'   add information beyond the visible caption, OR explain a
'   non-obvious consequence (ESC shortcut, OldVersions skip, stale
'   condition, double-click behaviour). Tooltips that merely
'   restate the caption are noise and are not applied.
'
'   v9.19.9: trimmed from 15 to 12 tooltips. Commented out:
'     - lblBrowse  — "Browse..." is self-evident
'     - lblCopy    — "Copy" is self-evident on a report form
'     - txtPath    — a path textbox does not need a tooltip
Public Sub PRJ_ApplyTooltips(oFrm As Object)
    ' Buttons — action semantics
    ' PRJ_SetTip oFrm, "lblBrowse", L("tip.browse")   ' v9.19.9: redundant
    PRJ_SetTip oFrm, "lblRun", L("tip.run")
    PRJ_SetTip oFrm, "lblTheme", L("tip.theme")
    PRJ_SetTip oFrm, "lblReset", L("tip.reset")
    PRJ_SetTip oFrm, "lblCancel", L("tip.cancel")

    ' Report actions
    PRJ_SetTip oFrm, "lblRetry", L("tip.retry")
    ' PRJ_SetTip oFrm, "lblCopy", L("tip.copy")       ' v9.19.9: redundant
    PRJ_SetTip oFrm, "lblExp", L("tip.export")

    ' Filter — only NoIdw (cryptic label). Others are self-evident.
    PRJ_SetTip oFrm, "lblFltNoIdw", L("tip.filter.nodrawing")

    ' Interactive controls
    PRJ_SetTip oFrm, "lstReport", L("tip.list")
    ' PRJ_SetTip oFrm, "txtPath", L("tip.txtpath")    ' v9.19.9: redundant
    PRJ_SetTip oFrm, "txtSearch", L("tip.search")

    ' Scan option toggles — each explains a non-obvious consequence.
    PRJ_SetTip oFrm, "lblOptSub", L("tip.subfolders")
    PRJ_SetTip oFrm, "lblOptForce", L("tip.force")
    PRJ_SetTip oFrm, "lblOptAudit", L("tip.audit")
End Sub

Public Sub PRJ_SetupAccelerators(ByVal oFrm As Object)
    On Error Resume Next

    Dim sFrm As String
    sFrm = oFrm.Name

    Select Case sFrm
        Case "frmScan"
            PRJ_SetAccelerator oFrm, "lblRun", "R"
            PRJ_SetAccelerator oFrm, "lblBrowse", "B"
            PRJ_SetAccelerator oFrm, "lblTheme", "T"
            PRJ_SetAccelerator oFrm, "lblReset", "E"
            PRJ_SetAccelerator oFrm, "lblCancel", "C"
            PRJ_SetAccelerator oFrm, "lblOptSub", "S"
            PRJ_SetAccelerator oFrm, "lblOptForce", "F"
            PRJ_SetAccelerator oFrm, "lblOptAudit", "A"

        Case "frmReport"
            PRJ_SetAccelerator oFrm, "lblBack", "B"
            PRJ_SetAccelerator oFrm, "lblCopy", "C"
            PRJ_SetAccelerator oFrm, "lblRetry", "R"
            PRJ_SetAccelerator oFrm, "lblExp", "E"
            PRJ_SetAccelerator oFrm, "lblCancel", "Q"
            PRJ_SetAccelerator oFrm, "lblFltAll", "1"
            PRJ_SetAccelerator oFrm, "lblFltErr", "2"
            PRJ_SetAccelerator oFrm, "lblFltSkip", "3"
            PRJ_SetAccelerator oFrm, "lblFltOk", "4"
            PRJ_SetAccelerator oFrm, "lblFltNoIdw", "5"

        Case "frmProgress"
            PRJ_SetAccelerator oFrm, "lblCancel", "C"
    End Select

    On Error GoTo 0
End Sub

Private Sub PRJ_SetAccelerator(ByVal oFrm As Object, _
                                ByVal sControl As String, _
                                ByVal sKey As String)
    On Error Resume Next
    Err.Clear
    oFrm.Controls(sControl).Accelerator = sKey
    If Err.Number <> 0 Then
        PRJ_LogDebug "PRJ_SetAccelerator", _
            "Cannot set accelerator on " & sControl & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub PRJ_LocalizeCaptions(ByVal oFrm As Object)
    On Error Resume Next

    Dim sFrm As String
    sFrm = oFrm.Name

    Select Case sFrm
        Case "frmScan"
            SetCap oFrm, "lblTitle", "scan.title.no_folder"
            SetCap oFrm, "lblOptSub", "lbl.subfolders"
            SetCap oFrm, "lblOptForce", "chk.force"
            SetCap oFrm, "lblOptAudit", "chk.audit"
            SetCap oFrm, "lblCancelTxt", "btn.cancel"
            SetCap oFrm, "lblResetTxt", "btn.reset"
            SetCap oFrm, "lblRunTxt", "btn.run"

        Case "frmReport"
            SetCap oFrm, "lblTitle", "report.form.title"
            SetCap oFrm, "lblBackTxt", "btn.back"
            SetCap oFrm, "lblCopyTxt", "btn.copy"
            SetCap oFrm, "lblExpTxt", "btn.export"
            SetCap oFrm, "lblRetryTxt", "btn.retry"
            SetCap oFrm, "lblCancelTxt", "btn.cancel"
            SetCap oFrm, "lblFltAllTxt", "report.filter.all"
            SetCap oFrm, "lblFltErrTxt", "report.filter.err"
            SetCap oFrm, "lblFltSkipTxt", "report.filter.skip"
            SetCap oFrm, "lblFltOkTxt", "report.filter.ok"
            SetCap oFrm, "lblFltNoIdwTxt", "report.filter.noidw"
            SetCap oFrm, "lblSearch", "report.search.label"

        Case "frmProgress"
            SetCap oFrm, "lblCancelTxt", "btn.cancel"
    End Select

    On Error GoTo 0
End Sub

Private Sub SetCap(ByVal oFrm As Object, _
                    ByVal sControl As String, _
                    ByVal sKey As String)
    On Error Resume Next
    Err.Clear
    Dim oCtl As Object
    Set oCtl = oFrm.Controls(sControl)
    If Err.Number = 0 Then
        If Not oCtl Is Nothing Then
            oCtl.Caption = L(sKey)
        End If
    Else
        PRJ_LogDebug "SetCap", "Control not found: " & sControl
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub SC_SetLabelText(ByVal oLbl As Object, ByVal sText As String, ByVal lColor As Long)
    If oLbl Is Nothing Then Exit Sub
    If oLbl.ForeColor <> lColor Then oLbl.ForeColor = lColor
    If oLbl.Caption <> sText Then oLbl.Caption = sText
End Sub

Public Sub SC_ShowSum2(ByVal oFrm As Object)
    ' No-op — lblSum2 is always Visible.
End Sub

Public Sub SC_HideSum2(ByVal oFrm As Object)
    Dim oSum2 As Object
    On Error Resume Next
    Set oSum2 = oFrm.Controls("lblSum2")
    On Error GoTo 0
    If oSum2 Is Nothing Then Exit Sub
    If Len(oSum2.Caption) > 0 Then oSum2.Caption = ""
    oSum2.ControlTipText = ""
End Sub

Public Sub SC_ApplyDrawingBadge(ByVal oFrm As Object, ByRef t As TBadgeCounts)
    Dim oSum As Object
    Dim sText As String

    g_tLastBadge.nTot = t.nTot
    g_tLastBadge.nReady = t.nReady
    g_tLastBadge.nMissing = t.nMissing
    g_tLastBadge.nStale = t.nStale
    g_tLastBadge.nMismatch = t.nMismatch
    g_tLastBadge.nErr = t.nErr

    sText = SC_BuildBadgeText(t)
    g_sLastBadgeText = sText

    On Error Resume Next
    Set oSum = oFrm.Controls("lblSum")
    On Error GoTo EH

    If oSum Is Nothing Then Exit Sub

    SC_SetLabelText oSum, sText, _
        SC_PickBadgeColor(t.nErr, t.nMismatch, t.nMissing, t.nStale)

    oSum.ControlTipText = SC_BuildBadgeTooltip(t)
    Exit Sub

EH:
    PRJ_LogError "SC_ApplyDrawingBadge", Err.Description
End Sub

Private Function SC_PickBadgeColor(ByVal nErr As Long, ByVal nMismatch As Long, _
                                    ByVal nMissing As Long, ByVal nStale As Long) As Long
    If nErr > 0 Then
        SC_PickBadgeColor = CLR_ERROR
    ElseIf nMismatch > 0 Or nMissing > 0 Or nStale > 0 Then
        SC_PickBadgeColor = CLR_WARN
    Else
        SC_PickBadgeColor = CLR_OK
    End If
End Function

' SC_BuildBadgeText   compose the visible lblSum caption.
'   Format depends on g_bForceOverwrite:
'     OFF  — 27 drawings · 3 to create · 1 to recreate · 23 current
'     ON   — 27 drawings · 27 to recreate (force)
'   "all current" shortcut only when Force is OFF and nothing to do.
'   Errors are always shown last, before the (force) marker.
Private Function SC_BuildBadgeText(ByRef t As TBadgeCounts) As String
    Dim sep As String
    sep = " " & Chr(183) & " "

    Dim sSum As String
    sSum = L2("scan.sum.drawings", CStr(t.nTot))

    Dim nCreate As Long, nRecreate As Long, nCurrent As Long
    nCreate = t.nMissing
    If g_bForceOverwrite Then
        nRecreate = t.nStale + t.nMismatch + t.nReady
        nCurrent = 0
    Else
        nRecreate = t.nStale + t.nMismatch
        nCurrent = t.nReady
    End If

    ' All-current shortcut: Force OFF, nothing to do.
    If Not g_bForceOverwrite Then
        If t.nErr = 0 And t.nMismatch = 0 And t.nMissing = 0 And t.nStale = 0 Then
            SC_BuildBadgeText = sSum & sep & L("scan.sum.all.current")
            Exit Function
        End If
    End If

    If nCreate > 0 Then
        sSum = sSum & sep & L2("scan.sum.to_create", CStr(nCreate))
    End If
    If nRecreate > 0 Then
        sSum = sSum & sep & L2("scan.sum.to_recreate", CStr(nRecreate))
    End If
    If nCurrent > 0 Then
        sSum = sSum & sep & L2("scan.sum.ok", CStr(nCurrent))
    End If
    If t.nErr > 0 Then
        sSum = sSum & sep & L2("scan.sum.errors", CStr(t.nErr))
    End If
    If g_bForceOverwrite Then
        sSum = sSum & sep & L("scan.sum.force_suffix")
    End If

    SC_BuildBadgeText = sSum
End Function

' SC_BuildBadgeTooltip   diagnostic breakdown for the ControlTipText.
'   Uses the missing/stale/mismatch vocabulary (technicians read the
'   tooltip; users read the visible caption). Appends (force) when
'   Force is on, so the diagnostic view mirrors the visible badge.
Private Function SC_BuildBadgeTooltip(ByRef t As TBadgeCounts) As String
    Dim s As String
    s = LN("scan.badge.tooltip", _
        CStr(t.nTot), CStr(t.nReady), CStr(t.nMissing), _
        CStr(t.nStale), CStr(t.nMismatch), CStr(t.nErr))
    If g_bForceOverwrite Then
        s = s & " " & Chr(183) & " " & L("scan.sum.force_suffix")
    End If
    SC_BuildBadgeTooltip = s
End Function

Public Sub SC_CenterAll(ByVal oFrm As Object)
    PRJ_AlignTxt oFrm, "lblRun", "lblRunTxt"
    PRJ_AlignTxt oFrm, "lblBrowse", "lblBrowseTxt"
    PRJ_AlignTxt oFrm, "lblTheme", "lblThemeTxt"
    PRJ_AlignTxt oFrm, "lblReset", "lblResetTxt"
    PRJ_AlignTxt oFrm, "lblCancel", "lblCancelTxt"
End Sub

Public Sub RP_CenterAll(ByVal oFrm As Object)
    PRJ_AlignTxt oFrm, "lblBack", "lblBackTxt"
    PRJ_AlignTxt oFrm, "lblCopy", "lblCopyTxt"
    PRJ_AlignTxt oFrm, "lblExp", "lblExpTxt"
    PRJ_AlignTxt oFrm, "lblRetry", "lblRetryTxt"
    PRJ_AlignTxt oFrm, "lblCancel", "lblCancelTxt"
    PRJ_AlignTxt oFrm, "lblFltAll", "lblFltAllTxt"
    PRJ_AlignTxt oFrm, "lblFltErr", "lblFltErrTxt"
    PRJ_AlignTxt oFrm, "lblFltSkip", "lblFltSkipTxt"
    PRJ_AlignTxt oFrm, "lblFltOk", "lblFltOkTxt"
    PRJ_AlignTxt oFrm, "lblFltNoIdw", "lblFltNoIdwTxt"
End Sub

' RP_RepositionSearch   place txtSearch to the right of lblSearch.
'   Sets ONLY Left. Top is left untouched so the designer-set
'   vertical position stays exactly as designed — no drift on
'   repeated calls, no dependency on lblSearch.Top.
Public Sub RP_RepositionSearch(ByVal oFrm As Object)
    On Error Resume Next
    Dim oL As Object, oT As Object
    Set oL = oFrm.Controls("lblSearch")
    Set oT = oFrm.Controls("txtSearch")
    If oL Is Nothing Or oT Is Nothing Then Exit Sub
    oT.Left = oL.Left + oL.Width + 6
    On Error GoTo 0
End Sub

' SC_UpdateSum2   update lblSum2 with the audit summary.
'   v9.19.7: audit mode gate. Prior version ran the audit summary
'   unconditionally, so "N models (X assemblies, Y parts) ..."
'   appeared even when the Audit toggle was off. SC_Validate
'   already hides lblSum2 when audit is off, but the direct
'   SC_UpdateSum2 calls in frmScan (UserForm_Activate,
'   OptAuditClick, txtPath_AfterUpdate, OptSubClick,
'   BrowseClick, ResetClick) undid the hide.
'   v9.19.6: check the modValidate audit cache first. If the cache
'   is fresh for the same folder and subfolder flag, use the cached
'   counts. Otherwise fall back to PRJ_GetAuditSummary.
Public Sub SC_UpdateSum2(ByVal oFrm As Object)
    On Error Resume Next
    Dim oSum2 As Object
    Set oSum2 = oFrm.Controls("lblSum2")
    If oSum2 Is Nothing Then Exit Sub

    ' v9.19.7: audit mode gate.
    If Not g_bAuditMode Then
        If Len(oSum2.Caption) > 0 Then oSum2.Caption = ""
        oSum2.ControlTipText = ""
        Exit Sub
    End If

    Dim sPath As String
    sPath = SN_SanitizePath(oFrm.Controls("txtPath").Text)
    If Len(sPath) = 0 Then
        oSum2.Caption = ""
        Exit Sub
    End If
    If Not PRJ_FSO().FolderExists(sPath) Then
        oSum2.Caption = ""
        Exit Sub
    End If

    ' v9.19.6: cache-first path.
    Dim nIam As Long, nIpt As Long, nWith As Long, nWithout As Long
    If SC_GetCachedAudit(sPath, g_bSubfolders, nIam, nIpt, nWith, nWithout) Then
        Dim sCached As String
        sCached = LN("summary.hdr.models_full", _
                     CStr(nIam + nIpt), CStr(nIam), CStr(nIpt))
        sCached = sCached & " " & Chr(183) & " " _
                & L2("summary.hdr.with_idw", CStr(nWith))
        sCached = sCached & " " & Chr(183) & " " _
                & L2("summary.hdr.no_idw", CStr(nWithout))

        oSum2.ForeColor = CLR_TEMP_INFO
        oSum2.Caption = sCached
        Exit Sub
    End If

    ' Fallback: 4-pass traversal via modSummary.
    Dim sAudit As String
    sAudit = PRJ_GetAuditSummary(sPath, g_bSubfolders)

    oSum2.ForeColor = CLR_TEMP_INFO
    oSum2.Caption = sAudit
    On Error GoTo 0
End Sub

Public Sub SC_ReapplyColors(ByVal oFrm As Object)
    On Error Resume Next
    Dim oSum  As Object
    Dim oSum2 As Object

    Set oSum = oFrm.Controls("lblSum")
    Set oSum2 = oFrm.Controls("lblSum2")

    If Not oSum Is Nothing Then
        If g_tLastBadge.nTot > 0 And Len(oSum.Caption) > 0 Then
            If oSum.Caption = g_sLastBadgeText Then
                oSum.ForeColor = SC_PickBadgeColor(g_tLastBadge.nErr, _
                                                    g_tLastBadge.nMismatch, _
                                                    g_tLastBadge.nMissing, _
                                                    g_tLastBadge.nStale)
            End If
        End If
    End If

    If Not oSum2 Is Nothing Then
        If Len(oSum2.Caption) > 0 Then
            oSum2.ForeColor = CLR_TEMP_INFO
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub UI_SleepResponsive(ByVal lMs As Long)
    Dim dStart   As Double
    Dim dElapsed As Double
    Dim dLeft    As Double
    Dim dTarget  As Double

    On Error GoTo EH

    If lMs < 0 Then lMs = 0

    dStart = TM_Now()
    dTarget = lMs / MS_PER_SECOND

    Do
        DoEvents

        dElapsed = TM_Elapsed(dStart, TM_Now())
        dLeft = dTarget - dElapsed

        If dLeft <= 0# Then Exit Do

        If dLeft > UI_SLEEP_CHUNK_MS / MS_PER_SECOND Then
            Sleep UI_SLEEP_CHUNK_MS
        Else
            Sleep CLng(dLeft * MS_PER_SECOND)
        End If
    Loop
    Exit Sub

EH:
    PRJ_LogWarn "UI_SleepResponsive", "Sleep failed: " & Err.Description
End Sub

' END: modUIHelpers v9.19.9

