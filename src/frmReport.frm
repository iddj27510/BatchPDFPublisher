VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "Batch PDF Publisher"
   ClientHeight    =   10320
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   13560
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Component : frmReport (UserForm)
' Purpose   : Processing report form — summary lines, filterable list
'             of per-file results, copy/export/retry actions, and
'             back-to-scan navigation. Hosts the CKeyPreview hook
'             so Esc clears the search box or unloads the form.
' Reference : Appendix A.13
' Version   : 9.19.11 (Inventor 2027)
'
' v9.19.11 (UI fixes + audit-summary mirror):
'   1. lstReport ColumnWidths reduced to 625pt total so the
'      vertical scrollbar does not trigger a horizontal one.
'   2. lblTitle shows folder + session state, no prefix. FIT
'      call preserves the designer Width so TextAlign=2 stays
'      visually centered across the full form (previously the
'      internal FIT_MeasureText left the label at natural text
'      width and the caption appeared left-of-center).
'   4. txtSearch filters as the user types (case-insensitive
'      substring on the relative path).
'   6. lblSum2 mirrors frmScan.lblSum2 — audit summary from the
'      scan-phase cache. RP_UpdateSum2 reads g_sLastFolder
'      directly (frmReport has no txtPath control) and consults
'      SC_GetCachedAudit, falling back to PRJ_GetAuditSummary.
'   8. CopyClick no longer calls Me.Repaint (flicker over the
'      24+ row list).
'  10. RetryClick "no files in error" uses CLR_TEMP_OK (green)
'      so it stands apart from lblSum2's CLR_TEMP_INFO (blue).
'  11. lstReport size column padded with Unicode FIGURE SPACE
'      (U+2007, digit-width) so KB/MB/B suffixes align
'      vertically. Regular spaces are narrower than digits in
'      Segoe UI, producing visible misalignment.
'   All click handlers restructured to one statement per line.
'
' v9.19.7:
'   UserForm_MouseMove: FW_ApplyTheme removed (flicker fix).
'   UserForm_Activate: FIT_FitAllLabels moved before RP_CenterAll
'   on both activation paths.
'   CKeyPreview now uses concrete hook classes.
'
' v9.19.6 (Header standardization):
'   Added Component/Purpose/Reference/Version header fields.
'   No functional changes.
'
' v9.19.6:
'   A-13: Font.bOld -> Font.Bold.
'   BUG-01: ColumnCount set before ColumnWidths.
'   BUG-02: RP_UpdateTitle removed (dead code).
'   V-01: CopyClick uses UI_SleepResponsive.
'   V-02: UserForm_KeyDown no longer unloads on KeyCode=69 ('E').
' v9.19.0:
'   lblSel removed. All row details in lstReport columns.
'   lblTitle static ("report.form.title"); RP_UpdateTitle deleted.
' v9.18.8: CKeyPreview disabled (now re-enabled in v9.19.6).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Private m_sOrigCap   As String
Private m_bActivated As Boolean
Private m_oKeyPreview As CKeyPreview

Private Sub UserForm_Initialize()
    m_sOrigCap = Me.Caption
    Me.StartUpPosition = 1
    PRJ_SetupAccelerators Me
    Set m_oKeyPreview = New CKeyPreview
    m_oKeyPreview.Initialize Me, "OnFormKeyDown"
End Sub

Private Sub UserForm_Activate()
    Dim sCap As String
    sCap = L("title")
    If Len(sCap) = 0 Then sCap = Trim$(m_sOrigCap)
    Dim sVerSuffix As String
    sVerSuffix = " v" & PRJ_VERSION
    If Right$(sCap, Len(sVerSuffix)) <> sVerSuffix Then
        Me.Caption = sCap & sVerSuffix
    End If

    If g_colFiles Is Nothing Then
        Unload Me
        Exit Sub
    End If

    If m_bActivated Then
        FW_ApplyTheme Me
        PRJ_LocalizeCaptions Me
        FIT_FitAllLabels Me
        RP_CenterAll Me
        RP_RepositionSearch Me
        PRJ_ApplyTooltips Me
        FW_ResetHover
        RP_SetTitle Me
        RP_UpdateSum Me
        RP_UpdateSum2 Me
        RP_LoadList Me
        Exit Sub
    End If
    m_bActivated = True

    On Error Resume Next
    Me.Controls("lstReport").ScrollBars = 2
    Me.Controls("lstReport").ColumnCount = 5
    Me.Controls("lstReport").ColumnWidths = "130 pt;55 pt;100 pt;340 pt;0 pt"
    On Error GoTo 0

    g_bThemeApplied = False
    FW_ApplyTheme Me
    PRJ_LocalizeCaptions Me
    FIT_FitAllLabels Me
    RP_CenterAll Me
    RP_RepositionSearch Me
    PRJ_ApplyTooltips Me
    PRJ_InitHover Me
    RP_SetTitle Me
    RP_UpdateSum Me
    RP_UpdateSum2 Me
    RP_LoadList Me
    Me.Repaint
End Sub

Public Sub OnFormKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        Dim oSearch As Object
        Set oSearch = Me.Controls("txtSearch")
        If Not oSearch Is Nothing Then
            If Len(oSearch.Text) > 0 Then
                oSearch.Text = ""
                KeyCode = 0
                Exit Sub
            End If
        End If
        Unload Me
        KeyCode = 0
        Exit Sub
    End If

    If KeyCode = 70 And (Shift And 2) <> 0 Then
        On Error Resume Next
        Me.Controls("txtSearch").SetFocus
        On Error GoTo 0
        KeyCode = 0
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' v9.19.7: FW_ApplyTheme removed (flicker fix).
    FW_ResetHover
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' v9.19.6 V-02: only Esc unloads. Bare 'E' removed.
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If g_bIsProcessing Then
        Cancel = True
        RP_ShowStatus Me, L("report.msg.close_blocked"), CLR_WARN, UI_MSG_LONG_MS
    End If
End Sub

Private Sub UserForm_Terminate()
    FW_ReleaseAllHoverButtons
    If Not m_oKeyPreview Is Nothing Then
        m_oKeyPreview.Terminate
        Set m_oKeyPreview = Nothing
    End If
End Sub

' RP_SetTitle   compose the visible header:
'   <Folder>   ·   <Subfolders: X> · <Force: X> · <Audit: X>
' When no folder is known the label shows only the session state.
'
' FIT_ELL_MIDDLE path handles overflow by shrinking font down to
' 7pt, then adding middle-ellipsis. The designer Width is saved
' and restored around the FIT call: FIT_MeasureText temporarily
' sets AutoSize=True, and MSForms leaves Width at the natural
' text width when AutoSize goes back to False. Without restore
' the label shrinks to the caption width and TextAlign=2
' centers within that narrow box — visually left-of-center
' relative to the form. Restoring Width keeps the caption
' centered across the full label bounds.
Private Sub RP_SetTitle(ByVal oFrm As Object)
    On Error Resume Next
    Dim oTitle As Object
    Set oTitle = oFrm.Controls("lblTitle")
    If oTitle Is Nothing Then Exit Sub

    Dim sep As String
    sep = " " & Chr(183) & " "
    Dim sepWide As String
    sepWide = "   " & Chr(183) & "   "

    Dim sState As String
    sState = L("report.hdr.sub") & ": " & RP_OnOffLocal(g_bSubfolders)
    sState = sState & sep _
             & L("report.hdr.force_short") & ": " & RP_OnOffLocal(g_bForceOverwrite)
    sState = sState & sep _
             & L("report.hdr.audit") & ": " & RP_OnOffLocal(g_bAuditMode)

    Dim sTitle As String
    If Len(g_sLastFolder) > 0 Then
        sTitle = g_sLastFolder & sepWide & sState
    Else
        sTitle = sState
    End If

    Dim dOrigW As Double
    dOrigW = oTitle.Width

    FIT_FitLabel oTitle, sTitle, FIT_ELL_MIDDLE, True, dOrigW

    ' Restore designer Width and TextAlign — FIT_MeasureText leaves
    ' the label at natural text width after its AutoSize round-trip.
    If oTitle.Width <> dOrigW Then oTitle.Width = dOrigW
    If oTitle.TextAlign <> 2 Then oTitle.TextAlign = 2
    On Error GoTo 0
End Sub

Private Function RP_OnOffLocal(ByVal bState As Boolean) As String
    If bState Then
        RP_OnOffLocal = L("report.state.on")
    Else
        RP_OnOffLocal = L("report.state.off")
    End If
End Function

Private Sub RP_LoadList(ByVal oFrm As Object)
    On Error Resume Next
    Dim lb As Object
    Set lb = oFrm.Controls("lstReport")
    If lb Is Nothing Then Exit Sub
    lb.Clear
    If g_colFiles Is Nothing Then Exit Sub

    ' v9.19.11: read search filter (case-insensitive substring on rel path).
    Dim sSearch As String
    sSearch = ""
    On Error Resume Next
    sSearch = LCase$(Trim$(oFrm.Controls("txtSearch").Text))
    On Error GoTo 0

    Dim i As Long
    For i = 0 To g_colFiles.Count - 1
        Dim a As Variant
        a = g_colFiles(i)
        Dim sPath As String, sRes As String
        sPath = CStr(a(0))
        sRes = CStr(a(1))

        If RP_MatchesFilter(sRes) Then
            Dim sRel As String
            sRel = PRJ_GetRelPath(sPath)

            Dim bSearchHit As Boolean
            bSearchHit = (Len(sSearch) = 0)
            If Not bSearchHit Then
                bSearchHit = (InStr(1, LCase$(sRel), sSearch) > 0)
            End If

            If bSearchHit Then
                Dim sMarker As String
                Select Case STS_Classify(sRes)
                    Case ST_ERROR:                     sMarker = RPT_MARK_ERROR
                    Case ST_SKIPPED, ST_SKIPPED_EMPTY: sMarker = RPT_MARK_SKIP
                    Case Else:                         sMarker = RPT_MARK_OK
                End Select

                Dim sStatus As String
                sStatus = sMarker & " " & sRes

                Dim sSize As String
                sSize = PRJ_GetPDFSize(sPath)
                ' v9.19.11: pad with Unicode FIGURE SPACE (U+2007) rather
                ' than a regular space. Figure space is digit-width by
                ' definition, so the KB/MB/B suffix and the right-aligned
                ' digits line up vertically. Regular spaces are narrower
                ' than digits in Segoe UI, producing visible misalignment.
                Do While Len(sSize) < 10
                    sSize = ChrW$(&H2007) & sSize
                Loop

                Dim sDate As String
                sDate = PRJ_FormatDate(PRJ_GetFileDate(sPath))

                lb.AddItem sStatus
                On Error Resume Next
                lb.List(lb.ListCount - 1, 1) = sSize
                lb.List(lb.ListCount - 1, 2) = sDate
                lb.List(lb.ListCount - 1, 3) = sRel
                lb.List(lb.ListCount - 1, 4) = sPath
                On Error GoTo 0
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub RP_UpdateSum(ByVal oFrm As Object)
    On Error Resume Next
    If g_colFiles Is Nothing Then Exit Sub

    Dim t As TReportCounts
    RP_CountFiles t

    Dim sSum As String
    sSum = RP_BuildSummaryText(t, g_dTimeEnd - g_dTimeStart)

    Dim oSum As Object
    Set oSum = oFrm.Controls("lblSum")
    If oSum Is Nothing Then Exit Sub

    oSum.ForeColor = RP_PickSummaryColor(t)
    oSum.Caption = sSum
    On Error GoTo 0
End Sub

' RP_UpdateSum2   mirror frmScan.lblSum2 — audit summary from the
'   scan-phase cache. Reads g_sLastFolder directly (frmReport has
'   no txtPath control, which is why the prior delegation to
'   modUIHelpers.SC_UpdateSum2 silently emptied the label).
'
'   Cache-first: SC_GetCachedAudit checks folder + subfolder flag
'   against modValidate module state (persists across form unloads).
'   Cache miss -> PRJ_GetAuditSummary does a 4-pass FSO scan.
'   Audit-mode gate: empty when g_bAuditMode = False.
Private Sub RP_UpdateSum2(ByVal oFrm As Object)
    On Error Resume Next
    Dim oSum2 As Object
    Set oSum2 = oFrm.Controls("lblSum2")
    If oSum2 Is Nothing Then Exit Sub

    If Not g_bAuditMode Then
        If Len(oSum2.Caption) > 0 Then oSum2.Caption = ""
        oSum2.ControlTipText = ""
        Exit Sub
    End If

    Dim sPath As String
    sPath = SN_SanitizePath(g_sLastFolder)
    If Len(sPath) = 0 Then
        If Len(oSum2.Caption) > 0 Then oSum2.Caption = ""
        Exit Sub
    End If

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
        oSum2.ControlTipText = L3("scan.models.detail", CStr(nWith), CStr(nWithout))
    Else
        Dim sAudit As String
        sAudit = PRJ_GetAuditSummary(sPath, g_bSubfolders)
        oSum2.ForeColor = CLR_TEMP_INFO
        oSum2.Caption = sAudit
    End If
    On Error GoTo 0
End Sub

Private Sub RP_SetFilter(ByVal oFrm As Object, ByVal sF As String)
    If g_bIsProcessing Then Exit Sub
    g_sActiveFilter = sF
    FW_ResetHover
    oFrm.Repaint
    RP_LoadList oFrm
End Sub

Private Sub RP_ShowStatus(ByVal oFrm As Object, ByVal sText As String, ByVal lColor As Long, ByVal lMs As Long)
    On Error Resume Next
    Dim oSum As Object
    Set oSum = oFrm.Controls("lblSum")
    If oSum Is Nothing Then Exit Sub

    Dim sOld As String, lOld As Long, bWasBold As Boolean
    sOld = oSum.Caption
    lOld = oSum.ForeColor
    bWasBold = oSum.Font.Bold

    oSum.ForeColor = lColor
    oSum.Caption = sText
    oSum.Font.Bold = True
    oFrm.Repaint
    UI_SleepResponsive lMs

    oSum.Caption = sOld
    oSum.ForeColor = lOld
    oSum.Font.Bold = bWasBold
    oFrm.Repaint
    On Error GoTo 0
End Sub

Private Sub RetryClick()
    On Error GoTo EH
    If g_bIsProcessing Then Exit Sub
    If g_colFiles Is Nothing Then Exit Sub

    Dim nErr As Long
    nErr = RP_CountErrorRows()

    If nErr = 0 Then
        ' v9.19.11: CLR_TEMP_OK (green) — "no errors" is a positive
        ' state. Previously CLR_TEMP_INFO (blue) collided with
        ' lblSum2, which also uses CLR_TEMP_INFO.
        RP_ShowStatus Me, L("report.msg.retry.none"), CLR_TEMP_OK, UI_MSG_LONG_MS
        Exit Sub
    End If

    g_bIsProcessing = True
    Me.MousePointer = 11
    DoEvents

    Dim nOK As Long, nStill As Long
    RP_RetryPass nOK, nStill

    Me.MousePointer = 0
    g_bIsProcessing = False
    RP_LoadList Me
    RP_UpdateSum Me
    FW_ResetHover
    Me.Repaint

    If nStill > 0 Then
        RP_ShowStatus Me, L3("report.msg.retry.partial", CStr(nOK), CStr(nStill)), CLR_ERROR, UI_MSG_ERROR_MS
    Else
        RP_ShowStatus Me, L2("report.msg.retry.all_fixed", CStr(nOK)), CLR_TEMP_OK, UI_MSG_ERROR_MS
    End If
    Exit Sub

EH:
    Me.MousePointer = 0
    g_bIsProcessing = False
    RP_ShowStatus Me, L2("report.msg.retry.error", Err.Description), CLR_ERROR, UI_MSG_ERROR_MS
End Sub

' CopyClick   flash "Copied!" on the label, then restore. No
'   Me.Repaint — MSForms repaints the changed label automatically
'   and the full-form repaint caused a visible flicker over the
'   24+ row list.
Private Sub CopyClick()
    If g_bIsProcessing Then Exit Sub

    Dim sLog As String
    sLog = PRJ_GetReportText()

    Dim bOK As Boolean
    bOK = PRJ_PutClipboard(sLog)
    If Not bOK Then bOK = PRJ_PutClipboardFallback(sLog)

    Dim oTxt As Object
    Set oTxt = Me.Controls("lblCopyTxt")
    If oTxt Is Nothing Then Exit Sub

    Dim sOld As String, lOld As Long, bWasBold As Boolean
    sOld = oTxt.Caption
    lOld = oTxt.ForeColor
    bWasBold = oTxt.Font.Bold

    Dim sCopied As String
    sCopied = L("report.msg.copied")

    oTxt.Caption = sCopied
    oTxt.ForeColor = CLR_OK
    oTxt.Font.Bold = True
    PRJ_AlignTxt Me, "lblCopy", "lblCopyTxt"
    UI_SleepResponsive UI_MSG_SHORT_MS

    oTxt.Caption = sOld
    oTxt.ForeColor = lOld
    oTxt.Font.Bold = bWasBold
    PRJ_AlignTxt Me, "lblCopy", "lblCopyTxt"
End Sub

Private Sub ExportClick()
    If g_bIsProcessing Then Exit Sub

    Dim bOK As Boolean
    bOK = PRJ_ExportReport()
    If bOK Then
        RP_ShowStatus Me, L("report.msg.exported"), CLR_TEMP_OK, UI_MSG_MEDIUM_MS
    Else
        RP_ShowStatus Me, L("report.msg.export_failed"), CLR_ERROR, UI_MSG_ERROR_MS
    End If
End Sub

Private Sub lstReport_Click()
    ' Details are now in the list columns. Click does nothing.
End Sub

Private Sub lstReport_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If g_bIsProcessing Then Exit Sub

    On Error Resume Next
    Dim lb As Object
    Set lb = Me.Controls("lstReport")
    If lb Is Nothing Then Exit Sub
    If lb.ListIndex < 0 Then Exit Sub

    Dim sPath As String
    sPath = RP_GetSelectedPath(lb)
    If Len(sPath) = 0 Then Exit Sub

    If Not PRJ_OpenPathOrSelect(sPath) Then
        RP_ShowStatus Me, L("report.msg.open_failed"), CLR_ERROR, UI_MSG_ERROR_MS
    End If
    On Error GoTo 0
End Sub

Private Sub txtSearch_Enter()
    On Error Resume Next
    Me.Controls("txtSearch").BorderColor = CLR_ACCENT
    On Error GoTo 0
End Sub

Private Sub txtSearch_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Me.Controls("txtSearch").BorderColor = CLR_SEP
    On Error GoTo 0
End Sub

Private Sub txtSearch_Change()
    If m_bActivated Then RP_LoadList Me
End Sub

Private Sub lblFltAll_Click()
    RP_SetFilter Me, "All"
End Sub

Private Sub lblFltAllTxt_Click()
    RP_SetFilter Me, "All"
End Sub

Private Sub lblFltErr_Click()
    RP_SetFilter Me, "Err"
End Sub

Private Sub lblFltErrTxt_Click()
    RP_SetFilter Me, "Err"
End Sub

Private Sub lblFltSkip_Click()
    RP_SetFilter Me, "Skip"
End Sub

Private Sub lblFltSkipTxt_Click()
    RP_SetFilter Me, "Skip"
End Sub

Private Sub lblFltOk_Click()
    RP_SetFilter Me, "Ok"
End Sub

Private Sub lblFltOkTxt_Click()
    RP_SetFilter Me, "Ok"
End Sub

Private Sub lblFltNoIdw_Click()
    RP_SetFilter Me, "NoIdw"
End Sub

Private Sub lblFltNoIdwTxt_Click()
    RP_SetFilter Me, "NoIdw"
End Sub

Private Sub lblBack_Click()
    RP_BackClick
End Sub

Private Sub lblBackTxt_Click()
    RP_BackClick
End Sub

Private Sub RP_BackClick()
    Unload Me
    PRJ_ShowScan
End Sub

Private Sub lblRetry_Click()
    RetryClick
End Sub

Private Sub lblRetryTxt_Click()
    RetryClick
End Sub

Private Sub lblCopy_Click()
    CopyClick
End Sub

Private Sub lblCopyTxt_Click()
    CopyClick
End Sub

Private Sub lblExp_Click()
    ExportClick
End Sub

Private Sub lblExpTxt_Click()
    ExportClick
End Sub

Private Sub lblCancel_Click()
    Unload Me
End Sub

Private Sub lblCancelTxt_Click()
    Unload Me
End Sub

' END: frmReport v9.19.11

