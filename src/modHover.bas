Attribute VB_Name = "modHover"
Option Explicit
' Module    : modHover
' Purpose   : Hover state machine and hover button registry
' Reference : Appendix A.13, A.103
' Version   : 9.19.11 (Inventor 2027)
'
' v9.19.11 (initial resting state):
'   PRJ_InitHover now applies resting state to every registered
'   button as the final step, so filter buttons (All/Err/Skip/Ok/
'   NoIdw) and Run show the correct color on first paint.
'   Without this pass, they kept their designer default until the
'   user hovered the mouse over them — the active filter (All)
'   only turned blue on first hover.
'
' v9.19.6 (Documentation accuracy + header completion):
'   Added missing Requires section (CHoverBtn class, modState
'   globals, modConstants tag prefixes, modState/modPalette colors).
'   Notes section corrected: this module DOES use line-continuation
'   characters (in the FW_PickHoverColor signature).
'   Changelog entries translated from Serbian to English
'   ("Faza" -> "Phase"). No functional changes.
'
' v9.19.6: Header sync to PRJ_VERSION. No functional changes.
' v9.19.5:
'   FW_ReleaseAllHoverButtons now resets m_sCurrentHover to "".
'   Prior revision cleared g_colHoverBtns but left the cached hover
'   button name intact; the next FW_ResetHover would skip the
'   resting-state pass because it saw a stale non-empty name.
'   (Fix #10 from the v9.19.5 bugfix batch.)
' v9.18.0 (§6.11 Memory leak):
'   PRJ_InitHover now calls FW_ReleaseAllHoverButtons before
'   creating a new collection. Prevents orphaned CHoverBtn
'   instances with dangling WithEvents references on form
'   re-initialization.
' v9.16.8 (7.6b): XML doc annotations; inline WHY in FW_PickHoverColor.
' v9.15.4 (Phase 6D.1, M.11): FW_PickHoverColor replaces 3 hover
'          color functions.
' v9.14.7 (Phase 3, H.14): FW_ReleaseAllHoverButtons.
' v9.14.0 (ARCH-32): FW_SetResting refactored.
' v9.13.7 (RULE S.2): removed decorative separator comment lines.
'
' Requires  : CHoverBtn (class).
'             g_colHoverBtns, g_bHoverRunEnabled, g_sActiveFilter
'             (modState).
'             CLR_*, BTN_* palette variables (modState / modPalette).
'             TAG_BTN_PREFIX, TAG_FILTER_PREFIX, TAG_TXT_SUFFIX
'             (modConstants).
'
' Notes     : Extracted from modTheme v9.13.2 (ARCH-04).
'             No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Private m_sCurrentHover As String

Public Sub FW_SetResting(oFrm As Object, sBtn As String, bRun As Boolean, bFilter As Boolean)
    Dim b As Object, t As Object
    If Not FW_GetBtnPair(oFrm, sBtn, b, t) Then Exit Sub
    Dim clr As Long, clrTxt As Long
    If bRun Then
        FW_PickRunColor clr, clrTxt
    ElseIf bFilter Then
        FW_PickFilterColor b.Tag, clr, clrTxt
    Else
        FW_PickNeutralColor clr, clrTxt
    End If
    FW_ApplyColors b, t, clr, clrTxt
End Sub

Private Sub FW_ApplyColors(ByRef b As Object, ByRef t As Object, ByVal clr As Long, ByVal clrTxt As Long)
    If b.BackColor <> clr Then b.BackColor = clr
    If t Is Nothing Then Exit Sub
    If t.BackColor <> clr Then t.BackColor = clr
    If t.ForeColor <> clrTxt Then t.ForeColor = clrTxt
End Sub

Private Sub FW_PickRunColor(ByRef clr As Long, ByRef clrTxt As Long)
    If g_bHoverRunEnabled Then
        clr = CLR_RUN_BG
        clrTxt = CLR_RUN_TXT
    Else
        clr = CLR_RUN_DIS
        clrTxt = CLR_RUN_DIS_TXT
    End If
End Sub

Private Sub FW_PickFilterColor(ByVal sTag As String, ByRef clr As Long, ByRef clrTxt As Long)
    Dim sKey As String
    sKey = Mid(sTag, Len(TAG_FILTER_PREFIX) + 1)
    If g_sActiveFilter = sKey Then
        clr = CLR_ACCENT
        clrTxt = CLR_ACCENT_TXT
    Else
        clr = CLR_BTN_SEC
        clrTxt = CLR_BTN_SEC_TXT
    End If
End Sub

Private Sub FW_PickNeutralColor(ByRef clr As Long, ByRef clrTxt As Long)
    clr = CLR_BTN_SEC
    clrTxt = CLR_BTN_SEC_TXT
End Sub

Private Function FW_GetBtnPair(oFrm As Object, sBtn As String, ByRef b As Object, ByRef t As Object) As Boolean
    FW_GetBtnPair = False
    Set b = Nothing
    Set t = Nothing
    On Error Resume Next
    Set b = oFrm.Controls(sBtn)
    On Error GoTo 0
    If b Is Nothing Then Exit Function
    On Error Resume Next
    Set t = oFrm.Controls(sBtn & "Txt")
    On Error GoTo 0
    FW_GetBtnPair = True
End Function

Public Function FW_PickHoverColor(ByVal bIn As Boolean, ByVal bPress As Boolean, _
    ByVal clrHover As Long, ByVal clrPress As Long, _
    ByVal clrNeutral As Long) As Long

    If bPress Then
        FW_PickHoverColor = clrPress
    ElseIf bIn Then
        FW_PickHoverColor = clrHover
    Else
        FW_PickHoverColor = clrNeutral
    End If
End Function

Public Sub FW_Hover(oFrm As Object, sBtn As String, bIn As Boolean, Optional bPress As Boolean = False, Optional bRunEnabled As Boolean = True)
    If sBtn = "lblRun" And Not bRunEnabled Then Exit Sub
    Dim clr As Long
    If sBtn = "lblRun" Then
        clr = FW_PickHoverColor(bIn, bPress, BTN_HOVER_RUN, BTN_PRESS_RUN, CLR_RUN_BG)
    Else
        clr = FW_PickHoverColor(bIn, bPress, BTN_HOVER_SEC, BTN_PRESS_SEC, CLR_BTN_SEC)
    End If
    Dim b As Object, t As Object
    If Not FW_GetBtnPair(oFrm, sBtn, b, t) Then Exit Sub
    If b.BackColor <> clr Then b.BackColor = clr
    If Not t Is Nothing Then
        If t.BackColor <> clr Then t.BackColor = clr
    End If
End Sub

Public Sub FW_ApplyFilterHover(oFrm As Object, sBtn As String, bActive As Boolean, bPress As Boolean)
    Dim clr As Long
    If bActive Then
        clr = FW_PickHoverColor(True, bPress, BTN_HOVER_ACCENT, BTN_PRESS_ACCENT, CLR_ACCENT)
    Else
        clr = FW_PickHoverColor(True, bPress, BTN_HOVER_SEC, BTN_PRESS_SEC, CLR_BTN_SEC)
    End If
    Dim b As Object, t As Object
    If Not FW_GetBtnPair(oFrm, sBtn, b, t) Then Exit Sub
    If b.BackColor <> clr Then b.BackColor = clr
    If Not t Is Nothing Then
        If t.BackColor <> clr Then t.BackColor = clr
    End If
End Sub

Public Sub FW_HandleHover(oBtn As CHoverBtn, bIn As Boolean, bPress As Boolean)
    If g_colHoverBtns Is Nothing Then Exit Sub
    If Not bIn Then
        m_sCurrentHover = ""
        FW_SetResting oBtn.FormRef, oBtn.BtnName, oBtn.IsRun, oBtn.IsFilter
        Exit Sub
    End If
    If m_sCurrentHover <> oBtn.BtnName Then
        FW_ResetOtherHoverButtons oBtn
        m_sCurrentHover = oBtn.BtnName
    End If
    FW_ApplyHoverToButton oBtn, bPress
End Sub

Private Sub FW_ResetOtherHoverButtons(oBtn As CHoverBtn)
    Dim o As CHoverBtn
    For Each o In g_colHoverBtns
        If o.BtnName <> oBtn.BtnName Then
            FW_SetResting o.FormRef, o.BtnName, o.IsRun, o.IsFilter
        End If
    Next
End Sub

Private Sub FW_ApplyHoverToButton(oBtn As CHoverBtn, bPress As Boolean)
    If oBtn.IsFilter Then
        FW_ApplyFilterHoverButton oBtn, bPress
    ElseIf oBtn.IsRun Then
        If g_bHoverRunEnabled Then FW_Hover oBtn.FormRef, oBtn.BtnName, True, bPress
    Else
        FW_Hover oBtn.FormRef, oBtn.BtnName, True, bPress
    End If
End Sub

Private Sub FW_ApplyFilterHoverButton(oBtn As CHoverBtn, bPress As Boolean)
    Dim sKey As String
    sKey = Mid(oBtn.BtnRole, Len(TAG_FILTER_PREFIX) + 1)
    FW_ApplyFilterHover oBtn.FormRef, oBtn.BtnName, (g_sActiveFilter = sKey), bPress
End Sub

Public Sub FW_ResetHover()
    If g_colHoverBtns Is Nothing Then Exit Sub
    If Len(m_sCurrentHover) = 0 Then Exit Sub
    m_sCurrentHover = ""
    Dim o As CHoverBtn
    For Each o In g_colHoverBtns
        FW_SetResting o.FormRef, o.BtnName, o.IsRun, o.IsFilter
    Next
End Sub

Public Sub FW_ReleaseAllHoverButtons()
    On Error Resume Next
    Dim o As CHoverBtn
    If Not g_colHoverBtns Is Nothing Then
        For Each o In g_colHoverBtns
            o.Cleanup
        Next
    End If
    Set g_colHoverBtns = Nothing
    m_sCurrentHover = ""
    On Error GoTo 0
End Sub

Public Sub PRJ_InitHover(oFrm As Object)
    ' §6.11: Release any previous hover button references before
    ' creating a new collection. Without this, re-initialization
    ' would leave orphaned CHoverBtn instances with dangling
    ' WithEvents references keeping the previous form instance
    ' alive in memory.
    If Not g_colHoverBtns Is Nothing Then
        FW_ReleaseAllHoverButtons
    End If

    Set g_colHoverBtns = New Collection
    g_bHoverRunEnabled = False
    g_sActiveFilter = "All"
    m_sCurrentHover = ""

    Dim ctrl As Object
    For Each ctrl In oFrm.Controls
        PRJ_TryAddHoverButton oFrm, ctrl
    Next

    ' v9.19.11: apply initial resting state to every registered
    ' button so filter buttons (All/Err/Skip/Ok/NoIdw) and Run
    ' show their correct colors on first paint, not just on first
    ' mouse hover. Without this pass, filter buttons used the
    ' designer default until the user moved the mouse over them.
    Dim oHover As CHoverBtn
    For Each oHover In g_colHoverBtns
        FW_SetResting oHover.FormRef, oHover.BtnName, oHover.IsRun, oHover.IsFilter
    Next oHover
End Sub

Private Sub PRJ_TryAddHoverButton(oFrm As Object, oCtl As Object)
    Dim sTag As String
    sTag = oCtl.Tag
    If Not PRJ_IsHoverFrameTag(sTag) Then Exit Sub
    Dim bRun As Boolean, bFilter As Boolean
    PRJ_ClassifyHoverTag sTag, bRun, bFilter
    PRJ_AddHover oFrm, oCtl.Name, bRun, bFilter
End Sub

Private Function PRJ_IsHoverFrameTag(sTag As String) As Boolean
    PRJ_IsHoverFrameTag = False
    If Left(sTag, Len(TAG_BTN_PREFIX)) <> TAG_BTN_PREFIX Then Exit Function
    If Right(sTag, Len(TAG_TXT_SUFFIX)) = TAG_TXT_SUFFIX Then Exit Function
    If sTag = "btn-run" Or sTag = "btn-secondary" Then
        PRJ_IsHoverFrameTag = True
    ElseIf Left(sTag, Len(TAG_FILTER_PREFIX)) = TAG_FILTER_PREFIX Then
        PRJ_IsHoverFrameTag = True
    End If
End Function

Private Sub PRJ_ClassifyHoverTag(sTag As String, ByRef bRun As Boolean, ByRef bFilter As Boolean)
    bRun = (sTag = "btn-run")
    bFilter = (Left(sTag, Len(TAG_FILTER_PREFIX)) = TAG_FILTER_PREFIX)
End Sub

Private Sub PRJ_AddHover(oFrm As Object, sBtn As String, bRun As Boolean, bFilter As Boolean)
    Dim o As CHoverBtn
    Set o = New CHoverBtn
    On Error Resume Next
    Set o.lbl = oFrm.Controls(sBtn)
    On Error GoTo 0
    If o.lbl Is Nothing Then Exit Sub
    o.BtnRole = o.lbl.Tag
    On Error Resume Next
    Set o.txt = oFrm.Controls(sBtn & "Txt")
    On Error GoTo 0
    o.BtnName = sBtn
    o.IsRun = bRun
    o.IsFilter = bFilter
    Set o.FormRef = oFrm
    g_colHoverBtns.Add o
End Sub

' END: modHover v9.19.11

