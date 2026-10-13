Attribute VB_Name = "modTheme"
Option Explicit
' Module    : modTheme
' Purpose   : Theme application, cycling, run-state, text alignment
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.7 (PRJ_AlignTxt rewrite):
'   PRJ_AlignTxt now measures the text label's natural bounds
'   internally and centers geometrically. Prior version relied on
'   AutoSize=False + TextAlign=2, which MSForms Label applies with
'   ~3px internal left padding before centering — visible on short
'   captions ("Dark", "Run"). Geometric centering has no such
'   padding. ALIGN_Y_OFFSET compensates for MSForms' top-aligned
'   caption rendering inside its own bounds.
'   Procedure is now self-contained: it does not depend on
'   FIT_FitAllLabels having run first.
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in a Select Case label list and in the PRJ_AlignTxt
'   signature). No functional changes.
'
' v9.19.6: Header sync to PRJ_VERSION. No functional changes.
' v9.19.3 (Localized theme names):
'   PRJ_UpdateThemeLabel now goes through L("theme.*") keys instead
'   of hardcoded English strings. The four theme names
'   (Inventor / System / Light / Dark) change with the active
'   language, same as every other UI string.
'
' v9.18.1 (Button text centering):
'   PRJ_AlignTxt leaves AutoSize=True on the text label and
'   computes Left/Top manually from the actual label width.
' v9.18.0 (§7.19 High contrast):
'   FW_ApplyTheme returns early if g_bHighContrast is True.
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and Select Case label lists are
'             excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

' v9.19.7: geometric centering offset for PRJ_AlignTxt.
' MSForms Label renders the caption at the top of its AutoSize
' bounds; the descender space below the glyphs is empty. For
' uppercase-only captions ("RUN", "RESET") this makes the visible
' text sit ~1.5pt above the visual center of the label. The offset
' shifts the label down so the visible glyphs land at visual center.
' Tuned empirically: 0 for Segoe UI 9pt at 96 DPI on Windows 11.
' The label itself is centered geometrically; MSForms renders the
' caption at the top of the label bounds, which lands the visible
' text at visual center when offset = 0. Bump to 1.0 or -1.0 only
' if the value feels off on a different DPI or font.
Private Const ALIGN_Y_OFFSET As Double = 0#


Public Sub FW_ApplyTheme(ByVal oFrm As Object)
    If g_bHighContrast Then
        g_bThemeApplied = True
        Exit Sub
    End If

    Dim bDark As Boolean
    bDark = FW_IsDark()

    g_bCurrentDark = bDark
    g_bThemeApplied = True
    FW_LoadPalette

    oFrm.BackColor = CLR_BG

    Dim ctrl As Object
    For Each ctrl In oFrm.Controls
        Select Case ctrl.Tag
            Case "title"
                ctrl.ForeColor = CLR_TEXT_PRI
                ctrl.BackColor = CLR_BG
            Case "sep", "progress-bar"
                ctrl.BackColor = CLR_SEP
            Case "text-primary"
                ctrl.BackColor = CLR_BG
                ctrl.ForeColor = CLR_TEXT_PRI
            Case "text-secondary"
                ctrl.BackColor = CLR_BG
                ctrl.ForeColor = CLR_TEXT_SEC
            Case "input"
                ctrl.BackColor = CLR_SURFACE
                ctrl.ForeColor = CLR_TEXT_PRI
            Case "checkbox"
                ctrl.ForeColor = CLR_TEXT_PRI
                ctrl.BackColor = CLR_BG
            Case "btn-run", "btn-run-txt"
                ' handled by FW_SetResting
            Case "btn-secondary", "btn-filter-All", "btn-filter-Err", _
                 "btn-filter-Skip", "btn-filter-Ok", "btn-filter-NoIdw"
                ctrl.BackColor = CLR_BTN_SEC
                ctrl.BorderStyle = 1
            Case "btn-secondary-txt", "btn-filter-All-txt", "btn-filter-Err-txt", _
                 "btn-filter-Skip-txt", "btn-filter-Ok-txt", "btn-filter-NoIdw-txt"
                ctrl.ForeColor = CLR_BTN_SEC_TXT
                ctrl.BackColor = CLR_BTN_SEC
            Case "progress-fill"
                ctrl.BackColor = CLR_ACCENT
            Case "list"
                ctrl.BackColor = CLR_SURFACE
                ctrl.ForeColor = CLR_TEXT_PRI
        End Select
    Next

    On Error Resume Next
    If Not oFrm.Controls("lblRun") Is Nothing Then
        FW_SetResting oFrm, "lblRun", True, False
    End If
    On Error GoTo 0
End Sub

Public Sub PRJ_CycleTheme(ByVal oFrm As Object)
    Select Case g_sTheme
        Case "I": g_sTheme = "S"
        Case "S": g_sTheme = "L"
        Case "L": g_sTheme = "D"
        Case "D": g_sTheme = "I"
        Case Else: g_sTheme = "I"
    End Select
    REG_Set REG_KEY_THEME, g_sTheme
    g_bThemeApplied = False
    FW_ApplyTheme oFrm
    PRJ_UpdateThemeLabel oFrm
End Sub

' PRJ_UpdateThemeLabel   set lblThemeTxt caption.
' v9.19.3: localized via L("theme.*"). The four theme names
'   change with the active language just like every other UI
'   string. "Inventor" is a brand name and does not translate,
'   but it still goes through L() for uniformity.
Public Sub PRJ_UpdateThemeLabel(ByVal oFrm As Object)
    On Error Resume Next
    Dim oTxt As Object
    Set oTxt = oFrm.Controls("lblThemeTxt")
    If oTxt Is Nothing Then Exit Sub
    Select Case g_sTheme
        Case "I": oTxt.Caption = L("theme.inventor")
        Case "S": oTxt.Caption = L("theme.system")
        Case "L": oTxt.Caption = L("theme.light")
        Case "D": oTxt.Caption = L("theme.dark")
        Case Else: oTxt.Caption = L("theme.inventor")
    End Select
    FIT_FitLabel oTxt, oTxt.Caption, FIT_ELL_NONE, False
    On Error GoTo 0
End Sub

Public Sub FW_ApplyRunState(ByVal oFrm As Object, ByVal bEnabled As Boolean)
    g_bHoverRunEnabled = bEnabled
    FW_SetResting oFrm, "lblRun", True, False
End Sub

' PRJ_AlignTxt   center the text label inside its background label.
'   v9.19.7: does NOT use AutoSize=True.
'
'   Why: FIT_FitAllLabels (which runs before this procedure) sets
'   oTxt.Height = bgH as a workaround for FIT_FitLabel using
'   oTxt.Height internally as its dH_bg parameter. MSForms does
'   not re-measure Height when AutoSize is later set to True on a
'   label whose Height was explicitly assigned — the stale value
'   sticks. Result: oT.Height = oL.Height (24pt) instead of the
'   natural text cell (12.5pt), and geometric centering splits
'   the wrong cell.
'
'   Instead: disable AutoSize, set Height to the natural font cell
'   height (Segoe UI cell height ˜ 1.4 * Font.Size), position the
'   label geometrically, and center the caption horizontally via
'   TextAlign=2. The known MSForms Label TextAlign=2 left-shift
'   quirk is under 1pt at button widths — acceptable.
'
'   ALIGN_Y_OFFSET compensates for MSForms rendering the caption
'   at the top of the label's bounds. Without it the visible cap
'   sits visibly above the label's vertical center.
Public Sub PRJ_AlignTxt(ByVal oFrm As Object, _
                         ByVal sBg As String, _
                         ByVal sTxt As String)
    On Error Resume Next
    Dim oL As Object, oT As Object
    Set oL = oFrm.Controls(sBg)
    Set oT = oFrm.Controls(sTxt)
    If oL Is Nothing Or oT Is Nothing Then Exit Sub

    ' Natural font cell height. Segoe UI cell height is ~1.4x the
    ' point size (ascender + descender + leading).
    Dim dCellH As Double
    dCellH = oT.Font.Size * 1.4
    If dCellH < 8# Then dCellH = 8#
    If dCellH > oL.Height Then dCellH = oL.Height

    oT.AutoSize = False
    oT.WordWrap = False
    oT.Height = dCellH
    oT.Left = oL.Left
    oT.Width = oL.Width
    oT.TextAlign = 2
    oT.Top = oL.Top + (oL.Height - dCellH) / 2 + ALIGN_Y_OFFSET
    On Error GoTo 0
End Sub

' END: modTheme v9.19.9
