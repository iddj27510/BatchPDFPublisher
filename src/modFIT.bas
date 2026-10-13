Attribute VB_Name = "modFIT"
Option Explicit
' Module    : modFIT
' Purpose   : Fit-to-Width Label Framework. Shrinks, word-wraps or
'             ellipsis-truncates label captions so they fit inside
'             their paired background label without clipping.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Documentation accuracy + header standardization):
'   Header converted from Style A (free-form) to Style B
'   (Module/Purpose/Reference/Version).
'   Added missing Requires section (FIT_* constants from
'   modConstants).
'   Added missing Notes section; documented that this module DOES
'   use line-continuation characters in procedure signatures.
'   No functional changes.
'
' v9.19.6: Header sync to PRJ_VERSION. No functional changes.
' v9.19.0 (Wordwrap-first algorithm):
'   FIT_FitLabel now tries wordwrap at the base font before shrinking.
'   Algorithm per font size (9 down to 7, step 0.5):
'     1. AutoSize=True, WordWrap=False. If Width fits in
'        constraint (bg width - 7) -> use this font, single line.
'     2. AutoSize=True, WordWrap=True, Width = bg width - 7.
'        If Height fits in bg height -> use this font, two lines.
'     3. Shrink font by FIT_SHRINK_STEP and retry.
'   If even FIT_MIN_FONT does not fit, clip: AutoSize=False,
'   WordWrap=True, Width=bg-7, Height=bg. ControlTipText gets
'   the full text.
'
'   For FIT_ELL_MIDDLE / FIT_ELL_END (paths, titles), the old
'   single-line + ellipsis logic is preserved.
'
'   New optional parameter dConstraintW: the bg label width.
'   When 0, falls back to oTxt.Width (legacy behavior). FIT_FitAllLabels
'   passes the bg width automatically.
'
' v9.18.2 (TextAlign preservation):
'   FIT_MeasureText saves and restores TextAlign.
' v9.18.1: FIT_ELLIPSIS_TAIL_CHARS from modConstants.
' v9.18.0: FIT_START_FONT / FIT_MIN_FONT / FIT_SHRINK_STEP.
'
' Requires  : FIT_ELLIPSIS_TAIL_CHARS, FIT_START_FONT, FIT_MIN_FONT,
'             FIT_SHRINK_STEP (modConstants).
'
' Notes     : Called by modTheme (PRJ_UpdateThemeLabel) and
'             modEntry / modHover paths via FIT_FitAllLabels.
'             Self-contained otherwise: no cross-module function
'             calls. The skip list (FIT_IsSkipped) hardcodes the
'             (form, control) pairs that must not be touched; add
'             new entries there.
'             No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public Enum FIT_EllipsisMode
    FIT_ELL_NONE = 0
    FIT_ELL_MIDDLE = 1
    FIT_ELL_END = 2
End Enum

Private Const FIT_MARGIN As Double = 7#

Public Sub FIT_MeasureText(ByVal oLabel As Object, _
                            ByRef dWidth As Double, _
                            ByRef dHeight As Double)
    On Error GoTo EH
    Dim bWasAuto  As Boolean
    Dim nWasAlign As Long
    bWasAuto = oLabel.AutoSize
    nWasAlign = oLabel.TextAlign

    oLabel.AutoSize = True
    DoEvents
    dWidth = oLabel.Width
    dHeight = oLabel.Height
    oLabel.AutoSize = bWasAuto
    oLabel.TextAlign = nWasAlign
    Exit Sub
EH:
    On Error Resume Next
    oLabel.AutoSize = bWasAuto
    oLabel.TextAlign = nWasAlign
    dWidth = oLabel.Width
    dHeight = oLabel.Height
    On Error GoTo 0
End Sub

Public Function FIT_EllipsisMiddle(ByVal sFull As String, _
                                    ByVal dMaxWidth As Double, _
                                    ByVal oLabel As Object) As String
    On Error GoTo EH
    Dim dW As Double, dH As Double
    Dim sCand As String
    Dim nRight As Long, nLeft As Long

    oLabel.Caption = sFull
    FIT_MeasureText oLabel, dW, dH
    If dW <= dMaxWidth Then FIT_EllipsisMiddle = sFull: Exit Function

    nRight = FIT_ELLIPSIS_TAIL_CHARS
    If Len(sFull) < nRight * 2 Then nRight = Len(sFull) \ 2
    nLeft = Len(sFull) - nRight

    Do While nLeft > 3
        sCand = Left$(sFull, nLeft) & "..." & Right$(sFull, nRight)
        oLabel.Caption = sCand
        FIT_MeasureText oLabel, dW, dH
        If dW <= dMaxWidth Then FIT_EllipsisMiddle = sCand: Exit Function
        nLeft = nLeft - 1
    Loop

    FIT_EllipsisMiddle = "..." & Right$(sFull, nRight)
    Exit Function
EH:
    FIT_EllipsisMiddle = sFull
End Function

Public Function FIT_EllipsisEnd(ByVal sFull As String, _
                                 ByVal dMaxWidth As Double, _
                                 ByVal oLabel As Object) As String
    On Error GoTo EH
    Dim dW As Double, dH As Double
    Dim n As Long
    Dim sCand As String

    oLabel.Caption = sFull
    FIT_MeasureText oLabel, dW, dH
    If dW <= dMaxWidth Then FIT_EllipsisEnd = sFull: Exit Function

    n = Len(sFull)
    Do While n > 3
        n = n - 1
        sCand = Left$(sFull, n) & "..."
        oLabel.Caption = sCand
        FIT_MeasureText oLabel, dW, dH
        If dW <= dMaxWidth Then FIT_EllipsisEnd = sCand: Exit Function
    Loop

    FIT_EllipsisEnd = "..."
    Exit Function
EH:
    FIT_EllipsisEnd = sFull
End Function

' FIT_FitLabel   fit sFullText into oTxt.
'   dConstraintW: bg label width (0 = legacy: use oTxt.Width)
Public Sub FIT_FitLabel(ByVal oTxt As Object, _
                         ByVal sFullText As String, _
                         Optional ByVal eEllipsis As FIT_EllipsisMode = FIT_ELL_NONE, _
                         Optional ByVal bBold As Boolean = False, _
                         Optional ByVal dConstraintW As Double = 0)
    On Error GoTo EH

    Dim dMaxW As Double
    Dim dH_bg As Double
    If dConstraintW > 0 Then
        dMaxW = dConstraintW - FIT_MARGIN
    Else
        dMaxW = oTxt.Width
    End If
    dH_bg = oTxt.Height

    If Len(sFullText) = 0 Then Exit Sub

    ' ---- Path/title mode: single line + ellipsis ----
    If eEllipsis <> FIT_ELL_NONE Then
        Dim dFontP As Double
        Dim dWP As Double, dHP As Double
        Dim sShownP As String

        oTxt.AutoSize = False
        oTxt.WordWrap = False
        oTxt.Font.Bold = bBold
        oTxt.Caption = sFullText

        dFontP = FIT_START_FONT
        oTxt.Font.Size = dFontP

        Do While dFontP > FIT_MIN_FONT
            FIT_MeasureText oTxt, dWP, dHP
            If dWP <= dMaxW Then oTxt.ControlTipText = "": Exit Sub
            dFontP = dFontP - FIT_SHRINK_STEP
            oTxt.Font.Size = dFontP
        Loop

        FIT_MeasureText oTxt, dWP, dHP
        If dWP <= dMaxW Then oTxt.ControlTipText = "": Exit Sub

        Select Case eEllipsis
            Case FIT_ELL_MIDDLE
                sShownP = FIT_EllipsisMiddle(sFullText, dMaxW, oTxt)
            Case FIT_ELL_END
                sShownP = FIT_EllipsisEnd(sFullText, dMaxW, oTxt)
            Case Else
                sShownP = sFullText
        End Select

        oTxt.Caption = sShownP
        If sShownP <> sFullText Then
            oTxt.ControlTipText = sFullText
        Else
            oTxt.ControlTipText = ""
        End If
        Exit Sub
    End If

    ' ---- Button label mode: wordwrap-first ----
    oTxt.Font.Bold = bBold
    oTxt.Caption = sFullText

    Dim dFont As Double
    Dim dW As Double
    Dim dH As Double

    dFont = FIT_START_FONT
    Do While dFont >= FIT_MIN_FONT
        ' Try single line at this font
        oTxt.AutoSize = True
        oTxt.WordWrap = False
        oTxt.Font.Size = dFont
        FIT_MeasureText oTxt, dW, dH
        If dW <= dMaxW Then
            oTxt.ControlTipText = ""
            Exit Sub
        End If

        ' Try two lines with wordwrap at this font
        oTxt.AutoSize = True
        oTxt.WordWrap = True
        oTxt.Width = dMaxW
        FIT_MeasureText oTxt, dW, dH
        If dH <= dH_bg Then
            ' Keep AutoSize=True + WordWrap=True + Width=dMaxW
            oTxt.ControlTipText = ""
            Exit Sub
        End If

        dFont = dFont - FIT_SHRINK_STEP
    Loop

    ' Nothing fit. Clip at min font, wordwrap on, tooltip full text.
    oTxt.AutoSize = False
    oTxt.WordWrap = True
    oTxt.Width = dMaxW
    oTxt.Height = dH_bg
    oTxt.Font.Size = FIT_MIN_FONT
    oTxt.Caption = sFullText
    oTxt.ControlTipText = sFullText
    Exit Sub

EH:
    oTxt.Caption = sFullText
End Sub

Public Function FIT_GetEllipsisForTag(ByVal sTag As String) As FIT_EllipsisMode
    Select Case LCase$(sTag)
        Case "title", "text-primary"
            FIT_GetEllipsisForTag = FIT_ELL_MIDDLE
        Case "text-secondary"
            FIT_GetEllipsisForTag = FIT_ELL_END
        Case Else
            FIT_GetEllipsisForTag = FIT_ELL_NONE
    End Select
End Function

' FIT_FitAllLabels   walk the form, fit every "Txt" label.
'   For each lblXxxTxt, look up lblXxx (background) and pass its
'   Width as the constraint. Controls listed in the skip list
'   are not touched.
Public Sub FIT_FitAllLabels(ByVal oFrm As Object)
    On Error Resume Next
    Dim ctrl As Object
    Dim sName As String
    Dim oTxt As Object
    Dim oBg As Object
    Dim bgW As Double
    Dim bgH As Double
    Dim sFull As String
    Dim eMode As FIT_EllipsisMode
    Dim sFrm As String

    sFrm = oFrm.Name

For Each ctrl In oFrm.Controls
    sName = ctrl.Name
    If Right$(sName, 3) = "Txt" Then
        Set oTxt = ctrl
        sFull = oTxt.Caption
        If Len(sFull) > 0 Then
            If Not FIT_IsSkipped(sFrm, sName) Then
                eMode = FIT_GetEllipsisForTag(oTxt.Tag)

                    bgW = 0
                    bgH = 0
                    Set oBg = Nothing
                    On Error Resume Next
                    Set oBg = oFrm.Controls(Left$(sName, Len(sName) - 3))
                    On Error GoTo 0
                    If Not oBg Is Nothing Then
                        bgW = oBg.Width
                        bgH = oBg.Height
                    End If

                    If bgW > 0 Then
                        oTxt.Height = bgH
                        FIT_FitLabel oTxt, sFull, eMode, oTxt.Font.Bold, bgW
                    Else
                        FIT_FitLabel oTxt, sFull, eMode, oTxt.Font.Bold
                    End If
                End If
            End If
        End If
    Next ctrl
    On Error GoTo 0
End Sub

' FIT_IsSkipped   true if the (form, control) pair is on the
'   explicit skip list. Skipped controls keep their
'   designer-captured font, AutoSize and WordWrap because the
'   designer chose a specific non-default style that FIT would
'   destroy.
'
'   To add a new entry, add one Case below. Form name and control
'   name are case-sensitive and must match the designer-assigned
'   values.
Private Function FIT_IsSkipped(ByVal sForm As String, _
                                ByVal sControl As String) As Boolean
    Dim sKey As String
    sKey = sForm & "|" & sControl
    Select Case sKey
        Case "frmScan|lblBrowseTxt"
            FIT_IsSkipped = True
        Case Else
            FIT_IsSkipped = False
    End Select
End Function

' END: modFIT v9.19.9

