Attribute VB_Name = "modFont"
Option Explicit
' Module    : modFont
' Purpose   : Central font policy + precise text measurement via GDI.
' Reference : Appendix A.13
' Version   : 9.25.20
'
' v9.25.20 : Version bump; no functional change.
' v1.4     : Added FONT_UI_SIZE() - returns the host message font point
'             size as configured in Windows Display settings. Reads
'             NONCLIENTMETRICSW.lfMessageFont.lfHeight, same source as
'             FONT_UI reads the face name. Fallback 9pt.
' v1.3     : Removed FONT_ScrollAndBorderPt and FONT_CellInsetPt.
' v1.2     : Added FONT_CellInsetPt and FONT_CeilPxPt.
' v1.1     : Added FONT_DPI, FONT_CellPt, FONT_ScrollAndBorderPt.
' v1.0     : FONT_UI / FONT_GRID / FONT_TextWidthPt.

#If VBA7 Then
Private Declare PtrSafe Function fnGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function fnReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function fnGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function fnCreateFont Lib "gdi32" Alias "CreateFontW" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As LongPtr
Private Declare PtrSafe Function fnSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function fnDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function fnGetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE_STRUCT) As Long
Private Declare PtrSafe Function fnGetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As LongPtr, lpMetrics As TEXTMETRIC_STRUCT) As Long
Private Declare PtrSafe Function fnSystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uiAction As Long, ByVal uiParam As Long, ByRef pvParam As NONCLIENTMETRICSW, ByVal fWinIni As Long) As Long
#Else
Private Declare Function fnGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
Private Declare Function fnReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function fnGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function fnCreateFont Lib "gdi32" Alias "CreateFontW" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function fnSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function fnDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function fnGetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE_STRUCT) As Long
Private Declare Function fnGetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, lpMetrics As TEXTMETRIC_STRUCT) As Long
Private Declare Function fnSystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uiAction As Long, ByVal uiParam As Long, ByRef pvParam As NONCLIENTMETRICSW, ByVal fWinIni As Long) As Long
#End If

Private Type SIZE_STRUCT
    cx As Long
    cy As Long
End Type

Private Type TEXTMETRIC_STRUCT
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Integer
    tmLastChar As Integer
    tmDefaultChar As Integer
    tmBreakChar As Integer
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Private Type LOGFONTW
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(0 To 31) As Integer
End Type

Private Type NONCLIENTMETRICSW
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONTW
    iSmCaptionWidth As Long
    iSmCaptionHeight As Long
    lfSmCaptionFont As LOGFONTW
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONTW
    lfStatusFont As LOGFONTW
    lfMessageFont As LOGFONTW
End Type

Private Const SPI_GETNONCLIENTMETRICS As Long = 41
Private Const DEFAULT_CHARSET As Long = 1
Private Const DEFAULT_QUALITY As Long = 0

Private m_sUICache As String
Private m_nDPICache As Long
Private m_dUISizeCache As Double

Public Function FONT_DPI() As Long
    If m_nDPICache > 0 Then
        FONT_DPI = m_nDPICache
        Exit Function
    End If

    #If VBA7 Then
    Dim hDC As LongPtr
    #Else
    Dim hDC As Long
    #End If

    On Error Resume Next
    hDC = fnGetDC(0)
    If hDC <> 0 Then
        m_nDPICache = fnGetDeviceCaps(hDC, 90)
        fnReleaseDC 0, hDC
    End If
    On Error GoTo 0

    If m_nDPICache <= 0 Then m_nDPICache = 96
    FONT_DPI = m_nDPICache
End Function

Public Function FONT_UI() As String
    If Len(m_sUICache) > 0 Then
        FONT_UI = m_sUICache
        Exit Function
    End If

    Dim s As String

    If Len(FONT_UI_OVERRIDE) > 0 Then
        s = FONT_UI_OVERRIDE
    Else
        s = fnSystemMessageFont()
    End If

    If Len(s) = 0 Then s = "Segoe UI"
    m_sUICache = s
    FONT_UI = s
End Function

' FONT_UI_SIZE   returns point size of host message font from Windows
'   Display settings (NONCLIENTMETRICSW.lfMessageFont.lfHeight).
'   Falls back to 9pt.
Public Function FONT_UI_SIZE() As Double
    If m_dUISizeCache > 0 Then
        FONT_UI_SIZE = m_dUISizeCache
        Exit Function
    End If

    Dim ncm As NONCLIENTMETRICSW
    Dim nHeightPx As Long
    Dim nDPI As Long
    Dim dPt As Double

    On Error GoTo fail

    ncm.cbSize = LenB(ncm)
    If fnSystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, ncm, 0) <> 0 Then
        nHeightPx = ncm.lfMessageFont.lfHeight
        If nHeightPx <> 0 Then
            nDPI = FONT_DPI()
            If nDPI <= 0 Then nDPI = 96
            dPt = -nHeightPx * 72# / nDPI
            If dPt >= 5# And dPt <= 30# Then
                m_dUISizeCache = dPt
                FONT_UI_SIZE = dPt
                Exit Function
            End If
        End If
    End If

fail:
    On Error GoTo 0
    m_dUISizeCache = 9#
    FONT_UI_SIZE = 9#
End Function

Public Function FONT_GRID() As String
    If RP_GRID_MONO Then
        FONT_GRID = "Consolas"
    Else
        FONT_GRID = FONT_UI()
    End If
End Function

' FONT_TextWidthPt   exact rendered width of sText in points
Public Function FONT_TextWidthPt(ByVal sText As String, ByVal sFontName As String, ByVal dSizePt As Double, ByVal bBold As Boolean) As Double
    If Len(sText) = 0 Then
        FONT_TextWidthPt = 0
        Exit Function
    End If

    #If VBA7 Then
    Dim hDC As LongPtr
    Dim hFont As LongPtr
    Dim hOld As LongPtr
    #Else
    Dim hDC As Long
    Dim hFont As Long
    Dim hOld As Long
    #End If
    Dim sz As SIZE_STRUCT
    Dim tm As TEXTMETRIC_STRUCT
    Dim nDPI As Long
    Dim nHeightPx As Long
    Dim nWeight As Long

    FONT_TextWidthPt = -1
    nDPI = FONT_DPI()
    If bBold Then nWeight = 700 Else nWeight = 400

    On Error Resume Next
    hDC = fnGetDC(0)
    If hDC = 0 Then Exit Function

    nHeightPx = -CLng(dSizePt * nDPI / 72 + 0.5)
    hFont = fnCreateFont(nHeightPx, 0, 0, 0, nWeight, 0, 0, 0, DEFAULT_CHARSET, 0, 0, DEFAULT_QUALITY, 0, sFontName)
    If hFont <> 0 Then
        hOld = fnSelectObject(hDC, hFont)
        If fnGetTextExtentPoint32(hDC, sText, Len(sText), sz) <> 0 Then
            FONT_TextWidthPt = sz.cx * 72# / nDPI
        ElseIf fnGetTextMetrics(hDC, tm) <> 0 Then
            If tm.tmAveCharWidth > 0 Then
                FONT_TextWidthPt = (tm.tmAveCharWidth * Len(sText)) * 72# / nDPI
            End If
        End If
        fnSelectObject hDC, hOld
        fnDeleteObject hFont
    End If
    fnReleaseDC 0, hDC
    On Error GoTo 0
End Function

' FONT_CellPt   exact text cell height in points
Public Function FONT_CellPt(ByVal sFontName As String, ByVal dSizePt As Double, ByVal bBold As Boolean) As Double
    #If VBA7 Then
    Dim hDC As LongPtr
    Dim hFont As LongPtr
    Dim hOld As LongPtr
    #Else
    Dim hDC As Long
    Dim hFont As Long
    Dim hOld As Long
    #End If
    Dim tm As TEXTMETRIC_STRUCT
    Dim nDPI As Long
    Dim nWeight As Long

    FONT_CellPt = -1
    nDPI = FONT_DPI()
    If bBold Then nWeight = 700 Else nWeight = 400

    On Error Resume Next
    hDC = fnGetDC(0)
    If hDC = 0 Then Exit Function

    hFont = fnCreateFont(-CLng(dSizePt * nDPI / 72 + 0.5), 0, 0, 0, nWeight, 0, 0, 0, DEFAULT_CHARSET, 0, 0, DEFAULT_QUALITY, 0, sFontName)
    If hFont <> 0 Then
        hOld = fnSelectObject(hDC, hFont)
        If fnGetTextMetrics(hDC, tm) <> 0 Then
            FONT_CellPt = (tm.tmHeight + tm.tmExternalLeading) * 72# / nDPI
        End If
        fnSelectObject hDC, hOld
        fnDeleteObject hFont
    End If
    fnReleaseDC 0, hDC
    On Error GoTo 0
End Function

' FONT_CeilPxPt   round a point value UP to the next whole device pixel
Public Function FONT_CeilPxPt(ByVal dPt As Double) As Double
    Dim nPx As Long
    nPx = -Int(-(dPt * FONT_DPI() / 72#))
    If nPx < 1 Then nPx = 1
    FONT_CeilPxPt = nPx * 72# / FONT_DPI()
End Function

Private Function fnSystemMessageFont() As String
    Dim ncm As NONCLIENTMETRICSW
    Dim i As Long
    Dim s As String

    On Error Resume Next
    ncm.cbSize = LenB(ncm)
    If fnSystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, ncm, 0) <> 0 Then
        For i = 0 To 31
            If ncm.lfMessageFont.lfFaceName(i) = 0 Then Exit For
            s = s & ChrW$(ncm.lfMessageFont.lfFaceName(i))
        Next i
    End If
    On Error GoTo 0

    fnSystemMessageFont = s
End Function

' END: modFont v9.25.20
