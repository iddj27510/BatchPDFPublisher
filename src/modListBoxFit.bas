Attribute VB_Name = "modListBoxFit"
Option Explicit
' Module    : modListBoxFit
' Purpose   : Universal ListBox geometry — column widths, row height,
'             header alignment, form centering, interactive calibration,
'             built-in self-test, zero unnamed magic.
' Reference : Appendix A.14
' Version   : 9.25.31
'
' v9.25.31 : ADD — AutoFitTruncateForCellKeepFilename. Cuva PUNO ime
'             fajla + ekstenziju; skracuje FOLDER deo putanje. Prioritet:
'             (1) cela putanja, (2) "…\<puno ime>", (3) "<puno ime>",
'             (4) retko — skraceno ime sa ocuvanom ekstenzijom.
' v9.25.30 : AutoFitTruncateForCellKeepExt.
' v9.25.26 : HEIGHT DRIFT FIX.
' v9.25.25 : Payload-only columns hidden.
' v9.25.24 : AutoFitProbeRowHeight.
' v9.25.23 : Phase 2/3 rewrite.
' v9.25.22 : CELL_PADDING_PX raised 6 -> 12; AutoFitTruncateForCell added.
' v9.25.21 : OPTION B — px-grid geometry.
' v9.25.20 : AutoFitWidestOf replaced with GDI-based version.
'
' PUBLIC API
'   AutoFitListBox              column widths, header alignment, priority hiding
'   AutoFitListBoxHeight        row-exact height, never above designer height
'   AutoFitProbeRowHeight       one-shot true FM20 row-height measurement
'   AutoFitCenterForm           center form over Inventor main window
'   AutoFitWidestOf             GDI-widest string from a candidate array
'   AutoFitTruncateForCell      ellipsis-truncate a string to a column width
'   AutoFitTruncateForCellKeepExt       isto, ali cuva ekstenziju
'   AutoFitTruncateForCellKeepFilename  isto, ali cuva PUNO ime + ext
'   AutoFitClearCache           FULL reset (calibration/self-test only)
'   AutoFitCalibrateCellPadding interactive CELL_PADDING_PX calibration
'   AutoFitCalibrateRowPadding  interactive ROW_PADDING_PX calibration
'   AutoFitSelfTest             9 assertion groups
'   AutoFitReportBitness        prints which constant branch is ACTIVE

#If VBA7 And Win64 Then
    Private Const CELL_PADDING_PX   As Long = 24
    Private Const ROW_PADDING_PX    As Long = 8
    Private Const SAFETY_PADDING_PX As Long = 2
#Else
    Private Const CELL_PADDING_PX   As Long = 24
    Private Const ROW_PADDING_PX    As Long = 8
    Private Const SAFETY_PADDING_PX As Long = 2
#End If

Private Const POINTS_PER_INCH    As Single = 72#
Private Const DPI_AT_CALIBRATION As Single = 96#
Private Const MAX_CACHE          As Long = 512
Private Const WIDEST_CHAR_LIMIT  As Long = 32
Private Const AUTO_MIN_CHARS     As Long = 5
Private Const FW_NORMAL_C        As Long = 400
Private Const FW_BOLD_C          As Long = 700
Private Const AVG_CHAR_PX        As Single = 6.5
Private Const DEGRADE_MIN_COL_PX As Long = 16
Private Const AVAIL_MIN_PX       As Long = 8
Private Const RESIDUAL_TOL_PX    As Single = 0.5
Private Const CAL_STEP_PX        As Long = 1
Private Const MARGIN_INF         As Single = 1E+30
Private Const FALLBACK_EDGE_PX   As Long = 2
Private Const FALLBACK_SCROLL_PX As Long = 16
Private Const PROBE_ROWS         As Long = 4
Private Const PROBE_MIN_ROW_PX   As Single = 8
Private Const PROBE_MAX_ROW_PX   As Single = 128
Private Const MAX_ROWS_UNCLAMPED As Long = 2147483647

Private Const LF_TOTAL_BYTES   As Long = 92
Private Const LF_FACE_OFFSET   As Long = 28
Private Const LF_FACE_MAXB     As Long = 62
Private Const LF_W_OFFSET      As Long = 16
Private Const LF_I_OFFSET      As Long = 20
Private Const LF_CS_OFFSET     As Long = 23
Private Const LF_FACEBUF_CHARS As Long = 64

Private Const SM_CXVSCROLL  As Long = 2
Private Const SM_CYHSCROLL  As Long = 3
Private Const SM_CXEDGE     As Long = 45
Private Const SM_CYEDGE     As Long = 46
Private Const LOGPIXELSX    As Long = 88
Private Const DEFAULT_CHARSET As Long = 1

#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function CreateFontIndirectW Lib "gdi32" (ByRef lpLogFont As Any) As LongPtr
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function GetTextExtentPoint32W Lib "gdi32" (ByVal hDC As LongPtr, ByVal lpString As LongPtr, ByVal c As Long, ByRef lpSize As SIZEAPI) As Long
    Private Declare PtrSafe Function GetTextFaceW Lib "gdi32" (ByVal hDC As LongPtr, ByVal nCount As Long, ByVal lpFaceName As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function GetDpiForSystem Lib "user32" () As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Private Declare Function CreateFontIndirectW Lib "gdi32" (ByRef lpLogFont As Any) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hDC As Long, ByVal lpString As Long, ByVal c As Long, ByRef lpSize As SIZEAPI) As Long
    Private Declare Function GetTextFaceW Lib "gdi32" (ByVal hDC As Long, ByVal nCount As Long, ByVal lpFaceName As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#End If

Private Type SIZEAPI
    cx As Long
    cy As Long
End Type

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Type AutoFitColumn
    widestText As String
    headerText As String
    elastic    As Boolean
    priority   As Long
    minChars   As Long
End Type

Public Type AutoFitResult
    success     As Boolean
    degraded    As Boolean
    dpiUsed     As Single
    logMessages As String
End Type

Private m_cacheKeys()   As String
Private m_cacheWidths() As Single
Private m_cacheCount    As Long
Private m_cacheSig      As String
Private m_dpiCached     As Single
Private m_probeSig      As String
Private m_probeW        As Long
Private m_rowHPxCache   As Single
Private m_rowHDpi       As Single
Private m_designHPx     As Single
Private m_designDpi     As Single
Private m_fitN          As Long
Private m_fitDpi        As Single

Private Function DetectDpi() As Single
    If m_dpiCached > 0# Then
        DetectDpi = m_dpiCached
        Exit Function
    End If
    Dim dpiVal As Long
    dpiVal = 0
#If VBA7 Then
    On Error Resume Next
    dpiVal = GetDpiForSystem()
    On Error GoTo 0
#End If
    If dpiVal = 0 Then
#If VBA7 Then
        Dim hDC As LongPtr
#Else
        Dim hDC As Long
#End If
        hDC = GetDC(0)
        If hDC <> 0 Then
            dpiVal = GetDeviceCaps(hDC, LOGPIXELSX)
            ReleaseDC 0, hDC
        End If
    End If
    If dpiVal <= 0 Then dpiVal = CLng(DPI_AT_CALIBRATION)
    m_dpiCached = CSng(dpiVal)
    DetectDpi = m_dpiCached
End Function

Private Sub ComputeMarginsPx(ByRef borderPx As Long, _
                             ByRef scrollPx As Long, _
                             ByRef safetyPx As Long)
    Dim edgePx As Long: edgePx = GetSystemMetrics(SM_CXEDGE)
    Dim vScrPx As Long: vScrPx = GetSystemMetrics(SM_CXVSCROLL)
    If edgePx <= 0 Then edgePx = FALLBACK_EDGE_PX
    If vScrPx <= 0 Then vScrPx = FALLBACK_SCROLL_PX
    borderPx = edgePx * 2
    scrollPx = vScrPx
    safetyPx = SAFETY_PADDING_PX
End Sub

Private Function PtToPxFloor(ByVal pt As Single, ByVal dpi As Single) As Long
    PtToPxFloor = Int(pt * dpi / POINTS_PER_INCH)
End Function

Private Function PxToPtCeil(ByVal px As Single, ByVal dpi As Single) As Single
    If px <= 0# Then
        PxToPtCeil = 0#
        Exit Function
    End If
    Dim nPx As Long
    nPx = -Int(-px)
    Dim ptExact As Double
    ptExact = CDbl(nPx) * CDbl(POINTS_PER_INCH) / CDbl(dpi)
    Dim emit As Double
    emit = -Int(-ptExact * 100#) / 100#
    Do While CSng(emit) * dpi / POINTS_PER_INCH < CSng(nPx)
        emit = emit + 0.01
    Loop
    PxToPtCeil = CSng(emit)
End Function

Private Sub BuildLogFont(ByRef lf() As Byte, _
                         ByVal fontName As String, _
                         ByVal fontSizePt As Single, _
                         ByVal fontBold As Boolean, _
                         ByVal fontItalic As Boolean, _
                         ByVal dpi As Single)
    Dim i As Long
    For i = 0 To LF_TOTAL_BYTES - 1
        lf(i) = 0
    Next i
    Dim lfHeight As Long
    lfHeight = -CLng(fontSizePt * dpi / POINTS_PER_INCH)
    CopyMemory lf(0), lfHeight, 4
    Dim lfWeight As Long
    If fontBold Then lfWeight = FW_BOLD_C Else lfWeight = FW_NORMAL_C
    CopyMemory lf(LF_W_OFFSET), lfWeight, 4
    If fontItalic Then lf(LF_I_OFFSET) = 1
    lf(LF_CS_OFFSET) = DEFAULT_CHARSET
    Dim nameBytes As Long
    nameBytes = LenB(fontName)
    If nameBytes > LF_FACE_MAXB Then nameBytes = LF_FACE_MAXB
    If nameBytes > 0 Then
        CopyMemory lf(LF_FACE_OFFSET), ByVal StrPtr(fontName), nameBytes
    End If
End Sub

#If VBA7 Then
Private Function AcquireGdiFont(ByRef hDC_out As LongPtr, _
                                ByRef hFont_out As LongPtr, _
                                ByRef hOldFont_out As LongPtr, _
                                ByVal fontName As String, _
                                ByVal fontSizePt As Single, _
                                ByVal fontBold As Boolean, _
                                ByVal fontItalic As Boolean, _
                                ByVal dpi As Single, _
                                ByRef logMsg As String) As Boolean
#Else
Private Function AcquireGdiFont(ByRef hDC_out As Long, _
                                ByRef hFont_out As Long, _
                                ByRef hOldFont_out As Long, _
                                ByVal fontName As String, _
                                ByVal fontSizePt As Single, _
                                ByVal fontBold As Boolean, _
                                ByVal fontItalic As Boolean, _
                                ByVal dpi As Single, _
                                ByRef logMsg As String) As Boolean
#End If
    AcquireGdiFont = False
    hDC_out = 0
    hFont_out = 0
    hOldFont_out = 0
    hDC_out = GetDC(0)
    If hDC_out = 0 Then
        logMsg = logMsg & "ERROR: GetDC(0) returned 0" & vbLf
        Exit Function
    End If
    Dim lf(0 To LF_TOTAL_BYTES - 1) As Byte
    BuildLogFont lf, fontName, fontSizePt, fontBold, fontItalic, dpi
    hFont_out = CreateFontIndirectW(lf(0))
    If hFont_out = 0 Then
        logMsg = logMsg & "ERROR: CreateFontIndirectW returned 0" & vbLf
        ReleaseDC 0, hDC_out
        hDC_out = 0
        Exit Function
    End If
    hOldFont_out = SelectObject(hDC_out, hFont_out)
    Dim actualFace As String
    actualFace = LoadedFaceName(hDC_out)
    If Len(actualFace) > 0 Then
        If StrComp(actualFace, fontName, vbTextCompare) <> 0 Then
            logMsg = logMsg & "WARNING: font substitution - requested '" & fontName & _
                     "', GDI selected '" & actualFace & "'" & vbLf
        End If
    End If
    AcquireGdiFont = True
End Function

#If VBA7 Then
Private Sub ReleaseGdiFont(ByRef hDC As LongPtr, _
                           ByRef hFont As LongPtr, _
                           ByRef hOldFont As LongPtr)
#Else
Private Sub ReleaseGdiFont(ByRef hDC As Long, _
                           ByRef hFont As Long, _
                           ByRef hOldFont As Long)
#End If
    If hDC <> 0 Then
        If hOldFont <> 0 Then SelectObject hDC, hOldFont
        If hFont <> 0 Then DeleteObject hFont
        ReleaseDC 0, hDC
    End If
    hDC = 0
    hFont = 0
    hOldFont = 0
End Sub

#If VBA7 Then
Private Function LoadedFaceName(ByVal hDC As LongPtr) As String
#Else
Private Function LoadedFaceName(ByVal hDC As Long) As String
#End If
    Dim buf    As String: buf = Space$(LF_FACEBUF_CHARS)
    Dim nChars As Long:   nChars = GetTextFaceW(hDC, LF_FACEBUF_CHARS, StrPtr(buf))
    If nChars > 0 Then
        LoadedFaceName = Left$(buf, nChars - 1)
    Else
        LoadedFaceName = vbNullString
    End If
End Function

#If VBA7 Then
Private Function MeasureTextPx(ByVal hDC As LongPtr, ByVal txt As String) As Long
#Else
Private Function MeasureTextPx(ByVal hDC As Long, ByVal txt As String) As Long
#End If
    If Len(txt) = 0 Then Exit Function
    Dim sz As SIZEAPI
    If GetTextExtentPoint32W(hDC, StrPtr(txt), Len(txt), sz) <> 0 Then
        MeasureTextPx = sz.cx
    End If
End Function

#If VBA7 Then
Private Function MeasureTextHeightPx(ByVal hDC As LongPtr, ByVal txt As String) As Long
#Else
Private Function MeasureTextHeightPx(ByVal hDC As Long, ByVal txt As String) As Long
#End If
    If Len(txt) = 0 Then Exit Function
    Dim sz As SIZEAPI
    If GetTextExtentPoint32W(hDC, StrPtr(txt), Len(txt), sz) <> 0 Then
        MeasureTextHeightPx = sz.cy
    End If
End Function

Private Sub ClearWidthCache()
    ReDim m_cacheKeys(0 To MAX_CACHE - 1) As String
    ReDim m_cacheWidths(0 To MAX_CACHE - 1) As Single
    m_cacheCount = 0
    m_cacheSig = vbNullString
    m_probeSig = vbNullString
    m_probeW = 0
End Sub

#If VBA7 Then
Private Function GetCachedWidth(ByVal hDC As LongPtr, _
                                ByVal txt As String, _
                                ByVal fontName As String, _
                                ByVal fontSize As Single, _
                                ByVal fontBold As Boolean, _
                                ByVal dpi As Single) As Single
#Else
Private Function GetCachedWidth(ByVal hDC As Long, _
                                ByVal txt As String, _
                                ByVal fontName As String, _
                                ByVal fontSize As Single, _
                                ByVal fontBold As Boolean, _
                                ByVal dpi As Single) As Single
#End If
    Dim sig As String
    sig = fontName & "|" & CStr(fontSize) & "|" & CStr(fontBold) & "|" & CStr(dpi)
    If sig <> m_cacheSig Then
        ClearWidthCache
        m_cacheSig = sig
    End If
    Dim key As String: key = sig & "|" & txt
    Dim i   As Long
    For i = 0 To m_cacheCount - 1
        If m_cacheKeys(i) = key Then
            GetCachedWidth = m_cacheWidths(i)
            Exit Function
        End If
    Next i
    Dim w As Long: w = MeasureTextPx(hDC, txt)
    If w = 0 And Len(txt) > 0 Then Exit Function
    If m_cacheCount < MAX_CACHE Then
        m_cacheKeys(m_cacheCount) = key
        m_cacheWidths(m_cacheCount) = CSng(w)
        m_cacheCount = m_cacheCount + 1
    End If
    GetCachedWidth = CSng(w)
End Function

#If VBA7 Then
Private Function WidestCharPx(ByVal hDC As LongPtr, _
                              ByVal txt As String, _
                              ByVal fontSig As String) As Long
#Else
Private Function WidestCharPx(ByVal hDC As Long, _
                              ByVal txt As String, _
                              ByVal fontSig As String) As Long
#End If
    If fontSig <> m_probeSig Then
        m_probeW = 0
        m_probeSig = fontSig
        Dim probe As String: probe = "0123456789MWw ./-_:"
        Dim j As Long
        For j = 1 To Len(probe)
            Dim pw As Long: pw = MeasureTextPx(hDC, mid$(probe, j, 1))
            If pw > m_probeW Then m_probeW = pw
        Next j
    End If
    Dim maxW As Long: maxW = m_probeW
    If Len(txt) = 0 Then WidestCharPx = maxW: Exit Function
    Dim measured    As String: measured = vbNullString
    Dim uniqueCount As Long:   uniqueCount = 0
    Dim i  As Long
    Dim ch As String
    For i = 1 To Len(txt)
        If uniqueCount >= WIDEST_CHAR_LIMIT Then Exit For
        ch = mid$(txt, i, 1)
        If InStr(1, measured, ch, vbBinaryCompare) = 0 Then
            measured = measured & ch
            uniqueCount = uniqueCount + 1
            Dim w As Long: w = MeasureTextPx(hDC, ch)
            If w > maxW Then maxW = w
        End If
    Next i
    WidestCharPx = maxW
End Function

Private Function PtToInvariant(ByVal value As Single) As String
    If value < 0# Then value = 0#
    Dim total     As Long: total = CLng(Int(value * 100# + 0.5))
    Dim wholePart As Long: wholePart = total \ 100
    Dim fracPart  As Long: fracPart = total - wholePart * 100
    If fracPart < 0 Then fracPart = 0
    If fracPart = 0 Then
        PtToInvariant = CStr(wholePart)
    ElseIf fracPart Mod 10 = 0 Then
        PtToInvariant = CStr(wholePart) & "." & CStr(fracPart \ 10)
    Else
        Dim fracStr As String: fracStr = CStr(fracPart)
        If Len(fracStr) < 2 Then fracStr = "0" & fracStr
        PtToInvariant = CStr(wholePart) & "." & fracStr
    End If
End Function

#If VBA7 Then
Private Function TruncateText(ByVal hDC As LongPtr, _
                              ByVal txt As String, _
                              ByVal maxWidthPx As Long) As String
#Else
Private Function TruncateText(ByVal hDC As Long, _
                              ByVal txt As String, _
                              ByVal maxWidthPx As Long) As String
#End If
    Dim ellipsis As String: ellipsis = ChrW$(&H2026)
    If Len(txt) = 0 Then TruncateText = txt: Exit Function
    Dim fullW As Long: fullW = MeasureTextPx(hDC, txt)
    If fullW > 0 And fullW <= maxWidthPx Then
        TruncateText = txt
        Exit Function
    End If
    Dim ellW As Long: ellW = MeasureTextPx(hDC, ellipsis)
    If ellW = 0 Then ellW = MeasureTextPx(hDC, "...")
    If ellW = 0 Then
        Dim estChars As Long
        If fullW > 0 Then
            estChars = CLng(Len(txt) * (CDbl(maxWidthPx) / CDbl(fullW)))
        Else
            estChars = CLng(maxWidthPx / AVG_CHAR_PX)
        End If
        If estChars < 1 Then estChars = 1
        If estChars >= Len(txt) Then estChars = Len(txt) - 1
        TruncateText = Left$(txt, estChars) & ellipsis
        Exit Function
    End If
    Dim targetW As Long: targetW = maxWidthPx - ellW
    If targetW <= 0 Then TruncateText = ellipsis: Exit Function
    Dim lo   As Long: lo = 1
    Dim hi   As Long: hi = Len(txt)
    Dim best As Long: best = 0
    Do While lo <= hi
        Dim midIdx As Long: midIdx = (lo + hi) \ 2
        Dim testW  As Long: testW = MeasureTextPx(hDC, Left$(txt, midIdx))
        If testW = 0 Then
            best = CLng(CDbl(midIdx) * CDbl(targetW) / CDbl(maxWidthPx))
            If best < 1 Then best = 1
            Exit Do
        End If
        If testW <= targetW Then
            best = midIdx
            lo = midIdx + 1
        Else
            hi = midIdx - 1
        End If
    Loop
    If best >= Len(txt) Then
        TruncateText = txt
    ElseIf best > 0 Then
        TruncateText = Left$(txt, best) & ellipsis
    Else
        TruncateText = ellipsis
    End If
End Function

Private Function FloorRowCount(ByVal innerH As Single, _
                               ByVal rowH As Single, _
                               ByVal maxRows As Long) As Long
    If maxRows < 1 Then maxRows = 1
    If rowH <= 0# Or innerH <= 0# Then FloorRowCount = 1: Exit Function
    Dim n As Long: n = CLng(Int(innerH / rowH))
    If n < 1 Then n = 1
    If n > maxRows Then n = maxRows
    FloorRowCount = n
End Function

Private Function IsArrAllocated(ByRef baseIdx As Long, ByRef count As Long, _
                                ByVal lb As Long, ByVal ubOk As Boolean, _
                                ByVal ub As Long) As Boolean
    If ubOk Then
        count = ub - lb + 1
        If count > 0 Then
            baseIdx = lb
            IsArrAllocated = True
            Exit Function
        End If
    End If
    baseIdx = 0
    count = 0
    IsArrAllocated = False
End Function

Private Sub DegradeMode(ByVal lst As Object, _
                        ByVal nCols As Long, _
                        ByRef res As AutoFitResult, _
                        ByRef outWidths() As Single, _
                        ByRef outHidden() As Boolean, _
                        ByVal reason As String)
    res.degraded = True
    res.logMessages = res.logMessages & "DEGRADE: " & reason & vbLf
    If nCols < 1 Then nCols = 1
    ReDim outWidths(0 To nCols - 1)
    ReDim outHidden(0 To nCols - 1)
    Dim dpi As Single: dpi = DetectDpi()
    Dim borderPx As Long, scrollPx As Long, safetyPx As Long
    ComputeMarginsPx borderPx, scrollPx, safetyPx
    Dim usablePx As Single
    usablePx = PtToPxFloor(lst.Width, dpi) - borderPx - scrollPx - safetyPx
    If usablePx < nCols * DEGRADE_MIN_COL_PX Then usablePx = nCols * DEGRADE_MIN_COL_PX
    Dim eachPx As Single: eachPx = usablePx / CSng(nCols)
    If eachPx < 1# Then eachPx = 1#
    Dim colStr As String: colStr = vbNullString
    Dim i      As Long
    For i = 0 To nCols - 1
        outWidths(i) = PxToPtCeil(eachPx, dpi)
        outHidden(i) = False
        If i > 0 Then colStr = colStr & ";"
        colStr = colStr & PtToInvariant(outWidths(i))
    Next i
    lst.ColumnWidths = colStr
    res.success = True
End Sub

Public Sub AutoFitClearCache()
    ClearWidthCache
    m_dpiCached = 0#
    m_rowHPxCache = 0#
    m_rowHDpi = 0#
    m_designHPx = 0#
    m_designDpi = 0#
    m_fitN = 0
    m_fitDpi = 0#
End Sub

Public Sub AutoFitReportBitness()
#If VBA7 And Win64 Then
    Debug.Print "AutoFitReportBitness: 64-bit VBA host (Win64=True)"
    Debug.Print "  Active constants: Win64 branch"
    Debug.Print "  CELL_PADDING_PX   = " & CELL_PADDING_PX
    Debug.Print "  ROW_PADDING_PX    = " & ROW_PADDING_PX & " (fallback only)"
    Debug.Print "  SAFETY_PADDING_PX = " & SAFETY_PADDING_PX
#Else
    Debug.Print "AutoFitReportBitness: 32-bit VBA host (Win64=False)"
    Debug.Print "  (expected for Inventor VBA: out-of-process 32-bit host)"
    Debug.Print "  Active constants: x86 branch"
    Debug.Print "  CELL_PADDING_PX   = " & CELL_PADDING_PX
    Debug.Print "  ROW_PADDING_PX    = " & ROW_PADDING_PX & " (fallback only)"
    Debug.Print "  SAFETY_PADDING_PX = " & SAFETY_PADDING_PX
#End If
    Debug.Print "  DPI on this machine = " & DetectDpi()
End Sub

Public Sub AutoFitProbeRowHeight(ByVal lst As Object)
    If lst Is Nothing Then Exit Sub
    If lst.ListCount > 0 Then Exit Sub
    Dim dpi As Single: dpi = DetectDpi()
    If m_rowHPxCache > 0# And m_rowHDpi = dpi Then Exit Sub
    Dim edgeYPx As Long: edgeYPx = GetSystemMetrics(SM_CYEDGE)
    If edgeYPx <= 0 Then edgeYPx = FALLBACK_EDGE_PX
    Dim borderHPx As Long: borderHPx = edgeYPx * 2
    Dim origH As Single:         origH = lst.Height
    Dim origIntegral As Boolean: origIntegral = lst.IntegralHeight
    Dim flipClient As Long: flipClient = 0
    On Error GoTo ProbeFail
    lst.IntegralHeight = False
    lst.Clear
    Dim k As Long
    For k = 1 To PROBE_ROWS
        lst.AddItem "O"
    Next k
    Dim curPx As Long
    For curPx = 1 To PROBE_ROWS * 256
        lst.Height = PxToPtCeil(CSng(borderHPx + curPx), dpi)
        lst.ListIndex = PROBE_ROWS - 1
        If lst.TopIndex = 0 Then
            flipClient = curPx
            Exit For
        End If
    Next curPx
ProbeFail:
    On Error Resume Next
    lst.Clear
    lst.Height = origH
    lst.IntegralHeight = origIntegral
    lst.ListIndex = -1
    On Error GoTo 0
    If flipClient > 0 Then
        Dim rowH As Single: rowH = CSng(flipClient) / CSng(PROBE_ROWS)
        If rowH >= PROBE_MIN_ROW_PX And rowH <= PROBE_MAX_ROW_PX Then
            m_rowHPxCache = rowH
            m_rowHDpi = dpi
        End If
    End If
End Sub

Public Function AutoFitWidestOf(ByVal lst As Object, ByVal candidates As Variant) As String
    Dim bestTxt As String
    Dim bestLen As Long
    Dim i       As Long
    On Error GoTo LenFallback
#If VBA7 Then
    Dim hDC      As LongPtr
    Dim hFont    As LongPtr
    Dim hOldFont As LongPtr
#Else
    Dim hDC      As Long
    Dim hFont    As Long
    Dim hOldFont As Long
#End If
    Dim logMsg As String
    If Not AcquireGdiFont(hDC, hFont, hOldFont, lst.Font.Name, lst.Font.Size, _
                          lst.Font.bOld, lst.Font.Italic, DetectDpi(), logMsg) Then
        GoTo LenFallback
    End If
    On Error GoTo GdiCleanup
    Dim bestW As Long
    For i = LBound(candidates) To UBound(candidates)
        Dim w As Long: w = MeasureTextPx(hDC, CStr(candidates(i)))
        If w > bestW Then
            bestW = w
            bestTxt = CStr(candidates(i))
        End If
    Next i
    ReleaseGdiFont hDC, hFont, hOldFont
    If Len(bestTxt) > 0 Then
        AutoFitWidestOf = bestTxt
        Exit Function
    End If
    GoTo LenFallback

GdiCleanup:
    ReleaseGdiFont hDC, hFont, hOldFont
    Resume LenFallback

LenFallback:
    On Error Resume Next
    For i = LBound(candidates) To UBound(candidates)
        If Len(CStr(candidates(i))) > bestLen Then
            bestLen = Len(CStr(candidates(i)))
            bestTxt = CStr(candidates(i))
        End If
    Next i
    AutoFitWidestOf = bestTxt
End Function

' ---------------------------------------------------------------------------
' AutoFitTruncateForCell   v9.25.22.
' ---------------------------------------------------------------------------
Public Function AutoFitTruncateForCell(ByVal lst As Object, _
                                       ByVal txt As String, _
                                       ByVal colWidthPt As Single) As String
    AutoFitTruncateForCell = txt
    If lst Is Nothing Then Exit Function
    If Len(txt) = 0 Then Exit Function
    If colWidthPt <= 0# Then Exit Function
    Dim dpi As Single: dpi = DetectDpi()
#If VBA7 Then
    Dim hDC      As LongPtr
    Dim hFont    As LongPtr
    Dim hOldFont As LongPtr
#Else
    Dim hDC      As Long
    Dim hFont    As Long
    Dim hOldFont As Long
#End If
    Dim logMsg As String
    If Not AcquireGdiFont(hDC, hFont, hOldFont, lst.Font.Name, lst.Font.Size, _
                          lst.Font.bOld, lst.Font.Italic, dpi, logMsg) Then
        Exit Function
    End If
    Dim maxPx As Long
    maxPx = CLng(colWidthPt * dpi / POINTS_PER_INCH) - CELL_PADDING_PX
    If maxPx > 0 Then
        AutoFitTruncateForCell = TruncateText(hDC, txt, maxPx)
    End If
    ReleaseGdiFont hDC, hFont, hOldFont
End Function

' ---------------------------------------------------------------------------
' AutoFitTruncateForCellKeepExt   v9.25.30
'   Cuva ekstenziju na kraju reda. "…" ide pre ekstenzije.
' ---------------------------------------------------------------------------
Public Function AutoFitTruncateForCellKeepExt(ByVal lst As Object, _
                                              ByVal txt As String, _
                                              ByVal colWidthPt As Single) As String
    AutoFitTruncateForCellKeepExt = txt
    If lst Is Nothing Then Exit Function
    If Len(txt) = 0 Then Exit Function
    If colWidthPt <= 0# Then Exit Function

    Dim dpi As Single: dpi = DetectDpi()

#If VBA7 Then
    Dim hDC      As LongPtr
    Dim hFont    As LongPtr
    Dim hOldFont As LongPtr
#Else
    Dim hDC      As Long
    Dim hFont    As Long
    Dim hOldFont As Long
#End If

    Dim logMsg As String
    If Not AcquireGdiFont(hDC, hFont, hOldFont, lst.Font.Name, lst.Font.Size, _
                          lst.Font.bOld, lst.Font.Italic, dpi, logMsg) Then
        Exit Function
    End If

    Dim maxPx As Long
    maxPx = CLng(colWidthPt * dpi / POINTS_PER_INCH) - CELL_PADDING_PX
    If maxPx <= 0 Then
        ReleaseGdiFont hDC, hFont, hOldFont
        Exit Function
    End If

    Dim fullW As Long: fullW = MeasureTextPx(hDC, txt)
    If fullW > 0 And fullW <= maxPx Then
        ReleaseGdiFont hDC, hFont, hOldFont
        Exit Function
    End If

    Dim nDot   As Long: nDot = InStrRev(txt, ".")
    Dim nSlash As Long: nSlash = InStrRev(txt, "\")

    Dim sExt  As String
    Dim sBase As String
    If nDot > nSlash And nDot < Len(txt) Then
        sExt = mid$(txt, nDot)
        sBase = Left$(txt, nDot - 1)
    Else
        sExt = ""
        sBase = txt
    End If

    Dim ellipsis As String: ellipsis = ChrW$(&H2026)
    Dim tailW As Long: tailW = MeasureTextPx(hDC, ellipsis & sExt)
    Dim targetBaseW As Long: targetBaseW = maxPx - tailW

    If targetBaseW <= 0 Then
        AutoFitTruncateForCellKeepExt = TruncateText(hDC, txt, maxPx)
        ReleaseGdiFont hDC, hFont, hOldFont
        Exit Function
    End If

    Dim lo   As Long: lo = 1
    Dim hi   As Long: hi = Len(sBase)
    Dim best As Long: best = 0
    Do While lo <= hi
        Dim midIdx As Long: midIdx = (lo + hi) \ 2
        Dim testW  As Long: testW = MeasureTextPx(hDC, Left$(sBase, midIdx))
        If testW > 0 And testW <= targetBaseW Then
            best = midIdx
            lo = midIdx + 1
        Else
            hi = midIdx - 1
        End If
    Loop

    If best >= Len(sBase) Then
        AutoFitTruncateForCellKeepExt = txt
    ElseIf best > 0 Then
        AutoFitTruncateForCellKeepExt = Left$(sBase, best) & ellipsis & sExt
    Else
        AutoFitTruncateForCellKeepExt = ellipsis & sExt
    End If

    ReleaseGdiFont hDC, hFont, hOldFont
End Function

' ---------------------------------------------------------------------------
' AutoFitTruncateForCellKeepFilename   v9.25.31
'   Cuva PUNO ime fajla + ekstenziju. Skracuje FOLDER deo putanje.
'
'   Prioritet prikaza:
'     1. Cela putanja ako staje                -> vrati txt
'     2. "…\<puno ime fajla>" ako staje        -> skracen folder
'     3. "<puno ime fajla>" ako staje          -> bez foldera
'     4. Skraceno ime sa ocuvanom ekstenzijom  -> retko
' ---------------------------------------------------------------------------
Public Function AutoFitTruncateForCellKeepFilename(ByVal lst As Object, _
                                                   ByVal txt As String, _
                                                   ByVal colWidthPt As Single) As String
    AutoFitTruncateForCellKeepFilename = txt
    If lst Is Nothing Then Exit Function
    If Len(txt) = 0 Then Exit Function
    If colWidthPt <= 0# Then Exit Function

    Dim dpi As Single: dpi = DetectDpi()

#If VBA7 Then
    Dim hDC      As LongPtr
    Dim hFont    As LongPtr
    Dim hOldFont As LongPtr
#Else
    Dim hDC      As Long
    Dim hFont    As Long
    Dim hOldFont As Long
#End If

    Dim logMsg As String
    If Not AcquireGdiFont(hDC, hFont, hOldFont, lst.Font.Name, lst.Font.Size, _
                          lst.Font.bOld, lst.Font.Italic, dpi, logMsg) Then
        Exit Function
    End If

    Dim maxPx As Long
    maxPx = CLng(colWidthPt * dpi / POINTS_PER_INCH) - CELL_PADDING_PX
    If maxPx <= 0 Then
        ReleaseGdiFont hDC, hFont, hOldFont
        Exit Function
    End If

    ' 1. Cela putanja staje?
    If MeasureTextPx(hDC, txt) <= maxPx Then
        ReleaseGdiFont hDC, hFont, hOldFont
        Exit Function
    End If

    ' 2. Izdvoji ime fajla (sve posle poslednjeg backslasha).
    Dim nSlash As Long: nSlash = InStrRev(txt, "\")
    Dim sFile As String
    If nSlash > 0 And nSlash < Len(txt) Then
        sFile = mid$(txt, nSlash + 1)
    Else
        sFile = txt
    End If

    Dim ellipsis As String: ellipsis = ChrW$(&H2026)

    ' 3. "…\<puno ime>" staje?
    Dim sTry As String: sTry = ellipsis & "\" & sFile
    Dim wTry As Long: wTry = MeasureTextPx(hDC, sTry)
    If wTry > 0 And wTry <= maxPx Then
        AutoFitTruncateForCellKeepFilename = sTry
        ReleaseGdiFont hDC, hFont, hOldFont
        Exit Function
    End If

    ' 4. Samo ime fajla staje?
    Dim fileW As Long: fileW = MeasureTextPx(hDC, sFile)
    If fileW > 0 And fileW <= maxPx Then
        AutoFitTruncateForCellKeepFilename = sFile
        ReleaseGdiFont hDC, hFont, hOldFont
        Exit Function
    End If

    ' 5. Retko: i samo ime je predugacko. Skrati ime ali cuvaj ekstenziju.
    ReleaseGdiFont hDC, hFont, hOldFont
    AutoFitTruncateForCellKeepFilename = AutoFitTruncateForCellKeepExt(lst, sFile, colWidthPt)
End Function

Public Function AutoFitListBox(ByVal lst As Object, _
                               ByRef cfg() As AutoFitColumn, _
                               ByRef hdrs() As Object, _
                               ByRef res As AutoFitResult, _
                               ByRef outWidths() As Single, _
                               ByRef outHidden() As Boolean) As Boolean
    AutoFitListBox = False
    res.success = False
    res.degraded = False
    On Error GoTo FailSafe
    If lst Is Nothing Then
        res.logMessages = res.logMessages & "ERROR: lst Is Nothing" & vbLf
        Exit Function
    End If
    Dim hasCfg  As Boolean: hasCfg = False
    Dim hasHdrs As Boolean: hasHdrs = False
    Dim cfgBase As Long:    cfgBase = 0
    Dim hdrBase As Long:    hdrBase = 0
    Dim nCols   As Long:    nCols = 0
    Dim nHdrs   As Long:    nHdrs = 0
    Dim lbTmp   As Long
    Dim ubTmp   As Long
    On Error Resume Next
    lbTmp = LBound(cfg)
    If Err.Number = 0 Then
        ubTmp = UBound(cfg)
        If Err.Number = 0 Then hasCfg = IsArrAllocated(cfgBase, nCols, lbTmp, True, ubTmp)
    End If
    Err.Clear
    lbTmp = LBound(hdrs)
    If Err.Number = 0 Then
        ubTmp = UBound(hdrs)
        If Err.Number = 0 Then hasHdrs = IsArrAllocated(hdrBase, nHdrs, lbTmp, True, ubTmp)
    End If
    Err.Clear
    On Error GoTo FailSafe
    If Not hasCfg Then
        res.logMessages = res.logMessages & "ERROR: cfg is unallocated or empty" & vbLf
        Exit Function
    End If
    If hasHdrs And nHdrs <> nCols Then
        res.logMessages = res.logMessages & "WARNING: header count (" & nHdrs & _
                         ") <> column count (" & nCols & ")" & vbLf
        If nHdrs > nCols Then nHdrs = nCols
    End If
    Dim cc As Long
    On Error Resume Next
    cc = lst.ColumnCount
    If Err.Number = 0 Then
        If cc < nCols Then
            If lst.ListCount = 0 Then
                lst.ColumnCount = nCols
                res.logMessages = res.logMessages & "INFO: ColumnCount raised to " & _
                                 nCols & vbLf
            Else
                res.logMessages = res.logMessages & "WARNING: ColumnCount (" & cc & _
                                 ") < " & nCols & " - set BEFORE AddItem" & vbLf
            End If
        End If
    End If
    Err.Clear
    On Error GoTo FailSafe
    Dim dpi    As Single: dpi = DetectDpi()
    Dim pxToPt As Single: pxToPt = POINTS_PER_INCH / dpi
    res.dpiUsed = dpi
    Dim borderPx As Long, scrollPx As Long, safetyPx As Long
    ComputeMarginsPx borderPx, scrollPx, safetyPx
    Dim edgePx As Long: edgePx = GetSystemMetrics(SM_CXEDGE)
    If edgePx <= 0 Then edgePx = FALLBACK_EDGE_PX
    Dim borderInsetPt As Single: borderInsetPt = CSng(edgePx) * pxToPt
    res.logMessages = res.logMessages & "Margins: border=" & borderPx & _
                     " scroll=" & scrollPx & _
                     " safety=" & safetyPx & " px" & vbLf
    Dim cellFontName   As String:  cellFontName = lst.Font.Name
    Dim cellFontSize   As Single:  cellFontSize = lst.Font.Size
    Dim cellFontBold   As Boolean: cellFontBold = lst.Font.bOld
    Dim cellFontItalic As Boolean: cellFontItalic = lst.Font.Italic
    Dim hdrFontName    As String:  hdrFontName = cellFontName
    Dim hdrFontSize    As Single:  hdrFontSize = cellFontSize
    Dim hdrFontBold    As Boolean: hdrFontBold = cellFontBold
    Dim hdrFontItalic  As Boolean: hdrFontItalic = cellFontItalic
    Dim headerFontDiffers As Boolean: headerFontDiffers = False
    If hasHdrs Then
        Dim firstHdr As Object
        Set firstHdr = hdrs(hdrBase)
        If Not firstHdr Is Nothing Then
            On Error Resume Next
            hdrFontName = firstHdr.Font.Name
            hdrFontSize = firstHdr.Font.Size
            hdrFontBold = firstHdr.Font.bOld
            hdrFontItalic = firstHdr.Font.Italic
            Err.Clear
            On Error GoTo FailSafe
            headerFontDiffers = (hdrFontName <> cellFontName) Or _
                                (hdrFontSize <> cellFontSize) Or _
                                (hdrFontBold <> cellFontBold) Or _
                                (hdrFontItalic <> cellFontItalic)
            Dim hIdx As Long
            For hIdx = hdrBase + 1 To hdrBase + nHdrs - 1
                If Not hdrs(hIdx) Is Nothing Then
                    Dim chkName As String
                    Dim chkSize As Single
                    Dim chkBold As Boolean
                    On Error Resume Next
                    chkName = hdrs(hIdx).Font.Name
                    chkSize = hdrs(hIdx).Font.Size
                    chkBold = hdrs(hIdx).Font.bOld
                    Err.Clear
                    On Error GoTo FailSafe
                    If chkName <> hdrFontName Or chkSize <> hdrFontSize Or _
                       chkBold <> hdrFontBold Then
                        res.logMessages = res.logMessages & "WARNING: header " & _
                            (hIdx - hdrBase) & " font differs from header 0" & vbLf
                    End If
                End If
            Next hIdx
        End If
    End If
    If m_cacheCount = 0 Then
        If (Not Not m_cacheKeys) = 0 Then ClearWidthCache
    End If
#If VBA7 Then
    Dim hDC          As LongPtr
    Dim hFont        As LongPtr
    Dim hOldFont     As LongPtr
    Dim hDC_hdr      As LongPtr
    Dim hFont_hdr    As LongPtr
    Dim hOldFont_hdr As LongPtr
#Else
    Dim hDC          As Long
    Dim hFont        As Long
    Dim hOldFont     As Long
    Dim hDC_hdr      As Long
    Dim hFont_hdr    As Long
    Dim hOldFont_hdr As Long
#End If
    If Not AcquireGdiFont(hDC, hFont, hOldFont, cellFontName, cellFontSize, _
                          cellFontBold, cellFontItalic, dpi, res.logMessages) Then
        DegradeMode lst, nCols, res, outWidths, outHidden, "AcquireGdiFont failed (cell)"
        AutoFitListBox = True
        Exit Function
    End If
    Dim cellWidthsPx() As Single: ReDim cellWidthsPx(0 To nCols - 1)
    Dim maxCharW()     As Long:   ReDim maxCharW(0 To nCols - 1)
    Dim i      As Long
    Dim cfgIdx As Long
    Dim sig    As String
    sig = cellFontName & "|" & CStr(cellFontSize) & "|" & _
          CStr(cellFontBold) & "|" & CStr(dpi)
    For i = 0 To nCols - 1
        cfgIdx = cfgBase + i
        If Len(cfg(cfgIdx).widestText) > 0 Then
            cellWidthsPx(i) = GetCachedWidth(hDC, cfg(cfgIdx).widestText, _
                              cellFontName, cellFontSize, cellFontBold, dpi)
        End If
        maxCharW(i) = WidestCharPx(hDC, cfg(cfgIdx).widestText, sig)
    Next i
    Dim hdrWidthsPx() As Single: ReDim hdrWidthsPx(0 To nCols - 1)
    If headerFontDiffers Then
        ReleaseGdiFont hDC, hFont, hOldFont
        If AcquireGdiFont(hDC, hFont, hOldFont, hdrFontName, hdrFontSize, _
                          hdrFontBold, hdrFontItalic, dpi, res.logMessages) Then
            For i = 0 To nCols - 1
                If i < nHdrs And Len(cfg(cfgBase + i).headerText) > 0 Then
                    hdrWidthsPx(i) = GetCachedWidth(hDC, cfg(cfgBase + i).headerText, _
                                     hdrFontName, hdrFontSize, hdrFontBold, dpi)
                End If
            Next i
            ReleaseGdiFont hDC, hFont, hOldFont
        Else
            headerFontDiffers = False
            res.logMessages = res.logMessages & _
                "WARNING: header font acquisition failed; cell font used" & vbLf
        End If
        If Not AcquireGdiFont(hDC, hFont, hOldFont, cellFontName, cellFontSize, _
                              cellFontBold, cellFontItalic, dpi, res.logMessages) Then
            DegradeMode lst, nCols, res, outWidths, outHidden, "re-acquire cell font failed"
            AutoFitListBox = True
            Exit Function
        End If
    ElseIf hasHdrs Then
        For i = 0 To nCols - 1
            If i < nHdrs And Len(cfg(cfgBase + i).headerText) > 0 Then
                hdrWidthsPx(i) = GetCachedWidth(hDC, cfg(cfgBase + i).headerText, _
                                 cellFontName, cellFontSize, cellFontBold, dpi)
            End If
        Next i
    End If
    Dim cellPadPx As Single: cellPadPx = CSng(CELL_PADDING_PX)
    Dim idealWidths() As Single: ReDim idealWidths(0 To nCols - 1)
    Dim minWidths()   As Single: ReDim minWidths(0 To nCols - 1)
    Dim wPx()         As Single: ReDim wPx(0 To nCols - 1)
    Dim emitPx()      As Long:   ReDim emitPx(0 To nCols - 1)
    ReDim outWidths(0 To nCols - 1)
    ReDim outHidden(0 To nCols - 1)
    Dim elasticIdx As Long:   elasticIdx = -1
    Dim sumMin     As Single: sumMin = 0#
    For i = 0 To nCols - 1
        cfgIdx = cfgBase + i
        If Len(cfg(cfgIdx).widestText) = 0 And Len(cfg(cfgIdx).headerText) = 0 Then
            outHidden(i) = True
            idealWidths(i) = 0#
            minWidths(i) = 0#
        Else
            Dim idealPx As Single
            idealPx = cellWidthsPx(i)
            If hdrWidthsPx(i) > idealPx Then idealPx = hdrWidthsPx(i)
            idealWidths(i) = idealPx + cellPadPx
            Dim mc As Long: mc = cfg(cfgIdx).minChars
            If mc <= 0 Then mc = AUTO_MIN_CHARS
            If maxCharW(i) > 0 Then
                minWidths(i) = CSng(maxCharW(i)) * CSng(mc) + cellPadPx
            Else
                minWidths(i) = CSng(mc) * AVG_CHAR_PX + cellPadPx
            End If
            If minWidths(i) > idealWidths(i) Then minWidths(i) = idealWidths(i)
            If cfg(cfgIdx).elastic Then elasticIdx = i
        End If
        sumMin = sumMin + minWidths(i)
    Next i
    Dim wAvail As Single
    wAvail = PtToPxFloor(lst.Width, dpi) - borderPx - scrollPx - safetyPx
    If wAvail < AVAIL_MIN_PX Then
        res.logMessages = res.logMessages & "AVAIL FLOOR: wAvail=" & _
                         PtToInvariant(wAvail) & "px clamped to " & _
                         AVAIL_MIN_PX & "px" & vbLf
        wAvail = AVAIL_MIN_PX
    End If
    If sumMin > wAvail Then
        Do While sumMin > wAvail
            Dim bestPri    As Long:   bestPri = 0
            Dim bestHide   As Long:   bestHide = -1
            Dim bestMargin As Single: bestMargin = MARGIN_INF
            For i = 0 To nCols - 1
                cfgIdx = cfgBase + i
                If cfg(cfgIdx).priority > 0 And Not outHidden(i) And _
                   Not cfg(cfgIdx).elastic Then
                    Dim thisMargin As Single
                    thisMargin = idealWidths(i) - minWidths(i)
                    If cfg(cfgIdx).priority > bestPri Then
                        bestPri = cfg(cfgIdx).priority
                        bestMargin = thisMargin
                        bestHide = i
                    ElseIf cfg(cfgIdx).priority = bestPri And thisMargin < bestMargin Then
                        bestMargin = thisMargin
                        bestHide = i
                    End If
                End If
            Next i
            If bestHide < 0 Then
                ReleaseGdiFont hDC, hFont, hOldFont
                DegradeMode lst, nCols, res, outWidths, outHidden, _
                    "minimum widths exceed available width"
                AutoFitListBox = True
                Exit Function
            End If
            outHidden(bestHide) = True
            sumMin = sumMin - minWidths(bestHide)
            res.logMessages = res.logMessages & "HIDE: column " & bestHide & _
                             " priority=" & bestPri & " margin=" & _
                             PtToInvariant(bestMargin) & "px" & vbLf
        Loop
    End If
    Dim j As Long
    For j = 0 To nCols - 1
        If outHidden(j) Then wPx(j) = 0# Else wPx(j) = idealWidths(j)
    Next j
    Dim bElasticVis As Boolean
    bElasticVis = (elasticIdx >= 0 And Not outHidden(elasticIdx))
    Dim sumFixed As Single: sumFixed = 0#
    For j = 0 To nCols - 1
        If Not outHidden(j) And j <> elasticIdx Then sumFixed = sumFixed + idealWidths(j)
    Next j
    Dim bShrinkFixed As Boolean
    bShrinkFixed = True
    If bElasticVis Then
        Dim remPx As Single
        remPx = wAvail - sumFixed
        If remPx >= minWidths(elasticIdx) Then
            wPx(elasticIdx) = remPx
            bShrinkFixed = False
        Else
            wPx(elasticIdx) = minWidths(elasticIdx)
        End If
    End If
    If bShrinkFixed Then
        Dim targetPx As Single
        targetPx = wAvail
        If bElasticVis Then targetPx = wAvail - minWidths(elasticIdx)
        Dim sumVisFixed As Single: sumVisFixed = 0#
        For j = 0 To nCols - 1
            If Not outHidden(j) And j <> elasticIdx Then sumVisFixed = sumVisFixed + wPx(j)
        Next j
        If sumVisFixed > targetPx Then
            Dim locked()  As Boolean: ReDim locked(0 To nCols - 1)
            Dim iteration As Long
            For iteration = 1 To nCols * 2
                Dim sumUnlocked As Single: sumUnlocked = 0#
                Dim sumLocked   As Single: sumLocked = 0#
                For j = 0 To nCols - 1
                    If Not outHidden(j) And j <> elasticIdx Then
                        If locked(j) Then
                            sumLocked = sumLocked + wPx(j)
                        Else
                            sumUnlocked = sumUnlocked + wPx(j)
                        End If
                    End If
                Next j
                If sumUnlocked <= 0# Then Exit For
                Dim ratio As Single: ratio = (targetPx - sumLocked) / sumUnlocked
                If ratio >= 1# Then Exit For
                Dim anyNewLocked As Boolean: anyNewLocked = False
                For j = 0 To nCols - 1
                    If Not outHidden(j) And j <> elasticIdx And Not locked(j) Then
                        Dim newW As Single: newW = wPx(j) * ratio
                        If newW < minWidths(j) Then
                            newW = minWidths(j)
                            locked(j) = True
                            anyNewLocked = True
                        End If
                        wPx(j) = newW
                    End If
                Next j
                If Not anyNewLocked Then Exit For
            Next iteration
        End If
    End If
    Dim sumFinal As Single: sumFinal = 0#
    For j = 0 To nCols - 1
        If Not outHidden(j) Then sumFinal = sumFinal + wPx(j)
    Next j
    Dim residual As Single: residual = sumFinal - wAvail
    If Abs(residual) > RESIDUAL_TOL_PX Then
        Dim absorbIdx As Long: absorbIdx = elasticIdx
        If absorbIdx < 0 Or outHidden(absorbIdx) Then
            For j = nCols - 1 To 0 Step -1
                If Not outHidden(j) Then absorbIdx = j: Exit For
            Next j
        End If
        If absorbIdx >= 0 Then
            wPx(absorbIdx) = wPx(absorbIdx) - residual
            If wPx(absorbIdx) < 0# Then wPx(absorbIdx) = 0#
        End If
    End If
    Dim availIntPx As Long: availIntPx = CLng(Int(wAvail))
    Dim budgetIdx  As Long: budgetIdx = -1
    If elasticIdx >= 0 And Not outHidden(elasticIdx) Then
        budgetIdx = elasticIdx
    Else
        For j = 0 To nCols - 1
            If Not outHidden(j) Then budgetIdx = j
        Next j
    End If
    If budgetIdx >= 0 Then
        Dim sumOtherPx As Long: sumOtherPx = 0
        For j = 0 To nCols - 1
            If j <> budgetIdx And Not outHidden(j) Then
                sumOtherPx = sumOtherPx + (-Int(-wPx(j)))
            End If
        Next j
        Dim budgetPx As Long: budgetPx = availIntPx - sumOtherPx
        If budgetPx < 1 Then budgetPx = 1
        wPx(budgetIdx) = CSng(budgetPx)
    End If
    Dim colStr As String: colStr = vbNullString
    For i = 0 To nCols - 1
        If i > 0 Then colStr = colStr & ";"
        If outHidden(i) Then
            colStr = colStr & "0"
            outWidths(i) = 0#
            emitPx(i) = 0
        Else
            emitPx(i) = -Int(-wPx(i))
            outWidths(i) = PxToPtCeil(wPx(i), dpi)
            colStr = colStr & PtToInvariant(outWidths(i))
        End If
    Next i
    lst.ColumnWidths = colStr
    res.logMessages = res.logMessages & "ColumnWidths: " & colStr & vbLf
    If hasHdrs Then
        Dim xPos     As Single:  xPos = lst.Left + borderInsetPt
        Dim hdrGdiOk As Boolean: hdrGdiOk = False
        If headerFontDiffers Then
            hdrGdiOk = AcquireGdiFont(hDC_hdr, hFont_hdr, hOldFont_hdr, _
                        hdrFontName, hdrFontSize, hdrFontBold, hdrFontItalic, _
                        dpi, res.logMessages)
        Else
            hDC_hdr = hDC
            hdrGdiOk = (hDC <> 0)
        End If
        For i = 0 To nCols - 1
            If i < nHdrs Then
                Dim lbl As Object
                Set lbl = hdrs(hdrBase + i)
                If Not lbl Is Nothing Then
                    If outHidden(i) Then
                        lbl.Visible = False
                    Else
                        lbl.Visible = True
                        lbl.Left = xPos
                        lbl.Width = outWidths(i)
                        If hdrGdiOk And Len(cfg(cfgBase + i).headerText) > 0 Then
                            Dim maxHdrPx As Long
                            maxHdrPx = emitPx(i) - CELL_PADDING_PX
                            Dim cap As String
                            If maxHdrPx > 0 Then
                                cap = TruncateText(hDC_hdr, cfg(cfgBase + i).headerText, maxHdrPx)
                            Else
                                cap = ChrW$(&H2026)
                            End If
                            If cap <> cfg(cfgBase + i).headerText Then
                                res.logMessages = res.logMessages & "HDR TRUNC col " & _
                                                 i & " -> '" & cap & "'" & vbLf
                            End If
                            lbl.Caption = cap
                        End If
                    End If
                End If
            End If
            If Not outHidden(i) Then xPos = xPos + outWidths(i)
        Next i
        If headerFontDiffers And hdrGdiOk Then
            ReleaseGdiFont hDC_hdr, hFont_hdr, hOldFont_hdr
        End If
    End If
    ReleaseGdiFont hDC, hFont, hOldFont
    res.success = True
    AutoFitListBox = True
    Exit Function

FailSafe:
    res.logMessages = res.logMessages & "ERROR " & Err.Number & ": " & _
                     Err.Description & vbLf
    If hDC_hdr <> hDC Then ReleaseGdiFont hDC_hdr, hFont_hdr, hOldFont_hdr
    ReleaseGdiFont hDC, hFont, hOldFont
    Dim nSafe As Long: If hasCfg Then nSafe = nCols Else nSafe = 1
    DegradeMode lst, nSafe, res, outWidths, outHidden, "unhandled error " & Err.Number
    AutoFitListBox = True
End Function

Public Function AutoFitListBoxHeight(ByVal lst As Object, _
                                     ByRef res As AutoFitResult, _
                                     Optional ByVal reserveHScroll As Boolean = False) As Boolean
    AutoFitListBoxHeight = False
    res.success = False
    res.degraded = False
    res.logMessages = vbNullString
    On Error GoTo HeightFailed
    If lst Is Nothing Then
        res.logMessages = "ERROR: lst Is Nothing" & vbLf
        Exit Function
    End If
    If lst.ListCount = 0 Then
        res.logMessages = "WARNING: ListCount = 0 - height unchanged" & vbLf
        res.success = True
        AutoFitListBoxHeight = True
        Exit Function
    End If
    Dim dpi    As Single: dpi = DetectDpi()
    Dim pxToPt As Single: pxToPt = POINTS_PER_INCH / dpi
    res.dpiUsed = dpi
    Dim edgeYPx As Long: edgeYPx = GetSystemMetrics(SM_CYEDGE)
    If edgeYPx <= 0 Then
        edgeYPx = FALLBACK_EDGE_PX
        res.degraded = True
        res.logMessages = res.logMessages & "WARNING: SM_CYEDGE failed, using " & _
                         FALLBACK_EDGE_PX & vbLf
    End If
    Dim borderHPx As Long: borderHPx = edgeYPx * 2
    If reserveHScroll Then
        Dim hscrPx As Long: hscrPx = GetSystemMetrics(SM_CYHSCROLL)
        If hscrPx > 0 Then borderHPx = borderHPx + hscrPx
    End If
    Dim fontHPx As Long: fontHPx = 0
#If VBA7 Then
    Dim hDC      As LongPtr
    Dim hFont    As LongPtr
    Dim hOldFont As LongPtr
#Else
    Dim hDC      As Long
    Dim hFont    As Long
    Dim hOldFont As Long
#End If
    If AcquireGdiFont(hDC, hFont, hOldFont, lst.Font.Name, lst.Font.Size, _
                      lst.Font.bOld, lst.Font.Italic, dpi, res.logMessages) Then
        fontHPx = MeasureTextHeightPx(hDC, "Mgy|")
        ReleaseGdiFont hDC, hFont, hOldFont
    End If
    If fontHPx <= 0 Then
        fontHPx = CLng(lst.Font.Size * dpi / POINTS_PER_INCH)
        res.degraded = True
        res.logMessages = res.logMessages & "WARNING: font height estimated (" & _
                         fontHPx & " px)" & vbLf
    End If
    If m_designHPx = 0# Or m_designDpi <> dpi Then
        m_designHPx = PtToPxFloor(lst.Height, dpi)
        m_designDpi = dpi
    End If
    Dim designClientPx As Single
    designClientPx = m_designHPx - borderHPx
    If designClientPx <= 0# Then designClientPx = PtToPxFloor(lst.Height, dpi) - borderHPx
    Dim rowHPx As Single
    If m_rowHPxCache > 0# And m_rowHDpi = dpi Then
        rowHPx = m_rowHPxCache
        res.logMessages = res.logMessages & "RowSrc=probe "
    Else
        rowHPx = CSng(fontHPx) + CSng(ROW_PADDING_PX)
        res.logMessages = res.logMessages & "RowSrc=fallback "
    End If
    If designClientPx <= 0# Or rowHPx <= 0# Then
        res.logMessages = res.logMessages & "WARNING: innerH <= 0 - unchanged" & vbLf
        res.success = True
        AutoFitListBoxHeight = True
        Exit Function
    End If
    Dim nBudget As Long
    If m_fitN > 0 And m_fitDpi = dpi Then
        nBudget = m_fitN
    Else
        nBudget = FloorRowCount(designClientPx, rowHPx, MAX_ROWS_UNCLAMPED)
        m_fitN = nBudget
        m_fitDpi = dpi
    End If
    Dim nChosen As Long: nChosen = nBudget
    If nChosen > lst.ListCount Then nChosen = lst.ListCount
    If nChosen < 1 Then nChosen = 1
    Dim rowHPt As Single: rowHPt = PxToPtCeil(rowHPx, dpi)
    lst.Height = PxToPtCeil(CSng(nChosen) * rowHPx + CSng(borderHPx), dpi)
    lst.IntegralHeight = True
    res.logMessages = res.logMessages & "RowH=" & PtToInvariant(rowHPt) & _
                     "pt rows=" & nChosen & "/" & lst.ListCount & _
                     " budget=" & nBudget & _
                     " designH=" & m_designHPx & "px" & _
                     " newH=" & PtToInvariant(lst.Height) & "pt" & vbLf
    res.success = True
    AutoFitListBoxHeight = True
    Exit Function

HeightFailed:
    res.logMessages = res.logMessages & "ERROR " & Err.Number & ": " & _
                     Err.Description & vbLf
End Function

Public Sub AutoFitCenterForm(ByVal frm As Object, ByVal invApp As Object)
    On Error Resume Next
    If frm Is Nothing Or invApp Is Nothing Then Exit Sub
#If VBA7 Then
    Dim hwndMain As LongPtr
#Else
    Dim hwndMain As Long
#End If
    hwndMain = invApp.MainFrameHWND
    If hwndMain = 0 Then Exit Sub
    Dim rc As RECT
    If GetWindowRect(hwndMain, rc) = 0 Then Exit Sub
    Dim pxToPt As Single: pxToPt = POINTS_PER_INCH / DetectDpi()
    Dim cx As Single: cx = CSng(rc.Left + rc.Right) / 2# * pxToPt
    Dim cy As Single: cy = CSng(rc.Top + rc.Bottom) / 2# * pxToPt
    Dim formLeft As Single: formLeft = cx - frm.Width / 2#
    Dim formTop  As Single: formTop = cy - frm.Height / 2#
    Dim invL As Single: invL = CSng(rc.Left) * pxToPt
    Dim invT As Single: invT = CSng(rc.Top) * pxToPt
    Dim invR As Single: invR = CSng(rc.Right) * pxToPt - frm.Width
    Dim invB As Single: invB = CSng(rc.Bottom) * pxToPt - frm.Height
    If formLeft < invL Then formLeft = invL
    If formLeft > invR Then formLeft = invR
    If formTop < invT Then formTop = invT
    If formTop > invB Then formTop = invB
    frm.Left = formLeft
    frm.Top = formTop
End Sub

Public Function AutoFitSelfTest() As Boolean
    Dim pass As Long, fail As Long
    AutoFitClearCache
    If PtToInvariant(5.2) = "5.2" And PtToInvariant(6#) = "6" And _
       PtToInvariant(-3#) = "0" And PtToInvariant(5.25) = "5.25" And _
       PxToPtCeil(10, 144) = 5 And PxToPtCeil(7, 96) = 5.25 And _
       PxToPtCeil(0, 96) = 0 Then
        pass = pass + 1: Debug.Print "PASS T1 PtToInvariant+PxToPtCeil"
    Else
        fail = fail + 1: Debug.Print "FAIL T1 PtToInvariant+PxToPtCeil"
    End If
    Dim d As Single: d = DetectDpi()
    If d >= 60 And d <= 600 Then
        pass = pass + 1: Debug.Print "PASS T2 DPI=" & d
    Else
        fail = fail + 1: Debug.Print "FAIL T2 DPI=" & d
    End If
#If VBA7 Then
    Dim hDC As LongPtr, hF As LongPtr, hO As LongPtr
#Else
    Dim hDC As Long, hF As Long, hO As Long
#End If
    Dim lg As String
    If AcquireGdiFont(hDC, hF, hO, "Tahoma", 8, False, False, d, lg) Then
        If MeasureTextPx(hDC, "M") > 0 And MeasureTextHeightPx(hDC, "Mgy|") > 0 Then
            pass = pass + 1: Debug.Print "PASS T3 GDI measure"
        Else
            fail = fail + 1: Debug.Print "FAIL T3 zero measurement"
        End If
    Else
        fail = fail + 1: Debug.Print "FAIL T3 acquire: " & lg
    End If
    Dim c0 As Long: c0 = m_cacheCount
    Dim wA As Single: wA = GetCachedWidth(hDC, "AutoFitSelfTest", "Tahoma", 8, False, d)
    Dim wB As Single: wB = GetCachedWidth(hDC, "AutoFitSelfTest", "Tahoma", 8, False, d)
    If wA > 0 And wA = wB And m_cacheCount = c0 + 1 Then
        pass = pass + 1: Debug.Print "PASS T4 cache hit"
    Else
        fail = fail + 1: Debug.Print "FAIL T4 cache (wA=" & wA & " wB=" & wB & _
                                     " count=" & m_cacheCount & " expected " & c0 + 1 & ")"
    End If
    Dim wC As Single: wC = GetCachedWidth(hDC, "AutoFitSelfTest", "Tahoma", 8, True, d)
    If wC > 0 And m_cacheCount = 1 Then
        pass = pass + 1: Debug.Print "PASS T5 cache invalidation"
    Else
        fail = fail + 1: Debug.Print "FAIL T5 cache (wC=" & wC & _
                                     " count=" & m_cacheCount & " expected 1)"
    End If
    If FloorRowCount(100, 10, 50) = 10 And _
       FloorRowCount(95, 10, 50) = 9 And _
       FloorRowCount(105, 10, 50) = 10 And _
       FloorRowCount(5, 10, 3) = 1 And _
       FloorRowCount(0, 10, 3) = 1 Then
        pass = pass + 1: Debug.Print "PASS T6 FloorRowCount"
    Else
        fail = fail + 1: Debug.Print "FAIL T6 FloorRowCount"
    End If
    Dim tt As String: tt = TruncateText(hDC, String$(200, "M"), 60)
    If Len(tt) < 200 And Right$(tt, 1) = ChrW$(&H2026) Then
        pass = pass + 1: Debug.Print "PASS T7 TruncateText"
    Else
        fail = fail + 1: Debug.Print "FAIL T7 TruncateText"
    End If
    Dim bPx As Long, sPx As Long, sfPx As Long
    ComputeMarginsPx bPx, sPx, sfPx
    If bPx > 0 And sPx > 0 And sfPx >= 0 Then
        pass = pass + 1: Debug.Print "PASS T8 margins"
    Else
        fail = fail + 1: Debug.Print "FAIL T8 margins"
    End If
    If DPI_AT_CALIBRATION > 0 And CELL_PADDING_PX > 0 And _
       ROW_PADDING_PX > 0 And SAFETY_PADDING_PX >= 0 And _
       AVG_CHAR_PX > 0 And _
       LF_TOTAL_BYTES = 92 And LF_FACE_OFFSET = 28 And _
       LF_TOTAL_BYTES - LF_FACE_OFFSET = LF_FACEBUF_CHARS And _
       FALLBACK_EDGE_PX > 0 And FALLBACK_SCROLL_PX > 0 And _
       PROBE_ROWS > 0 And PROBE_MIN_ROW_PX > 0 And _
       PROBE_MAX_ROW_PX > PROBE_MIN_ROW_PX And _
       MAX_ROWS_UNCLAMPED > 0 Then
        pass = pass + 1: Debug.Print "PASS T9 constants layout"
    Else
        fail = fail + 1: Debug.Print "FAIL T9 constants layout"
    End If
    ReleaseGdiFont hDC, hF, hO
    AutoFitClearCache
    Debug.Print "AutoFitSelfTest: " & pass & "/" & (pass + fail) & " PASS"
    AutoFitSelfTest = (fail = 0)
End Function

Public Sub AutoFitCalibrateCellPadding(ByRef lst As Object)
    If lst Is Nothing Then Debug.Print "ERROR: lst Is Nothing": Exit Sub
    Dim dpi    As Single: dpi = DetectDpi()
    Dim pxToPt As Single: pxToPt = POINTS_PER_INCH / dpi
    Dim txt As String
    On Error Resume Next
    txt = lst.List(0, 0)
    On Error GoTo 0
    If Len(txt) = 0 Then txt = "MMMMMMMMMM"
#If VBA7 Then
    Dim hDC As LongPtr, hFont As LongPtr, hOldFont As LongPtr
#Else
    Dim hDC As Long, hFont As Long, hOldFont As Long
#End If
    Dim logMsg As String
    If Not AcquireGdiFont(hDC, hFont, hOldFont, lst.Font.Name, lst.Font.Size, _
                          lst.Font.bOld, lst.Font.Italic, dpi, logMsg) Then
        Debug.Print "ERROR: " & logMsg: Exit Sub
    End If
    Dim wPx As Long: wPx = MeasureTextPx(hDC, txt)
    ReleaseGdiFont hDC, hFont, hOldFont
    Debug.Print "--- AutoFitCalibrateCellPadding ---"
    Debug.Print "Text  : '" & txt & "'"
    Debug.Print "W_gdi : " & wPx & " px (" & PtToInvariant(CSng(wPx) * pxToPt) & " pt)"
    Debug.Print "DPI   : " & dpi
    Debug.Print "=> CELL_PADDING_PX = lastGoodPx - " & wPx
    Dim origWpt As Single: origWpt = lst.Width
    Dim origWPx As Long: origWPx = PtToPxFloor(origWpt, dpi)
    If wPx >= origWPx Then
        Debug.Print "NOTE: probe text already exceeds list width - nothing to sweep."
        Exit Sub
    End If
    Dim origCW As String
    On Error Resume Next
    origCW = lst.ColumnWidths
    On Error GoTo 0
    Dim curPx As Long
    For curPx = origWPx To wPx Step -CAL_STEP_PX
        lst.Width = PxToPtCeil(curPx, dpi)
        lst.ColumnWidths = PtToInvariant(PxToPtCeil(curPx, dpi))
        DoEvents
        Debug.Print "  Width = " & curPx & " px"
    Next curPx
    lst.Width = origWpt
    If Len(origCW) > 0 Then lst.ColumnWidths = origCW
    Debug.Print "Done. Original width and ColumnWidths restored."
End Sub

Public Sub AutoFitCalibrateRowPadding(ByRef lst As Object)
    If lst Is Nothing Then Debug.Print "ERROR: lst Is Nothing": Exit Sub
    Dim dpi    As Single: dpi = DetectDpi()
    Dim pxToPt As Single: pxToPt = POINTS_PER_INCH / dpi
    Dim edgeYPx As Long: edgeYPx = GetSystemMetrics(SM_CYEDGE)
    If edgeYPx <= 0 Then edgeYPx = FALLBACK_EDGE_PX
    Dim borderHPx As Long: borderHPx = edgeYPx * 2
#If VBA7 Then
    Dim hDC As LongPtr, hFont As LongPtr, hOldFont As LongPtr
#Else
    Dim hDC As Long, hFont As Long, hOldFont As Long
#End If
    Dim logMsg  As String
    Dim fontHPx As Long
    If AcquireGdiFont(hDC, hFont, hOldFont, lst.Font.Name, lst.Font.Size, _
                      lst.Font.bOld, lst.Font.Italic, dpi, logMsg) Then
        fontHPx = MeasureTextHeightPx(hDC, "Mgy|")
        ReleaseGdiFont hDC, hFont, hOldFont
    End If
    If fontHPx <= 0 Then fontHPx = CLng(lst.Font.Size * dpi / POINTS_PER_INCH)
    Debug.Print "--- AutoFitCalibrateRowPadding ---"
    Debug.Print "FontH   : " & fontHPx & " px (" & PtToInvariant(CSng(fontHPx) * pxToPt) & " pt)"
    Debug.Print "BorderH : " & borderHPx & " px"
    Debug.Print "Rows    : " & lst.ListCount
    Debug.Print "DPI     : " & dpi
    Debug.Print "Ctrl+Break when the bottom row starts to clip."
    Debug.Print "=> ROW_PADDING_PX = (lastGoodPx - " & borderHPx & _
                ") / nVisibleRows - " & fontHPx
    Dim origHpt As Single: origHpt = lst.Height
    Dim origHPx As Long: origHPx = PtToPxFloor(origHpt, dpi)
    Dim floorPx As Long: floorPx = borderHPx + fontHPx
    If floorPx >= origHPx Then
        Debug.Print "NOTE: single row already exceeds list height - nothing to sweep."
        Exit Sub
    End If
    Dim curPx As Long
    For curPx = origHPx To floorPx Step -CAL_STEP_PX
        lst.Height = PxToPtCeil(curPx, dpi)
        DoEvents
        Debug.Print "  Height = " & curPx & " px"
    Next curPx
    lst.Height = origHpt
    Debug.Print "Done. Original height restored."
End Sub

' END: modListBoxFit v9.25.31

