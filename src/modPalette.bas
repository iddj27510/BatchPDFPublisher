Attribute VB_Name = "modPalette"
Option Explicit
' Module    : modPalette
' Purpose   : CLR_* palette, FW_IsDark, FW_LoadPalette, high contrast detect
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + header completion + language):
'   Added missing Option Explicit at the top of the module.
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Added missing Reference and Notes sections for consistency
'   with the rest of the project.
'   Serbian section labels inside FW_LoadPalette translated to
'   English ("Chrome / struktura" -> "Chrome / structure", etc.)
'   for consistency with the rest of the module.
'   No functional changes.
'
' v9.19.3 (Header language cleanup):
'   Serbian fragments in v9.19.0 note replaced with English.
'   No code changes — cosmetic only.
' v9.19.0 (Inventor brand colors, WCAG AA):
'   Palette aligned with Autodesk Inventor document "Color Register
'   and Semantics" v4.1:
'     Dark:  chrome #2D2D2D, ribbon #333333, surface #252526,
'            separator #4A4A4A, accent #2A8FD4 (Inventor accent)
'     Light: chrome #FAFAFA, panel #FFFFFF, separator #D0D0D0,
'            nav-widget #E8E8E8, accent #2A8FD4
'   All text colors pass WCAG AA (>= 4.5:1) on their background.
'   Contrast values are in comments below each color.
'
'   Note on the RUN button: Inventor accent #2A8FD4 with white
'   text gives contrast 3.51:1 — passes AA Large (18pt+) but
'   NOT AA for normal text (9pt). Therefore RUN_BG uses a darker
'   variant #1E5FA0 (6.66:1) in dark theme, and #0066CC (5.57:1)
'   in light theme. Accent #2A8FD4 remains for focus, hover and
'   borders — where WCAG threshold is 3:1 (UI components).
'
'   CAD constants from the same document (§2.1) — Triad X/Y/Z,
'   add/cut preview, Sick magenta, grip orange, transient surface,
'   supplemental purple, DOF free, drafting Full-Constrain —
'   are NOT part of this palette because they are not used in
'   VBA UI code. If needed, hardcode them in modConstants
'   (never bind them to UI theme).
' v9.18.1 (AA contrast):
'   Light theme — fixed colors that did not pass WCAG AA.
' v9.18.0 (§7.5 Palette documentation, §7.19 High contrast):
'   BGR/RGB comments for every CLR_* value.
'   PRJ_IsHighContrast using SystemParametersInfoW.
'
' Requires  : g_sTheme, g_bHighContrast, g_bCurrentDark (modState).
'             CLR_* and BTN_* palette variables (modState).
'             ThisApplication.ThemeManager, ThisApplication.ColorSchemes
'             (Inventor host).
'
' Notes     : No line-continuation characters in procedure bodies.
'             API Declare statements and multi-line string
'             concatenations are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

#If VBA7 Then
    Private Declare PtrSafe Function SystemParametersInfoW Lib "user32" _
        (ByVal uiAction As Long, ByVal uiParam As Long, _
         ByRef pvParam As Any, ByVal fWinIni As Long) As Long
#Else
    Private Declare Function SystemParametersInfoW Lib "user32" _
        (ByVal uiAction As Long, ByVal uiParam As Long, _
         ByRef pvParam As Any, ByVal fWinIni As Long) As Long
#End If

Private Type HIGHCONTRAST
    cbSize            As Long
    dwFlags           As Long
    lpszDefaultScheme As Long
End Type

Private Const SPI_GETHIGHCONTRAST As Long = 66
Private Const HCF_HIGHCONTRASTON As Long = &H1

Public Function PRJ_IsHighContrast() As Boolean
    PRJ_IsHighContrast = False
    Dim hc As HIGHCONTRAST
    hc.cbSize = LenB(hc)
    Dim lRet As Long
    On Error Resume Next
    Err.Clear
    lRet = SystemParametersInfoW(SPI_GETHIGHCONTRAST, LenB(hc), hc, 0)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If lRet = 0 Then Exit Function
    PRJ_IsHighContrast = ((hc.dwFlags And HCF_HIGHCONTRASTON) <> 0)
End Function

Public Function PRJ_IsInventorDark() As Boolean
    PRJ_IsInventorDark = False
    Dim sName As String
    On Error Resume Next
    Err.Clear
    sName = LCase$(ThisApplication.ThemeManager.ActiveTheme.Name)
    If Err.Number <> 0 Then
        Err.Clear
        sName = ""
    End If
    On Error GoTo 0

    If Len(sName) > 0 Then
        PRJ_IsInventorDark = (InStr(sName, "dark") > 0)
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    sName = LCase$(ThisApplication.ColorSchemes.ActiveColorScheme.Name)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    PRJ_IsInventorDark = (InStr(sName, "dark") > 0)
End Function

Public Function PRJ_IsWindowsDark() As Boolean
    PRJ_IsWindowsDark = False
    Dim oShell As Object
    On Error Resume Next
    Err.Clear
    Set oShell = CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If oShell Is Nothing Then Exit Function

    Dim v As Variant
    On Error Resume Next
    Err.Clear
    v = oShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\" & _
                       "Themes\Personalize\AppsUseLightTheme")
    If Err.Number <> 0 Then
        Err.Clear
        Set oShell = Nothing
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    Set oShell = Nothing
    If Not IsNumeric(v) Then Exit Function
    PRJ_IsWindowsDark = (CInt(v) = 0)
End Function

Public Function FW_IsDark() As Boolean
    Select Case g_sTheme
        Case "L": FW_IsDark = False
        Case "D": FW_IsDark = True
        Case "S": FW_IsDark = PRJ_IsWindowsDark()
        Case "I": FW_IsDark = PRJ_IsInventorDark()
        Case Else: FW_IsDark = PRJ_IsInventorDark()
    End Select
End Function

Public Sub FW_LoadPalette()
    ' §7.19: If user has high contrast mode active, skip custom palette.
    ' MSForms will use system colors which are already HC-aware.
    g_bHighContrast = PRJ_IsHighContrast()
    If g_bHighContrast Then
        g_bCurrentDark = False
        Exit Sub
    End If

    Dim bD As Boolean
    bD = FW_IsDark()
    g_bCurrentDark = bD

    If bD Then
        ' ============================================================
        ' DARK THEME — Inventor Dark UI
        ' VBA hex literals are BGR (0x00BBGGRR). Comments show the
        ' rendered #RRGGBB value and WCAG AA contrast on its surface.
        ' ============================================================

        ' --- Chrome / structure ---
        CLR_BG = &H2D2D2D               ' #2D2D2D  Inventor panel background
        CLR_SURFACE = &H262525          ' #252526  Inventor input field
        CLR_SEP = &H4A4A4A              ' #4A4A4A  Inventor separator

        ' --- Text (AA on #2D2D2D) ---
        CLR_TEXT_PRI = &HF0F0F0         ' #F0F0F0  AA 13.6:1
        CLR_TEXT_SEC = &HC8C8C8         ' #C8C8C8  AA  7.9:1

        ' --- Accent (Inventor blue) ---
        CLR_ACCENT = &HD48F2A           ' #2A8FD4  Inventor accent (focus)
        CLR_ACCENT_TXT = &HFFFFFF       ' #FFFFFF

        ' --- Secondary buttons ---
        CLR_BTN_SEC = &H3C3C3C          ' #3C3C3C  Inventor section header
        CLR_BTN_SEC_TXT = &HF0F0F0      ' #F0F0F0  AA 10.9:1 on #3C3C3C

        ' --- Semantic colors (AA on #2D2D2D) ---
        CLR_ERROR = &H6B6BFF            ' #FF6B6B  AA  4.84:1
        CLR_OK = &H50AF4C               ' #4CAF50  AA  4.84:1
        CLR_WARN = &H47B3FF             ' #FFB347  AA  9.58:1

        ' --- RUN button (darker blue for AA with white text) ---
        CLR_RUN_BG = &HA05F1E           ' #1E5FA0  AA  6.66:1 with #FFFFFF
        CLR_RUN_TXT = &HFFFFFF          ' #FFFFFF
        CLR_RUN_DIS = &H3A3A3A          ' #3A3A3A
        CLR_RUN_DIS_TXT = &H707070      ' #707070  (disabled, WCAG exempt)

        ' --- Hover / press states ---
        BTN_HOVER_RUN = &HD48F2A        ' #2A8FD4  accent as hover
        BTN_PRESS_RUN = &H804A15        ' #154A80  darker
        BTN_HOVER_SEC = &H524A4A        ' #4A4A52  lighter than #3C3C3C
        BTN_PRESS_SEC = &H322A2A        ' #2A2A32  darker than #3C3C3C
        BTN_HOVER_ACCENT = &HD48F2A     ' #2A8FD4
        BTN_PRESS_ACCENT = &H804A15     ' #154A80

        ' --- Temp status (AA on #2D2D2D) ---
        CLR_TEMP_OK = &H7FE07F          ' #7FE07F  AA 10.35:1
        CLR_TEMP_WARN = &HD7FF          ' #FFD700  AA 12.08:1
        CLR_TEMP_INFO = &HF6B564        ' #64B5F6  AA  7.74:1

        ' --- Process (same as TEMP_INFO) ---
        CLR_PROCESS = &HF6B564          ' #64B5F6
    Else
        ' ============================================================
        ' LIGHT THEME — Inventor Light UI
        ' AA threshold >= 4.5:1 for normal text on #FAFAFA.
        ' ============================================================

        ' --- Chrome / structure ---
        CLR_BG = &HFAFAFA               ' #FAFAFA  Inventor ribbon base
        CLR_SURFACE = &HFFFFFF          ' #FFFFFF  Inventor panel/input
        CLR_SEP = &HD0D0D0              ' #D0D0D0  Inventor separator

        ' --- Text (AA on #FAFAFA) ---
        CLR_TEXT_PRI = &H1A1A1A         ' #1A1A1A  AA 16.7:1
        CLR_TEXT_SEC = &H555555         ' #555555  AA  7.2:1

        ' --- Accent (same in both themes) ---
        CLR_ACCENT = &HD48F2A           ' #2A8FD4  Inventor accent (focus)
        CLR_ACCENT_TXT = &HFFFFFF       ' #FFFFFF

        ' --- Secondary buttons (Inventor nav widget) ---
        CLR_BTN_SEC = &HE8E8E8          ' #E8E8E8  Inventor nav-widget bg
        CLR_BTN_SEC_TXT = &H333333      ' #333333  AA 10.4:1 on #E8E8E8

        ' --- Semantic colors (AA on #FAFAFA) ---
        CLR_ERROR = &H2000B0            ' #B00020  AA  7.06:1
        CLR_OK = &H6400                 ' #006400  AA  7.17:1
        CLR_WARN = &H953B4              ' #B45309  AA  4.85:1

        ' --- RUN button (#0066CC for AA with white text) ---
        CLR_RUN_BG = &HCC6600           ' #0066CC  AA  5.57:1 with #FFFFFF
        CLR_RUN_TXT = &HFFFFFF          ' #FFFFFF
        CLR_RUN_DIS = &HE0E0E0          ' #E0E0E0
        CLR_RUN_DIS_TXT = &HA0A0A0      ' #A0A0A0  (disabled, WCAG exempt)

        ' --- Hover / press states ---
        BTN_HOVER_RUN = &HD68533        ' #3385D6
        BTN_PRESS_RUN = &HA35200        ' #0052A3
        BTN_HOVER_SEC = &HD8D8D8        ' #D8D8D8  lighter than #E8E8E8
        BTN_PRESS_SEC = &HC0C0C0        ' #C0C0C0  darker than #E8E8E8
        BTN_HOVER_ACCENT = &HD68533     ' #3385D6
        BTN_PRESS_ACCENT = &HA35200     ' #0052A3

        ' --- Temp status (AA on #FAFAFA) ---
        CLR_TEMP_OK = &H346516          ' #166534  AA  6.54:1
        CLR_TEMP_WARN = &HE4092         ' #92400E  AA  6.50:1
        CLR_TEMP_INFO = &HC06515        ' #1565C0  AA  5.27:1

        ' --- Process (same as TEMP_INFO) ---
        CLR_PROCESS = &HC06515          ' #1565C0
    End If
End Sub

' END: modPalette v9.19.9

