VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmScan 
   Caption         =   "Batch PDF Publisher"
   ClientHeight    =   4320
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8400.001
   OleObjectBlob   =   "frmScan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Component : frmScan (UserForm)
' Purpose   : Main scan form — folder selection, subfolder/force/audit
'             toggles, theme cycling, reset arm/confirm, and batch
'             launch trigger. Hosts the CKeyPreview hook so Esc
'             cancels the running batch or unloads the form.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.9 (path field undo fix):
'   Ctrl+Z (Undo) in an MSForms TextBox is consumed by the
'   underlying Windows Edit control before VBA sees it. The
'   Change event does not fire for Ctrl+Z, so the validation
'   label stayed stale after the user restored deleted text.
'
'   Fix: RevalidatePathIfChanged compares the current txtPath
'   text against g_sLastFolder and re-runs SC_QuickCheck +
'   SC_Validate + SC_UpdateSum2 when they differ. Called from
'   UserForm_MouseMove.
'
'   Known limitation: pure keyboard Ctrl+Z with no subsequent
'   mouse movement does not refresh the label until the user
'   moves the mouse. Three alternative approaches were
'   investigated and rejected:
'     - txtPath_KeyUp / txtPath_MouseDown handlers: reading
'       .Text during those events resets the Edit control's
'       undo buffer, breaking Ctrl+Z entirely.
'     - Win32 SetTimer + AddressOf callback: Inventor VBA
'       does not support the AddressOf operator (error 35010
'       at compile time).
'     - Subclassing the Edit control (GWL_WNDPROC + WM_UNDO):
'       same AddressOf dependency, not viable in Inventor.
'   In practice the user always moves the mouse after Ctrl+Z,
'   so the MouseMove hook covers every real workflow. Enter
'   in txtPath (txtPath_KeyDown) already runs the full
'   validation on its own.
'
' v9.19.9 (scan options: checkbox -> single toggle labels):
'   The three scan options (subfolders / force / audit) were
'   CheckBox controls, then briefly label pairs (lblOptX +
'   lblOptXTxt). Now: single-label toggles — one label carries
'   both the background colour and the localized caption.
'
'   Why single label:
'     1. No MSForms first-click quirk (no .Value).
'     2. No programmatic-Click side effect (no .Value).
'     3. No duplicate-Click debounce needed.
'     4. No m_bInitializing guard.
'     5. Half the controls, half the code.
'
'   Trade-off: screen readers will not announce the toggle state.
'   Keyboard access is preserved through Accelerators
'   (Alt+S / Alt+F / Alt+A) wired in PRJ_SetupAccelerators.
'
'   Layout — LayoutToggleLabels measures each label's natural
'   width via AutoSize, then:
'     Width  = Ceiling(natural width) + TOGGLE_WIDTH_PAD
'     Height = designer height (captured in Initialize)
'     Left   = previous.Left + previous.Width + TOGGLE_GAP
'     Top    = designer top (captured in Initialize)
'
'   Order-critical: AutoSize = False MUST be assigned BEFORE
'   .Width / .Height. MSForms ignores those assignments while
'   AutoSize is True.
'
' v9.19.9 (no-flash activation):
'   The registry read, caption setup, theme application, toggle
'   layout and toggle colouring all happen in UserForm_Initialize
'   (fires on Load, before Show, before the first paint).
'   UserForm_Activate no longer touches colours or layout — it
'   runs only the passes that need a visible form (FIT,
'   centering, hover registration, validation).
'
'   Visual state:
'     Active   -> CLR_ACCENT background, CLR_ACCENT_TXT text
'     Inactive -> CLR_BTN_SEC background, CLR_BTN_SEC_TXT text
'   Border: BorderStyle = 1 on all three, matching the secondary
'   button family (Dark / Reset / Cancel).
'
' v9.19.8 (checkbox first-click + programmatic-Click fix):
'   Superseded by the v9.19.9 label refactor.
'
' v9.19.6 (Header standardization):
'   Added Component/Purpose/Reference/Version header fields.
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Private Const TOGGLE_WIDTH_PAD As Single = 6
Private Const TOGGLE_GAP       As Single = 6

Private m_sOrigCap         As String
Private m_bActivated       As Boolean
Private m_dActivateTime    As Double
Private m_bResetArmed      As Boolean
Private m_dResetArmedAt    As Double
Private m_oKeyPreview      As CKeyPreview
Private m_dToggleStartLeft As Single
Private m_dToggleStartTop  As Single
Private m_dToggleHeight    As Single

Private Sub UserForm_Initialize()
    m_sOrigCap = Me.Caption
    Me.StartUpPosition = 1
    m_bResetArmed = False
    g_tLastBadge.nTot = 0
    g_sLastBadgeText = ""

    ' Capture designer positions so LayoutToggleLabels can place
    ' the row from a stable origin across repeated calls.
    On Error Resume Next
    m_dToggleStartLeft = Me.Controls("lblOptSub").Left
    m_dToggleStartTop = Me.Controls("lblOptSub").Top
    m_dToggleHeight = Me.Controls("lblOptSub").Height
    On Error GoTo 0

    ' v9.19.9: read state + apply colours HERE, before Show
    ' triggers the first paint. Moving this out of Activate
    ' eliminates the designer-default -> theme -> accent flash
    ' that used to be visible on every form open.
    REG_ClearCache

    Dim sCap As String
    sCap = L("title")
    If Len(sCap) = 0 Then sCap = Trim$(m_sOrigCap)
    Dim sVerSuffix As String
    sVerSuffix = " v" & PRJ_VERSION
    If Right$(sCap, Len(sVerSuffix)) <> sVerSuffix Then
        Me.Caption = sCap & sVerSuffix
    End If

    g_sTheme = REG_Get(REG_KEY_THEME, "I")
    g_bSubfolders = (LCase$(REG_Get(REG_KEY_SUB, "True")) = "true")
    g_bForceOverwrite = (LCase$(REG_Get(REG_KEY_FORCE, "False")) = "true")
    g_bAuditMode = (LCase$(REG_Get(REG_KEY_AUDIT, "False")) = "true")
    g_sLastFolder = REG_Get(REG_KEY_PATH, "")

    Me.Controls("txtPath").Text = g_sLastFolder

    g_bThemeApplied = False
    FW_ApplyTheme Me
    PRJ_LocalizeCaptions Me
    ApplyToggleColors
    PRJ_UpdateThemeLabel Me
    LayoutToggleLabels
    ApplyToggleColors

    PRJ_SetupAccelerators Me
    Set m_oKeyPreview = New CKeyPreview
    m_oKeyPreview.Initialize Me, "OnFormKeyDown"
End Sub

Private Sub UserForm_Activate()
    If m_bActivated Then Exit Sub
    m_bActivated = True

    m_dActivateTime = TM_Now()

    ' v9.19.9: colours and layout were applied in Initialize, so
    ' the first paint is already correct. Activate now runs only
    ' the passes that need a visible form.
    FIT_FitAllLabels Me
    SC_CenterAll Me
    PRJ_ApplyTooltips Me
    PRJ_InitHover Me
    SC_Validate Me
    SC_UpdateSum2 Me
End Sub

Public Sub OnFormKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        If g_bInProcessLoop Then
            g_bCancel = True
            KeyCode = 0
        Else
            Unload Me
            KeyCode = 0
        End If
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FW_ResetHover
    ResetDisarmCheck
    RevalidatePathIfChanged
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        If g_bInProcessLoop Then
            g_bCancel = True
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And g_bInProcessLoop Then
        Cancel = True
        g_bCancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
    FW_ReleaseAllHoverButtons
    m_bActivated = False
    If Not m_oKeyPreview Is Nothing Then
        m_oKeyPreview.Terminate
        Set m_oKeyPreview = Nothing
    End If
End Sub

' LayoutToggleLabels   place lblOptSub / lblOptForce / lblOptAudit
'   in a single horizontal row.
'   Each label: AutoSize = False (first), then
'     Width  = Ceiling(natural width) + TOGGLE_WIDTH_PAD
'     Height = designer height (captured in Initialize)
'   Spacing: TOGGLE_GAP between adjacent labels.
'   MSForms quirk: .Width / .Height are ignored while AutoSize is
'   True. AutoSize must be cleared before those assignments, hence
'   the explicit property order below.
Private Sub LayoutToggleLabels()
    On Error Resume Next

    Dim c1 As Object, c2 As Object, c3 As Object
    Set c1 = Me.Controls("lblOptSub")
    Set c2 = Me.Controls("lblOptForce")
    Set c3 = Me.Controls("lblOptAudit")
    If c1 Is Nothing Or c2 Is Nothing Or c3 Is Nothing Then Exit Sub

    Dim dW1 As Double, dH1 As Double
    Dim dW2 As Double, dH2 As Double
    Dim dW3 As Double, dH3 As Double

    MeasureToggleLabel c1, dW1, dH1
    MeasureToggleLabel c2, dW2, dH2
    MeasureToggleLabel c3, dW3, dH3

    c1.AutoSize = False
    c1.WordWrap = False
    c1.TextAlign = 2
    c1.Left = m_dToggleStartLeft
    c1.Top = m_dToggleStartTop
    c1.Width = CeilToSingle(dW1) + TOGGLE_WIDTH_PAD
    c1.Height = m_dToggleHeight

    c2.AutoSize = False
    c2.WordWrap = False
    c2.TextAlign = 2
    c2.Left = c1.Left + c1.Width + TOGGLE_GAP
    c2.Top = m_dToggleStartTop
    c2.Width = CeilToSingle(dW2) + TOGGLE_WIDTH_PAD
    c2.Height = m_dToggleHeight

    c3.AutoSize = False
    c3.WordWrap = False
    c3.TextAlign = 2
    c3.Left = c2.Left + c2.Width + TOGGLE_GAP
    c3.Top = m_dToggleStartTop
    c3.Width = CeilToSingle(dW3) + TOGGLE_WIDTH_PAD
    c3.Height = m_dToggleHeight

    On Error GoTo 0
End Sub

Private Sub MeasureToggleLabel(ByVal o As Object, ByRef dW As Double, ByRef dH As Double)
    Dim bWasAuto As Boolean
    bWasAuto = o.AutoSize
    o.WordWrap = False
    o.AutoSize = True
    DoEvents
    dW = o.Width
    dH = o.Height
    o.AutoSize = bWasAuto
End Sub

Private Function CeilToSingle(ByVal d As Double) As Single
    CeilToSingle = CSng(-Int(-d))
End Function

Private Sub ApplyToggleColors()
    ApplyOneToggle "lblOptSub", g_bSubfolders
    ApplyOneToggle "lblOptForce", g_bForceOverwrite
    ApplyOneToggle "lblOptAudit", g_bAuditMode
End Sub

Private Sub ApplyOneToggle(ByVal sName As String, ByVal bActive As Boolean)
    On Error Resume Next
    Dim o As Object
    Set o = Me.Controls(sName)
    If o Is Nothing Then Exit Sub

    Dim clrBg As Long, clrTxt As Long
    If bActive Then
        clrBg = CLR_ACCENT
        clrTxt = CLR_ACCENT_TXT
    Else
        clrBg = CLR_BTN_SEC
        clrTxt = CLR_BTN_SEC_TXT
    End If

    ' v9.19.9: 1px border to match the secondary button family
    ' (Dark / Reset / Cancel). BorderColor is left at the MSForms
    ' default (Window Frame system color).
    If o.BorderStyle <> 1 Then o.BorderStyle = 1

    If o.BackColor <> clrBg Then o.BackColor = clrBg
    If o.ForeColor <> clrTxt Then o.ForeColor = clrTxt
    On Error GoTo 0
End Sub

Private Sub OptSubClick()
    If g_bInProcessLoop Then Exit Sub
    g_bSubfolders = Not g_bSubfolders
    REG_Set REG_KEY_SUB, CStr(g_bSubfolders)
    ApplyToggleColors
    SC_Validate Me
    SC_UpdateSum2 Me
End Sub

Private Sub OptForceClick()
    If g_bInProcessLoop Then Exit Sub
    g_bForceOverwrite = Not g_bForceOverwrite
    REG_Set REG_KEY_FORCE, CStr(g_bForceOverwrite)
    ApplyToggleColors
    SC_ApplyDrawingBadge Me, g_tLastBadge
End Sub

Private Sub OptAuditClick()
    If g_bInProcessLoop Then Exit Sub
    g_bAuditMode = Not g_bAuditMode
    REG_Set REG_KEY_AUDIT, CStr(g_bAuditMode)
    ApplyToggleColors
    SC_UpdateSum2 Me
End Sub

Private Sub BrowseClick()
    If g_bInProcessLoop Then Exit Sub

    Dim sPicked As String
    sPicked = BrowseForFolderEx(L("scan.browse_title"), g_sLastFolder)
    If Len(sPicked) = 0 Then Exit Sub
    g_sLastFolder = SN_SanitizePath(sPicked)
    Me.Controls("txtPath").Text = g_sLastFolder
    REG_Set REG_KEY_PATH, g_sLastFolder
    SC_Validate Me
    SC_UpdateSum2 Me
End Sub

Private Sub RunClick()
    If g_bInProcessLoop Then Exit Sub
    If Not g_bHoverRunEnabled Then Exit Sub
    If Not PRJ_ValidateFolderPath(Me, True) Then Exit Sub
    PRJ_StartBatchFromScan
    Unload Me
    PRJ_ShowProgress
    PRJ_ShowReport
End Sub

Private Sub ThemeClick()
    If g_bInProcessLoop Then Exit Sub

    PRJ_CycleTheme Me
    SC_ReapplyColors Me
    PRJ_LocalizeCaptions Me
    PRJ_UpdateThemeLabel Me
    LayoutToggleLabels
    ApplyToggleColors
    FIT_FitAllLabels Me
    SC_CenterAll Me
    PRJ_ApplyTooltips Me
    Me.Repaint
End Sub

Private Sub SetResetText(ByVal sText As String, ByVal lColor As Long, ByVal bBold As Boolean)
    On Error Resume Next
    Dim oTxt As Object
    Set oTxt = Me.Controls("lblResetTxt")
    If oTxt Is Nothing Then Exit Sub
    oTxt.Caption = sText
    oTxt.ForeColor = lColor
    oTxt.Font.Bold = bBold
    PRJ_AlignTxt Me, "lblReset", "lblResetTxt"
    On Error GoTo 0
End Sub

Private Sub ResetClick()
    If g_bInProcessLoop Then Exit Sub

    On Error Resume Next
    Dim oLbl As Object
    Set oLbl = Me.Controls("lblReset")
    If oLbl Is Nothing Then Exit Sub

    If Not m_bResetArmed Then
        m_bResetArmed = True
        m_dResetArmedAt = TM_Now()
        SetResetText L("scan.reset.sure"), CLR_TEMP_WARN, True
        Me.Repaint
        Exit Sub
    End If

    m_bResetArmed = False
    STATE_ResetPreferences

    Me.Controls("txtPath").Text = g_sLastFolder
    ApplyToggleColors

    SetResetText L("scan.reset.done"), CLR_TEMP_OK, True
    Me.Repaint
    UI_SleepResponsive UI_MSG_BRIEF_MS

    SetResetText L("btn.reset"), CLR_RUN_TXT, False
    FW_ApplyTheme Me
    LayoutToggleLabels
    ApplyToggleColors
    Me.Repaint

    SC_Validate Me
    SC_UpdateSum2 Me
End Sub

Private Sub ResetDisarmCheck()
    On Error Resume Next
    If Not m_bResetArmed Then Exit Sub
    If TM_Elapsed(m_dResetArmedAt, TM_Now()) < RESET_ARM_TIMEOUT_SEC Then Exit Sub

    m_bResetArmed = False
    SetResetText L("btn.reset"), CLR_RUN_TXT, False
End Sub

' RevalidatePathIfChanged   re-run path validation when the text
'   has changed without firing the Change event.
'
'   Ctrl+Z (Undo) in an MSForms TextBox is consumed by the
'   underlying Windows Edit control before VBA sees it. Change
'   does not fire for Ctrl+Z on any host, so the validation label
'   and the audit summary both stayed stale after the user
'   restored deleted text.
'
'   Called from UserForm_MouseMove. The fast path is one string
'   compare and exit, so it is safe to call on every mouse move.
Private Sub RevalidatePathIfChanged()
    If g_bInProcessLoop Then Exit Sub

    Dim sNow As String
    sNow = SN_SanitizePath(Me.Controls("txtPath").Text)
    If sNow = g_sLastFolder Then Exit Sub

    g_sLastFolder = sNow
    SC_QuickCheck Me
    ResetDisarmCheck

    If Len(g_sLastFolder) > 3 Then
        If PRJ_FSO().FolderExists(g_sLastFolder) Then
            SC_Validate Me, True
            SC_UpdateSum2 Me
        End If
    End If
End Sub

Private Sub txtPath_Change()
    If g_bInProcessLoop Then Exit Sub
    g_sLastFolder = SN_SanitizePath(Me.Controls("txtPath").Text)
    SC_QuickCheck Me
    ResetDisarmCheck

    If Len(g_sLastFolder) > 3 Then
        If PRJ_FSO().FolderExists(g_sLastFolder) Then
            SC_Validate Me, True
        End If
    End If
End Sub

Private Sub txtPath_AfterUpdate()
    If g_bInProcessLoop Then Exit Sub
    REG_Set REG_KEY_PATH, g_sLastFolder
    SC_Validate Me
    SC_UpdateSum2 Me
End Sub

Private Sub txtPath_Enter()
    On Error Resume Next
    Me.Controls("txtPath").BorderColor = CLR_ACCENT
    On Error GoTo 0
End Sub

Private Sub txtPath_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Me.Controls("txtPath").BorderColor = CLR_SEP
    On Error GoTo 0
End Sub

Private Sub txtPath_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If g_bInProcessLoop Then Exit Sub
    If KeyCode = 13 Then
        If TM_Elapsed(m_dActivateTime, TM_Now()) < KEY_DEBOUNCE_SEC Then
            KeyCode = 0
            Exit Sub
        End If
        REG_Set REG_KEY_PATH, g_sLastFolder
        SC_Validate Me
        SC_UpdateSum2 Me
        If g_bHoverRunEnabled Then RunClick
    End If
End Sub

Private Sub lblOptSub_Click()
    OptSubClick
End Sub

Private Sub lblOptForce_Click()
    OptForceClick
End Sub

Private Sub lblOptAudit_Click()
    OptAuditClick
End Sub

Private Sub lblBrowse_Click()
    BrowseClick
End Sub

Private Sub lblBrowseTxt_Click()
    BrowseClick
End Sub

Private Sub lblRun_Click()
    RunClick
End Sub

Private Sub lblRunTxt_Click()
    RunClick
End Sub

Private Sub lblTheme_Click()
    ThemeClick
End Sub

Private Sub lblThemeTxt_Click()
    ThemeClick
End Sub

Private Sub lblReset_Click()
    ResetClick
End Sub

Private Sub lblResetTxt_Click()
    ResetClick
End Sub

Private Sub lblCancel_Click()
    If g_bInProcessLoop Then
        g_bCancel = True
    Else
        Unload Me
    End If
End Sub

Private Sub lblCancelTxt_Click()
    If g_bInProcessLoop Then
        g_bCancel = True
    Else
        Unload Me
    End If
End Sub

' END: frmScan v9.19.9
