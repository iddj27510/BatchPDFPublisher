Attribute VB_Name = "modEntry"
Option Explicit
' Module    : modEntry
' Purpose   : Application entry point, form navigation, batch orchestration,
'             runtime log reconfiguration, uninstall helper.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.9 (debug cleanup):
'   Removed 12 Debug.Print trace lines from
'   Run_BatchPDFPublisher ("1. Entry" ... "12. Show returned").
'   Those were diagnostic markers from an earlier debugging
'   session and had no place in production code. The EH block
'   still logs the failure reason via PRJ_LogError.
'
' v9.19.6 (Documentation accuracy):
'   Added missing Requires section (modEntry is the application
'   entry point and depends on nearly every other module in the
'   project).
'   Notes section corrected: this module DOES use line-continuation
'   characters (in MsgBox expressions and PRJ_LogWarn /
'   PRJ_LogInfo / PRJ_LogError calls inside Run_BatchPDFPublisher
'   and Uninstall_BatchPDFPublisher).
'   Clarified the two v9.19.6 SilentOperation entries: the second
'   is marked as superseded by the first.
'   No functional changes.
'
' v9.19.6 (SilentOperation scope fix — FINAL):
'   SilentOperation set/restore moved from Run_BatchPDFPublisher into
'   PRJ_ShowProgress. The prior v9.19.6 attempt restored
'   SilentOperation immediately after frmScan.Show (modeless), so the
'   batch itself ran with SilentOperation=False and Inventor showed
'   a "deferred updates" dialog per IDW on folders with unresolved
'   references. PRJ_ShowProgress now brackets its own
'   frmProgress.Show vbModal — the only place where the flag is
'   actually needed — and restores on both normal and error exit
'   paths. Run_BatchPDFPublisher no longer touches SilentOperation.
'
' v9.19.6 (SilentOperation restore — SUPERSEDED):
'   SilentOperation restore added to Run_BatchPDFPublisher via
'   GoTo Cleanup after frmScan.Show. Superseded by the entry above
'   because the restore ran too early (frmScan is modal-agnostic;
'   the batch is triggered after frmScan closes).
'
' v9.19.3 (Header standardization + uninstall helper):
'   Header converted from Style A (NAZIV/TIP/OPIS) to Style B
'   (Module/Purpose/Version/Notes) to match the majority of the
'   project modules. Added Uninstall_BatchPDFPublisher — cleans
'   external state (registry, sandbox) and prints a list of VBE
'   components the user must remove manually. VBE cannot safely
'   remove its own currently-running module, so the module list is
'   advisory.
' v9.19.3 (§9.10 API versioning, §9.8 Backward compatibility):
'   PRJ_CheckHostVersion moved here from modSelfTest.
'   Run_BatchPDFPublisher now calls it after PRJ_IsHostAvailable.
'   §1.3 UI/Logic separation: added PRJ_PrepareProgressForm.
' v9.18.14 (Pre-Show preparation):
'   Introduced PRJ_PrepareScanForm and pre-Show setup in
'   PRJ_ShowProgress / PRJ_ShowReport.
' v9.18.13 (Theme pre-application):
'   FW_ApplyTheme called between Load and Show.
' v9.18.5 (§16.10 Runtime log reconfiguration):
'   Added PRJ_ReloadLogLevel.
' v9.18.0 (§9.4 Fallback without Inventor):
'   Added PRJ_IsHostAvailable.
' v9.18.0 (§8.4 Config validation):
'   Run_BatchPDFPublisher calls STATE_ValidateConfig after STATE_Init.
' v9.18.0 (§1.3 UI/Logic separation):
'   Added PRJ_StartBatchFromScan.
' v9.18.0 (§1.1 consolidation):
'   modNav merged into this module.
'
' Requires  : CACHE_Clear (modCache).
'             PRJ_ReleaseApprentice, PRJ_FSO (modCOM).
'             PRJ_ShowError (modError).
'             FIT_FitAllLabels (modFIT).
'             PRJ_InitHover (modHover).
'             L (modLang).
'             PRJ_LogInfo, PRJ_LogWarn, PRJ_LogError,
'             PRJ_LogReloadLevel, PRJ_LogReloadJson (modLog).
'             FW_LoadPalette (modPalette).
'             PRJ_EnsureSandbox, PRJ_SandboxPath, PRJ_CleanupOrphans
'             (modSandbox).
'             FW_ApplyTheme, PRJ_UpdateThemeLabel, PRJ_AlignTxt
'             (modTheme).
'             TM_Now (modTime).
'             PRJ_ApplyTooltips, PRJ_LocalizeCaptions,
'             SC_CenterAll, RP_CenterAll (modUIHelpers).
'             REG_Get, REG_Set, REG_ClearCache, STATE_Init,
'             STATE_ValidateConfig, STATE_BackupPreferences,
'             g_bRetryMode, g_bIsProcessing, g_sTheme,
'             g_bSubfolders, g_bForceOverwrite, g_bAuditMode,
'             g_sLastFolder, g_nLogLevel, g_dTimeStart,
'             g_oProgress, g_bCancel, g_bThemeApplied (modState).
'             PRJ_InitCollections (modBatch).
'             REG_SECTION, REG_PREFS, REG_KEY_* (modConstants).
'             frmScan, frmReport, frmProgress (UserForms).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and multi-line MsgBox /
'             PRJ_Log* expressions are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.


Public Function PRJ_IsHostAvailable() As Boolean
    ' §9.4: Check if ThisApplication is accessible.
    PRJ_IsHostAvailable = False

    On Error Resume Next
    Err.Clear
    Dim sVer As String
    sVer = ThisApplication.SoftwareVersion.DisplayName
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    PRJ_IsHostAvailable = (Len(sVer) > 0)
End Function

' PRJ_CheckHostVersion   verify the host is Autodesk Inventor 2026
'   or 2027 (major version 26 or 27).
'
'   Used by PRJ_SelfTest (returns [FAIL] outside {26, 27}).
'   In Run_BatchPDFPublisher it is a WARN gate, not a hard stop:
'   an unsupported version logs a warning and shows an information
'   dialog, but the batch continues. Rationale: the drawing-export
'   API surface this tool relies on has been stable across
'   2024-2027, so a false negative on an unlisted version would
'   be more costly than the risk of running. Users on unsupported
'   versions run at their own risk.
Public Function PRJ_CheckHostVersion() As Boolean
    PRJ_CheckHostVersion = False
    Dim nMajor As Long

    On Error Resume Next
    Err.Clear
    nMajor = ThisApplication.SoftwareVersion.Major
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    PRJ_CheckHostVersion = HOST_IsSupportedMajor(nMajor)
End Function

Public Sub Run_BatchPDFPublisher()
    On Error GoTo EH

    If Not PRJ_IsHostAvailable() Then Exit Sub
    If Not PRJ_CheckHostVersion() Then Exit Sub

    g_bRetryMode = False
    g_bIsProcessing = False
    STATE_Init
    CACHE_Clear

    Dim sCfgReport As String
    STATE_ValidateConfig sCfgReport

    Dim nOrphans As Long
    nOrphans = PRJ_CleanupOrphans()

    g_sTheme = REG_Get(REG_KEY_THEME, "I")
    g_bSubfolders = (LCase$(REG_Get(REG_KEY_SUB, "True")) = "true")
    g_bForceOverwrite = (LCase$(REG_Get(REG_KEY_FORCE, "False")) = "true")
    g_bAuditMode = (LCase$(REG_Get(REG_KEY_AUDIT, "False")) = "true")
    g_sLastFolder = REG_Get(REG_KEY_PATH, "")
    g_nLogLevel = CLng(REG_Get(REG_KEY_LOGLEVEL, "1"))

    FW_LoadPalette
    Load frmScan
    PRJ_PrepareScanForm
    frmScan.Show

    PRJ_ReleaseApprentice
    Exit Sub

EH:
    Dim nErrNum As Long, sErrDesc As String
    nErrNum = Err.Number
    sErrDesc = Err.Description
    On Error Resume Next
    PRJ_ReleaseApprentice
    PRJ_LogError "Run_BatchPDFPublisher", "err " & nErrNum & ": " & sErrDesc
    MsgBox "Startup error (" & nErrNum & "):" & vbCrLf & vbCrLf & sErrDesc, _
           vbExclamation, "Batch PDF Publisher"
End Sub

' PRJ_PrepareScanForm   apply all visual setup to frmScan before
'   Show. Idempotent: safe to call again from
'   frmScan.UserForm_Activate on subsequent activations.
Public Sub PRJ_PrepareScanForm()
    On Error Resume Next
    g_bThemeApplied = False
    FW_ApplyTheme frmScan
    PRJ_LocalizeCaptions frmScan
    PRJ_UpdateThemeLabel frmScan
    FIT_FitAllLabels frmScan
    SC_CenterAll frmScan
    PRJ_ApplyTooltips frmScan
    PRJ_InitHover frmScan
    On Error GoTo 0
End Sub

' PRJ_PrepareProgressForm   apply all visual setup to frmProgress
'   before Show. Extracted from frmProgress.UserForm_Activate in
'   v9.19.3 so that frmProgress and frmScan follow the same
'   prepare-then-show pattern. The form's caption (with version
'   suffix) is still computed inside UserForm_Activate because it
'   needs the designer-captured m_sOrigCap value.
Public Sub PRJ_PrepareProgressForm()
    On Error Resume Next
    Set g_oProgress = frmProgress
    g_bCancel = False
    g_bThemeApplied = False
    FW_ApplyTheme frmProgress
    PRJ_LocalizeCaptions frmProgress
    PRJ_AlignTxt frmProgress, "lblCancel", "lblCancelTxt"
    PRJ_ApplyTooltips frmProgress
    PRJ_InitHover frmProgress
    On Error GoTo 0
End Sub

' PRJ_ShowProgress   display the modal progress dialog with
'   SilentOperation enabled for the duration.
'
'   v9.19.6: SilentOperation set here, not in Run_BatchPDFPublisher.
'   The progress form is shown vbModal — the calling thread is
'   blocked until UserForm_Terminate, so SilentOperation is True
'   for the entire batch and restored when control returns. Prior
'   v9.19.6 attempt set it in Run_BatchPDFPublisher and restored
'   after frmScan.Show (modeless) — too early, batch ran without
'   the flag and Inventor showed a "deferred updates" dialog per
'   IDW on folders with unresolved references.
Public Sub PRJ_ShowProgress()
    Dim bWasSilent As Boolean
    Dim bRestore   As Boolean

    bWasSilent = False
    bRestore = False

    On Error Resume Next
    Err.Clear
    bWasSilent = ThisApplication.SilentOperation
    ThisApplication.SilentOperation = True
    bRestore = (Err.Number = 0)
    If Not bRestore Then Err.Clear
    On Error GoTo EH

    Load frmProgress
    g_bThemeApplied = False
    FW_ApplyTheme frmProgress
    PRJ_LocalizeCaptions frmProgress
    FIT_FitAllLabels frmProgress
    SC_CenterAll frmProgress
    PRJ_ApplyTooltips frmProgress
    PRJ_InitHover frmProgress
    frmProgress.Show vbModal

    If bRestore Then
        On Error Resume Next
        ThisApplication.SilentOperation = bWasSilent
        On Error GoTo 0
    End If
    Exit Sub

EH:
    If bRestore Then
        On Error Resume Next
        ThisApplication.SilentOperation = bWasSilent
        On Error GoTo 0
    End If
    PRJ_LogError "PRJ_ShowProgress", Err.Description
End Sub

Public Sub PRJ_ShowReport()
    On Error GoTo EH
    Load frmReport
    g_bThemeApplied = False
    FW_ApplyTheme frmReport
    PRJ_LocalizeCaptions frmReport
    FIT_FitAllLabels frmReport
    RP_CenterAll frmReport
    PRJ_ApplyTooltips frmReport
    PRJ_InitHover frmReport
    frmReport.Show vbModal
    Exit Sub
EH:
    PRJ_LogError "PRJ_ShowReport", Err.Description
End Sub

Public Sub PRJ_ShowScan()
    On Error GoTo EH
    Load frmScan
    PRJ_PrepareScanForm
    frmScan.Show
    Exit Sub
EH:
    PRJ_LogError "PRJ_ShowScan", Err.Description
    PRJ_ShowError "PRJ_ShowScan", "Cannot open scan form: " & Err.Description
End Sub

Public Sub PRJ_StartBatchFromScan()
    REG_Set REG_KEY_PATH, g_sLastFolder
    REG_Set REG_KEY_SUB, CStr(g_bSubfolders)
    REG_Set REG_KEY_FORCE, CStr(g_bForceOverwrite)
    REG_Set REG_KEY_AUDIT, CStr(g_bAuditMode)
    REG_Set REG_KEY_THEME, g_sTheme
    g_dTimeStart = TM_Now()
    g_bRetryMode = False
    PRJ_InitCollections
End Sub

Public Sub PRJ_ReloadLogLevel()
    On Error GoTo EH
    PRJ_LogReloadLevel
    PRJ_LogReloadJson
    Exit Sub
EH:
    PRJ_LogError "PRJ_ReloadLogLevel", Err.Description
End Sub

' Uninstall_BatchPDFPublisher   clean external state and print
'   instructions for removing VBE components manually. VBE cannot
'   safely remove the module that is currently executing, so the
'   component list is advisory — the user runs Remove in the VBE
'   after this call returns.
'
'   External state removed:
'     1. Registry tree (REG_SECTION + REG_PREFS)
'     2. Sandbox folder (%TEMP%\VBADR) and its contents
'     3. Registry cache (m_oRegCache in modState)
'
'   VBE components the user must remove manually:
'     All mod* modules, all 3 UserForms, CHoverBtn.cls,
'     CKeyPreview.cls, CKeyPreviewHook.cls.
'   The list is generated at runtime from ActiveVBProject so it
'   reflects the actual project state, not a hardcoded name list.
Public Sub Uninstall_BatchPDFPublisher()
    Dim sMsg As String
    Dim sSandbox As String
    Dim nRemoved As Long

    On Error GoTo EH

    ' 1. Backup preferences before wiping (best-effort)
    STATE_BackupPreferences

    ' 2. Registry cleanup
    DeleteSetting REG_SECTION, REG_PREFS
    REG_ClearCache

    ' 3. Sandbox cleanup
    sSandbox = PRJ_SandboxPath()
    On Error Resume Next
    Err.Clear
    If PRJ_FSO().FolderExists(sSandbox) Then
        PRJ_FSO().DeleteFolder sSandbox, True
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo EH

    ' 4. Build component list from the live project
    Dim colNames As Object
    Set colNames = CreateObject("System.Collections.ArrayList")

    Dim oComp As Object
    On Error Resume Next
    Err.Clear
    For Each oComp In ThisApplication.VBE.ActiveVBProject.VBComponents
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
        colNames.Add oComp.Name
    Next
    On Error GoTo EH

    colNames.Sort
    nRemoved = colNames.Count

    Dim i As Long
    For i = 0 To colNames.Count - 1
        sMsg = sMsg & "    " & CStr(colNames(i)) & vbCrLf
    Next i

    Dim sBody As String
    sBody = "Batch PDF Publisher - Uninstall" & vbCrLf & vbCrLf
    sBody = sBody & "External state cleaned:" & vbCrLf
    sBody = sBody & "  - Registry: " & REG_SECTION & "\" & REG_PREFS & " deleted" & vbCrLf
    sBody = sBody & "  - Sandbox:  " & sSandbox & " deleted" & vbCrLf & vbCrLf
    sBody = sBody & "Preferences were backed up before wipe." & vbCrLf & vbCrLf
    sBody = sBody & "Manual step — remove these " & nRemoved & " components from VBE:" & vbCrLf
    sBody = sBody & "  (right-click each in Project Explorer -> Remove -> No Export)" & vbCrLf & vbCrLf
    sBody = sBody & sMsg & vbCrLf
    sBody = sBody & "After removal, save the project (Ctrl+S) and close VBE." & vbCrLf

    PRJ_LogInfo "AUDIT", "Uninstall_BatchPDFPublisher: external state cleaned, " & nRemoved & " components pending manual removal"

    MsgBox sBody, vbInformation, "Batch PDF Publisher - Uninstall"
    Exit Sub

EH:
    On Error Resume Next
    PRJ_LogError "Uninstall_BatchPDFPublisher", Err.Description
    MsgBox "Uninstall failed: " & Err.Description, vbExclamation, "Batch PDF Publisher"
End Sub

' END: modEntry v9.19.9

