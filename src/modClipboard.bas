Attribute VB_Name = "modClipboard"
Option Explicit
' Module    : modClipboard
' Purpose   : Clipboard write with htmlfile + OLE DataObject + Shell clip
'             fallback chain
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Critical fix + version alignment):
'   Added missing Option Explicit at the top of the module. Prior
'   revision had no Option Explicit declaration — a latent risk
'   for silent typos on undeclared variable names.
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Requires section corrected: SLEEP_CLIP_MS lives in
'   modConstants, not modState. Added PRJ_EnsureSandbox
'   (modSandbox), used in PRJ_PutClipboardFallback but previously
'   undocumented.
'   No functional changes.
'
' v9.19.6:
'   PRJ_PutClipboardFallback now calls PRJ_EnsureSandbox on entry
'   and closes the TextStream explicitly before SHELL_ClipStdinFile
'   reads it. Prior revision relied on FSO.CreateTextFile flush
'   ordering, which was not guaranteed on slow machines; the clip
'   command could read a partial or empty file.
' v9.18.0 (§8.11 File system paths):
'   PRJ_PutClipboardFallback uses PRJ_PathCombine and PRJ_SandboxPath.
' v9.18.0 (§1.12 API isolation):
'   PRJ_PutClipboardFallback delegates to modShellApp.SHELL_ClipStdinFile.
' v9.18.0: Err.Clear before each Err.Number = 0 probe.
'
' Requires  : PRJ_EnsureSandbox, PRJ_SandboxPath, PRJ_PathCombine
'             (modSandbox).
'             PRJ_FSO (modCOM).
'             SHELL_ClipStdinFile (modShellApp).
'             UI_SleepResponsive (modUIHelpers).
'             SLEEP_CLIP_MS (modConstants).
'             PRJ_LogError (modLog).
'
' Notes     : PRJ_PutClipboard is shared service used by:
'               - frmReport.CopyClick
'               - modHelp.HELP_CopyApprenticeREADMEToClipboard
'             Three fallback strategies:
'               1. htmlfile.ParentWindow.ClipboardData.SetData "Text"
'               2. OLE DataObject (SetText / PutInClipboard)
'               3. Shell "cmd.exe /c clip" via temp file
'             PRJ_PutClipboard does 1 then 2.
'             PRJ_PutClipboardFallback does 3.
'             No line-continuation characters in this module.
'             No decorative separator comment lines.

Public Function PRJ_PutClipboard(ByVal sText As String) As Boolean
    PRJ_PutClipboard = False

    Dim oHtml As Object
    Dim bHtmlOk As Boolean
    Dim oDO As Object
    Dim bDoOk As Boolean

    On Error Resume Next
    Err.Clear
    Set oHtml = CreateObject("htmlfile")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    If Not oHtml Is Nothing Then
        On Error Resume Next
        Err.Clear
        oHtml.ParentWindow.ClipboardData.SetData "Text", sText
        bHtmlOk = (Err.Number = 0)
        If Not bHtmlOk Then Err.Clear
        On Error GoTo 0
        Set oHtml = Nothing

        If bHtmlOk Then
            PRJ_PutClipboard = True
            Exit Function
        End If
    End If

    On Error Resume Next
    Err.Clear
    Set oDO = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    If Not oDO Is Nothing Then
        On Error Resume Next
        Err.Clear
        oDO.SetText sText
        oDO.PutInClipboard
        bDoOk = (Err.Number = 0)
        If Not bDoOk Then Err.Clear
        On Error GoTo 0
        Set oDO = Nothing

        If bDoOk Then PRJ_PutClipboard = True
    End If
End Function

Public Function PRJ_PutClipboardFallback(ByVal sText As String) As Boolean
    PRJ_PutClipboardFallback = False
    On Error GoTo EH

    PRJ_EnsureSandbox

    Dim sTmp As String
    sTmp = PRJ_PathCombine(PRJ_SandboxPath(), "clip_tmp.txt")

    Dim f As Object
    On Error Resume Next
    Err.Clear
    Set f = PRJ_FSO().CreateTextFile(sTmp, True)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    f.Write sText
    f.Close
    Set f = Nothing

    If Not SHELL_ClipStdinFile(sTmp) Then Exit Function

    UI_SleepResponsive SLEEP_CLIP_MS

    On Error Resume Next
    If PRJ_FSO().FileExists(sTmp) Then PRJ_FSO().DeleteFile sTmp, True
    On Error GoTo 0

    PRJ_PutClipboardFallback = True
    Exit Function

EH:
    PRJ_LogError "PRJ_PutClipboardFallback", Err.Description
End Function

' END: modClipboard v9.19.9
