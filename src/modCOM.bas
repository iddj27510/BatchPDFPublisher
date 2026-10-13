Attribute VB_Name = "modCOM"
Option Explicit
' Module    : modCOM
' Purpose   : Lazy COM singletons (FSO, Apprentice) with release semantics
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.7 (BUG-02 Host version guard fix):
'   PRJ_LogApprenticeFallback threshold corrected from
'   nMajor >= 26 to nMajor >= 30. The prior value treated
'   Inventor 2022 (Major 26) as "2026+". Inventor uses an
'   offset of 1996, so 2026 is Major 30 and 2027 is Major 31.
'   The message string now reads "Inventor 2026+" instead of
'   the raw "Inventor <n>+" — Major 30 is meaningless to a
'   user reading the log. No functional change to the warning
'   logic itself; only the threshold and the human-readable
'   text are corrected.
'
' v9.19.6 (Version alignment + header restructuring):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Reordered header sections to match the project convention
'   (Requires -> Reentrancy Guard -> Notes). The v9.13.2 historical
'   note moved from Notes to the changelog where it belongs.
'   No functional changes.
'
' v9.19.6: Header sync to PRJ_VERSION. No functional changes.
' v9.18.0 (§1.14 Circular Dependencies):
'   Documented the m_bInCreateFSO guard in the header.
' v9.17.0: PRJ_FSO now logs a warning if CreateObject fails.
' v9.14.0 (ARCH-33): PRJ_GetApprentice refactored. Helper functions
'          PRJ_TestApprentice, PRJ_CreateApprentice and
'          PRJ_LogApprenticeFallback extracted. Nesting depth reduced
'          from 6 to 3.
' v9.13.7 (RULE S.2): removed all decorative separator comment lines.
' v9.13.2 (ARCH-09): extracted from modUtils to satisfy SRP.
'
' Requires  : g_oFSO, g_oApprentice (modState).
'             PRJ_LogWarn (modLog).
'
' Reentrancy Guard (this module):
'   m_bInCreateFSO — breaks cycle 3. Set before CreateObject, cleared
'                    after. If a second PRJ_FSO call comes in while the
'                    first is creating the object, it returns Nothing
'                    immediately. Also cleared before PRJ_LogWarn so the
'                    log path can safely re-enter if needed.
'
' Notes     : No line-continuation characters anywhere in this module.
'             No decorative separator comment lines.

Private m_bInCreateFSO As Boolean

Public Function PRJ_FSO() As Object
    If g_oFSO Is Nothing Then

        ' Cycle 3 guard: PRJ_LogWarn below may trigger the log open
        ' path, which calls PRJ_EnsureSandbox, which calls PRJ_FSO again.
        ' Without this guard, a failing CreateObject would recurse until
        ' "Out of stack space".
        If m_bInCreateFSO Then
            Set PRJ_FSO = Nothing
            Exit Function
        End If

        m_bInCreateFSO = True

        On Error Resume Next
        Err.Clear
        Set g_oFSO = CreateObject("Scripting.FileSystemObject")

        Dim bFailed As Boolean
        bFailed = (Err.Number <> 0 Or g_oFSO Is Nothing)
        If bFailed Then
            Err.Clear
            Set g_oFSO = Nothing
        End If
        On Error GoTo 0

        m_bInCreateFSO = False

        If bFailed Then
            PRJ_LogWarn "PRJ_FSO", "Cannot create Scripting.FileSystemObject"
            Set PRJ_FSO = Nothing
            Exit Function
        End If
    End If
    Set PRJ_FSO = g_oFSO
End Function

Public Function PRJ_GetApprentice() As Object
    If Not g_oApprentice Is Nothing Then
        If PRJ_TestApprentice(g_oApprentice) Then
            Set PRJ_GetApprentice = g_oApprentice
            Exit Function
        End If
        Set g_oApprentice = Nothing
    End If
    Set g_oApprentice = PRJ_CreateApprentice()
    If g_oApprentice Is Nothing Then
        PRJ_LogApprenticeFallback
    Else
        On Error Resume Next
        g_oApprentice.SilentOperation = True
        On Error GoTo 0
    End If
    Set PRJ_GetApprentice = g_oApprentice
End Function

Private Function PRJ_TestApprentice(ByRef oApp As Object) As Boolean
    On Error Resume Next
    Err.Clear
    Dim sTest As String
    sTest = oApp.SoftwareVersion.DisplayName
    PRJ_TestApprentice = (Err.Number = 0)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Function PRJ_CreateApprentice() As Object
    On Error Resume Next
    Err.Clear
    Set PRJ_CreateApprentice = CreateObject("Inventor.ApprenticeServerComponent")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Sub PRJ_LogApprenticeFallback()
    Dim nMajor As Long
    Dim sWarnMsg As String
    On Error Resume Next
    Err.Clear
    nMajor = ThisApplication.SoftwareVersion.Major
    If Err.Number <> 0 Then
        sWarnMsg = "Apprentice Server not available and Inventor version could not be read. Using full Inventor (slower)."
        Err.Clear
    ElseIf nMajor >= 30 Then
        sWarnMsg = "Apprentice Server not registered. Inventor 2026+ requires standalone Apprentice install. Using full Inventor (slower)."
    Else
        sWarnMsg = "Apprentice Server not available. Using full Inventor."
    End If
    On Error GoTo 0
    PRJ_LogWarn "PRJ_GetApprentice", sWarnMsg
End Sub

Public Sub PRJ_ReleaseApprentice()
    On Error Resume Next
    Set g_oApprentice = Nothing
    On Error GoTo 0
End Sub

' END: modCOM v9.19.9
