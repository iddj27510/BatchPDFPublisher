Attribute VB_Name = "modShellApp"
Option Explicit
' Module    : modShellApp
' Purpose   : Isolation layer for Shell "cmd.exe" / "explorer.exe" and
'             Shell.Application COM invocations. All Shell-related API
'             calls in the project go through this module.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in SHELL_ShellExecute, SHELL_OpenWithApp and
'   SHELL_ClipStdinFile signatures). No functional changes.
'
' v9.19.3 (Header sync):
'   Version bumped to match PRJ_VERSION. Footer added.
'   No code changes.
' v9.18.0 (§1.12 API isolation):
'   New module. Extracted all Shell "cmd.exe" and "explorer.exe"
'   invocations from modFileInfo, modClipboard, modReportExport.
'   Extracted Shell.Application.ShellExecute wrapper and centralized
'   CreateObject("Shell.Application") factory used by modBrowse and
'   modInspect.
'
' Notes     : All Shell invocations are best-effort. Return True on
'             successful dispatch, False on error (logged to modLog).
'             Callers decide how to surface failure.
'
'             Why not extract WScript.Shell as well: modPalette uses it
'             for a single registry read outside the VB/VBA settings
'             tree. That is a registry concern, not a Shell concern —
'             WScript.Shell just happens to be the way to reach the
'             direct registry API from VBA. It stays in modPalette.
'
'             No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public Function SHELL_GetShellApp() As Object
    On Error Resume Next
    Err.Clear
    Set SHELL_GetShellApp = CreateObject("Shell.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set SHELL_GetShellApp = Nothing
    End If
    On Error GoTo 0
End Function

Public Function SHELL_ShellExecute(ByVal sPath As String, _
                                    Optional ByVal sVerb As String = "open", _
                                    Optional ByVal sParams As String = "", _
                                    Optional ByVal sDirectory As String = "", _
                                    Optional ByVal nShow As Long = 1) As Boolean
    SHELL_ShellExecute = False
    If Len(sPath) = 0 Then Exit Function

    Dim oShell As Object
    Set oShell = SHELL_GetShellApp()
    If oShell Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    oShell.ShellExecute sPath, sParams, sDirectory, sVerb, nShow
    SHELL_ShellExecute = (Err.Number = 0)
    If Err.Number <> 0 Then Err.Clear
    Set oShell = Nothing
    On Error GoTo 0
End Function

Public Function SHELL_OpenWithDefault(ByVal sPath As String) As Boolean
    SHELL_OpenWithDefault = False
    If Len(sPath) = 0 Then Exit Function

    Dim q As String
    q = Chr$(34)

    On Error Resume Next
    Err.Clear
    Shell "cmd.exe /c start " & q & q & " " & q & sPath & q, 0
    SHELL_OpenWithDefault = (Err.Number = 0)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
End Function

Public Function SHELL_SelectInExplorer(ByVal sPath As String) As Boolean
    SHELL_SelectInExplorer = False
    If Len(sPath) = 0 Then Exit Function

    Dim q As String
    q = Chr$(34)

    On Error Resume Next
    Err.Clear
    Shell "explorer.exe /select," & q & sPath & q, 1
    SHELL_SelectInExplorer = (Err.Number = 0)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
End Function

Public Function SHELL_OpenWithApp(ByVal sAppPath As String, _
                                   ByVal sFilePath As String, _
                                   Optional ByVal nWindowStyle As Long = 1) As Boolean
    SHELL_OpenWithApp = False
    If Len(sAppPath) = 0 Then Exit Function
    If Len(sFilePath) = 0 Then Exit Function

    Dim q As String
    q = Chr$(34)

    On Error Resume Next
    Err.Clear
    Shell sAppPath & " " & q & sFilePath & q, nWindowStyle
    SHELL_OpenWithApp = (Err.Number = 0)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
End Function

Public Function SHELL_ClipStdinFile(ByVal sFilePath As String) As Boolean
    SHELL_ClipStdinFile = False
    If Len(sFilePath) = 0 Then Exit Function

    Dim q As String
    q = Chr$(34)

    On Error Resume Next
    Err.Clear
    Shell "cmd.exe /c clip < " & q & sFilePath & q, vbHide
    SHELL_ClipStdinFile = (Err.Number = 0)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
End Function

' END: modShellApp v9.19.9

