Attribute VB_Name = "modBrowse"
Option Explicit
' Module    : modBrowse
' Purpose   : Folder picker via Shell.Application.BrowseForFolder.
'             Wraps the COM call, makes the dialog modal to Inventor,
'             honors an initial folder when valid, and releases COM
'             references explicitly.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment + header standardization):
'   Header converted from Style A (free-form) to Style B
'   (Module/Purpose/Reference/Version). Version field aligned
'   with PRJ_VERSION and the footer.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in the two oShell.BrowseForFolder calls).
'   No functional changes.
'
' v9.19.6: Header sync to PRJ_VERSION. No functional changes.
' v9.18.0 (§1.12 API isolation):
'   CreateObject("Shell.Application") replaced with
'   modShellApp.SHELL_GetShellApp(). Only the factory is isolated;
'   oShell.Namespace / oShell.BrowseForFolder remain direct calls
'   on the returned object (a single-consumer usage pattern).
' v9.18.0: sInitial is now honored via oShell.Namespace() when valid.
'          hWnd from ThisApplication.MainFrameHWND makes the dialog
'          modal to Inventor (otherwise it can hide behind it).
'          Falls back to hWnd=0 and CSIDL_DRIVES on any failure.
' v9.14.5: Privacy fix - sInitial no longer logged; PRJ_LogInfo -> PRJ_LogDebug.
' v9.13.1: Header sync; failure handling; explicit release.
'
' Requires  : SHELL_GetShellApp (modShellApp).
'             PRJ_LogDebug, PRJ_LogError (modLog).
'             ThisApplication.MainFrameHWND (Inventor host).
'
' Notes     : No line-continuation characters in procedure bodies.
'             Multi-argument call continuation is excepted per
'             Doc 01 / Doc 06.
'             No decorative separator comment lines.

Private Const CSIDL_DRIVES         As Long = 17
Private Const BIF_RETURNONLYFSDIRS As Long = 1
Private Const BIF_NEWDIALOGSTYLE   As Long = &H40

Public Function BrowseForFolderEx(ByVal sTitle As String, ByVal sInitial As String) As String
    BrowseForFolderEx = ""
    On Error GoTo EH

    Dim oShell  As Object
    Dim oFolder As Object
    Dim oRoot   As Object

    PRJ_LogDebug "modBrowse", "BrowseForFolderEx invoked"

    Set oShell = SHELL_GetShellApp()
    If oShell Is Nothing Then
        PRJ_LogError "modBrowse.BrowseForFolderEx", "Shell.Application COM not available"
        Exit Function
    End If

    Set oRoot = Nothing
    If Len(sInitial) > 0 Then
        On Error Resume Next
        Err.Clear
        Set oRoot = oShell.Namespace(sInitial)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo EH
    End If

    Dim hWnd As Long
    hWnd = 0
    On Error Resume Next
    Err.Clear
    hWnd = ThisApplication.MainFrameHWND
    If Err.Number <> 0 Then
        Err.Clear
        hWnd = 0
    End If
    On Error GoTo EH

    If oRoot Is Nothing Then
        Set oFolder = oShell.BrowseForFolder(hWnd, sTitle, BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE, CSIDL_DRIVES)
    Else
        Set oFolder = oShell.BrowseForFolder(hWnd, sTitle, BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE, oRoot)
    End If

    If oFolder Is Nothing Then
        Set oRoot = Nothing
        Set oShell = Nothing
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    BrowseForFolderEx = oFolder.Self.Path
    If Err.Number <> 0 Then
        Err.Clear
        BrowseForFolderEx = ""
    End If
    On Error GoTo EH

    Set oFolder = Nothing
    Set oRoot = Nothing
    Set oShell = Nothing
    Exit Function

EH:
    On Error Resume Next
    Set oFolder = Nothing
    Set oRoot = Nothing
    Set oShell = Nothing
    PRJ_LogError "modBrowse.BrowseForFolderEx", "Error " & Err.Number & ": " & Err.Description
    On Error GoTo 0
    BrowseForFolderEx = ""
End Function

' END: modBrowse v9.19.9
