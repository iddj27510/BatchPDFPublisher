Attribute VB_Name = "modReportExport"
Option Explicit
' Module    : modReportExport
' Purpose   : Report file export (write text + launch default .txt editor)
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   No functional changes.
'
' v9.19.3 (§10.5 Duplicate logic):
'   PRJ_ExportReport now uses modSandbox.PRJ_WriteTextFile — the
'   canonical text file writer introduced in modSandbox v9.19.3.
'   The three private helpers (RP_WriteTextFile, RP_WriteViaAdo,
'   RP_WriteViaFso) were duplicates of the modSandbox logic and
'   are removed. One writer, one code path — matches the pattern
'   already applied to modError.Write_CrashReport.
' v9.18.0 (§8.11 File system paths):
'   PRJ_ExportReport uses PRJ_PathCombine for report path.
' v9.18.0 (§1.12 API isolation):
'   PRJ_ExportReport delegates to modShellApp.
' v9.16.4 (Faza 7.2a, L.8):
'   Uses ShellExecute instead of Shell "notepad.exe".
' v9.13.5 (ARCH-23): RP_WriteTextFile split.
' v9.13.4 (UTF-8 FIX): RP_WriteViaAdo uses ADODB.Stream Charset="utf-8".
'
' Requires  : PRJ_GetReportText (modReportBuild).
'             PRJ_EnsureSandbox, PRJ_SandboxPath, PRJ_PathCombine,
'             PRJ_WriteTextFile (modSandbox).
'             SHELL_ShellExecute, SHELL_OpenWithApp (modShellApp).
'             PRJ_LogError (modLog).
'
' Notes     : No line-continuation characters in this module.
'             No decorative separator comment lines.

Public Function PRJ_ExportReport() As Boolean
    PRJ_ExportReport = False
    On Error GoTo EH
    PRJ_EnsureSandbox

    Dim sPath As String
    sPath = PRJ_PathCombine(PRJ_SandboxPath(), "batch_report.txt")

    If Not PRJ_WriteTextFile(sPath, PRJ_GetReportText()) Then
        PRJ_LogError "PRJ_ExportReport", "Could not write " & sPath
        Exit Function
    End If

    If Not SHELL_ShellExecute(sPath, "open", "", "", 1) Then
        SHELL_OpenWithApp "notepad.exe", sPath, vbNormalFocus
    End If

    PRJ_ExportReport = True
    Exit Function

EH:
    PRJ_LogError "PRJ_ExportReport", Err.Description
End Function

' END: modReportExport v9.19.9
