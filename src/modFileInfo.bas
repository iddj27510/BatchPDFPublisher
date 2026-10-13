Attribute VB_Name = "modFileInfo"
Option Explicit
' Module    : modFileInfo
' Purpose   : File metadata queries, value formatters, disk-space queries,
'             open-file-or-select-in-explorer helper
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Documentation accuracy):
'   Requires section corrected: SECONDS_PER_DAY and RELPATH_MAX_LEN
'   live in modConstants, not modState; only g_sLastFolder is in
'   modState. Removed PRJ_LogError — the module never calls it
'   (only PRJ_LogWarn).
'   Notes section corrected: this module DOES use line-continuation
'   characters (in the GetDiskFreeSpaceExW API Declare statements
'   and in the PRJ_LogWarn call inside PRJ_CheckDiskSpace).
'   No functional changes.
'
' v9.19.6: Header sync to PRJ_VERSION. No functional changes.
' v9.19.5: Header sync to PRJ_VERSION. No functional changes.
' v9.19.4: Header sync to PRJ_VERSION. No functional changes.
' v9.19.3: Header sync to PRJ_VERSION. No functional changes.
' v9.19.2 (§8.4 Date/time locale):
'   PRJ_FormatDate now uses Windows Region settings via the
'   built-in VBA named formats "Short Date" and "Short Time".
'   The old L("report.date.format") indirection is gone: per-
'   language hardcoded templates ignored the user's actual
'   Windows locale (e.g. an English Windows user running the
'   PL UI was still shown "dd.MM.yyyy"). The named VBA formats
'   already consult the locale on every call.
'   Added PRJ_FormatNow as a convenience wrapper for report
'   headers and log timestamps.
'
' v9.18.0 (§3.19 Implicit conversions):
'   Format -> Format$ in PRJ_FormatElapsed / PRJ_FormatFileSize / PRJ_FormatDate.
' v9.18.0 (§1.12 API isolation):
'   PRJ_OpenPathOrSelect delegates to modShellApp.
' v9.18.0 (§1.3 UI/Logic separation):
'   Added PRJ_OpenPathOrSelect.
'
' Requires  : PRJ_FSO (modCOM).
'             SHELL_OpenWithDefault, SHELL_SelectInExplorer (modShellApp).
'             PRJ_LogWarn (modLog).
'             SECONDS_PER_DAY, RELPATH_MAX_LEN (modConstants).
'             g_sLastFolder (modState).
'
' Notes     : No line-continuation characters in procedure bodies.
'             API Declare statements and multi-line PRJ_LogWarn
'             calls are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

#If VBA7 Then
    Private Declare PtrSafe Function GetDiskFreeSpaceExW Lib "kernel32" (ByVal lpDirectoryName As LongPtr, _
                        ByRef lpFreeBytesAvailable As Currency, ByRef lpTotalNumberOfBytes As Currency, ByRef lpTotalNumberOfFreeBytes As Currency) As Long
#Else
    Private Declare Function GetDiskFreeSpaceExW Lib "kernel32" (ByVal lpDirectoryName As Long, _
                        ByRef lpFreeBytesAvailable As Currency, ByRef lpTotalNumberOfBytes As Currency, ByRef lpTotalNumberOfFreeBytes As Currency) As Long
#End If

Public Function PRJ_FormatElapsed(dSec As Double) As String
    Dim nSec As Long
    Dim h As Long
    Dim m As Long
    Dim s As Double

    If dSec < 0 Then dSec = dSec + SECONDS_PER_DAY

    nSec = Int(dSec)
    h = nSec \ 3600
    m = (nSec Mod 3600) \ 60
    s = dSec - (h * 3600) - (m * 60)
    If s < 0 Then s = 0

    If h > 0 Then
        PRJ_FormatElapsed = h & " h " & m & " min " & Format$(s, "0.00") & " s"
    ElseIf m > 0 Then
        PRJ_FormatElapsed = m & " min " & Format$(s, "0.00") & " s"
    Else
        PRJ_FormatElapsed = Format$(s, "0.00") & " s"
    End If
End Function

Public Function PRJ_FormatFileSize(bBytes As Double) As String
    If bBytes < 0 Then
        PRJ_FormatFileSize = "unknown"
        Exit Function
    End If
    If bBytes >= 1073741824# Then
        PRJ_FormatFileSize = Format$(bBytes / 1073741824#, "#,##0.0") & " GB"
    ElseIf bBytes >= 1048576# Then
        PRJ_FormatFileSize = Format$(bBytes / 1048576#, "#,##0.0") & " MB"
    ElseIf bBytes >= 1024# Then
        PRJ_FormatFileSize = Format$(bBytes / 1024#, "#,##0") & " KB"
    Else
        PRJ_FormatFileSize = Format$(bBytes, "#,##0") & " B"
    End If
End Function

' PRJ_FormatDate   format a file timestamp using the Windows
'   Region settings (Short Date + Short Time). The named VBA
'   formats "Short Date" / "Short Time" consult the current
'   user locale on every call, so the output matches Windows
'   Explorer's "Date modified" column and every other native
'   application on the machine.
' Params  : dDate — Date, file timestamp
' Returns : String — localized date + time, or "unknown" for
'           zero / pre-1980 sentinels (FAT epoch guard)
Public Function PRJ_FormatDate(dDate As Date) As String
    If dDate = 0 Then
        PRJ_FormatDate = "unknown"
        Exit Function
    End If
    If Year(dDate) < 1980 Then
        PRJ_FormatDate = "unknown"
        Exit Function
    End If
    PRJ_FormatDate = Format$(dDate, "Short Date") & " " & Format$(dDate, "Short Time")
End Function

' PRJ_FormatNow   current wall-clock time in Windows locale format.
'   Convenience for report headers and UI banners that used to
'   call Format$(Now, L("report.date.format")).
Public Function PRJ_FormatNow() As String
    PRJ_FormatNow = Format$(Now, "Short Date") & " " & Format$(Now, "Short Time")
End Function

Public Function PRJ_GetFileSize(sPath As String) As Double
    Dim oF As Object

    PRJ_GetFileSize = -1

    On Error Resume Next
    Err.Clear
    Set oF = PRJ_FSO().GetFile(sPath)
    On Error GoTo 0
    If oF Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    PRJ_GetFileSize = oF.Size
    On Error GoTo 0
End Function

Public Function PRJ_GetFileDate(sPath As String) As Date
    Dim oF As Object

    PRJ_GetFileDate = 0

    On Error Resume Next
    Err.Clear
    Set oF = PRJ_FSO().GetFile(sPath)
    On Error GoTo 0
    If oF Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    PRJ_GetFileDate = oF.DateLastModified
    On Error GoTo 0
End Function

Public Function PRJ_GetPDFSize(sIDW As String) As String
    Dim nDot As Long
    Dim sPDF As String

    nDot = InStrRev(sIDW, ".")
    If nDot = 0 Then
        PRJ_GetPDFSize = "-"
        Exit Function
    End If

    sPDF = Left$(sIDW, nDot) & "pdf"
    If Not PRJ_FSO().FileExists(sPDF) Then
        PRJ_GetPDFSize = "-"
        Exit Function
    End If

    PRJ_GetPDFSize = PRJ_FormatFileSize(PRJ_GetFileSize(sPDF))
End Function

Public Function PRJ_GetRelPath(sPath As String) As String
    Dim sRoot  As String
    Dim sRel   As String
    Dim sFile  As String
    Dim sDir   As String
    Dim parts() As String
    Dim n      As Long

    If Len(g_sLastFolder) = 0 Then
        PRJ_GetRelPath = sPath
        Exit Function
    End If

    sRoot = g_sLastFolder
    If Right$(sRoot, 1) <> "\" Then sRoot = sRoot & "\"

    If LCase$(Left$(sPath, Len(sRoot))) = LCase$(sRoot) Then
        sRel = Mid$(sPath, Len(sRoot) + 1)
    Else
        sRel = sPath
    End If

    If Len(sRel) > RELPATH_MAX_LEN Then
        sFile = PRJ_FSO().GetFileName(sRel)
        sDir = PRJ_FSO().GetParentFolderName(sRel)
        parts = Split(sDir, "\")
        n = UBound(parts)
        If n >= 1 Then
            sRel = parts(n - 1) & "\...\" & parts(n) & "\" & sFile
        End If
    End If

    PRJ_GetRelPath = sRel
End Function

Public Function PRJ_GetFreeSpaceBytes(sPath As String) As Double
    PRJ_GetFreeSpaceBytes = -1#

    If Len(sPath) = 0 Then Exit Function
    If Not PRJ_FSO().FolderExists(sPath) Then
        Dim sParent As String
        sParent = PRJ_FSO().GetParentFolderName(sPath)
        If Len(sParent) = 0 Then Exit Function
        If Not PRJ_FSO().FolderExists(sParent) Then Exit Function
        sPath = sParent
    End If

    Dim cFree      As Currency
    Dim cTotal     As Currency
    Dim cFreeTotal As Currency
    Dim lRet       As Long

    On Error Resume Next
    Err.Clear
    lRet = GetDiskFreeSpaceExW(StrPtr(sPath), cFree, cTotal, cFreeTotal)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If lRet = 0 Then Exit Function

    PRJ_GetFreeSpaceBytes = CDbl(cFree) * 10000#
End Function

Public Function PRJ_CheckDiskSpace(ByVal sPath As String, ByVal nRequiredBytes As Double, ByRef dFreeBytes As Double) As Boolean
    dFreeBytes = PRJ_GetFreeSpaceBytes(sPath)

    If dFreeBytes < 0# Then
        PRJ_CheckDiskSpace = True
        Exit Function
    End If

    If dFreeBytes = 0# Then
        PRJ_LogWarn "PRJ_CheckDiskSpace", "Free-space probe returned 0 for " & sPath & " (assuming OK)"
        PRJ_CheckDiskSpace = True
        Exit Function
    End If

    PRJ_CheckDiskSpace = (dFreeBytes >= nRequiredBytes)
End Function

Public Function PRJ_OpenPathOrSelect(ByVal sPath As String) As Boolean
    PRJ_OpenPathOrSelect = False
    If Len(sPath) = 0 Then Exit Function

    Dim nDot As Long
    nDot = InStrRev(sPath, ".")
    If nDot = 0 Then Exit Function
    Dim sPDF As String
    sPDF = Left$(sPath, nDot) & "pdf"

    If PRJ_FSO().FileExists(sPDF) Then
        PRJ_OpenPathOrSelect = SHELL_OpenWithDefault(sPDF)
    Else
        PRJ_OpenPathOrSelect = SHELL_SelectInExplorer(sPath)
    End If
End Function

' END: modFileInfo v9.19.9
