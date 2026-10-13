Attribute VB_Name = "modSelfTest"
Option Explicit
' Module    : modSelfTest
' Purpose   : Startup health checks - environment verification and
'             accessibility audit.
' Reference : Appendix A.13, A.61
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Notes section corrected: this module DOES use line-continuation
'   characters (in the CheckFormA11y signature and in several
'   multi-line string concatenations). No functional changes.
'
' v9.19.6 (Memory leak fix):
'   CheckFormA11y no longer leaks form instances. Prior to this
'   fix, `VBA.UserForms.Add(sFormName)` created a new instance on
'   every call and the reference was dropped without `Unload`.
'   Every PRJ_CheckAccessibility call leaked 3 form instances.
'   Now: check whether the form is already loaded; if not, load
'   it, use it, and unload it. If it is already loaded, use the
'   existing instance without unloading.
'
'   PRJ_CheckMSForms uses the same pattern (Load + Unload) instead
'   of `New frmScan`.
'
' v9.19.3 (Header standardization).
' v9.19.3 (§9.10): PRJ_CheckHostVersion moved to modEntry.
' v9.18.6 (§17.3, §17.10): Added PRJ_CheckAccessibility.
' v9.18.0: PRJ_CheckMSForms no longer calls Unload on `New`.
'
' Requires  : modConstants, modCOM, modSandbox, modFileInfo,
'             modState, modLog.
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures and multi-line string
'             concatenations are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public Function PRJ_SelfTest(ByRef sReport As String) As Boolean
    Dim colResults As Object
    Dim nFail      As Long
    Dim sOut       As String
    Dim v          As Variant
    Dim sVer       As String
    Dim nMajor     As Long
    Dim dFree      As Double
    Dim oApp       As Object
    Dim sA11y      As String

    Set colResults = CreateObject("System.Collections.ArrayList")
    nFail = 0

    If PRJ_CheckFSO() Then
        colResults.Add "[PASS] Scripting.FileSystemObject"
    Else
        colResults.Add "[FAIL] Scripting.FileSystemObject"
        nFail = nFail + 1
    End If

    If PRJ_CheckSandbox() Then
        colResults.Add "[PASS] Sandbox: " & PRJ_SandboxPath()
    Else
        colResults.Add "[FAIL] Sandbox: " & PRJ_SandboxPath()
        nFail = nFail + 1
    End If

    If PRJ_CheckWriteAccess() Then
        colResults.Add "[PASS] Write access"
    Else
        colResults.Add "[FAIL] Write access"
        nFail = nFail + 1
    End If

    If PRJ_CheckDiskSpace(PRJ_SandboxPath(), MIN_FREE_DISK_BYTES, dFree) Then
        colResults.Add "[PASS] Disk space: " & PRJ_FormatFileSize(dFree)
    Else
        colResults.Add "[WARN] Low disk space: " & PRJ_FormatFileSize(dFree)
    End If

    On Error Resume Next
    Err.Clear
    Set oApp = PRJ_GetApprentice()
    If Err.Number <> 0 Then
        Err.Clear
        Set oApp = Nothing
    End If
    On Error GoTo 0
    If Not oApp Is Nothing Then
        colResults.Add "[PASS] Apprentice Server available"
        Set oApp = Nothing
    Else
        colResults.Add "[INFO] Apprentice not available (fallback to full Inventor)"
    End If

    On Error Resume Next
    Err.Clear
    sVer = ThisApplication.SoftwareVersion.DisplayName
    nMajor = ThisApplication.SoftwareVersion.Major
    If Err.Number <> 0 Then
        sVer = "unknown"
        nMajor = 0
        Err.Clear
    End If
    On Error GoTo 0

    If nMajor = 0 Then
        colResults.Add "[FAIL] Inventor version could not be read"
        nFail = nFail + 1
    ElseIf HOST_IsSupportedMajor(nMajor) Then
      colResults.Add "[PASS] Inventor version: " & sVer
    Else
        colResults.Add "[FAIL] Unsupported Inventor version: " & sVer & " (this add-in requires 2026 or 2027)"
     nFail = nFail + 1
    End If

    If PRJ_CheckMSForms() Then
        colResults.Add "[PASS] MSForms available"
    Else
        colResults.Add "[FAIL] MSForms not available"
        nFail = nFail + 1
    End If

    If PRJ_CheckRegistry() Then
        colResults.Add "[PASS] Registry access"
    Else
        colResults.Add "[FAIL] Registry access"
        nFail = nFail + 1
    End If

    If PRJ_CheckAdoStream() Then
        colResults.Add "[PASS] ADODB.Stream available"
    Else
        colResults.Add "[WARN] ADODB.Stream not available (UTF-8 fallback to FSO)"
    End If

    If PRJ_CheckShellApp() Then
        colResults.Add "[PASS] Shell.Application available"
    Else
        colResults.Add "[WARN] Shell.Application not available"
    End If

    If PRJ_CheckVBAEnabler() Then
        colResults.Add "[PASS] VBA project object model access"
    Else
        colResults.Add "[WARN] VBA project object model access disabled (diagnostics unavailable)"
    End If

    sA11y = ""
    PRJ_CheckAccessibility sA11y
    colResults.Add "[INFO] Accessibility audit:" & vbCrLf & sA11y

    sOut = "Self-test (" & colResults.Count & " checks, " & nFail & " fail):" & vbCrLf
    For Each v In colResults
        sOut = sOut & "  " & CStr(v) & vbCrLf
    Next v

    sReport = sOut
    PRJ_SelfTest = (nFail = 0)
End Function

' PRJ_CheckAccessibility   walk every form's controls and report
'   accessibility status. Advisory — does not count as FAIL.
'   v9.19.6: does not leak form instances. See module header.
Public Sub PRJ_CheckAccessibility(ByRef sReport As String)
    Dim sOut As String
    Dim nIssue As Long

    sOut = ""
    nIssue = 0

    nIssue = nIssue + CheckFormA11y("frmScan", sOut)
    nIssue = nIssue + CheckFormA11y("frmReport", sOut)
    nIssue = nIssue + CheckFormA11y("frmProgress", sOut)

    If nIssue = 0 Then
        sReport = "    All form controls pass accessibility baseline."
    Else
        sReport = "    " & nIssue & " accessibility issue(s) found:"
        sReport = sReport & vbCrLf & sOut
    End If
End Sub

' CheckFormA11y   walk one form's controls and report gaps.
'
'   v9.19.6: memory-safe. If the form is already loaded, uses the
'   existing instance without unloading it. If not loaded, loads it,
'   uses it, and unloads it before returning.
Private Function CheckFormA11y(ByVal sFormName As String, _
                                ByRef sOut As String) As Long
    Dim oFrm As Object
    Dim nIssue As Long
    Dim bWeLoaded As Boolean
    Dim i As Long

    nIssue = 0
    bWeLoaded = False

    ' Check whether the form is already in the loaded collection.
    On Error Resume Next
    Err.Clear
    For i = 0 To VBA.UserForms.Count - 1
        If VBA.UserForms(i).Name = sFormName Then
            Set oFrm = VBA.UserForms(i)
            Exit For
        End If
    Next i
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    ' If not loaded, load it and mark it for unload at the end.
    If oFrm Is Nothing Then
        On Error Resume Next
        Err.Clear
        Load VBA.UserForms.Add(sFormName)
        Set oFrm = VBA.UserForms(VBA.UserForms.Count - 1)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            sOut = sOut & "    " & sFormName & ": (could not load)" & vbCrLf
            Exit Function
        End If
        On Error GoTo 0
        bWeLoaded = True
    End If

    If oFrm Is Nothing Then Exit Function

    Dim oCtl As Object
    Dim sCtlName As String
    Dim sTag As String
    Dim sType As String
    Dim sTip As String
    Dim bTabStop As Boolean

    On Error Resume Next
    For Each oCtl In oFrm.Controls
        sCtlName = ""
        sTag = ""
        sType = ""
        sTip = ""
        bTabStop = False

        sCtlName = CStr(oCtl.Name)
        sType = TypeName(oCtl)
        sTag = CStr(oCtl.Tag)
        bTabStop = CBool(oCtl.TabStop)
        sTip = CStr(oCtl.ControlTipText)

        Dim bInteractive As Boolean
        bInteractive = False
        Select Case sType
            Case "TextBox", "CheckBox", "OptionButton", "ComboBox", "ListBox"
                bInteractive = True
            Case "Label"
                If Left$(sTag, Len(TAG_BTN_PREFIX)) = TAG_BTN_PREFIX Then
                    bInteractive = True
                End If
        End Select

        If bInteractive Then
            If Len(sTip) = 0 Then
                sOut = sOut & "    - " & sFormName & "." & sCtlName & _
                       ": no ControlTipText" & vbCrLf
                nIssue = nIssue + 1
            End If
            If Not bTabStop And sType <> "Label" Then
                sOut = sOut & "    - " & sFormName & "." & sCtlName & _
                       ": TabStop = False on interactive control" & vbCrLf
                nIssue = nIssue + 1
            End If
        End If
    Next oCtl
    On Error GoTo 0

    ' Unload only if we loaded the form ourselves.
    If bWeLoaded Then
        On Error Resume Next
        Unload oFrm
        Set oFrm = Nothing
        On Error GoTo 0
    End If

    CheckFormA11y = nIssue
End Function

Private Function PRJ_CheckFSO() As Boolean
    PRJ_CheckFSO = False
    On Error Resume Next
    Err.Clear
    PRJ_CheckFSO = Not (PRJ_FSO() Is Nothing)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Function PRJ_CheckSandbox() As Boolean
    PRJ_CheckSandbox = False
    On Error Resume Next
    Err.Clear
    PRJ_EnsureSandbox
    PRJ_CheckSandbox = PRJ_FSO().FolderExists(PRJ_SandboxPath())
    If Err.Number <> 0 Then
        Err.Clear
        PRJ_CheckSandbox = False
    End If
    On Error GoTo 0
End Function

Public Function PRJ_CheckWriteAccess() As Boolean
    PRJ_CheckWriteAccess = False
    Dim sProbe As String
    Dim ff As Integer
    On Error GoTo Fail

    PRJ_EnsureSandbox
    sProbe = PRJ_SandboxPath() & "\_selftest_probe.tmp"
    ff = FreeFile
    Open sProbe For Output As #ff
    Print #ff, "probe"
    Close #ff
    ff = 0

    On Error Resume Next
    Err.Clear
    PRJ_FSO().DeleteFile sProbe, True
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    PRJ_CheckWriteAccess = True
    Exit Function

Fail:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    On Error GoTo 0
End Function

' PRJ_CheckMSForms   verify MSForms is usable.
'   v9.19.6: uses Load + Unload instead of `New frmScan` to avoid
'   leaking a form instance per call.
Private Function PRJ_CheckMSForms() As Boolean
    PRJ_CheckMSForms = False
    Dim bAlreadyLoaded As Boolean
    Dim i As Long

    ' Check whether frmScan is already loaded.
    bAlreadyLoaded = False
    On Error Resume Next
    Err.Clear
    For i = 0 To VBA.UserForms.Count - 1
        If VBA.UserForms(i).Name = "frmScan" Then
            bAlreadyLoaded = True
            Exit For
        End If
    Next i
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    If bAlreadyLoaded Then
        PRJ_CheckMSForms = True
        Exit Function
    End If

    ' Load, verify, unload.
    On Error Resume Next
    Err.Clear
    Load frmScan
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    PRJ_CheckMSForms = True

    On Error Resume Next
    Unload frmScan
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Function PRJ_CheckRegistry() As Boolean
    PRJ_CheckRegistry = False
    Dim sOrig As String
    Dim sTest As String

    On Error Resume Next
    Err.Clear
    sOrig = REG_Get("SelfTestProbe", "")
    REG_Set "SelfTestProbe", "ok"
    sTest = REG_Get("SelfTestProbe", "")
    REG_Set "SelfTestProbe", sOrig
    PRJ_CheckRegistry = (sTest = "ok")
    If Err.Number <> 0 Then
        Err.Clear
        PRJ_CheckRegistry = False
    End If
    On Error GoTo 0
End Function

Private Function PRJ_CheckAdoStream() As Boolean
    PRJ_CheckAdoStream = False
    Dim oS As Object

    On Error Resume Next
    Err.Clear
    Set oS = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    PRJ_CheckAdoStream = Not (oS Is Nothing)
    Set oS = Nothing
End Function

Private Function PRJ_CheckShellApp() As Boolean
    PRJ_CheckShellApp = False
    Dim oS As Object

    On Error Resume Next
    Err.Clear
    Set oS = CreateObject("Shell.Application")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    PRJ_CheckShellApp = Not (oS Is Nothing)
    Set oS = Nothing
End Function

Public Function PRJ_CheckVBAEnabler() As Boolean
    PRJ_CheckVBAEnabler = False
    Dim oVBE As Object

    On Error Resume Next
    Err.Clear
    Set oVBE = ThisApplication.VBE
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    PRJ_CheckVBAEnabler = Not (oVBE Is Nothing)
    Set oVBE = Nothing
End Function

' END: modSelfTest v9.19.9

