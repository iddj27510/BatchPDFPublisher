Attribute VB_Name = "modWinAPI"
Option Explicit
' Module    : modWinAPI
' Purpose   : Win32 API declarations shared across the project.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6:
'   Version aligned with PRJ_VERSION. No functional changes.
'
' v9.19.3 (Header standardization):
'   Header converted from Style A (NAZIV/TIP/OPIS) to Style B
'   (Module/Purpose/Reference/Version). Serbian comments in the
'   header were converted to English. Version bumped to match
'   PRJ_VERSION. No code changes.
'
' v9.18.2: Initial.
'
' Sleep was called from modRetry.RETRY_Pause and
' modUIHelpers.UI_SleepResponsive but never declared anywhere.
' This module provides it as Public so callers do not need to
' re-declare it.
'
' Other Win32 API usage in the project is declared locally in the
' module that needs it:
'   GetDiskFreeSpaceExW        modFileInfo
'   GetCurrentProcessId        modLog, modPDFFileIO
'   QueryPerformanceCounter    modTime
'   QueryPerformanceFrequency  modTime
'   SystemParametersInfoW      modPalette
' Those are kept local because they are single-consumer. Sleep has
' two consumers, so it lives here.
'
' Both VBA7 (64-bit Office) and pre-VBA7 (32-bit Office) signatures
' are provided. The kernel32 Sleep signature is identical in both
' cases — the PtrSafe keyword and the Long parameter type are the
' only differences, and Long is the correct width for Sleep's
' dwMilliseconds parameter in both.
'
' Requires  : kernel32.dll
' Notes     : No line-continuation characters in this module.
'             No decorative separator comment lines.

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' END: modWinAPI v9.19.9
