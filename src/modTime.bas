Attribute VB_Name = "modTime"
Option Explicit
' Module    : modTime
' Purpose   : Monotonic high-resolution time source.
' Reference : Appendix A.13
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6 (Version alignment):
'   Header Version field aligned with PRJ_VERSION and the footer.
'   Notes section corrected: this module DOES use line-continuation
'   characters in the Win32 API Declare statements below (the prior
'   note claimed otherwise). No functional changes.
'
' v9.19.3 (Header standardization):
'   Header converted from Style A (NAZIV/TIP/OPIS) to Style B
'   (Module/Purpose/Reference/Version). Serbian comments in the
'   header were converted to English. Version bumped to match
'   PRJ_VERSION. No code changes.
'
' v9.18.2: Initial (extracted from inline Timer-based timing).
'
' TM_Now returns seconds since an arbitrary epoch (system boot on
' Windows via QueryPerformanceCounter). Monotonic: never goes
' backward, unaffected by DST, NTP corrections, or midnight
' rollover. Falls back to VBA.Timer (seconds-since-midnight) if
' QPC is unavailable — callers that cross midnight should still
' add SECONDS_PER_DAY on negative elapsed, as they already do.
'
' TM_Elapsed returns dEnd - dStart in seconds, clamped to >= 0.
' Compatible with both QPC and Timer fallback because both
' produce values in seconds.
'
' Extracted from inline Timer-based timing across modRetry,
' modValidate, modPDFExport, modPDFProgress, frmProgress, frmScan
' in v9.18.0.
'
' Requires  : kernel32.QueryPerformanceCounter / QueryPerformanceFrequency
' Notes     : No line-continuation characters in procedure bodies.
'             API Declare statements are excepted (multi-line
'             signatures improve readability of the Win32 bindings).
'             No decorative separator comment lines.

#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter _
        Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency _
        Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter _
        Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency _
        Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

Private m_bFreqProbed As Boolean
Private m_dFreq       As Double

' TM_Now   monotonic clock in seconds (Double).
'   QPC path: seconds since boot, high resolution (~microsecond).
'   Fallback: VBA.Timer, seconds since midnight (0..86400).
' Returns : Double — monotonic seconds; never goes backward within
'           a session under QPC; may wrap at midnight under fallback
' Throws  : (none) — QPC failure falls through to Timer
Public Function TM_Now() As Double
    Dim c As Currency

    If TM_Frequency() > 0# Then
        If QueryPerformanceCounter(c) <> 0 Then
            TM_Now = CDbl(c) / m_dFreq
            Exit Function
        End If
    End If

    ' Fallback: wall clock within current day.
    TM_Now = Timer
End Function

' TM_Elapsed   seconds between two TM_Now values, clamped to >= 0.
'   The clamp hides harmless clock weirdness (e.g. a counter reset
'   after resume-from-hibernate) without producing a negative
'   duration for logging / formatting code.
' Params  : dStart — Double, earlier TM_Now() value
'           dEnd   — Double, later TM_Now() value
' Returns : Double — dEnd - dStart, always >= 0
' Throws  : (none)
Public Function TM_Elapsed(ByVal dStart As Double, ByVal dEnd As Double) As Double
    Dim d As Double
    d = dEnd - dStart
    If d < 0# Then d = 0#
    TM_Elapsed = d
End Function

' TM_Frequency   cached QPC frequency. Returns 0 if QPC unavailable.
'   Probe runs once per session. The Currency parameter is the
'   standard VBA idiom for receiving a Windows LARGE_INTEGER
'   (both are 8 bytes). The /10000 scaling VBA applies to Currency
'   cancels out because both the counter and the frequency go
'   through the same conversion.
' Returns : Double — QPC frequency in Hz, or 0# if unavailable
Private Function TM_Frequency() As Double
    If Not m_bFreqProbed Then
        Dim c As Currency
        If QueryPerformanceFrequency(c) <> 0 Then
            m_dFreq = CDbl(c)
        Else
            m_dFreq = 0#
        End If
        m_bFreqProbed = True
    End If
    TM_Frequency = m_dFreq
End Function

' END: modTime v9.19.9

