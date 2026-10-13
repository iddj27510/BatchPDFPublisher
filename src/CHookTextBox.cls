VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHookTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class     : CHookTextBox
' Purpose   : WithEvents wrapper for MSForms.TextBox KeyDown
' Reference : Appendix A.13, TD-19
' Version   : 9.19.9 (Inventor 2027)
'
' Part of the v9.19.7 CKeyPreview refactor. The prior single
' CKeyPreviewHook class declared WithEvents on MSForms.Control —
' a base type that does not expose KeyDown. VBA accepted the
' declaration at compile time but raised error 459 ("Object or
' class does not support the set of events") the moment a
' concrete control was assigned. The fix is one class per
' concrete control type that actually has a KeyDown event.
'
' Requires  : CKeyPreview (Friend DispatchKeyDown)
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.

Private WithEvents m_ctl As MSForms.TextBox
Attribute m_ctl.VB_VarHelpID = -1
Private m_oParent As CKeyPreview
Private m_sCallback As String

Public Sub Attach(ByVal oCtl As MSForms.TextBox, _
                   ByVal oParent As CKeyPreview, _
                   ByVal sCallback As String)
    Set m_ctl = oCtl
    Set m_oParent = oParent
    m_sCallback = sCallback
End Sub

Public Sub Detach()
    Set m_ctl = Nothing
    Set m_oParent = Nothing
    m_sCallback = ""
End Sub

Private Sub m_ctl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                           ByVal Shift As Integer)
    On Error Resume Next
    If Not m_oParent Is Nothing Then
        m_oParent.DispatchKeyDown KeyCode, Shift
    End If
    If Err.Number <> 0 Then
        Debug.Print "[CHookTextBox] DispatchKeyDown error " & _
                    Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' END: CHookTextBox v9.19.9
