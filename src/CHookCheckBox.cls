VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHookCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class     : CHookCheckBox
' Purpose   : WithEvents wrapper for MSForms.CheckBox KeyDown
' Reference : Appendix A.13, TD-19
' Version   : 9.19.9 (Inventor 2027)
'
' Requires  : CKeyPreview (Friend DispatchKeyDown)
' Notes     : No line-continuation characters in procedure bodies.

Private WithEvents m_ctl As MSForms.CheckBox
Attribute m_ctl.VB_VarHelpID = -1
Private m_oParent As CKeyPreview
Private m_sCallback As String

Public Sub Attach(ByVal oCtl As MSForms.CheckBox, _
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
        Debug.Print "[CHookCheckBox] DispatchKeyDown error " & _
                    Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' END: CHookCheckBox v9.19.9
