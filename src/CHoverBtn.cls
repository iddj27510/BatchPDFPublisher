VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHoverBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class     : CHoverBtn
' Purpose   : Hover handler for a lblXxx + lblXxxTxt label pair
' Reference : Appendix A.13, A.103
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.6:
'   Version aligned with PRJ_VERSION.
'   Header changelog completed (was stuck at v9.18.0 while the
'   footer already claimed v9.19.6). No functional changes.
'
' v9.18.0: Header comment corrected: method is named Cleanup(), not
'          Release() as the v9.14.7 line implied. Header sync.
' v9.14.7 (Faza 3, H.14): Added Cleanup() method (originally described
'          as Release() in the header comment).
' v9.13.1: Header sync, Class_Terminate, FormRef explanation.
' v9.6: version bump.
' v9.0: Event handlers renamed; MouseUp fixed.
'
' Public WithEvents fields are required to receive MouseMove /
' MouseDown / MouseUp from the paired labels. They cannot be
' Private because the owner (modHover.PRJ_AddHover) assigns them
' from outside the class.
'
' Requires  : FW_HandleHover (modHover).
'             PRJ_LogError (modLog).
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Public WithEvents lbl As MSForms.Label
Attribute lbl.VB_VarHelpID = -1
Public WithEvents txt As MSForms.Label
Attribute txt.VB_VarHelpID = -1

Public BtnName  As String
Public BtnRole  As String
Public IsRun    As Boolean
Public IsFilter As Boolean

Public FormRef  As Object

Private Sub lbl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo EH
    FW_HandleHover Me, True, False
    Exit Sub
EH:
    PRJ_LogError "CHoverBtn.lbl_MouseMove", "[" & BtnName & "] " & Err.Description
End Sub

Private Sub lbl_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo EH
    FW_HandleHover Me, True, True
    Exit Sub
EH:
    PRJ_LogError "CHoverBtn.lbl_MouseDown", "[" & BtnName & "] " & Err.Description
End Sub

Private Sub lbl_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo EH
    FW_HandleHover Me, True, False
    Exit Sub
EH:
    PRJ_LogError "CHoverBtn.lbl_MouseUp", "[" & BtnName & "] " & Err.Description
End Sub

Private Sub txt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo EH
    FW_HandleHover Me, True, False
    Exit Sub
EH:
    PRJ_LogError "CHoverBtn.txt_MouseMove", "[" & BtnName & "] " & Err.Description
End Sub

Private Sub txt_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo EH
    FW_HandleHover Me, True, True
    Exit Sub
EH:
    PRJ_LogError "CHoverBtn.txt_MouseDown", "[" & BtnName & "] " & Err.Description
End Sub

Private Sub txt_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo EH
    FW_HandleHover Me, True, False
    Exit Sub
EH:
    PRJ_LogError "CHoverBtn.txt_MouseUp", "[" & BtnName & "] " & Err.Description
End Sub

Private Sub Class_Terminate()
    Cleanup
End Sub

' Cleanup   disconnect WithEvents references so the form can be
'   garbage-collected after Unload without dangling hover-button
'   references keeping it alive.
Public Sub Cleanup()
    On Error Resume Next
    Set lbl = Nothing
    Set txt = Nothing
    Set FormRef = Nothing
    On Error GoTo 0
End Sub

' END: CHoverBtn v9.19.9
