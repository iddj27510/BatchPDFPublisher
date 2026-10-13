VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseOverControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class     : MouseOverControl
' Purpose   : Holds a single reference to a MSForms.UserForm or a
'             MSForms.Control that is capable of having a MouseMove event.
'             Used as an asynchronous callback via the Terminate event.
' Reference : Appendix A.13
' Version   : 9.25.5 (Inventor 2027)
'
' Notes     : This class implements Scroll events for the container
'             types (frame, form and multiPage) in order to suppress
'             any scrolls that are related to the controls getting
'             the focus.
'             No decorative separator comment lines.

#Const x64 = Win64

#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Private WithEvents m_CheckBox As MSForms.CheckBox
Attribute m_CheckBox.VB_VarHelpID = -1
Private WithEvents m_ComboBox As MSForms.ComboBox
Attribute m_ComboBox.VB_VarHelpID = -1
Private WithEvents m_CommandButton As MSForms.CommandButton
Attribute m_CommandButton.VB_VarHelpID = -1
Private WithEvents m_Frame As MSForms.Frame
Attribute m_Frame.VB_VarHelpID = -1
Private WithEvents m_Image As MSForms.Image
Attribute m_Image.VB_VarHelpID = -1
Private WithEvents m_Label As MSForms.Label
Attribute m_Label.VB_VarHelpID = -1
Private WithEvents m_ListBox As MSForms.ListBox
Attribute m_ListBox.VB_VarHelpID = -1
Private WithEvents m_MultiPage As MSForms.MultiPage
Attribute m_MultiPage.VB_VarHelpID = -1
Private WithEvents m_OptionButton As MSForms.OptionButton
Attribute m_OptionButton.VB_VarHelpID = -1
Private WithEvents m_TabStrip As MSForms.TabStrip
Attribute m_TabStrip.VB_VarHelpID = -1
Private WithEvents m_TextBox As MSForms.TextBox
Attribute m_TextBox.VB_VarHelpID = -1
Private WithEvents m_ToggleButton As MSForms.ToggleButton
Attribute m_ToggleButton.VB_VarHelpID = -1
Private WithEvents m_UserForm As MSForms.UserForm
Attribute m_UserForm.VB_VarHelpID = -1

Private m_control As Object
Private m_hWndForm As LongPtr

Friend Sub MouseProc(ByVal ncode As Long, ByVal wParam As Long, ByRef lParam As MOUSEHOOKSTRUCTEX)
    ProcessMouseData ncode, wParam, lParam
End Sub

#If VBA7 Then
Public Function CreateFromControl(ByVal formControl As MSForms.Control, ByVal hWndForm As LongPtr) As MouseOverControl
#Else
Public Function CreateFromControl(ByVal formControl As MSForms.Control, ByVal hWndForm As Long) As MouseOverControl
#End If
    Dim moControl As MouseOverControl
    Set moControl = New MouseOverControl
    moControl.InitFromControl formControl, hWndForm
    Set CreateFromControl = moControl
End Function

#If VBA7 Then
Public Function CreateFromForm(ByVal uForm As MSForms.UserForm, ByVal hWndForm As LongPtr) As MouseOverControl
#Else
Public Function CreateFromForm(ByVal uForm As MSForms.UserForm, ByVal hWndForm As Long) As MouseOverControl
#End If
    Dim moControl As MouseOverControl
    Set moControl = New MouseOverControl
    moControl.InitFromForm uForm, hWndForm
    Set CreateFromForm = moControl
End Function

#If VBA7 Then
Public Sub InitFromControl(ByVal formControl As MSForms.Control, ByVal hWndForm As LongPtr)
#Else
Public Sub InitFromControl(ByVal formControl As MSForms.Control, ByVal hWndForm As Long)
#End If
    Set m_control = formControl
    m_hWndForm = hWndForm
    
    Select Case TypeName(formControl)
        Case "CheckBox":      Set m_CheckBox = formControl
        Case "ComboBox":      Set m_ComboBox = formControl
        Case "CommandButton": Set m_CommandButton = formControl
        Case "Frame":         Set m_Frame = formControl
        Case "Image":         Set m_Image = formControl
        Case "Label":         Set m_Label = formControl
        Case "ListBox":       Set m_ListBox = formControl
        Case "MultiPage":     Set m_MultiPage = formControl
        Case "OptionButton":  Set m_OptionButton = formControl
        Case "TabStrip":      Set m_TabStrip = formControl
        Case "TextBox":       Set m_TextBox = formControl
        Case "ToggleButton":  Set m_ToggleButton = formControl
    End Select
End Sub

#If VBA7 Then
Public Sub InitFromForm(ByVal uForm As MSForms.UserForm, ByVal hWndForm As LongPtr)
#Else
Public Sub InitFromForm(ByVal uForm As MSForms.UserForm, ByVal hWndForm As Long)
#End If
    Set m_control = uForm
    m_hWndForm = hWndForm
    Set m_UserForm = uForm
End Sub

Private Sub m_CheckBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_ComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_CommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_Frame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_Image_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_Label_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_ListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_MultiPage_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_OptionButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub M_TabStrip_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_TextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_ToggleButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SetHoveredControl Me
End Sub

Private Sub m_UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    If ActionY = fmScrollActionFocusRequest Or ActionY = fmScrollActionControlRequest Then ActualDy = 0
    If ActionX = fmScrollActionFocusRequest Or ActionX = fmScrollActionControlRequest Then ActualDx = 0
End Sub

Private Sub m_Frame_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    If ActionY = fmScrollActionFocusRequest Or ActionY = fmScrollActionControlRequest Then ActualDy = 0
    If ActionX = fmScrollActionFocusRequest Or ActionX = fmScrollActionControlRequest Then ActualDx = 0
End Sub

Private Sub m_MultiPage_Scroll(ByVal Index As Long, ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
    If ActionY = fmScrollActionFocusRequest Or ActionY = fmScrollActionControlRequest Then ActualDy = 0
    If ActionX = fmScrollActionFocusRequest Or ActionX = fmScrollActionControlRequest Then ActualDx = 0
End Sub

Public Function GetControl() As Object
    Set GetControl = m_control
End Function

#If VBA7 Then
Public Property Get FormHandle() As LongPtr
#Else
Public Property Get FormHandle() As Long
#End If
    FormHandle = m_hWndForm
End Property

Private Sub Class_Terminate()
    Set m_CheckBox = Nothing
    Set m_ComboBox = Nothing
    Set m_CommandButton = Nothing
    Set m_Frame = Nothing
    Set m_Image = Nothing
    Set m_Label = Nothing
    Set m_ListBox = Nothing
    Set m_MultiPage = Nothing
    Set m_OptionButton = Nothing
    Set m_TabStrip = Nothing
    Set m_TextBox = Nothing
    Set m_ToggleButton = Nothing
    Set m_UserForm = Nothing
    Set m_control = Nothing
End Sub

' END: MouseOverControl v9.25.5
