VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CKeyPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Class     : CKeyPreview
' Purpose   : Hooks KeyDown on all focusable controls of a UserForm
'             and forwards them to a form-level callback.
' Reference : Appendix A.13, TD-19
' Version   : 9.19.9 (Inventor 2027)
'
' v9.19.7 (TD-19 fix — concrete WithEvents types):
'   Prior implementation used a single CKeyPreviewHook class with
'   Private WithEvents m_ctl As MSForms.Control. MSForms.Control is
'   a base type that does not expose KeyDown; VBA raised error 459
'   ("Object or class does not support the set of events") the
'   moment a concrete control was assigned. Load frmScan failed at
'   runtime on every Inventor 2026 / 2027 host.
'
'   Fix: one hook class per concrete control type. CKeyPreview now
'   dispatches in HookOne based on TypeName(oCtl) and instantiates
'   the matching class. Each hook class declares its WithEvents
'   field with the concrete MSForms type, which VBA binds correctly.
'
'   Supported types (all present in this project's forms):
'     TextBox, CheckBox, ListBox, ComboBox, OptionButton
'   Additional control types that have KeyDown (CommandButton,
'   ToggleButton, ScrollBar, SpinButton, MultiPage, TabStrip) can
'   be added later — one class each, one Case each in HookOne.
'
' v9.19.6:
'   Version aligned with PRJ_VERSION.
'   DispatchKeyDown logs unexpected callback errors via Debug.Print
'   instead of swallowing them silently.
'
' v9.19.3 (Header standardization):
'   Header converted from Style A to Style B.
' v9.18.1: Initial.
'
' VBA MSForms.UserForm does not have a KeyPreview property. That is
' a Windows Forms (.NET) feature; FM20.DLL does not implement it.
' The only way to get form-level keyboard handling is to subscribe
' to the KeyDown event of every focusable control on the form.
'
' The callback procedure in the form must have this signature:
'
'     Public Sub OnFormKeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
'                              ByVal Shift As Integer)
'
' The callback can inspect KeyCode and, if it handles the key,
' set KeyCode = 0 to consume it.
'
' Requires  : CHookTextBox, CHookCheckBox, CHookListBox,
'             CHookComboBox, CHookOptionButton (hook classes)
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Private m_oForm     As Object
Private m_sCallback As String
Private m_colHooks  As Collection

Public Sub Initialize(ByVal oForm As Object, ByVal sCallback As String)
    Set m_oForm = oForm
    m_sCallback = sCallback
    Set m_colHooks = New Collection
    HookControlsRecursive oForm.Controls
End Sub

Public Sub Terminate()
    Dim i As Long
    If m_colHooks Is Nothing Then Exit Sub
    For i = 1 To m_colHooks.Count
        On Error Resume Next
        m_colHooks(i).Detach
        On Error GoTo 0
    Next i
    Set m_colHooks = Nothing
    Set m_oForm = Nothing
    m_sCallback = ""
End Sub

Private Sub HookControlsRecursive(ByVal oControls As Object)
    Dim oCtl As Object
    Dim oChild As Object
    Dim i As Long

    On Error Resume Next
    For i = 0 To oControls.Count - 1
        Set oCtl = oControls.Item(i)
        If Not oCtl Is Nothing Then
            If HasKeyDown(oCtl) Then
                HookOne oCtl
            End If

            ' Recurse into containers so Frame / MultiPage inner
            ' controls also forward their KeyDown events.
            Set oChild = Nothing
            On Error Resume Next
            Set oChild = oCtl.Controls
            On Error GoTo 0
            If Not oChild Is Nothing Then
                HookControlsRecursive oChild
                Set oChild = Nothing
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function HasKeyDown(ByVal oCtl As Object) As Boolean
    ' Only controls that can receive focus and expose KeyDown
    ' are hooked. Labels and Images do not.
    Dim sType As String
    On Error Resume Next
    sType = TypeName(oCtl)
    On Error GoTo 0

    Select Case sType
        Case "TextBox", "CheckBox", "ListBox", "ComboBox", "OptionButton"
            HasKeyDown = True
        Case Else
            HasKeyDown = False
    End Select
End Function

Private Sub HookOne(ByVal oCtl As Object)
    Dim sType As String
    On Error Resume Next
    sType = TypeName(oCtl)
    On Error GoTo 0

    Select Case sType
        Case "TextBox":      HookTextBox oCtl
        Case "CheckBox":     HookCheckBox oCtl
        Case "ListBox":      HookListBox oCtl
        Case "ComboBox":     HookComboBox oCtl
        Case "OptionButton": HookOptionButton oCtl
    End Select
End Sub

Private Sub HookTextBox(ByVal oCtl As Object)
    Dim oT As MSForms.TextBox
    Set oT = oCtl
    Dim h As CHookTextBox
    Set h = New CHookTextBox
    h.Attach oT, Me, m_sCallback
    m_colHooks.Add h
End Sub

Private Sub HookCheckBox(ByVal oCtl As Object)
    Dim oC As MSForms.CheckBox
    Set oC = oCtl
    Dim h As CHookCheckBox
    Set h = New CHookCheckBox
    h.Attach oC, Me, m_sCallback
    m_colHooks.Add h
End Sub

Private Sub HookListBox(ByVal oCtl As Object)
    Dim oL As MSForms.ListBox
    Set oL = oCtl
    Dim h As CHookListBox
    Set h = New CHookListBox
    h.Attach oL, Me, m_sCallback
    m_colHooks.Add h
End Sub

Private Sub HookComboBox(ByVal oCtl As Object)
    Dim oC As MSForms.ComboBox
    Set oC = oCtl
    Dim h As CHookComboBox
    Set h = New CHookComboBox
    h.Attach oC, Me, m_sCallback
    m_colHooks.Add h
End Sub

Private Sub HookOptionButton(ByVal oCtl As Object)
    Dim oO As MSForms.OptionButton
    Set oO = oCtl
    Dim h As CHookOptionButton
    Set h = New CHookOptionButton
    h.Attach oO, Me, m_sCallback
    m_colHooks.Add h
End Sub

' DispatchKeyDown   called by each hook class when a KeyDown
'   arrives on a hooked control. Forwards to the form's named
'   callback via CallByName.
Friend Sub DispatchKeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                            ByVal Shift As Integer)
    On Error Resume Next
    CallByName m_oForm, m_sCallback, VbMethod, KeyCode, Shift
    If Err.Number <> 0 Then
        Debug.Print "[CKeyPreview] DispatchKeyDown error " & _
                    Err.Number & ": " & Err.Description & _
                    " (callback='" & m_sCallback & "')"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' END: CKeyPreview v9.19.9

