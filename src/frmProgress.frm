VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "Batch PDF Publisher"
   ClientHeight    =   4320
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8400.001
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Component : frmProgress (UserForm)
' Purpose   : Modal progress dialog shown during batch processing.
'             Displays the current file path, per-file status, a
'             completion counter, a percentage, and a progress bar.
'             Esc cancels the running batch. Terminates by unloading
'             itself after PRJ_RunScanSession completes.
' Reference : Appendix A.13
' Version   : 9.19.11 (Inventor 2027)
'
' v9.19.11 (structure):
'   All click handlers restructured to one statement per line
'   (previously lblCancel_Click and lblCancelTxt_Click were
'   inline colon-separated subs). No functional changes.
'
' v9.19.7:
'   UserForm_MouseMove: FW_ApplyTheme removed (flicker fix).
'   CKeyPreview now uses concrete hook classes.
'
' v9.19.6 (Header standardization):
'   Added Component/Purpose/Reference/Version header fields.
'   No functional changes.
'
' v9.19.6:
'   Font.bOld -> Font.Bold (A-13).
' v9.19.3:
'   UserForm_Activate delegates form setup to PRJ_PrepareProgressForm.
' v9.19.0:
'   UserForm_Activate no longer calls FIT_FitLabel on lblFile.
'
' Notes     : No line-continuation characters in procedure bodies.
'             Procedure signatures are excepted per Doc 01 / Doc 06.
'             No decorative separator comment lines.

Private m_sOrigCap As String
Private m_bStarted  As Boolean
Private m_oKeyPreview As CKeyPreview

Private Sub UserForm_Initialize()
    m_sOrigCap = Me.Caption
    Me.StartUpPosition = 1
    PRJ_SetupAccelerators Me
    Set m_oKeyPreview = New CKeyPreview
    m_oKeyPreview.Initialize Me, "OnFormKeyDown"
End Sub

Private Sub UserForm_Activate()
    If m_bStarted Then Exit Sub
    m_bStarted = True

    Dim sCap As String
    sCap = L("title")
    If Len(sCap) = 0 Then sCap = Trim$(m_sOrigCap)
    Dim sVerSuffix As String
    sVerSuffix = " v" & PRJ_VERSION
    If Right$(sCap, Len(sVerSuffix)) <> sVerSuffix Then
        Me.Caption = sCap & sVerSuffix
    End If

    PRJ_PrepareProgressForm

    On Error Resume Next
    Me.Controls("lblFill").ZOrder 0
    On Error GoTo 0

    Me.Controls("lblFile").Caption = L("progress.init")

    Me.Controls("lblStat").Caption = ""
    Me.Controls("lblStat").ForeColor = CLR_TEXT_SEC

    Me.Controls("lblCount").Caption = "0 / 0"
    Me.Controls("lblCount").ForeColor = CLR_ERROR
    Me.Controls("lblPct").Caption = "0%"
    Me.Controls("lblPct").ForeColor = CLR_ERROR
    Me.Controls("lblPct").Font.Bold = False
    Me.Controls("lblFill").Width = 0
    DoEvents

    PRJ_RunScanSession
    Set g_oProgress = Nothing
    Unload Me
End Sub

Public Sub OnFormKeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then
        g_bCancel = True
        KeyCode = 0
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' v9.19.7: FW_ApplyTheme removed (flicker fix).
    FW_ResetHover
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 27 Then g_bCancel = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        g_bCancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
    FW_ReleaseAllHoverButtons
    PROG_ReleaseControls
    Set g_oProgress = Nothing
    If Not m_oKeyPreview Is Nothing Then
        m_oKeyPreview.Terminate
        Set m_oKeyPreview = Nothing
    End If
End Sub

Private Sub lblCancel_Click()
    g_bCancel = True
End Sub

Private Sub lblCancelTxt_Click()
    g_bCancel = True
End Sub

' END: frmProgress v9.19.11
