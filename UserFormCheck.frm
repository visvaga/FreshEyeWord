VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormCheck 
   Caption         =   "Ñâåæèé âçãëÿä"
   ClientHeight    =   2905
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4298
   OleObjectBlob   =   "UserFormCheck.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmdCheck_Click()
    ' Retrieve values from the form
    Dim sensitivity As Double
    Dim context As Integer
    Dim excludeNames As Boolean
    
        ' Initialize options if not already done
    Dim options As Object
    Set options = CreateObject("Scripting.Dictionary")

    ' Validate and get inputs
    On Error Resume Next
    sensitivity = CDbl(Me.txtSensitivity.text)
    context = CInt(Me.txtContextSize.text)
''    excludeNames = CBool(Me.chkExcludeProperNames.value)
    On Error GoTo 0

    ' Populate options dictionary with values from the UserForm
    options("sensitivity_threshold") = CDbl(Me.txtSensitivity.text)
    options("context_size") = CInt(Me.txtContextSize.text)
''    options("exclude_proper_names") = Me.chkExcludeProperNames.value
    
    ' Run the analysis with updated settings
    Call AnalyzeSelectedText

    ' Close the form
    Me.Hide
End Sub

Private Sub cmdClear_Click()
    ' Clear highlights in the selected text
    Dim wordRange As range
    For Each wordRange In Selection.words
        wordRange.HighlightColorIndex = wdNoHighlight
        wordRange.Shading.BackgroundPatternColor = wdColorAutomatic
    Next wordRange
    MsgBox "Î÷èùàåì òåêñò"
End Sub

Private Sub UserForm_Initialize()
    ' Set initial values for TextBoxes and SpinButtons
    Me.txtSensitivity.text = "370"
    Me.spinSensitivity.value = 370
    Me.txtContextSize.text = "30"
    Me.spinContextSize.value = 30
'    Me.chkExcludeProperNames.value = True ' Default to True
End Sub

' Update txtSensitivity based on spinSensitivity value
Private Sub spinSensitivity_Change()
    Me.txtSensitivity.text = Me.spinSensitivity.value
End Sub

' Update txtContextSize based on spinContextSize value
Private Sub spinContextSize_Change()
    Me.txtContextSize.text = Me.spinContextSize.value
End Sub

' Ensure manual input in txtSensitivity is within range and updates the spin button
Private Sub txtSensitivity_AfterUpdate()
    If IsNumeric(Me.txtSensitivity.text) Then
        Dim sensitivity As Long
        sensitivity = CLng(Me.txtSensitivity.text)
        ' Restrict to range 1-2000
        If sensitivity < 1 Then sensitivity = 1
        If sensitivity > 2000 Then sensitivity = 2000
        Me.txtSensitivity.text = sensitivity
        Me.spinSensitivity.value = sensitivity
    Else
        ' Reset to default if non-numeric
        Me.txtSensitivity.text = Me.spinSensitivity.value
    End If
End Sub

' Ensure manual input in txtContextSize is within range and updates the spin button
Private Sub txtContextSize_AfterUpdate()
    If IsNumeric(Me.txtContextSize.text) Then
        Dim contextSize As Long
        contextSize = CLng(Me.txtContextSize.text)
        ' Restrict to range 2-100
        If contextSize < 2 Then contextSize = 2
        If contextSize > 100 Then contextSize = 100
        Me.txtContextSize.text = contextSize
        Me.spinContextSize.value = contextSize
    Else
        ' Reset to default if non-numeric
        Me.txtContextSize.text = Me.spinContextSize.value
    End If
End Sub

