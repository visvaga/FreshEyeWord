Attribute VB_Name = "Module1"

Option Explicit

' version 11.0 of VBA macro FreshEye - search of similar words within selected text

' Variables for linguistic tables
Dim sim_ch_stride As Integer
Dim sim_ch(1 To 34, 1 To 34) As Integer
Dim inf_letters_stride As Integer
Dim inf_letters(1 To 34, 1 To 2) As Integer
Dim exceptions_voc As Object
Dim exceptions_voc_first As Object
Dim highlight_color_counter As Integer
Dim options As Object
Dim start_time As Date
Dim badwords As Collection
Dim highlightColors() As WdColorIndex



' Main subroutine to initialize and run the analysis on selected text
Sub AnalyzeSelectedText()
    ' Check if there is a text selection
    If Selection.Type = wdSelectionIP Then
        MsgBox "Òåêñò íå âûäåëåí."
        Exit Sub
    End If

    ' Initialize data arrays and dictionaries
    InitializeData
    
    ' Define highlight colors
    InitializeHighlightColors
    
    Dim options As Object
    Set options = CreateObject("Scripting.Dictionary")
    
    ' Get the selected text
    Dim selectedText As String
    selectedText = Selection.text

    ' Set analysis parameters
    Dim sensitivity_threshold As Double
    Dim context_size As Integer

    sensitivity_threshold = CDbl(UserFormCheck.txtSensitivity.text)
    context_size = CInt(UserFormCheck.txtContextSize.text)
    
    ' Initialize collection to store bad words
    Set badwords = New Collection

    ' Run the Fresheye analysis on the selected text
    Dim result As String

    result = Fresheye(selectedText, sensitivity_threshold, context_size)
    MsgBox result
    
    ' Highlight all identified bad words in the selected text
    HighlightBadWords
    
End Sub

' Define and initialize highlight colors

Sub InitializeHighlightColors()
    ' Define a set of soft colors to cycle through
    ReDim highlightColors(1 To 5)
    highlightColors(1) = RGB(51, 203, 203) ' Blue-green
    highlightColors(2) = RGB(0, 175, 239) ' Blue
    highlightColors(3) = RGB(152, 203, 254) ' Pale Blue
    highlightColors(4) = RGB(203, 254, 203) ' Light Green
    highlightColors(5) = RGB(254, 254, 152) ' Light Yellow
End Sub


' Initialization of data arrays and dictionaries
Sub InitializeData()
    sim_ch_stride = 34
    inf_letters_stride = 2

    ' Define and populate similarity map (sim_ch array) as Integer
    Dim sim_ch_data(1 To 34, 1 To 34) As Integer
    Dim i As Integer, j As Integer
    
    ' Populate similarity map values for sim_ch_data
' à
sim_ch_data(1, 1) = 9
sim_ch_data(1, 2) = 0
sim_ch_data(1, 3) = 0
sim_ch_data(1, 4) = 0
sim_ch_data(1, 5) = 0
sim_ch_data(1, 6) = 1
sim_ch_data(1, 7) = 0
sim_ch_data(1, 8) = 0
sim_ch_data(1, 9) = 1
sim_ch_data(1, 10) = 0
sim_ch_data(1, 11) = 0
sim_ch_data(1, 12) = 0
sim_ch_data(1, 13) = 0
sim_ch_data(1, 14) = 0
sim_ch_data(1, 15) = 2
sim_ch_data(1, 16) = 0
sim_ch_data(1, 17) = 0
sim_ch_data(1, 18) = 0
sim_ch_data(1, 19) = 0
sim_ch_data(1, 20) = 1
sim_ch_data(1, 21) = 0
sim_ch_data(1, 22) = 0
sim_ch_data(1, 23) = 0
sim_ch_data(1, 24) = 0
sim_ch_data(1, 25) = 0
sim_ch_data(1, 26) = 0
sim_ch_data(1, 27) = 0
sim_ch_data(1, 28) = 1
sim_ch_data(1, 29) = 0
sim_ch_data(1, 30) = 1
sim_ch_data(1, 31) = 1
sim_ch_data(1, 32) = 2
sim_ch_data(1, 33) = 0
sim_ch_data(1, 34) = 1

' á
sim_ch_data(2, 1) = 0
sim_ch_data(2, 2) = 9
sim_ch_data(2, 3) = 1
sim_ch_data(2, 4) = 0
sim_ch_data(2, 5) = 0
sim_ch_data(2, 6) = 0
sim_ch_data(2, 7) = 0
sim_ch_data(2, 8) = 0
sim_ch_data(2, 9) = 0
sim_ch_data(2, 10) = 0
sim_ch_data(2, 11) = 0
sim_ch_data(2, 12) = 0
sim_ch_data(2, 13) = 0
sim_ch_data(2, 14) = 0
sim_ch_data(2, 15) = 0
sim_ch_data(2, 16) = 3
sim_ch_data(2, 17) = 0
sim_ch_data(2, 18) = 0
sim_ch_data(2, 19) = 0
sim_ch_data(2, 20) = 0
sim_ch_data(2, 21) = 1
sim_ch_data(2, 22) = 0
sim_ch_data(2, 23) = 0
sim_ch_data(2, 24) = 0
sim_ch_data(2, 25) = 0
sim_ch_data(2, 26) = 0
sim_ch_data(2, 27) = 0
sim_ch_data(2, 28) = 0
sim_ch_data(2, 29) = 0
sim_ch_data(2, 30) = 0
sim_ch_data(2, 31) = 0
sim_ch_data(2, 32) = 0
sim_ch_data(2, 33) = 0
sim_ch_data(2, 34) = 0

' â
sim_ch_data(3, 1) = 0
sim_ch_data(3, 2) = 1
sim_ch_data(3, 3) = 9
sim_ch_data(3, 4) = 1
sim_ch_data(3, 5) = 0
sim_ch_data(3, 6) = 0
sim_ch_data(3, 7) = 0
sim_ch_data(3, 8) = 0
sim_ch_data(3, 9) = 0
sim_ch_data(3, 10) = 0
sim_ch_data(3, 11) = 0
sim_ch_data(3, 12) = 1
sim_ch_data(3, 13) = 1
sim_ch_data(3, 14) = 1
sim_ch_data(3, 15) = 0
sim_ch_data(3, 16) = 1
sim_ch_data(3, 17) = 0
sim_ch_data(3, 18) = 0
sim_ch_data(3, 19) = 0
sim_ch_data(3, 20) = 1
sim_ch_data(3, 21) = 3
sim_ch_data(3, 22) = 0
sim_ch_data(3, 23) = 0
sim_ch_data(3, 24) = 0
sim_ch_data(3, 25) = 0
sim_ch_data(3, 26) = 0
sim_ch_data(3, 27) = 0
sim_ch_data(3, 28) = 0
sim_ch_data(3, 29) = 0
sim_ch_data(3, 30) = 0
sim_ch_data(3, 31) = 0
sim_ch_data(3, 32) = 0
sim_ch_data(3, 33) = 0
sim_ch_data(3, 34) = 0

' ã
sim_ch_data(4, 1) = 0
sim_ch_data(4, 2) = 0
sim_ch_data(4, 3) = 1
sim_ch_data(4, 4) = 9
sim_ch_data(4, 5) = 0
sim_ch_data(4, 6) = 0
sim_ch_data(4, 7) = 3
sim_ch_data(4, 8) = 0
sim_ch_data(4, 9) = 0
sim_ch_data(4, 10) = 0
sim_ch_data(4, 11) = 3
sim_ch_data(4, 12) = 0
sim_ch_data(4, 13) = 0
sim_ch_data(4, 14) = 0
sim_ch_data(4, 15) = 0
sim_ch_data(4, 16) = 0
sim_ch_data(4, 17) = 0
sim_ch_data(4, 18) = 0
sim_ch_data(4, 19) = 0
sim_ch_data(4, 20) = 0
sim_ch_data(4, 21) = 0
sim_ch_data(4, 22) = 1
sim_ch_data(4, 23) = 0
sim_ch_data(4, 24) = 0
sim_ch_data(4, 25) = 0
sim_ch_data(4, 26) = 0
sim_ch_data(4, 27) = 0
sim_ch_data(4, 28) = 0
sim_ch_data(4, 29) = 0
sim_ch_data(4, 30) = 0
sim_ch_data(4, 31) = 0
sim_ch_data(4, 32) = 0
sim_ch_data(4, 33) = 0
sim_ch_data(4, 34) = 0


' ä
sim_ch_data(5, 1) = 0
sim_ch_data(5, 2) = 0
sim_ch_data(5, 3) = 0
sim_ch_data(5, 4) = 0
sim_ch_data(5, 5) = 9
sim_ch_data(5, 6) = 0
sim_ch_data(5, 7) = 0
sim_ch_data(5, 8) = 1
sim_ch_data(5, 9) = 0
sim_ch_data(5, 10) = 0
sim_ch_data(5, 11) = 0
sim_ch_data(5, 12) = 0
sim_ch_data(5, 13) = 0
sim_ch_data(5, 14) = 0
sim_ch_data(5, 15) = 0
sim_ch_data(5, 16) = 0
sim_ch_data(5, 17) = 0
sim_ch_data(5, 18) = 1
sim_ch_data(5, 19) = 3
sim_ch_data(5, 20) = 0
sim_ch_data(5, 21) = 0
sim_ch_data(5, 22) = 0
sim_ch_data(5, 23) = 1
sim_ch_data(5, 24) = 0
sim_ch_data(5, 25) = 0
sim_ch_data(5, 26) = 0
sim_ch_data(5, 27) = 0
sim_ch_data(5, 28) = 0
sim_ch_data(5, 29) = 0
sim_ch_data(5, 30) = 0
sim_ch_data(5, 31) = 0
sim_ch_data(5, 32) = 0
sim_ch_data(5, 33) = 0
sim_ch_data(5, 34) = 0


' å
sim_ch_data(6, 1) = 1
sim_ch_data(6, 2) = 0
sim_ch_data(6, 3) = 0
sim_ch_data(6, 4) = 0
sim_ch_data(6, 5) = 0
sim_ch_data(6, 6) = 9
sim_ch_data(6, 7) = 0
sim_ch_data(6, 8) = 0
sim_ch_data(6, 9) = 2
sim_ch_data(6, 10) = 0
sim_ch_data(6, 11) = 0
sim_ch_data(6, 12) = 0
sim_ch_data(6, 13) = 0
sim_ch_data(6, 14) = 0
sim_ch_data(6, 15) = 1
sim_ch_data(6, 16) = 0
sim_ch_data(6, 17) = 0
sim_ch_data(6, 18) = 0
sim_ch_data(6, 19) = 0
sim_ch_data(6, 20) = 1
sim_ch_data(6, 21) = 0
sim_ch_data(6, 22) = 0
sim_ch_data(6, 23) = 0
sim_ch_data(6, 24) = 0
sim_ch_data(6, 25) = 0
sim_ch_data(6, 26) = 0
sim_ch_data(6, 27) = 0
sim_ch_data(6, 28) = 1
sim_ch_data(6, 29) = 0
sim_ch_data(6, 30) = 2
sim_ch_data(6, 31) = 1
sim_ch_data(6, 32) = 1
sim_ch_data(6, 33) = 0
sim_ch_data(6, 34) = 9


' æ
sim_ch_data(7, 1) = 0
sim_ch_data(7, 2) = 0
sim_ch_data(7, 3) = 0
sim_ch_data(7, 4) = 3
sim_ch_data(7, 5) = 0
sim_ch_data(7, 6) = 0
sim_ch_data(7, 7) = 9
sim_ch_data(7, 8) = 3
sim_ch_data(7, 9) = 0
sim_ch_data(7, 10) = 0
sim_ch_data(7, 11) = 0
sim_ch_data(7, 12) = 0
sim_ch_data(7, 13) = 0
sim_ch_data(7, 14) = 0
sim_ch_data(7, 15) = 0
sim_ch_data(7, 16) = 0
sim_ch_data(7, 17) = 0
sim_ch_data(7, 18) = 0
sim_ch_data(7, 19) = 0
sim_ch_data(7, 20) = 0
sim_ch_data(7, 21) = 0
sim_ch_data(7, 22) = 0
sim_ch_data(7, 23) = 0
sim_ch_data(7, 24) = 3
sim_ch_data(7, 25) = 3
sim_ch_data(7, 26) = 3
sim_ch_data(7, 27) = 0
sim_ch_data(7, 28) = 0
sim_ch_data(7, 29) = 0
sim_ch_data(7, 30) = 0
sim_ch_data(7, 31) = 0
sim_ch_data(7, 32) = 0
sim_ch_data(7, 33) = 0
sim_ch_data(7, 34) = 0


' ç
sim_ch_data(8, 1) = 0
sim_ch_data(8, 2) = 0
sim_ch_data(8, 3) = 0
sim_ch_data(8, 4) = 0
sim_ch_data(8, 5) = 1
sim_ch_data(8, 6) = 0
sim_ch_data(8, 7) = 3
sim_ch_data(8, 8) = 9
sim_ch_data(8, 9) = 0
sim_ch_data(8, 10) = 0
sim_ch_data(8, 11) = 0
sim_ch_data(8, 12) = 0
sim_ch_data(8, 13) = 0
sim_ch_data(8, 14) = 0
sim_ch_data(8, 15) = 0
sim_ch_data(8, 16) = 0
sim_ch_data(8, 17) = 0
sim_ch_data(8, 18) = 3
sim_ch_data(8, 19) = 1
sim_ch_data(8, 20) = 0
sim_ch_data(8, 21) = 0
sim_ch_data(8, 22) = 0
sim_ch_data(8, 23) = 3
sim_ch_data(8, 24) = 1
sim_ch_data(8, 25) = 1
sim_ch_data(8, 26) = 1
sim_ch_data(8, 27) = 0
sim_ch_data(8, 28) = 0
sim_ch_data(8, 29) = 0
sim_ch_data(8, 30) = 0
sim_ch_data(8, 31) = 0
sim_ch_data(8, 32) = 0
sim_ch_data(8, 33) = 0
sim_ch_data(8, 34) = 0


' è
sim_ch_data(9, 1) = 1
sim_ch_data(9, 2) = 0
sim_ch_data(9, 3) = 0
sim_ch_data(9, 4) = 0
sim_ch_data(9, 5) = 0
sim_ch_data(9, 6) = 2
sim_ch_data(9, 7) = 0
sim_ch_data(9, 8) = 0
sim_ch_data(9, 9) = 9
sim_ch_data(9, 10) = 3
sim_ch_data(9, 11) = 0
sim_ch_data(9, 12) = 0
sim_ch_data(9, 13) = 0
sim_ch_data(9, 14) = 0
sim_ch_data(9, 15) = 1
sim_ch_data(9, 16) = 0
sim_ch_data(9, 17) = 0
sim_ch_data(9, 18) = 0
sim_ch_data(9, 19) = 0
sim_ch_data(9, 20) = 1
sim_ch_data(9, 21) = 0
sim_ch_data(9, 22) = 0
sim_ch_data(9, 23) = 0
sim_ch_data(9, 24) = 0
sim_ch_data(9, 25) = 0
sim_ch_data(9, 26) = 0
sim_ch_data(9, 27) = 0
sim_ch_data(9, 28) = 2
sim_ch_data(9, 29) = 0
sim_ch_data(9, 30) = 1
sim_ch_data(9, 31) = 1
sim_ch_data(9, 32) = 1
sim_ch_data(9, 33) = 0
sim_ch_data(9, 34) = 2


' é
sim_ch_data(10, 1) = 0
sim_ch_data(10, 2) = 0
sim_ch_data(10, 3) = 0
sim_ch_data(10, 4) = 0
sim_ch_data(10, 5) = 0
sim_ch_data(10, 6) = 0
sim_ch_data(10, 7) = 0
sim_ch_data(10, 8) = 0
sim_ch_data(10, 9) = 2
sim_ch_data(10, 10) = 9
sim_ch_data(10, 11) = 0
sim_ch_data(10, 12) = 1
sim_ch_data(10, 13) = 1
sim_ch_data(10, 14) = 1
sim_ch_data(10, 15) = 0
sim_ch_data(10, 16) = 0
sim_ch_data(10, 17) = 1
sim_ch_data(10, 18) = 0
sim_ch_data(10, 19) = 0
sim_ch_data(10, 20) = 0
sim_ch_data(10, 21) = 0
sim_ch_data(10, 22) = 0
sim_ch_data(10, 23) = 0
sim_ch_data(10, 24) = 0
sim_ch_data(10, 25) = 0
sim_ch_data(10, 26) = 0
sim_ch_data(10, 27) = 0
sim_ch_data(10, 28) = 0
sim_ch_data(10, 29) = 0
sim_ch_data(10, 30) = 0
sim_ch_data(10, 31) = 0
sim_ch_data(10, 32) = 0
sim_ch_data(10, 33) = 0
sim_ch_data(10, 34) = 0


' ê
sim_ch_data(11, 1) = 0
sim_ch_data(11, 2) = 0
sim_ch_data(11, 3) = 0
sim_ch_data(11, 4) = 3
sim_ch_data(11, 5) = 0
sim_ch_data(11, 6) = 0
sim_ch_data(11, 7) = 0
sim_ch_data(11, 8) = 0
sim_ch_data(11, 9) = 0
sim_ch_data(11, 10) = 0
sim_ch_data(11, 11) = 9
sim_ch_data(11, 12) = 0
sim_ch_data(11, 13) = 0
sim_ch_data(11, 14) = 0
sim_ch_data(11, 15) = 0
sim_ch_data(11, 16) = 0
sim_ch_data(11, 17) = 0
sim_ch_data(11, 18) = 0
sim_ch_data(11, 19) = 0
sim_ch_data(11, 20) = 0
sim_ch_data(11, 21) = 0
sim_ch_data(11, 22) = 1
sim_ch_data(11, 23) = 0
sim_ch_data(11, 24) = 0
sim_ch_data(11, 25) = 0
sim_ch_data(11, 26) = 0
sim_ch_data(11, 27) = 0
sim_ch_data(11, 28) = 0
sim_ch_data(11, 29) = 0
sim_ch_data(11, 30) = 0
sim_ch_data(11, 31) = 0
sim_ch_data(11, 32) = 0
sim_ch_data(11, 33) = 0
sim_ch_data(11, 34) = 0


' ë
sim_ch_data(12, 1) = 0
sim_ch_data(12, 2) = 0
sim_ch_data(12, 3) = 1
sim_ch_data(12, 4) = 0
sim_ch_data(12, 5) = 0
sim_ch_data(12, 6) = 0
sim_ch_data(12, 7) = 0
sim_ch_data(12, 8) = 0
sim_ch_data(12, 9) = 0
sim_ch_data(12, 10) = 1
sim_ch_data(12, 11) = 0
sim_ch_data(12, 12) = 9
sim_ch_data(12, 13) = 1
sim_ch_data(12, 14) = 1
sim_ch_data(12, 15) = 0
sim_ch_data(12, 16) = 0
sim_ch_data(12, 17) = 1
sim_ch_data(12, 18) = 0
sim_ch_data(12, 19) = 0
sim_ch_data(12, 20) = 0
sim_ch_data(12, 21) = 0
sim_ch_data(12, 22) = 0
sim_ch_data(12, 23) = 0
sim_ch_data(12, 24) = 0
sim_ch_data(12, 25) = 0
sim_ch_data(12, 26) = 0
sim_ch_data(12, 27) = 0
sim_ch_data(12, 28) = 0
sim_ch_data(12, 29) = 0
sim_ch_data(12, 30) = 0
sim_ch_data(12, 31) = 0
sim_ch_data(12, 32) = 0
sim_ch_data(12, 33) = 0
sim_ch_data(12, 34) = 0


' ì
sim_ch_data(13, 1) = 0
sim_ch_data(13, 2) = 0
sim_ch_data(13, 3) = 1
sim_ch_data(13, 4) = 0
sim_ch_data(13, 5) = 0
sim_ch_data(13, 6) = 0
sim_ch_data(13, 7) = 0
sim_ch_data(13, 8) = 0
sim_ch_data(13, 9) = 0
sim_ch_data(13, 10) = 1
sim_ch_data(13, 11) = 0
sim_ch_data(13, 12) = 1
sim_ch_data(13, 13) = 9
sim_ch_data(13, 14) = 3
sim_ch_data(13, 15) = 0
sim_ch_data(13, 16) = 0
sim_ch_data(13, 17) = 1
sim_ch_data(13, 18) = 0
sim_ch_data(13, 19) = 0
sim_ch_data(13, 20) = 0
sim_ch_data(13, 21) = 0
sim_ch_data(13, 22) = 0
sim_ch_data(13, 23) = 0
sim_ch_data(13, 24) = 0
sim_ch_data(13, 25) = 0
sim_ch_data(13, 26) = 0
sim_ch_data(13, 27) = 0
sim_ch_data(13, 28) = 0
sim_ch_data(13, 29) = 0
sim_ch_data(13, 30) = 0
sim_ch_data(13, 31) = 0
sim_ch_data(13, 32) = 0
sim_ch_data(13, 33) = 0
sim_ch_data(13, 34) = 0


' í
sim_ch_data(14, 1) = 0
sim_ch_data(14, 2) = 0
sim_ch_data(14, 3) = 1
sim_ch_data(14, 4) = 0
sim_ch_data(14, 5) = 0
sim_ch_data(14, 6) = 0
sim_ch_data(14, 7) = 0
sim_ch_data(14, 8) = 0
sim_ch_data(14, 9) = 0
sim_ch_data(14, 10) = 1
sim_ch_data(14, 11) = 0
sim_ch_data(14, 12) = 1
sim_ch_data(14, 13) = 3
sim_ch_data(14, 14) = 9
sim_ch_data(14, 15) = 0
sim_ch_data(14, 16) = 0
sim_ch_data(14, 17) = 1
sim_ch_data(14, 18) = 0
sim_ch_data(14, 19) = 0
sim_ch_data(14, 20) = 0
sim_ch_data(14, 21) = 0
sim_ch_data(14, 22) = 0
sim_ch_data(14, 23) = 0
sim_ch_data(14, 24) = 0
sim_ch_data(14, 25) = 0
sim_ch_data(14, 26) = 0
sim_ch_data(14, 27) = 0
sim_ch_data(14, 28) = 0
sim_ch_data(14, 29) = 0
sim_ch_data(14, 30) = 0
sim_ch_data(14, 31) = 0
sim_ch_data(14, 32) = 0
sim_ch_data(14, 33) = 0
sim_ch_data(14, 34) = 0


' î
sim_ch_data(15, 1) = 2
sim_ch_data(15, 2) = 0
sim_ch_data(15, 3) = 0
sim_ch_data(15, 4) = 0
sim_ch_data(15, 5) = 0
sim_ch_data(15, 6) = 1
sim_ch_data(15, 7) = 0
sim_ch_data(15, 8) = 0
sim_ch_data(15, 9) = 1
sim_ch_data(15, 10) = 0
sim_ch_data(15, 11) = 0
sim_ch_data(15, 12) = 0
sim_ch_data(15, 13) = 0
sim_ch_data(15, 14) = 0
sim_ch_data(15, 15) = 9
sim_ch_data(15, 16) = 0
sim_ch_data(15, 17) = 0
sim_ch_data(15, 18) = 0
sim_ch_data(15, 19) = 0
sim_ch_data(15, 20) = 1
sim_ch_data(15, 21) = 0
sim_ch_data(15, 22) = 0
sim_ch_data(15, 23) = 0
sim_ch_data(15, 24) = 0
sim_ch_data(15, 25) = 0
sim_ch_data(15, 26) = 0
sim_ch_data(15, 27) = 0
sim_ch_data(15, 28) = 1
sim_ch_data(15, 29) = 0
sim_ch_data(15, 30) = 1
sim_ch_data(15, 31) = 1
sim_ch_data(15, 32) = 1
sim_ch_data(15, 33) = 0
sim_ch_data(15, 34) = 1


' ï
sim_ch_data(16, 1) = 0
sim_ch_data(16, 2) = 3
sim_ch_data(16, 3) = 1
sim_ch_data(16, 4) = 0
sim_ch_data(16, 5) = 0
sim_ch_data(16, 6) = 0
sim_ch_data(16, 7) = 0
sim_ch_data(16, 8) = 0
sim_ch_data(16, 9) = 0
sim_ch_data(16, 10) = 0
sim_ch_data(16, 11) = 0
sim_ch_data(16, 12) = 0
sim_ch_data(16, 13) = 0
sim_ch_data(16, 14) = 0
sim_ch_data(16, 15) = 0
sim_ch_data(16, 16) = 9
sim_ch_data(16, 17) = 0
sim_ch_data(16, 18) = 0
sim_ch_data(16, 19) = 0
sim_ch_data(16, 20) = 0
sim_ch_data(16, 21) = 1
sim_ch_data(16, 22) = 0
sim_ch_data(16, 23) = 0
sim_ch_data(16, 24) = 0
sim_ch_data(16, 25) = 0
sim_ch_data(16, 26) = 0
sim_ch_data(16, 27) = 0
sim_ch_data(16, 28) = 0
sim_ch_data(16, 29) = 0
sim_ch_data(16, 30) = 0
sim_ch_data(16, 31) = 0
sim_ch_data(16, 32) = 0
sim_ch_data(16, 33) = 0
sim_ch_data(16, 34) = 0


' ð
sim_ch_data(17, 1) = 0
sim_ch_data(17, 2) = 0
sim_ch_data(17, 3) = 0
sim_ch_data(17, 4) = 0
sim_ch_data(17, 5) = 0
sim_ch_data(17, 6) = 0
sim_ch_data(17, 7) = 0
sim_ch_data(17, 8) = 0
sim_ch_data(17, 9) = 0
sim_ch_data(17, 10) = 1
sim_ch_data(17, 11) = 0
sim_ch_data(17, 12) = 1
sim_ch_data(17, 13) = 1
sim_ch_data(17, 14) = 1
sim_ch_data(17, 15) = 0
sim_ch_data(17, 16) = 0
sim_ch_data(17, 17) = 9
sim_ch_data(17, 18) = 0
sim_ch_data(17, 19) = 0
sim_ch_data(17, 20) = 0
sim_ch_data(17, 21) = 0
sim_ch_data(17, 22) = 1
sim_ch_data(17, 23) = 0
sim_ch_data(17, 24) = 0
sim_ch_data(17, 25) = 0
sim_ch_data(17, 26) = 0
sim_ch_data(17, 27) = 0
sim_ch_data(17, 28) = 0
sim_ch_data(17, 29) = 0
sim_ch_data(17, 30) = 0
sim_ch_data(17, 31) = 0
sim_ch_data(17, 32) = 0
sim_ch_data(17, 33) = 0
sim_ch_data(17, 34) = 0


' ñ
sim_ch_data(18, 1) = 0
sim_ch_data(18, 2) = 0
sim_ch_data(18, 3) = 0
sim_ch_data(18, 4) = 0
sim_ch_data(18, 5) = 1
sim_ch_data(18, 6) = 0
sim_ch_data(18, 7) = 0
sim_ch_data(18, 8) = 3
sim_ch_data(18, 9) = 0
sim_ch_data(18, 10) = 0
sim_ch_data(18, 11) = 0
sim_ch_data(18, 12) = 0
sim_ch_data(18, 13) = 0
sim_ch_data(18, 14) = 0
sim_ch_data(18, 15) = 0
sim_ch_data(18, 16) = 0
sim_ch_data(18, 17) = 0
sim_ch_data(18, 18) = 9
sim_ch_data(18, 19) = 1
sim_ch_data(18, 20) = 0
sim_ch_data(18, 21) = 0
sim_ch_data(18, 22) = 0
sim_ch_data(18, 23) = 3
sim_ch_data(18, 24) = 1
sim_ch_data(18, 25) = 0
sim_ch_data(18, 26) = 0
sim_ch_data(18, 27) = 0
sim_ch_data(18, 28) = 0
sim_ch_data(18, 29) = 0
sim_ch_data(18, 30) = 0
sim_ch_data(18, 31) = 0
sim_ch_data(18, 32) = 0
sim_ch_data(18, 33) = 0
sim_ch_data(18, 34) = 0


' ò
sim_ch_data(19, 1) = 0
sim_ch_data(19, 2) = 0
sim_ch_data(19, 3) = 0
sim_ch_data(19, 4) = 0
sim_ch_data(19, 5) = 3
sim_ch_data(19, 6) = 0
sim_ch_data(19, 7) = 0
sim_ch_data(19, 8) = 1
sim_ch_data(19, 9) = 0
sim_ch_data(19, 10) = 0
sim_ch_data(19, 11) = 0
sim_ch_data(19, 12) = 0
sim_ch_data(19, 13) = 0
sim_ch_data(19, 14) = 0
sim_ch_data(19, 15) = 0
sim_ch_data(19, 16) = 0
sim_ch_data(19, 17) = 0
sim_ch_data(19, 18) = 1
sim_ch_data(19, 19) = 9
sim_ch_data(19, 20) = 0
sim_ch_data(19, 21) = 0
sim_ch_data(19, 22) = 0
sim_ch_data(19, 23) = 1
sim_ch_data(19, 24) = 1
sim_ch_data(19, 25) = 0
sim_ch_data(19, 26) = 0
sim_ch_data(19, 27) = 0
sim_ch_data(19, 28) = 0
sim_ch_data(19, 29) = 0
sim_ch_data(19, 30) = 0
sim_ch_data(19, 31) = 0
sim_ch_data(19, 32) = 0
sim_ch_data(19, 33) = 0
sim_ch_data(19, 34) = 0


' ó
sim_ch_data(20, 1) = 1
sim_ch_data(20, 2) = 0
sim_ch_data(20, 3) = 1
sim_ch_data(20, 4) = 0
sim_ch_data(20, 5) = 0
sim_ch_data(20, 6) = 1
sim_ch_data(20, 7) = 0
sim_ch_data(20, 8) = 0
sim_ch_data(20, 9) = 1
sim_ch_data(20, 10) = 0
sim_ch_data(20, 11) = 0
sim_ch_data(20, 12) = 0
sim_ch_data(20, 13) = 0
sim_ch_data(20, 14) = 0
sim_ch_data(20, 15) = 1
sim_ch_data(20, 16) = 0
sim_ch_data(20, 17) = 0
sim_ch_data(20, 18) = 0
sim_ch_data(20, 19) = 0
sim_ch_data(20, 20) = 9
sim_ch_data(20, 21) = 0
sim_ch_data(20, 22) = 0
sim_ch_data(20, 23) = 0
sim_ch_data(20, 24) = 0
sim_ch_data(20, 25) = 0
sim_ch_data(20, 26) = 0
sim_ch_data(20, 27) = 0
sim_ch_data(20, 28) = 1
sim_ch_data(20, 29) = 0
sim_ch_data(20, 30) = 1
sim_ch_data(20, 31) = 2
sim_ch_data(20, 32) = 1
sim_ch_data(20, 33) = 0
sim_ch_data(20, 34) = 1


' ô
sim_ch_data(21, 1) = 0
sim_ch_data(21, 2) = 1
sim_ch_data(21, 3) = 3
sim_ch_data(21, 4) = 0
sim_ch_data(21, 5) = 0
sim_ch_data(21, 6) = 0
sim_ch_data(21, 7) = 0
sim_ch_data(21, 8) = 0
sim_ch_data(21, 9) = 0
sim_ch_data(21, 10) = 0
sim_ch_data(21, 11) = 0
sim_ch_data(21, 12) = 0
sim_ch_data(21, 13) = 0
sim_ch_data(21, 14) = 0
sim_ch_data(21, 15) = 0
sim_ch_data(21, 16) = 1
sim_ch_data(21, 17) = 0
sim_ch_data(21, 18) = 0
sim_ch_data(21, 19) = 0
sim_ch_data(21, 20) = 0
sim_ch_data(21, 21) = 9
sim_ch_data(21, 22) = 0
sim_ch_data(21, 23) = 0
sim_ch_data(21, 24) = 0
sim_ch_data(21, 25) = 0
sim_ch_data(21, 26) = 0
sim_ch_data(21, 27) = 0
sim_ch_data(21, 28) = 0
sim_ch_data(21, 29) = 0
sim_ch_data(21, 30) = 0
sim_ch_data(21, 31) = 0
sim_ch_data(21, 32) = 0
sim_ch_data(21, 33) = 0
sim_ch_data(21, 34) = 0


' õ
sim_ch_data(22, 1) = 0
sim_ch_data(22, 2) = 0
sim_ch_data(22, 3) = 0
sim_ch_data(22, 4) = 1
sim_ch_data(22, 5) = 0
sim_ch_data(22, 6) = 0
sim_ch_data(22, 7) = 0
sim_ch_data(22, 8) = 0
sim_ch_data(22, 9) = 0
sim_ch_data(22, 10) = 0
sim_ch_data(22, 11) = 1
sim_ch_data(22, 12) = 0
sim_ch_data(22, 13) = 0
sim_ch_data(22, 14) = 0
sim_ch_data(22, 15) = 0
sim_ch_data(22, 16) = 0
sim_ch_data(22, 17) = 1
sim_ch_data(22, 18) = 0
sim_ch_data(22, 19) = 0
sim_ch_data(22, 20) = 0
sim_ch_data(22, 21) = 0
sim_ch_data(22, 22) = 9
sim_ch_data(22, 23) = 0
sim_ch_data(22, 24) = 1
sim_ch_data(22, 25) = 0
sim_ch_data(22, 26) = 0
sim_ch_data(22, 27) = 0
sim_ch_data(22, 28) = 0
sim_ch_data(22, 29) = 0
sim_ch_data(22, 30) = 0
sim_ch_data(22, 31) = 0
sim_ch_data(22, 32) = 0
sim_ch_data(22, 33) = 0
sim_ch_data(22, 34) = 0


' ö
sim_ch_data(23, 1) = 0
sim_ch_data(23, 2) = 0
sim_ch_data(23, 3) = 0
sim_ch_data(23, 4) = 0
sim_ch_data(23, 5) = 1
sim_ch_data(23, 6) = 0
sim_ch_data(23, 7) = 0
sim_ch_data(23, 8) = 3
sim_ch_data(23, 9) = 0
sim_ch_data(23, 10) = 0
sim_ch_data(23, 11) = 0
sim_ch_data(23, 12) = 0
sim_ch_data(23, 13) = 0
sim_ch_data(23, 14) = 0
sim_ch_data(23, 15) = 0
sim_ch_data(23, 16) = 0
sim_ch_data(23, 17) = 0
sim_ch_data(23, 18) = 3
sim_ch_data(23, 19) = 1
sim_ch_data(23, 20) = 0
sim_ch_data(23, 21) = 0
sim_ch_data(23, 22) = 0
sim_ch_data(23, 23) = 9
sim_ch_data(23, 24) = 0
sim_ch_data(23, 25) = 0
sim_ch_data(23, 26) = 0
sim_ch_data(23, 27) = 0
sim_ch_data(23, 28) = 0
sim_ch_data(23, 29) = 0
sim_ch_data(23, 30) = 0
sim_ch_data(23, 31) = 0
sim_ch_data(23, 32) = 0
sim_ch_data(23, 33) = 0
sim_ch_data(23, 34) = 0


' ÷
sim_ch_data(24, 1) = 0
sim_ch_data(24, 2) = 0
sim_ch_data(24, 3) = 0
sim_ch_data(24, 4) = 0
sim_ch_data(24, 5) = 0
sim_ch_data(24, 6) = 0
sim_ch_data(24, 7) = 3
sim_ch_data(24, 8) = 1
sim_ch_data(24, 9) = 0
sim_ch_data(24, 10) = 0
sim_ch_data(24, 11) = 0
sim_ch_data(24, 12) = 0
sim_ch_data(24, 13) = 0
sim_ch_data(24, 14) = 0
sim_ch_data(24, 15) = 0
sim_ch_data(24, 16) = 0
sim_ch_data(24, 17) = 0
sim_ch_data(24, 18) = 1
sim_ch_data(24, 19) = 1
sim_ch_data(24, 20) = 0
sim_ch_data(24, 21) = 0
sim_ch_data(24, 22) = 1
sim_ch_data(24, 23) = 0
sim_ch_data(24, 24) = 9
sim_ch_data(24, 25) = 3
sim_ch_data(24, 26) = 3
sim_ch_data(24, 27) = 0
sim_ch_data(24, 28) = 0
sim_ch_data(24, 29) = 0
sim_ch_data(24, 30) = 0
sim_ch_data(24, 31) = 0
sim_ch_data(24, 32) = 0
sim_ch_data(24, 33) = 0
sim_ch_data(24, 34) = 0


' ø
sim_ch_data(25, 1) = 0
sim_ch_data(25, 2) = 0
sim_ch_data(25, 3) = 0
sim_ch_data(25, 4) = 0
sim_ch_data(25, 5) = 0
sim_ch_data(25, 6) = 0
sim_ch_data(25, 7) = 3
sim_ch_data(25, 8) = 1
sim_ch_data(25, 9) = 0
sim_ch_data(25, 10) = 0
sim_ch_data(25, 11) = 0
sim_ch_data(25, 12) = 0
sim_ch_data(25, 13) = 0
sim_ch_data(25, 14) = 0
sim_ch_data(25, 15) = 0
sim_ch_data(25, 16) = 0
sim_ch_data(25, 17) = 0
sim_ch_data(25, 18) = 0
sim_ch_data(25, 19) = 0
sim_ch_data(25, 20) = 0
sim_ch_data(25, 21) = 0
sim_ch_data(25, 22) = 0
sim_ch_data(25, 23) = 0
sim_ch_data(25, 24) = 3
sim_ch_data(25, 25) = 9
sim_ch_data(25, 26) = 3
sim_ch_data(25, 27) = 0
sim_ch_data(25, 28) = 0
sim_ch_data(25, 29) = 0
sim_ch_data(25, 30) = 0
sim_ch_data(25, 31) = 0
sim_ch_data(25, 32) = 0
sim_ch_data(25, 33) = 0
sim_ch_data(25, 34) = 0


' ù
sim_ch_data(26, 1) = 0
sim_ch_data(26, 2) = 0
sim_ch_data(26, 3) = 0
sim_ch_data(26, 4) = 0
sim_ch_data(26, 5) = 0
sim_ch_data(26, 6) = 0
sim_ch_data(26, 7) = 3
sim_ch_data(26, 8) = 1
sim_ch_data(26, 9) = 0
sim_ch_data(26, 10) = 0
sim_ch_data(26, 11) = 0
sim_ch_data(26, 12) = 0
sim_ch_data(26, 13) = 0
sim_ch_data(26, 14) = 0
sim_ch_data(26, 15) = 0
sim_ch_data(26, 16) = 0
sim_ch_data(26, 17) = 0
sim_ch_data(26, 18) = 0
sim_ch_data(26, 19) = 0
sim_ch_data(26, 20) = 0
sim_ch_data(26, 21) = 0
sim_ch_data(26, 22) = 0
sim_ch_data(26, 23) = 0
sim_ch_data(26, 24) = 3
sim_ch_data(26, 25) = 3
sim_ch_data(26, 26) = 9
sim_ch_data(26, 27) = 0
sim_ch_data(26, 28) = 0
sim_ch_data(26, 29) = 0
sim_ch_data(26, 30) = 0
sim_ch_data(26, 31) = 0
sim_ch_data(26, 32) = 0
sim_ch_data(26, 33) = 0
sim_ch_data(26, 34) = 0


' ú
sim_ch_data(27, 1) = 0
sim_ch_data(27, 2) = 0
sim_ch_data(27, 3) = 0
sim_ch_data(27, 4) = 0
sim_ch_data(27, 5) = 0
sim_ch_data(27, 6) = 0
sim_ch_data(27, 7) = 0
sim_ch_data(27, 8) = 0
sim_ch_data(27, 9) = 0
sim_ch_data(27, 10) = 0
sim_ch_data(27, 11) = 0
sim_ch_data(27, 12) = 0
sim_ch_data(27, 13) = 0
sim_ch_data(27, 14) = 0
sim_ch_data(27, 15) = 0
sim_ch_data(27, 16) = 0
sim_ch_data(27, 17) = 0
sim_ch_data(27, 18) = 0
sim_ch_data(27, 19) = 0
sim_ch_data(27, 20) = 0
sim_ch_data(27, 21) = 0
sim_ch_data(27, 22) = 0
sim_ch_data(27, 23) = 0
sim_ch_data(27, 24) = 0
sim_ch_data(27, 25) = 0
sim_ch_data(27, 26) = 0
sim_ch_data(27, 27) = 9
sim_ch_data(27, 28) = 0
sim_ch_data(27, 29) = 3
sim_ch_data(27, 30) = 0
sim_ch_data(27, 31) = 0
sim_ch_data(27, 32) = 0
sim_ch_data(27, 33) = 0
sim_ch_data(27, 34) = 0


' û
sim_ch_data(28, 1) = 1
sim_ch_data(28, 2) = 0
sim_ch_data(28, 3) = 0
sim_ch_data(28, 4) = 0
sim_ch_data(28, 5) = 0
sim_ch_data(28, 6) = 1
sim_ch_data(28, 7) = 0
sim_ch_data(28, 8) = 0
sim_ch_data(28, 9) = 2
sim_ch_data(28, 10) = 0
sim_ch_data(28, 11) = 0
sim_ch_data(28, 12) = 0
sim_ch_data(28, 13) = 0
sim_ch_data(28, 14) = 0
sim_ch_data(28, 15) = 1
sim_ch_data(28, 16) = 0
sim_ch_data(28, 17) = 0
sim_ch_data(28, 18) = 0
sim_ch_data(28, 19) = 0
sim_ch_data(28, 20) = 1
sim_ch_data(28, 21) = 0
sim_ch_data(28, 22) = 0
sim_ch_data(28, 23) = 0
sim_ch_data(28, 24) = 0
sim_ch_data(28, 25) = 0
sim_ch_data(28, 26) = 0
sim_ch_data(28, 27) = 0
sim_ch_data(28, 28) = 9
sim_ch_data(28, 29) = 0
sim_ch_data(28, 30) = 1
sim_ch_data(28, 31) = 1
sim_ch_data(28, 32) = 1
sim_ch_data(28, 33) = 0
sim_ch_data(28, 34) = 1


' ü
sim_ch_data(29, 1) = 0
sim_ch_data(29, 2) = 0
sim_ch_data(29, 3) = 0
sim_ch_data(29, 4) = 0
sim_ch_data(29, 5) = 0
sim_ch_data(29, 6) = 0
sim_ch_data(29, 7) = 0
sim_ch_data(29, 8) = 0
sim_ch_data(29, 9) = 0
sim_ch_data(29, 10) = 0
sim_ch_data(29, 11) = 0
sim_ch_data(29, 12) = 0
sim_ch_data(29, 13) = 0
sim_ch_data(29, 14) = 0
sim_ch_data(29, 15) = 0
sim_ch_data(29, 16) = 0
sim_ch_data(29, 17) = 0
sim_ch_data(29, 18) = 0
sim_ch_data(29, 19) = 0
sim_ch_data(29, 20) = 0
sim_ch_data(29, 21) = 0
sim_ch_data(29, 22) = 0
sim_ch_data(29, 23) = 0
sim_ch_data(29, 24) = 0
sim_ch_data(29, 25) = 0
sim_ch_data(29, 26) = 0
sim_ch_data(29, 27) = 3
sim_ch_data(29, 28) = 0
sim_ch_data(29, 29) = 9
sim_ch_data(29, 30) = 0
sim_ch_data(29, 31) = 0
sim_ch_data(29, 32) = 0
sim_ch_data(29, 33) = 0
sim_ch_data(29, 34) = 0


' ý
sim_ch_data(30, 1) = 1
sim_ch_data(30, 2) = 0
sim_ch_data(30, 3) = 0
sim_ch_data(30, 4) = 0
sim_ch_data(30, 5) = 0
sim_ch_data(30, 6) = 3
sim_ch_data(30, 7) = 0
sim_ch_data(30, 8) = 0
sim_ch_data(30, 9) = 1
sim_ch_data(30, 10) = 0
sim_ch_data(30, 11) = 0
sim_ch_data(30, 12) = 0
sim_ch_data(30, 13) = 0
sim_ch_data(30, 14) = 0
sim_ch_data(30, 15) = 1
sim_ch_data(30, 16) = 0
sim_ch_data(30, 17) = 0
sim_ch_data(30, 18) = 0
sim_ch_data(30, 19) = 0
sim_ch_data(30, 20) = 1
sim_ch_data(30, 21) = 0
sim_ch_data(30, 22) = 0
sim_ch_data(30, 23) = 0
sim_ch_data(30, 24) = 0
sim_ch_data(30, 25) = 0
sim_ch_data(30, 26) = 0
sim_ch_data(30, 27) = 0
sim_ch_data(30, 28) = 1
sim_ch_data(30, 29) = 0
sim_ch_data(30, 30) = 9
sim_ch_data(30, 31) = 1
sim_ch_data(30, 32) = 1
sim_ch_data(30, 33) = 0
sim_ch_data(30, 34) = 3


' þ
sim_ch_data(31, 1) = 1
sim_ch_data(31, 2) = 0
sim_ch_data(31, 3) = 0
sim_ch_data(31, 4) = 0
sim_ch_data(31, 5) = 0
sim_ch_data(31, 6) = 1
sim_ch_data(31, 7) = 0
sim_ch_data(31, 8) = 0
sim_ch_data(31, 9) = 1
sim_ch_data(31, 10) = 0
sim_ch_data(31, 11) = 0
sim_ch_data(31, 12) = 0
sim_ch_data(31, 13) = 0
sim_ch_data(31, 14) = 0
sim_ch_data(31, 15) = 1
sim_ch_data(31, 16) = 0
sim_ch_data(31, 17) = 0
sim_ch_data(31, 18) = 0
sim_ch_data(31, 19) = 0
sim_ch_data(31, 20) = 2
sim_ch_data(31, 21) = 0
sim_ch_data(31, 22) = 0
sim_ch_data(31, 23) = 0
sim_ch_data(31, 24) = 0
sim_ch_data(31, 25) = 0
sim_ch_data(31, 26) = 0
sim_ch_data(31, 27) = 0
sim_ch_data(31, 28) = 1
sim_ch_data(31, 29) = 0
sim_ch_data(31, 30) = 1
sim_ch_data(31, 31) = 9
sim_ch_data(31, 32) = 1
sim_ch_data(31, 33) = 0
sim_ch_data(31, 34) = 1


' ÿ
sim_ch_data(32, 1) = 2
sim_ch_data(32, 2) = 0
sim_ch_data(32, 3) = 0
sim_ch_data(32, 4) = 0
sim_ch_data(32, 5) = 0
sim_ch_data(32, 6) = 1
sim_ch_data(32, 7) = 0
sim_ch_data(32, 8) = 0
sim_ch_data(32, 9) = 1
sim_ch_data(32, 10) = 0
sim_ch_data(32, 11) = 0
sim_ch_data(32, 12) = 0
sim_ch_data(32, 13) = 0
sim_ch_data(32, 14) = 0
sim_ch_data(32, 15) = 1
sim_ch_data(32, 16) = 0
sim_ch_data(32, 17) = 0
sim_ch_data(32, 18) = 0
sim_ch_data(32, 19) = 0
sim_ch_data(32, 20) = 1
sim_ch_data(32, 21) = 0
sim_ch_data(32, 22) = 0
sim_ch_data(32, 23) = 0
sim_ch_data(32, 24) = 0
sim_ch_data(32, 25) = 0
sim_ch_data(32, 26) = 0
sim_ch_data(32, 27) = 0
sim_ch_data(32, 28) = 1
sim_ch_data(32, 29) = 0
sim_ch_data(32, 30) = 1
sim_ch_data(32, 31) = 1
sim_ch_data(32, 32) = 9
sim_ch_data(32, 33) = 0
sim_ch_data(32, 34) = 1


' .
sim_ch_data(33, 1) = 0
sim_ch_data(33, 2) = 0
sim_ch_data(33, 3) = 0
sim_ch_data(33, 4) = 0
sim_ch_data(33, 5) = 0
sim_ch_data(33, 6) = 0
sim_ch_data(33, 7) = 0
sim_ch_data(33, 8) = 0
sim_ch_data(33, 9) = 0
sim_ch_data(33, 10) = 0
sim_ch_data(33, 11) = 0
sim_ch_data(33, 12) = 0
sim_ch_data(33, 13) = 0
sim_ch_data(33, 14) = 0
sim_ch_data(33, 15) = 0
sim_ch_data(33, 16) = 0
sim_ch_data(33, 17) = 0
sim_ch_data(33, 18) = 0
sim_ch_data(33, 19) = 0
sim_ch_data(33, 20) = 0
sim_ch_data(33, 21) = 0
sim_ch_data(33, 22) = 0
sim_ch_data(33, 23) = 0
sim_ch_data(33, 24) = 0
sim_ch_data(33, 25) = 0
sim_ch_data(33, 26) = 0
sim_ch_data(33, 27) = 0
sim_ch_data(33, 28) = 0
sim_ch_data(33, 29) = 0
sim_ch_data(33, 30) = 0
sim_ch_data(33, 31) = 0
sim_ch_data(33, 32) = 0
sim_ch_data(33, 33) = 0
sim_ch_data(33, 34) = 0


' ¸
sim_ch_data(34, 1) = 1
sim_ch_data(34, 2) = 0
sim_ch_data(34, 3) = 0
sim_ch_data(34, 4) = 0
sim_ch_data(34, 5) = 0
sim_ch_data(34, 6) = 9
sim_ch_data(34, 7) = 0
sim_ch_data(34, 8) = 0
sim_ch_data(34, 9) = 2
sim_ch_data(34, 10) = 0
sim_ch_data(34, 11) = 0
sim_ch_data(34, 12) = 0
sim_ch_data(34, 13) = 0
sim_ch_data(34, 14) = 0
sim_ch_data(34, 15) = 1
sim_ch_data(34, 16) = 0
sim_ch_data(34, 17) = 0
sim_ch_data(34, 18) = 0
sim_ch_data(34, 19) = 0
sim_ch_data(34, 20) = 1
sim_ch_data(34, 21) = 0
sim_ch_data(34, 22) = 0
sim_ch_data(34, 23) = 0
sim_ch_data(34, 24) = 0
sim_ch_data(34, 25) = 0
sim_ch_data(34, 26) = 0
sim_ch_data(34, 27) = 0
sim_ch_data(34, 28) = 1
sim_ch_data(34, 29) = 0
sim_ch_data(34, 30) = 2
sim_ch_data(34, 31) = 1
sim_ch_data(34, 32) = 1
sim_ch_data(34, 33) = 0
sim_ch_data(34, 34) = 9

    ' Populate sim_ch with data
    For i = 1 To 34
        For j = 1 To 34
            sim_ch(i, j) = sim_ch_data(i, j)
        Next j
    Next i

    ' Define and populate information letters array (inf_letters) as Integer
Dim inf_letters_data(1 To 34, 1 To 2) As Integer
inf_letters_data(1, 1) = 802: inf_letters_data(1, 2) = 959 ' à
inf_letters_data(2, 1) = 1232: inf_letters_data(2, 2) = 1129 ' á
inf_letters_data(3, 1) = 944: inf_letters_data(3, 2) = 859 ' â
inf_letters_data(4, 1) = 1253: inf_letters_data(4, 2) = 1193 ' ã
inf_letters_data(5, 1) = 1064: inf_letters_data(5, 2) = 951 ' ä
inf_letters_data(6, 1) = 759: inf_letters_data(6, 2) = 1232 ' å
inf_letters_data(7, 1) = 1432: inf_letters_data(7, 2) = 1432 ' æ
inf_letters_data(8, 1) = 1193: inf_letters_data(8, 2) = 993 ' ç
inf_letters_data(9, 1) = 802: inf_letters_data(9, 2) = 767 ' è
inf_letters_data(10, 1) = 1329: inf_letters_data(10, 2) = 1993 ' é
inf_letters_data(11, 1) = 1032: inf_letters_data(11, 2) = 929 ' ê
inf_letters_data(12, 1) = 967: inf_letters_data(12, 2) = 1276 ' ë
inf_letters_data(13, 1) = 1053: inf_letters_data(13, 2) = 944 ' ì
inf_letters_data(14, 1) = 848: inf_letters_data(14, 2) = 711 ' í
inf_letters_data(15, 1) = 695: inf_letters_data(15, 2) = 853 ' î
inf_letters_data(16, 1) = 1088: inf_letters_data(16, 2) = 454 ' ï
inf_letters_data(17, 1) = 929: inf_letters_data(17, 2) = 1115 ' ð
inf_letters_data(18, 1) = 895: inf_letters_data(18, 2) = 793 ' ñ
inf_letters_data(19, 1) = 848: inf_letters_data(19, 2) = 1002 ' ò
inf_letters_data(20, 1) = 1115: inf_letters_data(20, 2) = 1129 ' ó
inf_letters_data(21, 1) = 1793: inf_letters_data(21, 2) = 1022 ' ô
inf_letters_data(22, 1) = 1259: inf_letters_data(22, 2) = 1329 ' õ ' [0] manually decreased! was 1359
inf_letters_data(23, 1) = 1593: inf_letters_data(23, 2) = 1393 ' ö
inf_letters_data(24, 1) = 1276: inf_letters_data(24, 2) = 1212 ' ÷
inf_letters_data(25, 1) = 1476: inf_letters_data(25, 2) = 1012 ' ø
inf_letters_data(26, 1) = 1676: inf_letters_data(26, 2) = 1676 ' ù
inf_letters_data(27, 1) = 1993: inf_letters_data(27, 2) = 3986 ' ú
inf_letters_data(28, 1) = 1193: inf_letters_data(28, 2) = 3986 ' û
inf_letters_data(29, 1) = 1253: inf_letters_data(29, 2) = 3986 ' ü
inf_letters_data(30, 1) = 1676: inf_letters_data(30, 2) = 1232 ' ý
inf_letters_data(31, 1) = 1476: inf_letters_data(31, 2) = 1793 ' þ
inf_letters_data(32, 1) = 1159: inf_letters_data(32, 2) = 967 ' ÿ
inf_letters_data(33, 1) = 0: inf_letters_data(33, 2) = 0 '
inf_letters_data(34, 1) = 1300: inf_letters_data(34, 2) = 900 ' ¸  ' set manually - dk


    ' Populate inf_letters with data
    For i = 1 To 34
        inf_letters(i, 1) = inf_letters_data(i, 1)
        inf_letters(i, 2) = inf_letters_data(i, 2)
    Next i

    ' Initialize exceptions vocabulary
    Set exceptions_voc = CreateObject("Scripting.Dictionary")
    Set exceptions_voc_first = CreateObject("Scripting.Dictionary")
    With exceptions_voc
.add "áåëûì áåëî", True
.add "áîëüøå ìåíüøå", True
.add "áîëüøå áîëåå", True
.add "áîëåå áîëüøå", True
.add "áû âû", True
.add "âèíû âèíîâàòûé", True
.add "âîëåé íåâîëåé", True
.add "âðåìÿ âðåìåíè", True
.add "âñåãî íàâñåãî", True
.add "âû áû", True
.add "äàæå óæå", True
.add "äðóã äðóãà", True
.add "äðóã äðóãå", True
.add "äðóã äðóãîì", True
.add "äðóã äðóãó", True
.add "äóðàê äóðàêîì", True
.add "åñëè åñëè", True
.add "çâîíêà çâîíêà", True
.add "èëè èëè", True
.add "êàê òàê", True
.add "êîíöå êîíöîâ", True
.add "êîðêè êîðêè", True
.add "êòî ÷òî", True
.add "ëèáî ëèáî", True
.add "ìàëî ïîìàëó", True
.add "ìåíüøå áîëüøå", True
.add "íà÷àòü ñíà÷àëà", True
.add "íå íà", True
.add "íå íå", True
.add "íå íè", True
.add "íåãî íåò", True
.add "íè íà", True
.add "íè íå", True
.add "íè íè", True
.add "íî íà", True
.add "íî íå", True
.add "íî íè", True
.add "íîâûå íîâûå", True
.add "îáúÿòü íåîáúÿòíîå", True
.add "îäíîìó òîìó", True
.add "ïîëíûì ïîëíî", True
.add "ïîñòîëüêó ïîñêîëüêó", True
.add "òàê êàê", True
.add "òåì ÷åì", True
.add "òî òî", True
.add "òîãäà êîãäà", True
.add "õà õà", True
.add "÷åì òåì", True
.add "÷òî òî", True
.add "÷óòü ÷óòü", True
.add "øàã øàãîì", True
.add "ýòîé ÷òî", True
.add "ýòîò ÷òî", True
        
    End With

    ' Populate exceptions_voc_first with the first words of the exception phrases
    Dim key As Variant
    For Each key In exceptions_voc.Keys
        exceptions_voc_first(Split(key, " ")(0)) = True
    Next key
End Sub

' CheckVoc function - checks if a word pair is an exception
Function CheckVoc(w1 As String, w2 As String) As Boolean
    If exceptions_voc_first.Exists(w1) Then
        CheckVoc = exceptions_voc.Exists(w1 & " " & w2)
    Else
        CheckVoc = False
    End If
End Function

' Function to calculate psychological length
Function Implen(x As Integer) As Double
    If x = 2 Then
        Implen = 5
    Else
        Implen = x - ((x - 1) * (x - 1) / 36) + (4.1 / x)
    End If
End Function

' Function to compare two words, a and b, to calculate two types of similarity metrics between them
Function InforSameDiff(a As String, b As String) As Variant
    Dim count_same As Integer
    Dim avg_res_same As Double
    Dim count_diff As Integer
    Dim avg_res_diff As Double
    Dim i As Integer, n As Integer

    ' Initialize values
    avg_res_same = 0
    avg_res_diff = 0
    count_same = 0
    count_diff = 0

    n = Len(a)
    For i = 1 To n
        Dim charIndexA As Integer
        charIndexA = AscW(Mid(a, i, 1)) - 1071 ' Convert character to array index

        If charIndexA >= 1 And charIndexA <= 34 Then
            If InStr(1, b, Mid(a, i, 1)) > 0 Then
                count_same = count_same + 1
                avg_res_same = avg_res_same + (inf_letters(charIndexA, IIf(i = 1, 2, 1)) - avg_res_same) / count_same
            Else
                count_diff = count_diff + 1
                avg_res_diff = avg_res_diff + (inf_letters(charIndexA, IIf(i = 1, 2, 1)) - avg_res_diff) / count_diff
            End If
        End If
    Next i

    For i = 1 To Len(b)
        Dim charIndexB As Integer
        charIndexB = AscW(Mid(b, i, 1)) - 1071
        If charIndexB >= 1 And charIndexB <= 34 Then
            If InStr(1, a, Mid(b, i, 1)) = 0 Then
                count_diff = count_diff + 1
                avg_res_diff = avg_res_diff + (inf_letters(charIndexB, IIf(i = 1, 2, 1)) - avg_res_diff) / count_diff
            End If
        End If
    Next i

    ' Return the average results safely
    InforSameDiff = Array(avg_res_same, avg_res_diff)
End Function


' Function to calculate similarity score between two words
Function SimWords(a As String, b As String) As Double
    ' Convert both words to lowercase to ignore capitalization differences
    a = LCase(a)
    b = LCase(b)
    
    ' Ensure neither input word is empty to avoid division by zero
    If Len(a) = 0 Or Len(b) = 0 Then
        SimWords = 0
        Exit Function
    End If

    Dim dissimilarity_threshold As Long
    dissimilarity_threshold = 24000
    Dim info As Variant
    info = InforSameDiff(a, b)
    
    ' Check if words are too dissimilar right away
    If info(1) >= dissimilarity_threshold Then
        SimWords = 0
        Exit Function
    End If

    Dim alen As Integer, blen As Integer
    alen = Len(a): blen = Len(b)

    ' Always set the shorter word as a for alignment
    If alen > blen Then
        Dim temp As String
        temp = a: a = b: b = temp
        temp = alen: alen = blen: blen = temp
    End If

    ' Reciprocal lengths
    Dim reciproc_3_alen As Double, reciproc_3_blen As Double
    reciproc_3_alen = 1 / (3 * alen)
    reciproc_3_blen = 1 / (3 * blen)

    Dim res As Double, resa As Double, partlen As Integer
    Dim ta As Integer, tb As Integer, prir As Double, dist As Integer
    Dim tx As Integer, ty As Integer

    ' Main similarity calculation loop
    For partlen = 1 To alen
        resa = 0
        For ta = 0 To alen - partlen
            For tb = 0 To blen - partlen
                prir = 0
                For tx = ta + 1 To ta + partlen
                    ty = tb + (tx - ta)
                    
                    ' Calculate indices for sim_ch lookup
                    Dim indexA As Integer
                    Dim indexB As Integer
                    indexA = AscW(Mid(a, tx, 1)) - 1071
                    indexB = AscW(Mid(b, ty, 1)) - 1071
                    
                    ' Ensure indices are within bounds of sim_ch (1 to 34)
                    If indexA >= 1 And indexA <= 34 And indexB >= 1 And indexB <= 34 Then
                        prir = prir + sim_ch(indexA, indexB)
                    End If
                Next tx
                If prir = 0 Then GoTo Next_tb
                If ta > 0 Then prir = prir - prir * ta * reciproc_3_alen
                If tb > 0 Then prir = prir - prir * tb * reciproc_3_blen

                dist = (blen - (tb + partlen)) + ta
                If dist < 3 Then prir = prir + prir * (2 - dist) * 0.333
                If prir > resa Then resa = prir
Next_tb:
            Next tb
        Next ta
        If resa > partlen * 6 Then
            prir = resa
            dist = (alen + blen) * 0.375 + 1
            res = res + resa + prir * (partlen - IIf(dist < alen, dist, alen)) / (2 * dist)
        End If
    Next partlen

    ' Verify pair boundary
    If IsNumeric(res) And res > 0 Then
        For partlen = 1 To alen
            resa = resa + 9 * partlen
        Next partlen
        res = (res * info(0) / resa) * (dissimilarity_threshold - info(1)) / dissimilarity_threshold
        res = res - (res * (blen - alen) / (2 * blen))
        SimWords = res * alen * blen / (Implen(alen) * Implen(blen))
    Else
        ' Boundary fallback to zero if any boundary issue detected
        SimWords = 0
    End If
End Function


' Main function
'Function Fresheye: includes the improved handling for capitalized words
Function Fresheye(sourceText As String, sensitivity_threshold As Double, context_size As Integer) As String
    start_time = now
    
    ' Initialize options dictionary for settings
    Set options = CreateObject("Scripting.Dictionary")
    options.add "sensitivity_threshold", sensitivity_threshold
    options.add "context_size", context_size


    Dim total_badness As Double
    Dim words_checked As Long
    total_badness = 0
    words_checked = 0

    Dim group_color_index As Integer
    group_color_index = 1  ' Initialize the color index for group coloring

    ' Preprocess source text to treat paragraph marks as word delimiters
    Dim textWithDelimiters As String
    textWithDelimiters = Replace(sourceText, vbCrLf, " ")
    textWithDelimiters = Replace(textWithDelimiters, vbCr, " ")
    textWithDelimiters = Replace(textWithDelimiters, vbLf, " ")

    ' Split the modified text into individual words
    Dim words() As String
    words = Split(textWithDelimiters, " ")

    ' Initialize the queue to store the context words and their positions
    Dim contextQueue As Collection
    Set contextQueue = New Collection

    ' Collection to store all unique bad word pairs and their color groups
    Set badwords = New Collection
    Dim badWordDetails As New Collection
    Dim seenPairs As Object
    Set seenPairs = CreateObject("Scripting.Dictionary")

    Dim i As Long, badness As Double
    For i = 0 To UBound(words)
        
        ' Remove the first word if the queue exceeds context size
        If contextQueue.count >= context_size Then
            contextQueue.Remove 1
        End If
        
        ' Add the current word and its position to the queue, converted to lowercase for comparison
        contextQueue.add Array(LCase(CStr(words(i))), i)
        
        ' Compare current word with all other words in the queue
        Dim k As Long
        For k = 1 To contextQueue.count
            Dim queueWord As Variant
            queueWord = contextQueue(k)
            
            Dim pairKey As String
            pairKey = CStr(queueWord(0)) & "-" & LCase(words(i))
            
            ' Skip if comparing the word with itself or if pair already processed
            If queueWord(1) = i Or seenPairs.Exists(pairKey) Then GoTo NextWord

            ' Increment words_checked for each comparison
            words_checked = words_checked + 1
            
            ' Calculate similarity between the two words, ignoring capitalization
            badness = SimWords(CStr(queueWord(0)), CStr(LCase(words(i))))

            If badness > sensitivity_threshold Then
                total_badness = total_badness + badness

                ' Assign color based on the group_color_index
                badwords.add Array(queueWord(0), queueWord(1), group_color_index)
                badwords.add Array(words(i), i, group_color_index)
                badWordDetails.add "Pair: " & queueWord(0) & " - " & words(i)

                seenPairs.add pairKey, True
            End If

NextWord:
        Next k

        ' Cycle color index after processing a word
        group_color_index = (group_color_index Mod 5) + 1
    Next i

    Dim average_badness As Double
    average_badness = IIf(words_checked > 0, total_badness / words_checked, 0)
    
    Fresheye = "Àíàëèç çàâåðø¸í:" & vbCrLf & _
               "Êîìáèíàöèé ñëîâ ïðîâåðåíî: " & words_checked & vbCrLf & _
               "Ñðåäíÿÿ ïëîõîñòü: " & Round(average_badness, 2) & vbCrLf & _
                vbCrLf & _
               "Òåïåðü ðàñêðàñèì ñëîâà"

    Dim badWordsSummary As String
    Dim detail As Variant
    badWordsSummary = "Bad Word Pairs Found:" & vbCrLf
    For Each detail In badWordDetails
        badWordsSummary = badWordsSummary & detail & vbCrLf
    Next detail
''    MsgBox badWordsSummary, vbInformation, "Bad Word Pairs Summary"

End Function


' Function to highlight bad words in the selected text
Sub HighlightBadWords()
    Dim wordInfo As Variant
    Dim badWord As String
    Dim colorIndex As Integer
    Dim wordRange As range
    Dim wordPosition As Long
    Dim characterPosition As Long
    Dim searchRange As range
    Dim i As Long

    ' Define the array of syntax symbols to exclude
    Dim syntaxSymbols As Variant
    syntaxSymbols = Array(",", ".", "/", "?", "!", ")", "(", "[", "]", "<", ">", """", "\", "«", "»", "“", "”", "„", "—", ";", ":")

    Set searchRange = Selection.range.Duplicate

    ' Initialize the character position counter for the start of the selected range
    characterPosition = 0

    ' Iterate through each recorded bad word position
    For Each wordInfo In badwords
        badWord = wordInfo(0)            ' The bad word itself
        wordPosition = wordInfo(1) + 1    ' Adjust for zero-based indexing
        colorIndex = wordInfo(2)          ' Color group index

        ' Reset the character position counter for this word
        characterPosition = 0

        ' Loop through words in searchRange to reach the target word by word count
        Dim currentWordIndex As Long
        currentWordIndex = 1
        
        Dim tempRange As range
        Set tempRange = searchRange.Duplicate
        
        ' Traverse words until reaching the specified word index
        Do While currentWordIndex < wordPosition And characterPosition < Len(searchRange.text)
            ' Find the next space or paragraph mark
            Dim nextSpace As Long
            Dim nextParagraph As Long

            nextSpace = InStr(characterPosition + 1, searchRange.text, " ")
            nextParagraph = InStr(characterPosition + 1, searchRange.text, vbCr)

            ' Determine the next boundary
            If (nextParagraph > 0 And (nextSpace = 0 Or nextParagraph < nextSpace)) Then
                characterPosition = nextParagraph
                currentWordIndex = currentWordIndex + 1 ' Count paragraph as a word
            ElseIf nextSpace > 0 Then
                characterPosition = nextSpace
                currentWordIndex = currentWordIndex + 1
            Else
                Exit Do ' No more spaces or paragraphs found
            End If
        Loop

        ' Handle the special case where the bad word is the first word in the text
        If wordPosition = 1 Then
            characterPosition = 0 ' Ensure characterPosition points to the start of the selection
        End If

        ' At this point, characterPosition should be the start of the desired word
        If characterPosition >= 0 Then
            Set wordRange = searchRange.Duplicate
            wordRange.SetRange start:=searchRange.start + characterPosition, _
                               End:=searchRange.start + characterPosition + Len(badWord)

            ' Check and remove syntax symbols from the start and end of the range if they exist
            ' Trim symbols from the end
            Do While InStr(1, Join(syntaxSymbols, ""), Right(wordRange.text, 1)) > 0
                wordRange.End = wordRange.End - 1
            Loop
            ' Trim symbols from the start
            Do While InStr(1, Join(syntaxSymbols, ""), Left(wordRange.text, 1)) > 0
                wordRange.start = wordRange.start + 1
            Loop

            ' Apply highlight color only if the range is valid (after trimming symbols)
            If Len(wordRange.text) > 0 Then
                wordRange.Shading.BackgroundPatternColor = highlightColors(colorIndex)
            End If
        End If
    Next wordInfo
End Sub

