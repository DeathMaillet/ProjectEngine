Attribute VB_Name = "mod_RunButtonsTrace"
Option Explicit

'===============================================================================
' MODULE : mod_RunButtonsTrace
' DOMAINE / DOMAIN : Shared Infrastructure
'
' FR
' Capture les checkpoints techniques du domaine lorsque la trace est explicitement activee.
' Est desactive par defaut et ne modifie pas le workflow trace.
'
' EN
' Captures domain technical checkpoints when tracing is explicitly enabled.
' Is disabled by default and does not alter the traced workflow.
'
' CONTRATS / CONTRACTS : RunButtonsTrace_Enable, RunButtonsTrace_Disable, RunButtonsTrace_Checkpoint
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gRunButtonsTraceEnabled As Boolean
Private gRunButtonsTracePath As String
Private gRunButtonsTraceMacro As String
Private gRunButtonsTraceStep As Long

'------------------------------------------------------------------------------
' FR: Active une trace fichier explicite pour les smokes non interactifs RunButtons.
' EN: Enables explicit file tracing for non-interactive RunButtons smokes.
'------------------------------------------------------------------------------
Public Sub RunButtonsTrace_Enable(ByVal tracePath As String, Optional ByVal macroName As String = "")

    gRunButtonsTracePath = tracePath
    gRunButtonsTraceMacro = macroName
    gRunButtonsTraceStep = 0
    gRunButtonsTraceEnabled = (Trim$(tracePath) <> "")
    RunButtonsTrace_Checkpoint "Trace", "Enabled"

End Sub

'------------------------------------------------------------------------------
' FR: Desactive la trace RunButtons sans modifier le workflow utilisateur.
' EN: Disables RunButtons tracing without changing the user workflow.
'------------------------------------------------------------------------------
Public Sub RunButtonsTrace_Disable()

    RunButtonsTrace_Checkpoint "Trace", "Disabled"
    gRunButtonsTraceEnabled = False
    gRunButtonsTracePath = vbNullString
    gRunButtonsTraceMacro = vbNullString
    gRunButtonsTraceStep = 0

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit un checkpoint inoffensif si la trace RunButtons est active.
' EN: Writes a harmless checkpoint when RunButtons tracing is active.
'------------------------------------------------------------------------------
Public Sub RunButtonsTrace_Checkpoint(ByVal domain As String, ByVal detail As String)

    Dim fileNo As Integer
    Dim lineText As String

    If Not gRunButtonsTraceEnabled Then Exit Sub
    If Trim$(gRunButtonsTracePath) = "" Then Exit Sub

    On Error Resume Next
    gRunButtonsTraceStep = gRunButtonsTraceStep + 1
    lineText = Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
        " | TRACE " & Format$(gRunButtonsTraceStep, "000") & _
        " | Macro=" & gRunButtonsTraceMacro & _
        " | Domain=" & domain & _
        " | " & detail

    fileNo = FreeFile
    Open gRunButtonsTracePath For Append As #fileNo
    Print #fileNo, lineText
    Close #fileNo
    On Error GoTo 0

End Sub
