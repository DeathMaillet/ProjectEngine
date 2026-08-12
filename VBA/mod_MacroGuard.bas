Attribute VB_Name = "mod_MacroGuard"
Option Explicit

'===============================================================================
' MODULE : mod_MacroGuard
' DOMAINE / DOMAIN : Shared Infrastructure
'
' FR
' Protege les macros contre les executions concurrentes et transporte les demandes d'abandon.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Protects macros from concurrent execution and carries abort requests.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : BeginMacroRun, EndMacroRun, RequestMacroAbort, IsMacroRunActive, IsMacroAbortRequested, AbortIfRequested, ShowAbortMessageOnce
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gMacroRunActive As Boolean
Private gMacroAbortRequested As Boolean
Private gMacroAbortSource As String
Private gMacroAbortMessageFR As String
Private gMacroAbortMessageEN As String
Private gMacroAbortPopupShown As Boolean

'------------------------------------------------------------------------------
' FR: Ouvre le cycle de traitement Macro Run.
' EN: Begins the Macro Run processing cycle.
'------------------------------------------------------------------------------
Public Sub BeginMacroRun(Optional ByVal sourceName As String = "")
    gMacroRunActive = True
    gMacroAbortRequested = False
    gMacroAbortSource = sourceName
    gMacroAbortMessageFR = ""
    gMacroAbortMessageEN = ""
    gMacroAbortPopupShown = False
    BeginPlanningWorkflow sourceName
End Sub

'------------------------------------------------------------------------------
' FR: Ferme le cycle de traitement Macro Run.
' EN: Ends the Macro Run processing cycle.
'------------------------------------------------------------------------------
Public Sub EndMacroRun()
    EndPlanningWorkflow
    gMacroRunActive = False
    gMacroAbortRequested = False
    gMacroAbortSource = ""
    gMacroAbortMessageFR = ""
    gMacroAbortMessageEN = ""
    gMacroAbortPopupShown = False
    EndPlanningEventRun
End Sub

'------------------------------------------------------------------------------
' FR: Enregistre ou applique la valeur Request Macro Abort dans le contexte MacroGuard courant.
' EN: Records or applies the Request Macro Abort value in the current MacroGuard context.
'------------------------------------------------------------------------------

Public Sub RequestMacroAbort( _
    ByVal sourceName As String, _
    ByVal frText As String, _
    ByVal enText As String)

    gMacroAbortRequested = True
    gMacroAbortSource = sourceName
    gMacroAbortMessageFR = frText
    gMacroAbortMessageEN = enText

    On Error Resume Next
    Application.EnableEvents = False
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' FR: Indique si Macro Run Active est vrai pour le contexte courant.
' EN: Returns whether Macro Run Active is true for the current context.
'------------------------------------------------------------------------------
Public Function IsMacroRunActive() As Boolean
    IsMacroRunActive = gMacroRunActive
End Function

'------------------------------------------------------------------------------
' FR: Indique si Macro Abort Requested est vrai pour le contexte courant.
' EN: Returns whether Macro Abort Requested is true for the current context.
'------------------------------------------------------------------------------
Public Function IsMacroAbortRequested() As Boolean
    IsMacroAbortRequested = gMacroAbortRequested
End Function

'------------------------------------------------------------------------------
' FR: Enregistre ou applique la valeur Abort If Requested dans le contexte MacroGuard courant.
' EN: Records or applies the Abort If Requested value in the current MacroGuard context.
'------------------------------------------------------------------------------

Public Sub AbortIfRequested(Optional ByVal sourceName As String = "")

    If Not gMacroAbortRequested Then Exit Sub

    If Trim$(sourceName) <> "" Then
        Err.Raise vbObjectError + 2901, sourceName, _
            gMacroAbortMessageFR & vbCrLf & gMacroAbortMessageEN
    Else
        Err.Raise vbObjectError + 2901, "MacroAbort", _
            gMacroAbortMessageFR & vbCrLf & gMacroAbortMessageEN
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Affiche Abort Message Once pour l'utilisateur ou le diagnostic.
' EN: Shows Abort Message Once for the user or diagnostics.
'------------------------------------------------------------------------------
Public Sub ShowAbortMessageOnce()

    Dim consoleMessages As Collection

    If Not gMacroAbortRequested Then Exit Sub
    If gMacroAbortPopupShown Then Exit Sub

    gMacroAbortPopupShown = True

    Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        "FR:" & vbCrLf & _
        gMacroAbortMessageFR & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        gMacroAbortMessageEN

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub



