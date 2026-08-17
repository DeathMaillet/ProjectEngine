Attribute VB_Name = "mod_GanttLockTrace"
Option Explicit

'===============================================================================
' MODULE : mod_GanttLockTrace
' DOMAINE / DOMAIN : Gantt
'
' FR
' Capture les checkpoints techniques du domaine lorsque la trace est explicitement activee.
' Est desactive par defaut et ne modifie pas le workflow trace.
'
' EN
' Captures domain technical checkpoints when tracing is explicitly enabled.
' Is disabled by default and does not alter the traced workflow.
'
' CONTRATS / CONTRACTS : GanttLockTrace_Begin, GanttLockTrace_End, GanttLockTrace_IsEnabled, GanttLockTrace_SetSuppressConsole, GanttLockTrace_ShowConsole, GanttLockTrace_Log
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private gLockTraceEnabled As Boolean
Private gLockTracePath As String
Private gLockTraceCounter As Long
Private gLockTraceSuppressConsole As Boolean

'------------------------------------------------------------------------------
' FR: Active la trace LOCK vers un fichier texte pour les smokes instrumentes.
' EN: Enables LOCK tracing to a text file for instrumented smokes.
'------------------------------------------------------------------------------
Public Sub GanttLockTrace_Begin(ByVal outputPath As String)

    gLockTracePath = outputPath
    gLockTraceCounter = 0
    gLockTraceSuppressConsole = False
    gLockTraceEnabled = (Len(Trim$(outputPath)) > 0)

    If gLockTraceEnabled Then
        GanttLockTrace_WriteRaw "TRACE 00 - Begin LOCK trace"
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Desactive la trace LOCK.
' EN: Disables LOCK tracing.
'------------------------------------------------------------------------------
Public Sub GanttLockTrace_End()

    If gLockTraceEnabled Then
        GanttLockTrace_WriteRaw "TRACE END - End LOCK trace"
    End If

    gLockTraceEnabled = False
    gLockTracePath = vbNullString
    gLockTraceSuppressConsole = False

End Sub

'------------------------------------------------------------------------------
' FR: Retourne True si la trace LOCK est activee.
' EN: Returns True when LOCK tracing is enabled.
'------------------------------------------------------------------------------
Public Function GanttLockTrace_IsEnabled() As Boolean

    GanttLockTrace_IsEnabled = gLockTraceEnabled

End Function

'------------------------------------------------------------------------------
' FR: Active/desactive le bypass console uniquement pour les smokes instrumentes.
' EN: Enables/disables console bypass only for instrumented smokes.
'------------------------------------------------------------------------------
Public Sub GanttLockTrace_SetSuppressConsole(ByVal enabled As Boolean)

    gLockTraceSuppressConsole = enabled
    GanttLockTrace_Log "Suppress console set to " & CStr(enabled)

End Sub

'------------------------------------------------------------------------------
' FR: Affiche normalement la console, ou la bypass en mode smoke instrumente.
' EN: Shows the console normally, or bypasses it in instrumented smoke mode.
'------------------------------------------------------------------------------
Public Sub GanttLockTrace_ShowConsole(ByVal messages As Collection, ByVal checkpoint As String)

    If gLockTraceEnabled And gLockTraceSuppressConsole Then
        GanttLockTrace_Log checkpoint & ": CalcBridge_ShowPlanningConsole skipped by instrumented smoke"
        Exit Sub
    End If

    CalcBridge_ShowPlanningConsole messages

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute un checkpoint LOCK sans effet si la trace est desactivee.
' EN: Adds a LOCK checkpoint with no effect when tracing is disabled.
'------------------------------------------------------------------------------
Public Sub GanttLockTrace_Log(ByVal message As String)

    If Not gLockTraceEnabled Then Exit Sub

    gLockTraceCounter = gLockTraceCounter + 1
    GanttLockTrace_WriteRaw "TRACE " & Format$(gLockTraceCounter, "000") & " - " & message

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Gantt Lock Trace Write Raw dans le stockage possede par le domaine.
' EN: Writes or synchronizes Gantt Lock Trace Write Raw in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub GanttLockTrace_WriteRaw(ByVal message As String)

    Dim f As Integer

    If Len(Trim$(gLockTracePath)) = 0 Then Exit Sub

    On Error Resume Next
    f = FreeFile
    Open gLockTracePath For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & message
    Close #f
    On Error GoTo 0

End Sub
