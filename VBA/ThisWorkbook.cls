Option Explicit

'===============================================================================
' MODULE : ThisWorkbook
' DOMAINE / DOMAIN : Workbook Lifecycle
'
' FR
' Recoit les callbacks Excel du classeur et initialise les services d'evenements et de drag.
' Reste un point d'entree Excel ; il ne fournit aucun service metier.
'
' EN
' Receives workbook Excel callbacks and initializes event and drag services.
' Remains an Excel entry point and provides no business service.
'
' CONTRATS / CONTRACTS : Init_AppEvents
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Workbook_Open, Workbook_Activate, Workbook_SheetActivate, Workbook_SheetDeactivate, Workbook_Deactivate, Workbook_BeforeClose
'===============================================================================


Dim gAppEvents As clsAppEvents

'------------------------------------------------------------------------------
' FR: Initialise les evenements, les reglages, les boutons WBS et le watcher Gantt a l'ouverture.
' EN: Initializes events, settings, WBS buttons, and the Gantt watcher on workbook open.
'------------------------------------------------------------------------------
Private Sub Workbook_Open()
    Init_AppEvents
    Settings_HydrateRuntimeState
    Ensure_WBS_Main_Buttons
    StartGanttDragWatchIfActive
End Sub

'------------------------------------------------------------------------------
' FR: Rebranche les evenements applicatifs et relance le watcher Gantt si necessaire.
' EN: Reconnects application events and restarts the Gantt watcher when needed.
'------------------------------------------------------------------------------
Private Sub Workbook_Activate()
    Init_AppEvents
    StartGanttDragWatchIfActive
End Sub

'------------------------------------------------------------------------------
' FR: Traite l'activation d'une feuille et demarre le watcher drag sur GANTT.
' EN: Handles sheet activation and starts the drag watcher on GANTT.
'------------------------------------------------------------------------------
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Init_AppEvents
    If Not Sh Is Nothing Then
        If CStr(Sh.Name) = "GANTT" Then
            Gantt_RepairPendingGeometryIfNeeded
            GanttUiControls_EnsureCanonical Sh
            GanttDrag_StartWatch
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' FR: Met en pause le watcher drag lorsque la feuille GANTT est quittee.
' EN: Pauses the drag watcher when the GANTT sheet is left.
'------------------------------------------------------------------------------
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    If Sh Is Nothing Then Exit Sub
    If CStr(Sh.Name) = "GANTT" Then GanttDrag_PauseForLifecycle
End Sub

'------------------------------------------------------------------------------
' FR: Met en pause le watcher drag lorsque le classeur perd le focus.
' EN: Pauses the drag watcher when the workbook loses focus.
'------------------------------------------------------------------------------
Private Sub Workbook_Deactivate()
    GanttDrag_PauseForLifecycle
End Sub

'------------------------------------------------------------------------------
' FR: Arrete le watcher drag avant la fermeture du classeur.
' EN: Stops the drag watcher before the workbook closes.
'------------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    GanttDrag_StopWatch False
End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Start Gantt Drag Watch If Active dans l'etat runtime du composant.
' EN: Activates or initializes Start Gantt Drag Watch If Active in the component runtime state.
'------------------------------------------------------------------------------

Private Sub StartGanttDragWatchIfActive()

    On Error GoTo SafeExit

    If Application.ActiveSheet Is Nothing Then Exit Sub
    If Not (Application.ActiveSheet.Parent Is Me) Then Exit Sub
    If CStr(Application.ActiveSheet.Name) <> "GANTT" Then Exit Sub

    Gantt_RepairPendingGeometryIfNeeded
    GanttUiControls_EnsureCanonical Application.ActiveSheet
    GanttDrag_StartWatch

SafeExit:

End Sub

'------------------------------------------------------------------------------
' FR: Initialise la valeur App Events et en prend une copie conforme au contrat du composant.
' EN: Initializes the App Events value and takes a contract-compliant copy of its input.
'------------------------------------------------------------------------------

Public Sub Init_AppEvents()

    If gAppEvents Is Nothing Then
        Set gAppEvents = New clsAppEvents
    End If

    If gAppEvents.App Is Nothing Then
        Set gAppEvents.App = Application
    End If

End Sub









