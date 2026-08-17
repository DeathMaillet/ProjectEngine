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
    Dim sheetBefore As String
    Dim startedAt As Double

    GanttOpenLifecycle_BeginTrace "Workbook_Open"
    sheetBefore = ThisWorkbook_ActiveSheetName()

    startedAt = Timer
    Init_AppEvents
    GanttOpenLifecycle_Record "Workbook_Open.Init_AppEvents", "", sheetBefore, ThisWorkbook_ActiveSheetName(), ThisWorkbook_MsSince(startedAt)

    startedAt = Timer
    Settings_HydrateRuntimeState
    GanttOpenLifecycle_Record "Workbook_Open.Settings_HydrateRuntimeState", "", ThisWorkbook_ActiveSheetName(), ThisWorkbook_ActiveSheetName(), ThisWorkbook_MsSince(startedAt)

    startedAt = Timer
    Ensure_WBS_Main_Buttons
    GanttOpenLifecycle_Record "Workbook_Open.Ensure_WBS_Main_Buttons", "", ThisWorkbook_ActiveSheetName(), ThisWorkbook_ActiveSheetName(), ThisWorkbook_MsSince(startedAt)

    If UCase$(ThisWorkbook_ActiveSheetName()) = "GANTT" Then
        StartGanttDragWatchIfActive
        GanttColdState_MarkPending "Workbook_OpenGanttPassive"
    Else
        GanttColdState_MarkPending "Workbook_OpenNonGantt"
    End If

    StartGanttDragWatchIfActive
    GanttOpenLifecycle_Record "Workbook_Open.Complete", "", sheetBefore, ThisWorkbook_ActiveSheetName(), 0#
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
            GanttDrag_EnsureArmed
            If GanttDeferredRender_IsPending() Then
                GanttDeferredRender_FinishIfPending "Workbook_SheetActivate"
            End If
            GanttDrag_EnsureArmed
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

    GanttDrag_EnsureArmed

SafeExit:

End Sub

Private Function ThisWorkbook_ActiveSheetName() As String

    On Error Resume Next
    If Not Application.ActiveSheet Is Nothing Then
        If Application.ActiveSheet.Parent Is Me Then
            ThisWorkbook_ActiveSheetName = CStr(Application.ActiveSheet.Name)
        End If
    End If
    On Error GoTo 0

End Function

Private Function ThisWorkbook_MsSince(ByVal startedAt As Double) As Double

    Dim nowValue As Double

    nowValue = Timer
    If nowValue < startedAt Then nowValue = nowValue + 86400#
    ThisWorkbook_MsSince = (nowValue - startedAt) * 1000#

End Function

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









