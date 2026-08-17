Attribute VB_Name = "mod_GanttDragWatch"
Option Explicit

'===============================================================================
' MODULE : mod_GanttDragWatch
' DOMAINE / DOMAIN : Gantt
'
' FR
' Surveille les Shapes Gantt, convertit drag/resize en inputs TEST/SCENARIO et gere le timer Win32.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Watches Gantt Shapes, converts drag/resize into TEST/SCENARIO inputs and manages the Win32 timer.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : GanttDrag_StartWatch, GanttDrag_StopWatch, GanttDrag_RebuildWatchMaps, GanttDrag_IsWatching, GanttDrag_ReconcileWatchState, GanttDrag_PauseForLifecycle
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : GanttDrag_StartWatch, GanttDrag_StopWatch, GanttDrag_TimerProc, GanttDrag_TimerProc
'===============================================================================


#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As LongPtr) As LongPtr

    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIDEvent As LongPtr) As Long

    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" ( _
        ByRef lpPerformanceCount As Currency) As Long

    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" ( _
        ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

    Private Declare Function KillTimer Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long) As Long

    Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
        ByRef lpPerformanceCount As Currency) As Long

    Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
        ByRef lpFrequency As Currency) As Long
#End If

Private Const GANTT_DRAG_SHEET As String = "GANTT"
Private Const GANTT_DRAG_SCALE_DAY As String = "DAY"
Private Const GANTT_DRAG_SCALE_WEEK As String = "WEEK"
Private Const GANTT_DRAG_SCALE_MONTH As String = "MONTH"
Private Const GANTT_DRAG_TIMER_MS As Long = 200
Private Const GANTT_DRAG_FIRST_TASK_ROW As Long = 5
Private Const GANTT_DRAG_HEADER_ROW As Long = 4
Private Const GANTT_DRAG_FIRST_TIMELINE_COL As Long = 11
Private Const GANTT_DRAG_COL_TEST_START As Long = 5
Private Const GANTT_DRAG_COL_TEST_FINISH As Long = 6
Private Const GANTT_DRAG_CHANGE_TOLERANCE As Double = 0.2
Private Const GANTT_DRAG_MAX_STRUCTURAL_ERRORS As Long = 3

#If VBA7 Then
    Private gTimerId As LongPtr
#Else
    Private gTimerId As Long
#End If

Private gWatchRequested As Boolean
Private gWatchEnabled As Boolean
Private gInTick As Boolean
Private gSuspendDepth As Long
Private gStructuralErrorCount As Long
Private gShapeState As Object
Private gLayoutSignature As String
Private gLayoutState(0 To 11) As Double
Private gHasLayoutState As Boolean
Private gLastDebugStatus As String
Private gTransactionActive As Boolean
Private gLastTransactionResult As String
Private gLastWrittenCells As String
Private gTransactionCount As Long

Private gMetricsEnabled As Boolean
Private gMetricTicks As Long
Private gMetricIdleTicks As Long
Private gMetricDeltaTicks As Long
Private gMetricActionTicks As Long
Private gMetricBusySkips As Long
Private gMetricReentrantSkips As Long
Private gMetricSuspendedSkips As Long
Private gMetricLayoutReconciles As Long
Private gMetricMapRebuilds As Long
Private gMetricMapShapesInspected As Long
Private gMetricTickShapesInspected As Long
Private gMetricLayoutReads As Long
Private gMetricGeometryReads As Long
Private gMetricCellsRead As Long
Private gMetricExcelWrites As Long
Private gMetricEngineCalls As Long
Private gMetricRendererCalls As Long
Private gMetricFallbacks As Long
Private gMetricStructuralErrors As Long
Private gMetricTotalMs As Double
Private gMetricMaxMs As Double
Private gMetricLastMs As Double
Private gMetricLayoutMs As Double
Private gMetricSelectionMs As Double
Private gMetricGeometryMs As Double

'------------------------------------------------------------------------------
' FR: Recoit le callback externe GanttDrag_StartWatch et le relaie vers le workflow proprietaire.
' EN: Receives the external GanttDrag_StartWatch callback and routes it to the owning workflow.
'------------------------------------------------------------------------------

Public Sub GanttDrag_StartWatch()

    gWatchRequested = True
    GanttDrag_StartRuntime

End Sub

'------------------------------------------------------------------------------
' FR: Reprend le watcher sans detruire son runtime si le timer existe deja.
' EN: Resumes the watcher without destroying runtime when the timer already exists.
'------------------------------------------------------------------------------
Public Sub GanttDrag_ResumeWatch()

    GanttDrag_EnsureArmed

End Sub

'------------------------------------------------------------------------------
' FR: Garantit que le watcher est arme quand GANTT est actif, sans reset inutile.
' EN: Ensures the watcher is armed when GANTT is active, without unnecessary reset.
'------------------------------------------------------------------------------
Public Sub GanttDrag_EnsureArmed()

    gWatchRequested = True
    If Not GanttDrag_CanStartWatch() Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    If GanttDrag_IsWatching() Then
        gSuspendDepth = 0
        gInTick = False
        gTransactionActive = False
        GanttDrag_RebuildWatchMaps
    Else
        GanttDrag_StartRuntime
    End If

    GanttDrag_PrimeLocalPathForFirstInteraction

End Sub

Private Sub GanttDrag_PrimeLocalPathForFirstInteraction()

    On Error Resume Next
    GanttLocal_PrimeNormalState
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Recoit le callback externe GanttDrag_StopWatch et le relaie vers le workflow proprietaire.
' EN: Receives the external GanttDrag_StopWatch callback and routes it to the owning workflow.
'------------------------------------------------------------------------------

Public Sub GanttDrag_StopWatch(Optional ByVal showStatus As Boolean = False)

    GanttDrag_StopRuntime True, showStatus

End Sub

'------------------------------------------------------------------------------
' FR: Traite la map Rebuild Watch Maps sans modifier les donnees d'entree.
' EN: Handles the Rebuild Watch Maps map without mutating input data.
'------------------------------------------------------------------------------

Public Sub GanttDrag_RebuildWatchMaps()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim rowIndex As Long
    Dim taskType As String

    Set gShapeState = CreateObject("Scripting.Dictionary")
    If gMetricsEnabled Then gMetricMapRebuilds = gMetricMapRebuilds + 1

    If Not GanttDrag_IsGanttSheetActive(ws) Then Exit Sub
    If Not GanttDrag_IsSupportedTimelineScale() Then Exit Sub

    For Each shp In ws.Shapes
        If gMetricsEnabled Then gMetricMapShapesInspected = gMetricMapShapesInspected + 1
        rowIndex = 0
        taskType = vbNullString

        If GanttDrag_IsEligibleShape(ws, shp, rowIndex, taskType) Then
            GanttDrag_SaveShapeState shp, rowIndex, taskType
        End If
    Next shp

    GanttDrag_CaptureLayoutState ws
    gLayoutSignature = GanttDrag_BuildLayoutSignature(ws)

End Sub

'------------------------------------------------------------------------------
' FR: Indique si la valeur Watching satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Watching value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Public Function GanttDrag_IsWatching() As Boolean

    GanttDrag_IsWatching = (gWatchEnabled And gTimerId <> 0)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Last Debug Status sans modifier les donnees d'entree.
' EN: Returns the Last Debug Status value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_LastDebugStatus() As String

    GanttDrag_LastDebugStatus = gLastDebugStatus

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Last Transaction Result sans modifier les donnees d'entree.
' EN: Returns the Last Transaction Result value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_LastTransactionResult() As String

    GanttDrag_LastTransactionResult = gLastTransactionResult

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Last Written Cells sans modifier les donnees d'entree.
' EN: Returns the Last Written Cells value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_LastWrittenCells() As String

    GanttDrag_LastWrittenCells = gLastWrittenCells

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Transaction Count sans modifier les donnees d'entree.
' EN: Returns the Transaction Count value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_TransactionCount() As Long

    GanttDrag_TransactionCount = gTransactionCount

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Shape Watched satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Shape Watched value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function GanttDrag_IsShapeWatched(ByVal shapeName As String) As Boolean

    If gShapeState Is Nothing Then Exit Function
    GanttDrag_IsShapeWatched = gShapeState.Exists(shapeName)

End Function

'------------------------------------------------------------------------------
' FR: Aligne la valeur Reconcile Watch State avec le lifecycle courant sans perdre l'etat possede.
' EN: Aligns the Reconcile Watch State value with the current lifecycle without losing owned state.
'------------------------------------------------------------------------------

Public Sub GanttDrag_ReconcileWatchState()

    If Not gWatchRequested Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    If Not GanttDrag_CanStartWatch() Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    If GanttDrag_IsWatching() Then
        GanttDrag_RebuildWatchMaps
    Else
        GanttDrag_StartRuntime
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Aligne la valeur Pause For Lifecycle avec le lifecycle courant sans perdre l'etat possede.
' EN: Aligns the Pause For Lifecycle value with the current lifecycle without losing owned state.
'------------------------------------------------------------------------------

Public Sub GanttDrag_PauseForLifecycle()

    GanttDrag_StopRuntime False, False

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Start Runtime dans l'etat runtime du composant.
' EN: Activates or initializes Start Runtime in the component runtime state.
' FR - Contrat externe : callback timer ; le nom et la signature doivent rester stables.
' EN - External contract: timer callback; name and signature must remain stable.
'------------------------------------------------------------------------------

Private Sub GanttDrag_StartRuntime()

    If Not GanttDrag_CanStartWatch() Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    GanttDrag_StopRuntime False, False
    GanttDrag_RebuildWatchMaps

    If gShapeState Is Nothing Then Exit Sub
    If gShapeState.Count = 0 Then Exit Sub

    gStructuralErrorCount = 0
    gInTick = False
    gTransactionActive = False
    gLastTransactionResult = vbNullString
    gLastWrittenCells = vbNullString
    gTransactionCount = 0

#If VBA7 Then
    gTimerId = SetTimer(0, 0, GANTT_DRAG_TIMER_MS, AddressOf GanttDrag_TimerProc)
#Else
    gTimerId = SetTimer(0, 0, GANTT_DRAG_TIMER_MS, AddressOf GanttDrag_TimerProc)
#End If

    gWatchEnabled = (gTimerId <> 0)
    If Not gWatchEnabled Then gWatchRequested = False

End Sub

'------------------------------------------------------------------------------
' FR: Termine Stop Runtime et restaure l'etat runtime possede par le composant.
' EN: Ends Stop Runtime and restores runtime state owned by the component.
'------------------------------------------------------------------------------

Private Sub GanttDrag_StopRuntime( _
    ByVal clearRequest As Boolean, _
    ByVal showStatus As Boolean)

    On Error Resume Next

    gWatchEnabled = False
    gInTick = False
    gSuspendDepth = 0
    gTransactionActive = False

    If gTimerId <> 0 Then
        KillTimer 0, gTimerId
        gTimerId = 0
    End If

    Set gShapeState = Nothing
    gLayoutSignature = vbNullString
    gHasLayoutState = False
    If clearRequest Then gWatchRequested = False

    If showStatus Then Debug.Print "Gantt Drag Watch stopped."

    On Error GoTo 0

End Sub

#If VBA7 Then
'------------------------------------------------------------------------------
' FR: Callback Win32 du timer qui declenche la surveillance du drag Gantt.
' EN: Win32 timer callback that triggers Gantt drag monitoring.
'------------------------------------------------------------------------------
Private Sub GanttDrag_TimerProc( _
    ByVal hwnd As LongPtr, _
    ByVal uMsg As Long, _
    ByVal idEvent As LongPtr, _
    ByVal dwTime As Long)
#Else
'------------------------------------------------------------------------------
' FR: Callback Win32 du timer qui declenche la surveillance du drag Gantt.
' EN: Win32 timer callback that triggers Gantt drag monitoring.
'------------------------------------------------------------------------------
Private Sub GanttDrag_TimerProc( _
    ByVal hwnd As Long, _
    ByVal uMsg As Long, _
    ByVal idEvent As Long, _
    ByVal dwTime As Long)
#End If

    On Error GoTo StructuralError

    GanttDrag_RunGuardedTick
    gStructuralErrorCount = 0
    Exit Sub

StructuralError:
    GanttDrag_RecordStructuralError

End Sub

'------------------------------------------------------------------------------
' FR: Enregistre une erreur structurelle et coupe le watcher apres erreurs repetees.
' EN: Records a structural error and stops the watcher after repeated failures.
'------------------------------------------------------------------------------
Private Sub GanttDrag_RecordStructuralError()

    gStructuralErrorCount = gStructuralErrorCount + 1
    If gMetricsEnabled Then
        gMetricStructuralErrors = gMetricStructuralErrors + 1
        gMetricFallbacks = gMetricFallbacks + 1
    End If
    If gStructuralErrorCount >= GANTT_DRAG_MAX_STRUCTURAL_ERRORS Then
        GanttDrag_StopRuntime True, False
    End If
    gInTick = False

End Sub

'------------------------------------------------------------------------------
' FR: Applique les gardes de reentrance et mesure un tick sans empiler de timer.
' EN: Applies reentrancy guards and measures one tick without stacking timers.
'------------------------------------------------------------------------------
Private Sub GanttDrag_RunGuardedTick(Optional ByVal requireEnabledRuntime As Boolean = True)

    Dim startedAt As Double
    Dim elapsedMs As Double

    If requireEnabledRuntime Then
        If Not gWatchEnabled Then Exit Sub
    End If
    If gInTick Then
        If gMetricsEnabled Then gMetricReentrantSkips = gMetricReentrantSkips + 1
        Exit Sub
    End If
    If gSuspendDepth > 0 Then
        If gMetricsEnabled Then gMetricSuspendedSkips = gMetricSuspendedSkips + 1
        Exit Sub
    End If

    If gMetricsEnabled Then
        gMetricTicks = gMetricTicks + 1
        startedAt = GanttDrag_PerformanceNowMs()
    End If

    gInTick = True
    On Error GoTo SafeExit
    GanttDrag_WatchTick

SafeExit:
    gInTick = False
    If gMetricsEnabled Then
        elapsedMs = GanttDrag_PerformanceNowMs() - startedAt
        gMetricLastMs = elapsedMs
        gMetricTotalMs = gMetricTotalMs + elapsedMs
        If elapsedMs > gMetricMaxMs Then gMetricMaxMs = elapsedMs
    End If
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Retourne un temps haute resolution en millisecondes pour l'instrumentation.
' EN: Returns a high-resolution millisecond timestamp for instrumentation.
'------------------------------------------------------------------------------
Private Function GanttDrag_PerformanceNowMs() As Double

    Static frequency As Currency
    Dim counter As Currency

    If frequency = 0 Then QueryPerformanceFrequency frequency
    QueryPerformanceCounter counter
    If frequency <> 0 Then
        GanttDrag_PerformanceNowMs = (CDbl(counter) / CDbl(frequency)) * 1000#
    End If

End Function

'------------------------------------------------------------------------------
' FR: Reinitialise les compteurs de preuve du watcher sans modifier son lifecycle.
' EN: Resets watcher proof counters without changing its lifecycle.
'------------------------------------------------------------------------------
Public Sub GanttDragHarness_ResetMetrics(Optional ByVal enableMetrics As Boolean = True)

    gMetricsEnabled = enableMetrics
    gMetricTicks = 0
    gMetricIdleTicks = 0
    gMetricDeltaTicks = 0
    gMetricActionTicks = 0
    gMetricBusySkips = 0
    gMetricReentrantSkips = 0
    gMetricSuspendedSkips = 0
    gMetricLayoutReconciles = 0
    gMetricMapRebuilds = 0
    gMetricMapShapesInspected = 0
    gMetricTickShapesInspected = 0
    gMetricLayoutReads = 0
    gMetricGeometryReads = 0
    gMetricCellsRead = 0
    gMetricExcelWrites = 0
    gMetricEngineCalls = 0
    gMetricRendererCalls = 0
    gMetricFallbacks = 0
    gMetricStructuralErrors = 0
    gMetricTotalMs = 0#
    gMetricMaxMs = 0#
    gMetricLastMs = 0#
    gMetricLayoutMs = 0#
    gMetricSelectionMs = 0#
    gMetricGeometryMs = 0#

End Sub

'------------------------------------------------------------------------------
' FR: Execute un tick garde pour les harnais sans creer de second timer.
' EN: Runs one guarded harness tick without creating a second timer.
'------------------------------------------------------------------------------
Public Sub GanttDragHarness_RunTick()

    GanttDrag_RunGuardedTick False

End Sub

'------------------------------------------------------------------------------
' FR: Prouve que la garde de reentrance absorbe un tick concurrent.
' EN: Proves that the reentrancy guard absorbs a concurrent tick.
'------------------------------------------------------------------------------
Public Sub GanttDragHarness_ProbeReentrance()

    Dim oldInTick As Boolean

    oldInTick = gInTick
    gInTick = True
    GanttDrag_RunGuardedTick False
    gInTick = oldInTick

End Sub

'------------------------------------------------------------------------------
' FR: Injecte une erreur structurelle pour prouver le fallback sans erreur runtime.
' EN: Injects one structural error to prove fallback handling without a runtime fault.
'------------------------------------------------------------------------------
Public Sub GanttDragHarness_ProbeStructuralError()

    GanttDrag_RecordStructuralError

End Sub

'------------------------------------------------------------------------------
' FR: Invalide la watch map afin de tester sa reconstruction au prochain tick.
' EN: Invalidates the watch map so the next tick can prove automatic recovery.
'------------------------------------------------------------------------------
Public Sub GanttDragHarness_InvalidateWatchMap()

    Set gShapeState = Nothing

End Sub

'------------------------------------------------------------------------------
' FR: Expose une metrique nommee en lecture seule au harnais permanent.
' EN: Exposes one named read-only metric to the permanent harness.
'------------------------------------------------------------------------------
Public Function GanttDragHarness_GetMetric(ByVal metricName As String) As Double

    Select Case UCase$(Trim$(metricName))
        Case "TICKS": GanttDragHarness_GetMetric = gMetricTicks
        Case "IDLE_TICKS": GanttDragHarness_GetMetric = gMetricIdleTicks
        Case "DELTA_TICKS": GanttDragHarness_GetMetric = gMetricDeltaTicks
        Case "ACTION_TICKS": GanttDragHarness_GetMetric = gMetricActionTicks
        Case "BUSY_SKIPS": GanttDragHarness_GetMetric = gMetricBusySkips
        Case "REENTRANT_SKIPS": GanttDragHarness_GetMetric = gMetricReentrantSkips
        Case "SUSPENDED_SKIPS": GanttDragHarness_GetMetric = gMetricSuspendedSkips
        Case "LAYOUT_RECONCILES": GanttDragHarness_GetMetric = gMetricLayoutReconciles
        Case "MAP_REBUILDS": GanttDragHarness_GetMetric = gMetricMapRebuilds
        Case "MAP_SHAPES_INSPECTED": GanttDragHarness_GetMetric = gMetricMapShapesInspected
        Case "TICK_SHAPES_INSPECTED": GanttDragHarness_GetMetric = gMetricTickShapesInspected
        Case "LAYOUT_READS": GanttDragHarness_GetMetric = gMetricLayoutReads
        Case "GEOMETRY_READS": GanttDragHarness_GetMetric = gMetricGeometryReads
        Case "CELLS_READ": GanttDragHarness_GetMetric = gMetricCellsRead
        Case "EXCEL_WRITES": GanttDragHarness_GetMetric = gMetricExcelWrites
        Case "ENGINE_CALLS": GanttDragHarness_GetMetric = gMetricEngineCalls
        Case "RENDERER_CALLS": GanttDragHarness_GetMetric = gMetricRendererCalls
        Case "FALLBACKS": GanttDragHarness_GetMetric = gMetricFallbacks
        Case "STRUCTURAL_ERRORS": GanttDragHarness_GetMetric = gMetricStructuralErrors
        Case "TOTAL_MS": GanttDragHarness_GetMetric = gMetricTotalMs
        Case "MAX_MS": GanttDragHarness_GetMetric = gMetricMaxMs
        Case "LAST_MS": GanttDragHarness_GetMetric = gMetricLastMs
        Case "LAYOUT_MS": GanttDragHarness_GetMetric = gMetricLayoutMs
        Case "SELECTION_MS": GanttDragHarness_GetMetric = gMetricSelectionMs
        Case "GEOMETRY_MS": GanttDragHarness_GetMetric = gMetricGeometryMs
        Case "WATCHED_SHAPES"
            If Not gShapeState Is Nothing Then GanttDragHarness_GetMetric = gShapeState.Count
        Case "TIMER_ACTIVE": GanttDragHarness_GetMetric = IIf(gTimerId <> 0, 1, 0)
        Case "SUSPEND_DEPTH": GanttDragHarness_GetMetric = gSuspendDepth
        Case "TRANSACTIONS": GanttDragHarness_GetMetric = gTransactionCount
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne un nom réellement surveillé pour les scénarios du harnais.
' EN: Returns one actually watched name for harness scenarios.
'------------------------------------------------------------------------------
Public Function GanttDragHarness_FirstWatchedShapeName() As String

    Dim shapeName As Variant

    If gShapeState Is Nothing Then Exit Function
    For Each shapeName In gShapeState.Keys
        GanttDragHarness_FirstWatchedShapeName = CStr(shapeName)
        Exit Function
    Next shapeName

End Function

'------------------------------------------------------------------------------
' FR: Retourne une tâche standard surveillée pour tester le resize horizontal.
' EN: Returns one watched standard task for horizontal resize testing.
'------------------------------------------------------------------------------
Public Function GanttDragHarness_FirstWatchedTaskShapeName() As String

    Dim shapeName As Variant

    If gShapeState Is Nothing Then Exit Function
    For Each shapeName In gShapeState.Keys
        If Left$(CStr(shapeName), 5) = "TASK_" Then
            GanttDragHarness_FirstWatchedTaskShapeName = CStr(shapeName)
            Exit Function
        End If
    Next shapeName

End Function

Public Function GanttDragHarness_IsShapeWatched(ByVal shapeName As String) As Boolean

    If gShapeState Is Nothing Then Exit Function
    GanttDragHarness_IsShapeWatched = gShapeState.Exists(CStr(shapeName))

End Function

'------------------------------------------------------------------------------
' FR: Traite la reference Watch Tick sans modifier les donnees d'entree.
' EN: Handles the Watch Tick reference without mutating input data.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Private Sub GanttDrag_WatchTick()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim state As Variant
    Dim geometryChanged As Boolean
    Dim shapeName As String
    Dim newLeft As Double
    Dim newRight As Double
    Dim phaseStartedAt As Double

    If Not GanttDrag_IsGanttSheetActive(ws) Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    If Not GanttDrag_IsSupportedTimelineScale() Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    If IsPlanningWorkflowActive() Then
        If gMetricsEnabled Then gMetricBusySkips = gMetricBusySkips + 1
        Exit Sub
    End If
    If GetGanttInternalWrite() Then
        If gMetricsEnabled Then gMetricBusySkips = gMetricBusySkips + 1
        Exit Sub
    End If
    If Application.CalculationState <> xlDone Then
        If gMetricsEnabled Then gMetricBusySkips = gMetricBusySkips + 1
        Exit Sub
    End If

    If gShapeState Is Nothing Then GanttDrag_RebuildWatchMaps
    If gShapeState Is Nothing Then Exit Sub
    If gShapeState.Count = 0 Then
        GanttDrag_StopRuntime False, False
        Exit Sub
    End If

    If gMetricsEnabled Then phaseStartedAt = GanttDrag_PerformanceNowMs()
    If Not GanttDrag_TryGetSelectedWatchedShape(ws, shp, shapeName, state) Then
        If gMetricsEnabled Then gMetricSelectionMs = gMetricSelectionMs + _
            (GanttDrag_PerformanceNowMs() - phaseStartedAt)

        If gMetricsEnabled Then phaseStartedAt = GanttDrag_PerformanceNowMs()
        If GanttDrag_LayoutStateChanged(ws) Then
            If gMetricsEnabled Then gMetricLayoutMs = gMetricLayoutMs + _
                (GanttDrag_PerformanceNowMs() - phaseStartedAt)
            gLastDebugStatus = "LAYOUT_RECONCILED"
            If gMetricsEnabled Then
                gMetricDeltaTicks = gMetricDeltaTicks + 1
                gMetricLayoutReconciles = gMetricLayoutReconciles + 1
            End If
            GanttDrag_RebuildWatchMaps
            Exit Sub
        End If
        If gMetricsEnabled Then gMetricLayoutMs = gMetricLayoutMs + _
            (GanttDrag_PerformanceNowMs() - phaseStartedAt)

        If gMetricsEnabled Then gMetricIdleTicks = gMetricIdleTicks + 1
        Exit Sub
    End If
    If gMetricsEnabled Then gMetricSelectionMs = gMetricSelectionMs + _
        (GanttDrag_PerformanceNowMs() - phaseStartedAt)

    If gMetricsEnabled Then
        gMetricTickShapesInspected = gMetricTickShapesInspected + 1
        gMetricGeometryReads = gMetricGeometryReads + 2
    End If

    If gMetricsEnabled Then phaseStartedAt = GanttDrag_PerformanceNowMs()
    newLeft = CDbl(shp.Left)
    newRight = newLeft + CDbl(shp.Width)
    If gMetricsEnabled Then gMetricGeometryMs = gMetricGeometryMs + _
        (GanttDrag_PerformanceNowMs() - phaseStartedAt)
    geometryChanged = _
           Abs(CDbl(state(0)) - newLeft) > GANTT_DRAG_CHANGE_TOLERANCE _
        Or Abs(CDbl(state(1)) - newRight) > GANTT_DRAG_CHANGE_TOLERANCE

    If Not geometryChanged Then
        If gMetricsEnabled Then gMetricIdleTicks = gMetricIdleTicks + 1
        Exit Sub
    End If

    If gMetricsEnabled Then
        gMetricDeltaTicks = gMetricDeltaTicks + 1
        gMetricActionTicks = gMetricActionTicks + 1
    End If

    gLastDebugStatus = _
        shapeName & _
        " Left " & Format$(CDbl(state(0)), "0.00") & " -> " & Format$(newLeft, "0.00") & _
        " | Right " & Format$(CDbl(state(1)), "0.00") & " -> " & Format$(newRight, "0.00")

    Debug.Print gLastDebugStatus
    GanttDrag_HandleShapeChange ws, shp, state

End Sub

'------------------------------------------------------------------------------
' FR: Construit la signature du layout qui peut reprojeter les shapes sans drag utilisateur.
' EN: Builds the layout signature that can reproject Shapes without a user drag.
'------------------------------------------------------------------------------
Private Function GanttDrag_BuildLayoutSignature(ByVal ws As Worksheet) As String

    Dim signature As String
    Dim colIndex As Long

    If ws Is Nothing Then Exit Function

    For colIndex = 1 To GANTT_DRAG_FIRST_TIMELINE_COL - 1
        signature = signature & _
            "|C" & CStr(colIndex) & "=" & _
            Format$(gLayoutState(colIndex - 1), "0.000")
    Next colIndex

    signature = signature & _
        "|TL_LEFT=" & Format$(gLayoutState(10), "0.000") & _
        "|TL_WIDTH=" & Format$(gLayoutState(11), "0.000")

    GanttDrag_BuildLayoutSignature = signature

End Function

'------------------------------------------------------------------------------
' FR: Capture les seules dimensions globales capables de reprojeter les shapes.
' EN: Captures only global dimensions that can reproject Shapes.
'------------------------------------------------------------------------------
Private Sub GanttDrag_CaptureLayoutState(ByVal ws As Worksheet)

    Dim colIndex As Long

    If ws Is Nothing Then Exit Sub

    For colIndex = 1 To GANTT_DRAG_FIRST_TIMELINE_COL - 1
        gLayoutState(colIndex - 1) = CDbl(ws.Columns(colIndex).Width)
    Next colIndex

    gLayoutState(10) = CDbl(ws.Cells(GANTT_DRAG_HEADER_ROW, GANTT_DRAG_FIRST_TIMELINE_COL).Left)
    gLayoutState(11) = CDbl(ws.Columns(GANTT_DRAG_FIRST_TIMELINE_COL).Width)
    gHasLayoutState = True

End Sub

'------------------------------------------------------------------------------
' FR: Detecte un changement global de layout avec douze lectures COM constantes.
' EN: Detects a global layout change with twelve constant COM reads.
'------------------------------------------------------------------------------
Private Function GanttDrag_LayoutStateChanged(ByVal ws As Worksheet) As Boolean

    Dim currentValue As Double
    Dim colIndex As Long

    If ws Is Nothing Then Exit Function
    If Not gHasLayoutState Then
        GanttDrag_LayoutStateChanged = True
        Exit Function
    End If

    For colIndex = 1 To GANTT_DRAG_FIRST_TIMELINE_COL - 1
        currentValue = CDbl(ws.Columns(colIndex).Width)
        If gMetricsEnabled Then gMetricLayoutReads = gMetricLayoutReads + 1
        If Abs(currentValue - gLayoutState(colIndex - 1)) > 0.001 Then
            GanttDrag_LayoutStateChanged = True
            Exit Function
        End If
    Next colIndex

    currentValue = CDbl(ws.Cells(GANTT_DRAG_HEADER_ROW, GANTT_DRAG_FIRST_TIMELINE_COL).Left)
    If gMetricsEnabled Then gMetricLayoutReads = gMetricLayoutReads + 1
    If Abs(currentValue - gLayoutState(10)) > 0.001 Then
        GanttDrag_LayoutStateChanged = True
        Exit Function
    End If

    currentValue = CDbl(ws.Columns(GANTT_DRAG_FIRST_TIMELINE_COL).Width)
    If gMetricsEnabled Then gMetricLayoutReads = gMetricLayoutReads + 1
    GanttDrag_LayoutStateChanged = _
        (Abs(currentValue - gLayoutState(11)) > 0.001)

End Function

'------------------------------------------------------------------------------
' FR: Retourne uniquement la shape Gantt actuellement selectionnee et surveillee.
' EN: Returns only the currently selected and watched Gantt Shape.
'------------------------------------------------------------------------------
Private Function GanttDrag_TryGetSelectedWatchedShape( _
    ByVal ws As Worksheet, _
    ByRef shp As Shape, _
    ByRef shapeName As String, _
    ByRef state As Variant) As Boolean

    Dim selectionObject As Object
    Dim selectedRange As ShapeRange

    On Error GoTo SafeExit

    If ws Is Nothing Then Exit Function
    Set selectionObject = Application.Selection
    If selectionObject Is Nothing Then Exit Function

    Set selectedRange = selectionObject.ShapeRange
    If selectedRange Is Nothing Then Exit Function
    If selectedRange.Count <> 1 Then Exit Function

    Set shp = selectedRange.Item(1)
    If shp Is Nothing Then Exit Function
    If Not (shp.Parent Is ws) Then Exit Function

    shapeName = CStr(shp.Name)
    If gShapeState Is Nothing Then Exit Function
    If Not gShapeState.Exists(shapeName) Then Exit Function

    state = gShapeState(shapeName)
    GanttDrag_TryGetSelectedWatchedShape = True

SafeExit:

End Function

'------------------------------------------------------------------------------
' FR: Traite la collection Handle Shape Change sans modifier les donnees d'entree.
' EN: Handles the Handle Shape Change collection without mutating input data.
'------------------------------------------------------------------------------

Private Sub GanttDrag_HandleShapeChange( _
    ByVal ws As Worksheet, _
    ByVal shp As Shape, _
    ByVal oldState As Variant)

    Dim writtenCells As Collection
    Dim consoleMessages As Collection
    Dim revertMessages As Collection
    Dim testSucceeded As Boolean
    Dim revertSucceeded As Boolean
    Dim ganttRebuilt As Boolean
    Dim revertGanttRebuilt As Boolean
    Dim shouldRunTest As Boolean
    Dim displayFailure As Boolean
    Dim displayConsole As Boolean
    Dim dragInfo As Object
    Dim simulationMode As String
    Dim engineScope As clsPerfScope
    Dim consoleScope As clsPerfScope

    On Error GoTo TransactionError

    If ws Is Nothing Then Exit Sub
    If shp Is Nothing Then Exit Sub
    If gTransactionActive Then Exit Sub

    Set writtenCells = New Collection
    shouldRunTest = GanttDrag_BuildTestInputs(ws, shp, oldState, writtenCells, dragInfo)

    If Not shouldRunTest Then
        GanttDrag_SaveShapeState shp, CLng(oldState(2)), CStr(oldState(3))
        Exit Sub
    End If

    'The released geometry is the user input. Keep it on screen until the
    'TEST/SCENARIO transaction converges; rollback restores oldState only on failure.
    Profiler_RecordOperation "GanttDragPreCommitRestoreSkipped", 1, 0#

    gTransactionActive = True
    gTransactionCount = gTransactionCount + 1
    gLastTransactionResult = "RUNNING"
    gLastWrittenCells = GanttDrag_CellList(writtenCells)
    GanttDrag_Suspend

    simulationMode = GanttDrag_NormalizedSimulationMode(GanttLive_GetActiveSimulationMode())
    GanttDrag_SetDragInfoMode dragInfo, simulationMode

    If Not GanttDrag_IsSupportedSimulationMode(simulationMode) Then
        GanttDrag_ClearWrittenCells writtenCells
        GanttDrag_RollbackShapeGeometry shp, oldState, "UnsupportedMode"
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        GanttDrag_AddUnsupportedModeMessage consoleMessages, dragInfo, simulationMode
        gLastTransactionResult = "NO_ACTIVE_MODE"
        displayConsole = True
        GoTo CleanExit
    End If

    Profiler_RecordOperation "GanttDrag_DatesWritten", writtenCells.Count, 0#
    Set engineScope = Profiler_BeginScope("GanttDrag_CommonEngineAfterDates", "Gantt Drag")
    testSucceeded = GanttDrag_RunSimulationTransactionByMode(simulationMode, consoleMessages, ganttRebuilt)
    Set engineScope = Nothing

    If testSucceeded Then
        gLastTransactionResult = "SUCCESS"
        GanttDrag_AddDragSuccessMessage consoleMessages, dragInfo
        displayConsole = True
    Else
        GanttDrag_ClearWrittenCells writtenCells
        revertSucceeded = GanttDrag_RunSimulationTransactionByMode(simulationMode, revertMessages, revertGanttRebuilt)

        If revertSucceeded Then
            gLastTransactionResult = "REVERTED"
        Else
            GanttDrag_RollbackShapeGeometry shp, oldState, "RevertFailed"
            gLastTransactionResult = "REVERT_FAILED"
        End If

        GanttDrag_AddDragFailureMessage consoleMessages, dragInfo
        displayFailure = True
        displayConsole = True
    End If

CleanExit:
    On Error Resume Next
    GanttDrag_RebuildWatchMaps
    GanttDrag_Resume
    gTransactionActive = False

    If displayFailure Or displayConsole Then
        If Not consoleMessages Is Nothing Then
            Set consoleScope = Profiler_BeginScope("GanttDrag_FinalConsole", "Gantt Drag")
            CalcBridge_ShowPlanningConsole consoleMessages
            Set consoleScope = Nothing
            GanttDrag_EnsureArmed
        End If
    End If
    On Error GoTo 0
    Exit Sub

TransactionError:
    gLastTransactionResult = "ERROR"
    If Not writtenCells Is Nothing Then GanttDrag_ClearWrittenCells writtenCells
    GanttDrag_RollbackShapeGeometry shp, oldState, "TransactionError"
    If consoleMessages Is Nothing Then Set consoleMessages = New Collection
    GanttDrag_SetDragInfoMode dragInfo, simulationMode
    GanttDrag_AddDragFailureMessage consoleMessages, dragInfo
    displayFailure = True
    displayConsole = True
    Resume CleanExit

End Sub
'------------------------------------------------------------------------------
' FR: Construit la collection Test Inputs a partir des donnees fournies par l'appelant.
' EN: Builds the Test Inputs collection from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function GanttDrag_BuildTestInputs( _
    ByVal ws As Worksheet, _
    ByVal shp As Shape, _
    ByVal oldState As Variant, _
    ByVal writtenCells As Collection, _
    ByRef dragInfo As Object) As Boolean

    Dim oldLeft As Double
    Dim oldRight As Double
    Dim oldWidth As Double
    Dim newLeft As Double
    Dim newRight As Double
    Dim leftChanged As Boolean
    Dim rightChanged As Boolean
    Dim widthChanged As Boolean
    Dim testStart As Date
    Dim testFinish As Date
    Dim milestoneDate As Date
    Dim ganttRow As Long
    Dim taskType As String

    On Error GoTo SafeExit

    oldLeft = CDbl(oldState(0))
    oldRight = CDbl(oldState(1))
    oldWidth = oldRight - oldLeft
    newLeft = CDbl(shp.Left)
    newRight = newLeft + CDbl(shp.Width)
    ganttRow = CLng(oldState(2))
    taskType = UCase$(Trim$(CStr(oldState(3))))

    leftChanged = Abs(newLeft - oldLeft) > GANTT_DRAG_CHANGE_TOLERANCE
    rightChanged = Abs(newRight - oldRight) > GANTT_DRAG_CHANGE_TOLERANCE
    widthChanged = Abs(CDbl(shp.Width) - oldWidth) > GANTT_DRAG_CHANGE_TOLERANCE

    If taskType = "MILESTONE" Then
        If Not leftChanged And Not rightChanged Then Exit Function
        If Not GanttDrag_DateFromX(ws, newLeft + (CDbl(shp.Width) / 2), 0, milestoneDate) Then Exit Function

        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_START), milestoneDate, writtenCells
        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_FINISH), milestoneDate, writtenCells
        Set dragInfo = GanttDrag_CreateDragInfo(ws, ganttRow, taskType, True, True, milestoneDate, milestoneDate)
        GanttDrag_BuildTestInputs = True
        Exit Function
    End If

    If taskType <> "TASK" Then Exit Function
    If Not leftChanged And Not rightChanged Then Exit Function

    If leftChanged And Not widthChanged Then
        If Not GanttDrag_DateFromX(ws, newLeft, -1, testStart) Then Exit Function
        If Not GanttDrag_DateFromX(ws, newRight, 1, testFinish) Then Exit Function

        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_START), testStart, writtenCells
        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_FINISH), testFinish, writtenCells
    ElseIf leftChanged And Not rightChanged Then
        If Not GanttDrag_DateFromX(ws, newLeft, -1, testStart) Then Exit Function
        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_START), testStart, writtenCells
    ElseIf Not leftChanged And rightChanged Then
        If Not GanttDrag_DateFromX(ws, newRight, 1, testFinish) Then Exit Function
        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_FINISH), testFinish, writtenCells
    Else
        If Not GanttDrag_DateFromX(ws, newLeft, -1, testStart) Then Exit Function
        If Not GanttDrag_DateFromX(ws, newRight, 1, testFinish) Then Exit Function

        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_START), testStart, writtenCells
        GanttDrag_WriteTestCell ws.Cells(ganttRow, GANTT_DRAG_COL_TEST_FINISH), testFinish, writtenCells
    End If

    If writtenCells.Count > 0 Then
        Set dragInfo = GanttDrag_CreateDragInfo(ws, ganttRow, taskType, _
            GanttDrag_WrittenCellsContainColumn(writtenCells, GANTT_DRAG_COL_TEST_START), _
            GanttDrag_WrittenCellsContainColumn(writtenCells, GANTT_DRAG_COL_TEST_FINISH), _
            testStart, testFinish)
    End If

    GanttDrag_BuildTestInputs = (writtenCells.Count > 0)

SafeExit:

End Function
'------------------------------------------------------------------------------
' FR: Retourne la reference Date From X sans modifier les donnees d'entree.
' EN: Returns the Date From X reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_DateFromX( _
    ByVal ws As Worksheet, _
    ByVal xPos As Double, _
    ByVal anchorSide As Long, _
    ByRef resultDate As Date) As Boolean

    Select Case UCase$(Trim$(GetGanttTimelineScaleMode()))
        Case GANTT_DRAG_SCALE_WEEK
            GanttDrag_DateFromX = GanttDrag_WeekDateFromX(ws, xPos, anchorSide, resultDate)
        Case GANTT_DRAG_SCALE_MONTH
            GanttDrag_DateFromX = GanttDrag_MonthDateFromX(ws, xPos, anchorSide, resultDate)
        Case Else
            GanttDrag_DateFromX = GanttDrag_DayDateFromX(ws, xPos, anchorSide, resultDate)
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Day Date From X sans modifier les donnees d'entree.
' EN: Returns the Day Date From X reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_DayDateFromX( _
    ByVal ws As Worksheet, _
    ByVal xPos As Double, _
    ByVal anchorSide As Long, _
    ByRef resultDate As Date) As Boolean

    Dim lastCol As Long
    Dim c As Long
    Dim headerValue As Variant
    Dim anchorX As Double
    Dim distance As Double
    Dim bestDistance As Double
    Dim foundDate As Boolean

    If ws Is Nothing Then Exit Function

    lastCol = ws.Cells(GANTT_DRAG_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function

    bestDistance = 1E+30

    For c = GANTT_DRAG_FIRST_TIMELINE_COL To lastCol
        headerValue = ws.Cells(GANTT_DRAG_HEADER_ROW, c).value

        If IsDate(headerValue) Then
            Select Case anchorSide
                Case -1
                    anchorX = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Left
                Case 1
                    anchorX = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Left + _
                        ws.Cells(GANTT_DRAG_HEADER_ROW, c).Width
                Case Else
                    anchorX = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Left + _
                        (ws.Cells(GANTT_DRAG_HEADER_ROW, c).Width / 2)
            End Select

            distance = Abs(xPos - anchorX)
            If distance < bestDistance Then
                bestDistance = distance
                resultDate = DateValue(CDate(headerValue))
                foundDate = True
            End If
        End If
    Next c

    GanttDrag_DayDateFromX = foundDate

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Week Date From X sans modifier les donnees d'entree.
' EN: Returns the Week Date From X reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_WeekDateFromX( _
    ByVal ws As Worksheet, _
    ByVal xPos As Double, _
    ByVal anchorSide As Long, _
    ByRef resultDate As Date) As Boolean

    Dim lastCol As Long
    Dim c As Long
    Dim targetCol As Long
    Dim cellLeft As Double
    Dim cellWidth As Double
    Dim cellRight As Double
    Dim distance As Double
    Dim bestDistance As Double
    Dim fraction As Double
    Dim dayOffset As Long
    Dim weekStart As Date

    If ws Is Nothing Then Exit Function

    lastCol = ws.Cells(GANTT_DRAG_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function

    bestDistance = 1E+30

    For c = GANTT_DRAG_FIRST_TIMELINE_COL To lastCol
        cellLeft = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Left
        cellWidth = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Width
        cellRight = cellLeft + cellWidth

        If (anchorSide = 1 And xPos >= cellLeft And xPos <= cellRight) Or _
           (anchorSide <> 1 And xPos >= cellLeft And (xPos < cellRight Or c = lastCol)) Then
            targetCol = c
            Exit For
        End If

        distance = WorksheetFunction.Min(Abs(xPos - cellLeft), Abs(xPos - cellRight))
        If distance < bestDistance Then
            bestDistance = distance
            targetCol = c
        End If
    Next c

    If targetCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function
    If Not GanttDrag_TimelineWeekStart(ws, targetCol, weekStart) Then Exit Function

    cellLeft = ws.Cells(GANTT_DRAG_HEADER_ROW, targetCol).Left
    cellWidth = ws.Cells(GANTT_DRAG_HEADER_ROW, targetCol).Width
    If cellWidth <= 0 Then Exit Function

    fraction = (xPos - cellLeft) / cellWidth
    If fraction < 0# Then fraction = 0#
    If fraction > 1# Then fraction = 1#

    Select Case anchorSide
        Case 1
            dayOffset = GanttDrag_CeilPositive(fraction * 7#) - 1
        Case Else
            dayOffset = CLng(Int(fraction * 7#))
    End Select

    If dayOffset < 0 Then dayOffset = 0
    If dayOffset > 6 Then dayOffset = 6

    resultDate = DateAdd("d", dayOffset, weekStart)
    GanttDrag_WeekDateFromX = True

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Week Start sans modifier les donnees d'entree.
' EN: Returns the Timeline Week Start reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_TimelineWeekStart( _
    ByVal ws As Worksheet, _
    ByVal timelineCol As Long, _
    ByRef weekStart As Date) As Boolean

    Dim weekNum As Long
    Dim isoYear As Long
    Dim yearValue As Variant
    Dim labelValue As String
    Dim janFourth As Date

    If ws Is Nothing Then Exit Function
    If timelineCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function

    labelValue = CStr(ws.Cells(GANTT_DRAG_HEADER_ROW, timelineCol).value)
    weekNum = GanttDrag_ExtractFirstLong(labelValue)
    If weekNum < 1 Or weekNum > 53 Then Exit Function

    yearValue = ws.Cells(GANTT_DRAG_HEADER_ROW - 1, timelineCol).MergeArea.Cells(1, 1).value
    If Not IsNumeric(yearValue) Then Exit Function
    isoYear = CLng(yearValue)
    If isoYear < 1900 Then Exit Function

    janFourth = DateSerial(isoYear, 1, 4)
    weekStart = DateAdd("d", (weekNum - 1) * 7, janFourth - Weekday(janFourth, vbMonday) + 1)
    GanttDrag_TimelineWeekStart = True

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Extract First Long sans modifier les donnees d'entree.
' EN: Returns the Extract First Long value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_ExtractFirstLong(ByVal textValue As String) As Long

    Dim i As Long
    Dim ch As String
    Dim digits As String

    For i = 1 To Len(textValue)
        ch = Mid$(textValue, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf digits <> "" Then
            Exit For
        End If
    Next i

    If digits <> "" Then GanttDrag_ExtractFirstLong = CLng(digits)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Ceil Positive sans modifier les donnees d'entree.
' EN: Returns the Ceil Positive value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_CeilPositive(ByVal value As Double) As Long

    If value <= 0# Then
        GanttDrag_CeilPositive = 0
    Else
        GanttDrag_CeilPositive = CLng(-Int(-value))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Month Date From X sans modifier les donnees d'entree.
' EN: Returns the Month Date From X reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_MonthDateFromX( _
    ByVal ws As Worksheet, _
    ByVal xPos As Double, _
    ByVal anchorSide As Long, _
    ByRef resultDate As Date) As Boolean

    Dim targetCol As Long
    Dim cellLeft As Double
    Dim cellWidth As Double
    Dim fraction As Double
    Dim monthStart As Date
    Dim daysInMonth As Long
    Dim dayOffset As Long

    If ws Is Nothing Then Exit Function
    If Not GanttDrag_TimelineColumnFromX(ws, xPos, anchorSide, targetCol) Then Exit Function
    If Not GanttDrag_TimelineMonthStart(ws, targetCol, monthStart) Then Exit Function

    cellLeft = ws.Cells(GANTT_DRAG_HEADER_ROW, targetCol).Left
    cellWidth = ws.Cells(GANTT_DRAG_HEADER_ROW, targetCol).Width
    If cellWidth <= 0 Then Exit Function

    fraction = (xPos - cellLeft) / cellWidth
    If fraction < 0# Then fraction = 0#
    If fraction > 1# Then fraction = 1#

    daysInMonth = Day(DateSerial(Year(monthStart), Month(monthStart) + 1, 0))

    Select Case anchorSide
        Case 1
            dayOffset = GanttDrag_CeilPositive(fraction * CDbl(daysInMonth)) - 1
        Case Else
            dayOffset = CLng(Int(fraction * CDbl(daysInMonth)))
    End Select

    If dayOffset < 0 Then dayOffset = 0
    If dayOffset > daysInMonth - 1 Then dayOffset = daysInMonth - 1

    resultDate = DateAdd("d", dayOffset, monthStart)
    GanttDrag_MonthDateFromX = True

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Column From X sans modifier les donnees d'entree.
' EN: Returns the Timeline Column From X reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_TimelineColumnFromX( _
    ByVal ws As Worksheet, _
    ByVal xPos As Double, _
    ByVal anchorSide As Long, _
    ByRef targetCol As Long) As Boolean

    Dim lastCol As Long
    Dim c As Long
    Dim cellLeft As Double
    Dim cellWidth As Double
    Dim cellRight As Double
    Dim distance As Double
    Dim bestDistance As Double

    If ws Is Nothing Then Exit Function

    lastCol = ws.Cells(GANTT_DRAG_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function

    bestDistance = 1E+30

    For c = GANTT_DRAG_FIRST_TIMELINE_COL To lastCol
        cellLeft = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Left
        cellWidth = ws.Cells(GANTT_DRAG_HEADER_ROW, c).Width
        cellRight = cellLeft + cellWidth

        If (anchorSide = 1 And xPos >= cellLeft And xPos <= cellRight) Or _
           (anchorSide <> 1 And xPos >= cellLeft And (xPos < cellRight Or c = lastCol)) Then
            targetCol = c
            GanttDrag_TimelineColumnFromX = True
            Exit Function
        End If

        distance = Abs(xPos - cellLeft)
        If Abs(xPos - cellRight) < distance Then distance = Abs(xPos - cellRight)

        If distance < bestDistance Then
            bestDistance = distance
            targetCol = c
        End If
    Next c

    GanttDrag_TimelineColumnFromX = (targetCol >= GANTT_DRAG_FIRST_TIMELINE_COL)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Month Start sans modifier les donnees d'entree.
' EN: Returns the Timeline Month Start reference without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_TimelineMonthStart( _
    ByVal ws As Worksheet, _
    ByVal timelineCol As Long, _
    ByRef monthStart As Date) As Boolean

    Dim yearValue As Variant
    Dim yearNum As Long
    Dim monthNum As Long

    If ws Is Nothing Then Exit Function
    If timelineCol < GANTT_DRAG_FIRST_TIMELINE_COL Then Exit Function

    yearValue = ws.Cells(GANTT_DRAG_HEADER_ROW - 1, timelineCol).MergeArea.Cells(1, 1).value
    If Not IsNumeric(yearValue) Then Exit Function
    yearNum = CLng(yearValue)
    If yearNum < 1900 Then Exit Function

    monthNum = GanttDrag_MonthNumberFromLabel(CStr(ws.Cells(GANTT_DRAG_HEADER_ROW, timelineCol).value))
    If monthNum < 1 Or monthNum > 12 Then Exit Function

    monthStart = DateSerial(yearNum, monthNum, 1)
    GanttDrag_TimelineMonthStart = True

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Month Number From Label sans modifier les donnees d'entree.
' EN: Returns the Month Number From Label value without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_MonthNumberFromLabel(ByVal monthLabel As String) As Long

    Dim normalizedLabel As String

    normalizedLabel = LCase$(Trim$(monthLabel))

    Select Case normalizedLabel
        Case "jan", "janv", "jan.", "janv.": GanttDrag_MonthNumberFromLabel = 1
        Case "feb", "fév", "fev", "févr", "fevr", "fév.", "fev.", "févr.", "fevr.": GanttDrag_MonthNumberFromLabel = 2
        Case "mar", "mars": GanttDrag_MonthNumberFromLabel = 3
        Case "apr", "avr", "avr.": GanttDrag_MonthNumberFromLabel = 4
        Case "may", "mai": GanttDrag_MonthNumberFromLabel = 5
        Case "jun", "juin": GanttDrag_MonthNumberFromLabel = 6
        Case "jul", "juil", "juil.": GanttDrag_MonthNumberFromLabel = 7
        Case "aug", "août", "aout": GanttDrag_MonthNumberFromLabel = 8
        Case "sep", "sept", "sept.": GanttDrag_MonthNumberFromLabel = 9
        Case "oct", "oct.": GanttDrag_MonthNumberFromLabel = 10
        Case "nov", "nov.": GanttDrag_MonthNumberFromLabel = 11
        Case "dec", "déc", "dec.", "déc.": GanttDrag_MonthNumberFromLabel = 12
    End Select

End Function
'------------------------------------------------------------------------------
' FR: Construit la map Drag Info a partir des donnees fournies par l'appelant.
' EN: Builds the Drag Info map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function GanttDrag_CreateDragInfo( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal taskType As String, _
    ByVal changedStart As Boolean, _
    ByVal changedFinish As Boolean, _
    ByVal requestedStart As Variant, _
    ByVal requestedFinish As Variant) As Object

    Dim info As Object

    Set info = CreateObject("Scripting.Dictionary")

    info("TaskName") = Trim$(CStr(ws.Cells(ganttRow, 2).value))
    info("WBS") = Trim$(CStr(ws.Cells(ganttRow, 1).value))
    info("TaskType") = UCase$(Trim$(taskType))
    info("ChangedStart") = changedStart
    info("ChangedFinish") = changedFinish
    info("RequestedStart") = requestedStart
    info("RequestedFinish") = requestedFinish

    Set GanttDrag_CreateDragInfo = info

End Function

'------------------------------------------------------------------------------
' FR: Retourne la collection Written Cells Contain Column sans modifier les donnees d'entree.
' EN: Returns the Written Cells Contain Column collection without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_WrittenCellsContainColumn( _
    ByVal writtenCells As Collection, _
    ByVal columnIndex As Long) As Boolean

    Dim item As Variant

    If writtenCells Is Nothing Then Exit Function

    For Each item In writtenCells
        If CLng(item.Column) = columnIndex Then
            GanttDrag_WrittenCellsContainColumn = True
            Exit Function
        End If
    Next item

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Normalized Simulation Mode selon le contrat canonique du composant.
' EN: Normalizes or formats Normalized Simulation Mode according to the component contract.
'------------------------------------------------------------------------------

Private Function GanttDrag_NormalizedSimulationMode(ByVal simulationMode As String) As String

    GanttDrag_NormalizedSimulationMode = UCase$(Trim$(simulationMode))
    If GanttDrag_NormalizedSimulationMode = "" Then GanttDrag_NormalizedSimulationMode = "TEST"

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Supported Simulation Mode satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Supported Simulation Mode value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function GanttDrag_IsSupportedSimulationMode(ByVal simulationMode As String) As Boolean

    Select Case UCase$(Trim$(simulationMode))
        Case "TEST", "SCENARIO"
            GanttDrag_IsSupportedSimulationMode = True
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Orchestre Run Simulation Transaction By Mode en preservant l'ordre contractuel des etapes du domaine.
' EN: Orchestrates Run Simulation Transaction By Mode while preserving the domain's contractual step order.
'------------------------------------------------------------------------------

Private Function GanttDrag_RunSimulationTransactionByMode( _
    ByVal simulationMode As String, _
    ByRef consoleMessages As Collection, _
    ByRef ganttRebuilt As Boolean) As Boolean

    If gMetricsEnabled Then gMetricEngineCalls = gMetricEngineCalls + 1

    Select Case UCase$(Trim$(simulationMode))
        Case "TEST"
            GanttDrag_RunSimulationTransactionByMode = GanttLive_RunTestTransaction(consoleMessages, ganttRebuilt)
        Case "SCENARIO"
            GanttDrag_RunSimulationTransactionByMode = GanttLive_RunScenarioTransaction(consoleMessages, ganttRebuilt)
    End Select

    If gMetricsEnabled And ganttRebuilt Then
        gMetricRendererCalls = gMetricRendererCalls + 1
    End If

End Function

'------------------------------------------------------------------------------
' FR: Restaure la geometrie d'origine uniquement en rollback d'echec.
' EN: Restores original geometry only as a failure rollback.
'------------------------------------------------------------------------------
Private Sub GanttDrag_RollbackShapeGeometry( _
    ByVal shp As Shape, _
    ByVal oldState As Variant, _
    Optional ByVal reason As String = "")

    On Error GoTo SafeExit

    If shp Is Nothing Then Exit Sub
    If IsEmpty(oldState) Then Exit Sub

    shp.Left = CDbl(oldState(0))
    shp.Width = CDbl(oldState(1)) - CDbl(oldState(0))
    GanttDrag_SaveShapeState shp, CLng(oldState(2)), CStr(oldState(3))
    Profiler_RecordOperation "GanttDragFailureRollbackGeometry", 1, 0#

SafeExit:

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Set Drag Info Mode dans l'etat runtime du composant.
' EN: Activates or initializes Set Drag Info Mode in the component runtime state.
'------------------------------------------------------------------------------

Private Sub GanttDrag_SetDragInfoMode( _
    ByVal dragInfo As Object, _
    ByVal simulationMode As String)

    If dragInfo Is Nothing Then Exit Sub
    dragInfo("EngineMode") = UCase$(Trim$(simulationMode))

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la collection Unsupported Mode Message a la structure cible fournie par l'appelant.
' EN: Adds the Unsupported Mode Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttDrag_AddUnsupportedModeMessage( _
    ByVal consoleMessages As Collection, _
    ByVal dragInfo As Object, _
    ByVal simulationMode As String)

    Dim taskLabel As String
    Dim modeLabel As String

    If consoleMessages Is Nothing Then Exit Sub

    taskLabel = GanttDrag_InfoTaskLabel(dragInfo)
    modeLabel = Trim$(simulationMode)
    If modeLabel = "" Then modeLabel = "NONE"

    CalcBridge_AddConsoleMessage consoleMessages, "WARNING", BiMsg( _
        "Drag ignoré : aucun moteur de simulation actif compatible n'a été trouvé." & vbCrLf & _
        "Tâche : " & taskLabel & vbCrLf & _
        "Mode actif : " & modeLabel & vbCrLf & _
        "Activez TEST ou SCENARIO avant de déplacer une tâche.", _
        "Drag ignored: no compatible active simulation engine was found." & vbCrLf & _
        "Task: " & taskLabel & vbCrLf & _
        "Active mode: " & modeLabel & vbCrLf & _
        "Activate TEST or SCENARIO before moving a task.")

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la collection Drag Success Message a la structure cible fournie par l'appelant.
' EN: Adds the Drag Success Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttDrag_AddDragSuccessMessage( _
    ByVal consoleMessages As Collection, _
    ByVal dragInfo As Object)

    If consoleMessages Is Nothing Then Exit Sub
    CalcBridge_AddConsoleMessage consoleMessages, "INFO", GanttDrag_BuildDragMessage(dragInfo, True)

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la collection Drag Failure Message a la structure cible fournie par l'appelant.
' EN: Adds the Drag Failure Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub GanttDrag_AddDragFailureMessage( _
    ByRef consoleMessages As Collection, _
    ByVal dragInfo As Object)

    If consoleMessages Is Nothing Then Set consoleMessages = New Collection
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", GanttDrag_BuildDragMessage(dragInfo, False)

End Sub

'------------------------------------------------------------------------------
' FR: Construit la map Drag Message a partir des donnees fournies par l'appelant.
' EN: Builds the Drag Message map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function GanttDrag_BuildDragMessage( _
    ByVal dragInfo As Object, _
    ByVal success As Boolean) As String

    Dim taskLabel As String
    Dim changesFr As String
    Dim changesEn As String
    Dim frText As String
    Dim enText As String
    Dim simulationMode As String

    taskLabel = GanttDrag_InfoTaskLabel(dragInfo)
    changesFr = GanttDrag_InfoChangesText(dragInfo, True)
    changesEn = GanttDrag_InfoChangesText(dragInfo, False)
    If Not dragInfo Is Nothing Then
        If dragInfo.Exists("EngineMode") Then
            simulationMode = UCase$(Trim$(CStr(dragInfo("EngineMode"))))
        End If
    End If

    If success Then
        frText = "Modification " & GanttDrag_InfoEngineLabel(dragInfo, "Drag") & " appliquée." & vbCrLf & _
            "Tâche : " & taskLabel & vbCrLf & _
            "Modification demandée : " & changesFr & vbCrLf & _
            "Le planning a été recalculé. Les conséquences éventuelles sur les autres tâches proviennent du moteur planning." & vbCrLf & vbCrLf & _
            "Pour abandonner cette simulation et revenir au dernier planning calculé, cliquez sur Réinitialiser."

        enText = GanttDrag_InfoEngineLabel(dragInfo, "Drag") & " modification applied." & vbCrLf & _
            "Task: " & taskLabel & vbCrLf & _
            "Requested modification: " & changesEn & vbCrLf & _
            "The schedule has been recalculated. Any consequences on other tasks come from the planning engine." & vbCrLf & vbCrLf & _
            "To abandon this simulation and return to the last calculated schedule, click Reset."

        Select Case simulationMode
            Case "TEST"
                frText = frText & vbCrLf & _
                    "Pour retirer uniquement une hypothèse, videz la cellule TEST jaune correspondante puis relancez TEST."
                enText = enText & vbCrLf & _
                    "To remove only one assumption, clear the corresponding yellow TEST cell and run TEST again."
            Case "SCENARIO"
                frText = frText & vbCrLf & _
                    "Pour retirer uniquement une hypothèse, videz la cellule jaune correspondante puis cliquez sur Scénario."
                enText = enText & vbCrLf & _
                    "To remove only one assumption, clear the corresponding yellow cell and click Scenario."
        End Select
    Else
        frText = "La modification demandée par Drag n'a pas pu être appliquée." & vbCrLf & _
            "Tâche : " & taskLabel & vbCrLf & _
            "Modification demandée : " & changesFr

        enText = "The modification requested by Drag could not be applied." & vbCrLf & _
            "Task: " & taskLabel & vbCrLf & _
            "Requested modification: " & changesEn
    End If

    GanttDrag_BuildDragMessage = BiMsg(frText, enText)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Info Engine Label sans modifier les donnees d'entree.
' EN: Returns the Info Engine Label map without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_InfoEngineLabel( _
    ByVal dragInfo As Object, _
    ByVal defaultLabel As String) As String

    Dim modeLabel As String

    If dragInfo Is Nothing Then
        GanttDrag_InfoEngineLabel = defaultLabel
        Exit Function
    End If

    If dragInfo.Exists("EngineMode") Then modeLabel = UCase$(Trim$(CStr(dragInfo("EngineMode"))))

    Select Case modeLabel
        Case "TEST"
            GanttDrag_InfoEngineLabel = "Drag/Test"
        Case "SCENARIO"
            GanttDrag_InfoEngineLabel = "Drag/Scenario"
        Case Else
            GanttDrag_InfoEngineLabel = defaultLabel
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Info Task Label sans modifier les donnees d'entree.
' EN: Returns the Info Task Label map without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_InfoTaskLabel(ByVal dragInfo As Object) As String

    Dim taskName As String
    Dim wbsVal As String

    If dragInfo Is Nothing Then
        GanttDrag_InfoTaskLabel = "(unknown task)"
        Exit Function
    End If

    taskName = Trim$(CStr(dragInfo("TaskName")))
    wbsVal = Trim$(CStr(dragInfo("WBS")))

    If wbsVal <> "" And taskName <> "" Then
        GanttDrag_InfoTaskLabel = wbsVal & " - " & taskName
    ElseIf taskName <> "" Then
        GanttDrag_InfoTaskLabel = taskName
    ElseIf wbsVal <> "" Then
        GanttDrag_InfoTaskLabel = wbsVal
    Else
        GanttDrag_InfoTaskLabel = "(unknown task)"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Info Changes Text sans modifier les donnees d'entree.
' EN: Returns the Info Changes Text map without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_InfoChangesText( _
    ByVal dragInfo As Object, _
    ByVal french As Boolean) As String

    Dim changedStart As Boolean
    Dim changedFinish As Boolean
    Dim startText As String
    Dim finishText As String

    If dragInfo Is Nothing Then
        If french Then
            GanttDrag_InfoChangesText = "modification non identifiée"
        Else
            GanttDrag_InfoChangesText = "unidentified modification"
        End If
        Exit Function
    End If

    changedStart = CBool(dragInfo("ChangedStart"))
    changedFinish = CBool(dragInfo("ChangedFinish"))
    startText = GanttDrag_FormatDateValue(dragInfo("RequestedStart"))
    finishText = GanttDrag_FormatDateValue(dragInfo("RequestedFinish"))

    If changedStart And changedFinish Then
        If startText = finishText Then
            If french Then
                GanttDrag_InfoChangesText = "début et fin = " & startText
            Else
                GanttDrag_InfoChangesText = "start and finish = " & startText
            End If
        ElseIf french Then
            GanttDrag_InfoChangesText = "début = " & startText & ", fin = " & finishText
        Else
            GanttDrag_InfoChangesText = "start = " & startText & ", finish = " & finishText
        End If
    ElseIf changedStart Then
        If french Then
            GanttDrag_InfoChangesText = "début = " & startText
        Else
            GanttDrag_InfoChangesText = "start = " & startText
        End If
    ElseIf changedFinish Then
        If french Then
            GanttDrag_InfoChangesText = "fin = " & finishText
        Else
            GanttDrag_InfoChangesText = "finish = " & finishText
        End If
    ElseIf french Then
        GanttDrag_InfoChangesText = "aucune date modifiée"
    Else
        GanttDrag_InfoChangesText = "no date changed"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Format Date Value selon le contrat canonique du composant.
' EN: Normalizes or formats Format Date Value according to the component contract.
'------------------------------------------------------------------------------

Private Function GanttDrag_FormatDateValue(ByVal value As Variant) As String

    If IsDate(value) Then
        GanttDrag_FormatDateValue = Format$(CDate(value), "dd/mm/yyyy")
    ElseIf Trim$(CStr(value)) <> "" Then
        GanttDrag_FormatDateValue = Trim$(CStr(value))
    Else
        GanttDrag_FormatDateValue = "-"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Write Test Cell dans le stockage possede par le domaine.
' EN: Writes or synchronizes Write Test Cell in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub GanttDrag_WriteTestCell( _
    ByVal targetCell As Range, _
    ByVal testDate As Date, _
    ByVal writtenCells As Collection)

    Dim oldEnableEvents As Boolean
    Dim oldInternalWrite As Boolean

    If targetCell Is Nothing Then Exit Sub
    If writtenCells Is Nothing Then Exit Sub

    oldEnableEvents = Application.EnableEvents
    oldInternalWrite = GetGanttInternalWrite()

    On Error GoTo CleanExit
    Application.EnableEvents = False
    SetGanttInternalWrite True

    targetCell.value = DateValue(testDate)
    If gMetricsEnabled Then gMetricExcelWrites = gMetricExcelWrites + 1
    writtenCells.Add targetCell

CleanExit:
    SetGanttInternalWrite oldInternalWrite
    Application.EnableEvents = oldEnableEvents

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Clear Written Cells dans le perimetre possede par le composant.
' EN: Resets Clear Written Cells within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub GanttDrag_ClearWrittenCells(ByVal writtenCells As Collection)

    Dim oldEnableEvents As Boolean
    Dim oldInternalWrite As Boolean
    Dim item As Variant

    If writtenCells Is Nothing Then Exit Sub

    oldEnableEvents = Application.EnableEvents
    oldInternalWrite = GetGanttInternalWrite()

    On Error GoTo CleanExit
    Application.EnableEvents = False
    SetGanttInternalWrite True

    For Each item In writtenCells
        item.ClearContents
        If gMetricsEnabled Then gMetricExcelWrites = gMetricExcelWrites + 1
    Next item

CleanExit:
    SetGanttInternalWrite oldInternalWrite
    Application.EnableEvents = oldEnableEvents

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la collection Cell List sans modifier les donnees d'entree.
' EN: Returns the Cell List collection without mutating input data.
'------------------------------------------------------------------------------

Private Function GanttDrag_CellList(ByVal cells As Collection) As String

    Dim item As Variant
    Dim result As String

    If cells Is Nothing Then Exit Function

    For Each item In cells
        If result <> "" Then result = result & ","
        result = result & item.Address(False, False)
    Next item

    GanttDrag_CellList = result

End Function

'------------------------------------------------------------------------------
' FR: Indique si la reference Eligible Shape satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Eligible Shape reference satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function GanttDrag_IsEligibleShape( _
    ByVal ws As Worksheet, _
    ByVal shp As Shape, _
    ByRef rowIndex As Long, _
    ByRef taskType As String) As Boolean

    Dim tblWBS As ListObject
    Dim shapeName As String
    Dim suffix As String
    Dim dataRow As Long
    Dim normalizedTaskType As String
    Dim isMilestoneShape As Boolean

    If ws Is Nothing Then Exit Function
    If shp Is Nothing Then Exit Function
    If shp.Visible = msoFalse Then Exit Function

    shapeName = CStr(shp.Name)

    If Left$(shapeName, 5) = "TASK_" Then
        suffix = Mid$(shapeName, 6)
        If suffix = "" Or suffix Like "*[!0-9]*" Then Exit Function
    ElseIf Left$(shapeName, 3) = "MS_" Then
        suffix = Mid$(shapeName, 4)
        If suffix = "" Or suffix Like "*[!0-9]*" Then Exit Function
        isMilestoneShape = True
    Else
        Exit Function
    End If

    dataRow = CLng(Val(suffix))
    If dataRow < 1 Then Exit Function

    Set tblWBS = ThisWorkbook.Worksheets("WBS").ListObjects("tbl_WBS")
    If tblWBS.DataBodyRange Is Nothing Then Exit Function
    If dataRow > tblWBS.ListRows.Count Then Exit Function

    normalizedTaskType = UCase$(Trim$(CStr( _
        tblWBS.DataBodyRange.Cells(dataRow, tblWBS.ListColumns("Task Type").Index).value)))
    If gMetricsEnabled Then gMetricCellsRead = gMetricCellsRead + 1

    If Not isMilestoneShape Then
        If normalizedTaskType <> "TASK" Then Exit Function
        taskType = "TASK"
    Else
        'MS_n is created only by DrawMilestone. The rendered shape type is the
        'authoritative watcher contract, independent of localized/input text.
        taskType = "MILESTONE"
    End If

    rowIndex = GANTT_DRAG_FIRST_TASK_ROW + dataRow - 1
    GanttDrag_IsEligibleShape = True

End Function

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Save Shape State dans le stockage possede par le domaine.
' EN: Writes or synchronizes Save Shape State in the store owned by the domain.
'------------------------------------------------------------------------------

Private Sub GanttDrag_SaveShapeState( _
    ByVal shp As Shape, _
    ByVal rowIndex As Long, _
    ByVal taskType As String)

    Dim state(0 To 3) As Variant
    Dim shapeLeft As Double

    If shp Is Nothing Then Exit Sub
    If gShapeState Is Nothing Then Set gShapeState = CreateObject("Scripting.Dictionary")

    shapeLeft = CDbl(shp.Left)
    state(0) = shapeLeft
    state(1) = shapeLeft + CDbl(shp.Width)
    state(2) = CLng(rowIndex)
    state(3) = CStr(taskType)

    gShapeState(CStr(shp.Name)) = state

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Shape State sans exposer de mutateur sur l'etat source.
' EN: Returns the Shape State value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function GanttDrag_GetShapeState(ByVal shapeName As String) As Variant

    If gShapeState Is Nothing Then Exit Function
    If Not gShapeState.Exists(shapeName) Then Exit Function

    GanttDrag_GetShapeState = gShapeState(shapeName)

End Function

'------------------------------------------------------------------------------
' FR: Aligne la valeur Suspend avec le lifecycle courant sans perdre l'etat possede.
' EN: Aligns the Suspend value with the current lifecycle without losing owned state.
'------------------------------------------------------------------------------

Private Sub GanttDrag_Suspend()

    gSuspendDepth = gSuspendDepth + 1

End Sub

'------------------------------------------------------------------------------
' FR: Aligne la valeur Resume avec le lifecycle courant sans perdre l'etat possede.
' EN: Aligns the Resume value with the current lifecycle without losing owned state.
'------------------------------------------------------------------------------

Private Sub GanttDrag_Resume()

    If gSuspendDepth > 0 Then gSuspendDepth = gSuspendDepth - 1
    If gSuspendDepth = 0 And GanttDrag_IsWatching() Then GanttDrag_RebuildWatchMaps

End Sub

'------------------------------------------------------------------------------
' FR: Indique si la valeur Supported Timeline Scale satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Supported Timeline Scale value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function GanttDrag_IsSupportedTimelineScale() As Boolean

    Select Case UCase$(Trim$(GetGanttTimelineScaleMode()))
        Case GANTT_DRAG_SCALE_DAY, GANTT_DRAG_SCALE_WEEK, GANTT_DRAG_SCALE_MONTH
            GanttDrag_IsSupportedTimelineScale = True
    End Select

End Function
'------------------------------------------------------------------------------
' FR: Indique si la reference Start Watch satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Start Watch reference satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function GanttDrag_CanStartWatch() As Boolean

    Dim ws As Worksheet

    If Not GanttDrag_IsGanttSheetActive(ws) Then Exit Function
    If Not GanttDrag_IsSupportedTimelineScale() Then Exit Function
    If IsPlanningWorkflowActive() Then Exit Function
    If GetGanttInternalWrite() Then Exit Function
    If Application.CalculationState <> xlDone Then Exit Function

    GanttDrag_CanStartWatch = True

End Function

'------------------------------------------------------------------------------
' FR: Indique si la reference Gantt Sheet Active satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Gantt Sheet Active reference satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function GanttDrag_IsGanttSheetActive(ByRef ws As Worksheet) As Boolean

    On Error GoTo SafeExit

    If Application.ActiveWorkbook Is Nothing Then Exit Function
    If Not (Application.ActiveWorkbook Is ThisWorkbook) Then Exit Function
    If Application.ActiveSheet Is Nothing Then Exit Function
    If Not (Application.ActiveSheet.Parent Is ThisWorkbook) Then Exit Function
    If CStr(Application.ActiveSheet.Name) <> GANTT_DRAG_SHEET Then Exit Function

    Set ws = Application.ActiveSheet
    GanttDrag_IsGanttSheetActive = True

SafeExit:

End Function
