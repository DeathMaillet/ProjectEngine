Attribute VB_Name = "mod_Gantt"
Option Explicit

'===============================================================================
' MODULE : mod_Gantt
' DOMAINE / DOMAIN : Gantt
'
' FR
' Expose les wrappers historiques du Gantt et orchestre les appels haut niveau restes apres decomposition.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Exposes historical Gantt wrappers and orchestrates the remaining high-level calls after decomposition.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : CommitGanttUpdate compatibility wrappers, Gantt_TryApplyTestDayPredictiveRegistry
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"

Private Const FIRST_TIMELINE_COL As Long = 11   ' K
Private Const TITLE_ROW As Long = 1
Private Const TOGGLE_ROW_TOP As Long = 2
Private Const TOGGLE_ROW_BOTTOM As Long = 3
Private Const HEADER_ROW_1 As Long = 3
Private Const HEADER_ROW_2 As Long = 4
Private Const FIRST_TASK_ROW As Long = 5

Private Const COL_WBS As Long = 1
Private Const COL_TASK As Long = 2
Private Const COL_START As Long = 3
Private Const COL_FINISH As Long = 4
Private Const COL_TEST_START As Long = 5
Private Const COL_TEST_FINISH As Long = 6
Private Const COL_DURATION As Long = 7
Private Const COL_PROGRESS As Long = 8
Private Const COL_TEST_PROGRESS As Long = 9
Private Const COL_LOGIC As Long = 10

Private Const BTN_VIEW_BG_NAME As String = "btn_Gantt_View_BG"
Private Const BTN_VIEW_KNOB_NAME As String = "btn_Gantt_View_Knob"
Private Const BTN_VIEW_LEFT_NAME As String = "btn_Gantt_View_Left"
Private Const BTN_VIEW_RIGHT_NAME As String = "btn_Gantt_View_Right"
Private Const BTN_VIEW_CELL_LEFT As String = "A2"
Private Const BTN_VIEW_CELL_RIGHT As String = "B2"

Private Const BTN_CP_BG_NAME As String = "btn_Gantt_CP_BG"
Private Const BTN_CP_KNOB_NAME As String = "btn_Gantt_CP_Knob"
Private Const BTN_CP_LEFT_NAME As String = "btn_Gantt_CP_Left"
Private Const BTN_CP_RIGHT_NAME As String = "btn_Gantt_CP_Right"
Private Const BTN_CP_CELL_LEFT As String = "A3"
Private Const BTN_CP_CELL_RIGHT As String = "B3"

Private Const GANTT_SCALE_DAY As String = "DAY"
Private Const GANTT_SCALE_WEEK As String = "WEEK"
Private Const GANTT_SCALE_MONTH As String = "MONTH"

Private Const BTN_CP_MULTI_BG_NAME As String = "shp_GANTT_CPMode_BG"
Private Const BTN_CP_MULTI_KNOB_NAME As String = "shp_GANTT_CPMode_Knob"
Private Const BTN_CP_MULTI_LEFT_NAME As String = "lbl_GANTT_CPMode_Toggle"

Private Const BTN_CONSTRAINT_BG_NAME As String = "btn_Gantt_Constraint_BG"
Private Const BTN_CONSTRAINT_KNOB_NAME As String = "btn_Gantt_Constraint_Knob"
Private Const BTN_CONSTRAINT_LEFT_NAME As String = "btn_Gantt_Constraint_Left"

Private Const GANTT_ROW_HEIGHT_HEADER_1 As Double = 21
Private Const GANTT_ROW_HEIGHT_HEADER_2 As Double = 18
Private Const GANTT_ROW_HEIGHT_TASK As Double = 18

Private Const BTN_SCALE_BG_NAME As String = "btn_Gantt_Scale_BG"
Private Const BTN_SCALE_KNOB_NAME As String = "btn_Gantt_Scale_Knob"
Private Const BTN_SCALE_LEFT_NAME As String = "btn_Gantt_Scale_Left"
Private Const BTN_SCALE_RIGHT_NAME As String = "btn_Gantt_Scale_Right"
Private Const BTN_SCALE_CELL_LEFT As String = "D2"
Private Const BTN_SCALE_CELL_RIGHT As String = "E2"

Private Const BTN_SCENARIO_NAME As String = "btn_Gantt_Scenario"
Private Const BTN_SCENARIO_CAPTION As String = "Scenario"

Private Const COLOR_TASK_BLUE As Long = 12874308
Private Const COLOR_TASK_CRITICAL As Long = 192

Private Const COLOR_PROGRESS_GREEN As Long = 4699504
Private Const COLOR_PROGRESS_ORANGE As Long = 3248093

Private Const LINK_STUB As Double = 6
Private Const LINK_MIN_CHANNEL_GAP As Double = 4
Private Const LINK_EDGE_PADDING As Double = 8

Private Const TEST_START_HEADER As String = "Test Start"
Private Const TEST_FINISH_HEADER As String = "Test Finish"
Private Const TEST_PROGRESS_HEADER As String = "Test %"

Private Const BTN_TEST_NAME As String = "btn_Gantt_Test"
Private Const BTN_LOCK_NAME As String = "btn_Gantt_Lock"

Private Const BTN_TEST_CAPTION As String = "Test"
Private Const BTN_LOCK_CAPTION As String = "Lock"

Private Const GANTT_VIEW_DETAIL As String = "DETAIL"
Private Const GANTT_VIEW_SUMMARY As String = "SUMMARY"

Private Const GANTT_ANALYTICS_PATH_NONE As String = "NONE"
Private Const GANTT_ANALYTICS_PATH_CP As String = "CP"
Private Const GANTT_ANALYTICS_PATH_LP As String = "LP"

Private Const LINK_ANCHOR_START As String = "START"
Private Const LINK_ANCHOR_FINISH As String = "FINISH"

Private Const GANTT_COLD_STATE_PENDING As String = "PENDING"
Private Const GANTT_COLD_STATE_READY As String = "READY"
Private Const GANTT_COLD_STATE_UNKNOWN As String = "UNKNOWN"

Public Const GANTT_ENSURE_NO_RENDER As String = "NO_RENDER"
Public Const GANTT_ENSURE_RENDER_OFFSCREEN As String = "RENDER_OFFSCREEN"
Public Const GANTT_ENSURE_RENDER_AND_SHOW As String = "RENDER_AND_SHOW"
Public Const GANTT_ENSURE_ON_ACTIVATION As String = "ENSURE_ON_ACTIVATION"
Public Const GANTT_ENSURE_LOCAL_UPDATE As String = "LOCAL_UPDATE"

Private gGanttColdState As String
Private gExplicitDeferredGanttRenderPending As Boolean
Private gExplicitDeferredGanttRenderReason As String
Private gOpenLifecycleTrace As Collection
Private gOpenLifecycleStart As Double
Private gOpenLifecycleSeq As Long

Public Sub GanttOpenLifecycle_BeginTrace(ByVal scenarioName As String)

    Set gOpenLifecycleTrace = New Collection
    gOpenLifecycleSeq = 0
    gOpenLifecycleStart = Timer
    GanttOpenLifecycle_Record "Begin", scenarioName, GanttOpenLifecycle_ActiveSheetName(), GanttOpenLifecycle_ActiveSheetName(), 0#

End Sub

Public Sub GanttOpenLifecycle_Record( _
    ByVal functionName As String, _
    Optional ByVal reason As String = "", _
    Optional ByVal sheetBefore As String = "", _
    Optional ByVal sheetAfter As String = "", _
    Optional ByVal elapsedMs As Double = 0#)

    Dim rowData(1 To 7) As Variant

    If gOpenLifecycleTrace Is Nothing Then Set gOpenLifecycleTrace = New Collection
    gOpenLifecycleSeq = gOpenLifecycleSeq + 1
    rowData(1) = gOpenLifecycleSeq
    rowData(2) = functionName
    rowData(3) = reason
    rowData(4) = sheetBefore
    rowData(5) = sheetAfter
    rowData(6) = elapsedMs
    rowData(7) = GanttOpenLifecycle_ElapsedMs()
    gOpenLifecycleTrace.Add rowData

End Sub

Public Function GanttOpenLifecycle_ExportTrace(ByVal outputPath As String) As Boolean

    Dim fso As Object
    Dim stream As Object
    Dim rowData As Variant

    On Error GoTo Failed
    If gOpenLifecycleTrace Is Nothing Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(outputPath, True, False)
    stream.WriteLine "Seq" & vbTab & "Function" & vbTab & "Reason" & vbTab & _
        "SheetBefore" & vbTab & "SheetAfter" & vbTab & "ElapsedMs" & vbTab & "SinceTraceStartMs"
    For Each rowData In gOpenLifecycleTrace
        stream.WriteLine CStr(rowData(1)) & vbTab & CStr(rowData(2)) & vbTab & _
            CStr(rowData(3)) & vbTab & CStr(rowData(4)) & vbTab & _
            CStr(rowData(5)) & vbTab & CStr(rowData(6)) & vbTab & CStr(rowData(7))
    Next rowData
    stream.Close
    GanttOpenLifecycle_ExportTrace = True
    Exit Function

Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0

End Function

Public Sub GanttColdState_MarkPending(Optional ByVal reason As String = "")

    gGanttColdState = GANTT_COLD_STATE_PENDING
    GanttOpenLifecycle_Record "GanttColdState_MarkPending", reason, _
        GanttOpenLifecycle_ActiveSheetName(), GanttOpenLifecycle_ActiveSheetName(), 0#

End Sub

Public Sub GanttDeferredRender_MarkPending(Optional ByVal reason As String = "")

    gExplicitDeferredGanttRenderPending = True
    gExplicitDeferredGanttRenderReason = reason
    GanttColdState_MarkPending "ExplicitDeferredRender|" & reason
    GanttOpenLifecycle_Record "GanttDeferredRender_MarkPending", reason, _
        GanttOpenLifecycle_ActiveSheetName(), GanttOpenLifecycle_ActiveSheetName(), 0#

End Sub

Public Sub GanttDeferredRender_Clear(Optional ByVal reason As String = "")

    gExplicitDeferredGanttRenderPending = False
    gExplicitDeferredGanttRenderReason = vbNullString
    GanttOpenLifecycle_Record "GanttDeferredRender_Clear", reason, _
        GanttOpenLifecycle_ActiveSheetName(), GanttOpenLifecycle_ActiveSheetName(), 0#

End Sub

Public Function GanttDeferredRender_IsPending() As Boolean

    GanttDeferredRender_IsPending = gExplicitDeferredGanttRenderPending

End Function

Public Function GanttDeferredRender_FinishIfPending(Optional ByVal reason As String = "") As Boolean

    Dim pendingReason As String
    Dim startedAt As Double

    If Not gExplicitDeferredGanttRenderPending Then Exit Function

    pendingReason = gExplicitDeferredGanttRenderReason
    startedAt = Timer
    GanttOpenLifecycle_Record "GanttDeferredRender_FinishIfPending", _
        "Start|" & reason & "|" & pendingReason, _
        GanttOpenLifecycle_ActiveSheetName(), GanttOpenLifecycle_ActiveSheetName(), 0#

    If EnsureGanttForCurrentPlanning( _
        GANTT_ENSURE_RENDER_OFFSCREEN, _
        "ExplicitDeferredActivation|" & reason & "|" & pendingReason) Then
        GanttDeferredRender_Clear "Completed|" & reason & "|" & pendingReason
        GanttDeferredRender_FinishIfPending = True
    End If

    GanttOpenLifecycle_Record "GanttDeferredRender_FinishIfPending", _
        "End|" & reason & "|Completed=" & CStr(GanttDeferredRender_FinishIfPending), _
        GanttOpenLifecycle_ActiveSheetName(), GanttOpenLifecycle_ActiveSheetName(), _
        GanttOpenLifecycle_MsSince(startedAt)

End Function

Public Function GanttColdState_Get() As String

    If Len(gGanttColdState) = 0 Then
        GanttColdState_Get = GANTT_COLD_STATE_UNKNOWN
    Else
        GanttColdState_Get = gGanttColdState
    End If

End Function

Public Sub GanttColdState_EnsureReadyIfGanttActive(Optional ByVal reason As String = "")

    Dim beforeSheet As String
    Dim startedAt As Double

    beforeSheet = GanttOpenLifecycle_ActiveSheetName()
    startedAt = Timer

    If UCase$(beforeSheet) <> GANTT_SHEET Then
        If Len(gGanttColdState) = 0 Then gGanttColdState = GANTT_COLD_STATE_PENDING
        GanttOpenLifecycle_Record "GanttColdState_EnsureReadyIfGanttActive", _
            "SkippedNonGantt|" & reason, beforeSheet, GanttOpenLifecycle_ActiveSheetName(), _
            GanttOpenLifecycle_MsSince(startedAt)
        Exit Sub
    End If

    EnsureGanttForCurrentPlanning GANTT_ENSURE_ON_ACTIVATION, reason
    GanttOpenLifecycle_Record "GanttColdState_EnsureReadyIfGanttActive", _
        GanttColdState_Get() & "|" & reason, beforeSheet, GanttOpenLifecycle_ActiveSheetName(), _
        GanttOpenLifecycle_MsSince(startedAt)

End Sub

Public Function EnsureGanttForCurrentPlanning( _
    ByVal mode As String, _
    Optional ByVal reason As String = "") As Boolean

    Dim normalizedMode As String
    Dim ws As Worksheet
    Dim beforeSheet As String
    Dim afterSheet As String
    Dim startedAt As Double
    Dim ready As Boolean

    startedAt = Timer
    beforeSheet = GanttOpenLifecycle_ActiveSheetName()
    normalizedMode = UCase$(Trim$(mode))

    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)

    Select Case normalizedMode
        Case GANTT_ENSURE_NO_RENDER
            ready = CommitGanttUpdate( _
                GanttEngine_ActiveDataSource(), GANTT_UPDATE_SCOPE_DEFER, _
                GANTT_RENDER_INTENT_DEFER, "EnsureNoRender|" & reason)
            EnsureGanttForCurrentPlanning = ready

        Case GANTT_ENSURE_RENDER_OFFSCREEN
            ready = CommitGanttUpdate( _
                GanttEngine_ActiveDataSource(), GANTT_UPDATE_SCOPE_PARTIAL, _
                GANTT_RENDER_INTENT_OFFSCREEN, "EnsureRenderOffscreen|" & reason)
            EnsureGanttForCurrentPlanning = ready

        Case GANTT_ENSURE_RENDER_AND_SHOW
            ready = CommitGanttUpdate( _
                GanttEngine_ActiveDataSource(), GANTT_UPDATE_SCOPE_PARTIAL, _
                GANTT_RENDER_INTENT_SHOW, "EnsureRenderAndShow|" & reason)
            EnsureGanttForCurrentPlanning = ready

        Case GANTT_ENSURE_ON_ACTIVATION
            Gantt_RepairColdStartSvgIfNeeded
            ready = Gantt_IsReadyStateValid()
            EnsureGanttForCurrentPlanning = True

        Case GANTT_ENSURE_LOCAL_UPDATE
            ready = CommitGanttUpdate( _
                GanttEngine_ActiveDataSource(), GANTT_UPDATE_SCOPE_PARTIAL, _
                GANTT_RENDER_INTENT_SHOW, "EnsureLocalUpdate|" & reason)
            EnsureGanttForCurrentPlanning = ready

        Case Else
            Err.Raise 5, "EnsureGanttForCurrentPlanning", "Unknown Gantt ensure mode: " & mode
    End Select

TraceExit:
    afterSheet = GanttOpenLifecycle_ActiveSheetName()
    GanttOpenLifecycle_Record "EnsureGanttForCurrentPlanning", _
        normalizedMode & "|" & reason & "|Ready=" & CStr(ready), _
        beforeSheet, afterSheet, GanttOpenLifecycle_MsSince(startedAt)
    Exit Function

Failed:
    gGanttColdState = GANTT_COLD_STATE_PENDING
    afterSheet = GanttOpenLifecycle_ActiveSheetName()
    GanttOpenLifecycle_Record "EnsureGanttForCurrentPlanning", _
        normalizedMode & "|" & reason & "|FAILED|Err=" & CStr(Err.Number), _
        beforeSheet, afterSheet, GanttOpenLifecycle_MsSince(startedAt)
    EnsureGanttForCurrentPlanning = False

End Function

Public Function Gantt_FinalizeReadyState(Optional ByVal reason As String = "") As Boolean

    Dim ws As Worksheet
    Dim startedAt As Double
    Dim beforeSheet As String
    Dim afterSheet As String
    Dim diag As Variant

    startedAt = Timer
    beforeSheet = GanttOpenLifecycle_ActiveSheetName()

    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)
    If ws Is Nothing Then GoTo Failed

    diag = GanttDependencySvg_GetPersistentCacheDiagnostics(ws)
    If Not Gantt_DependencyDiagnosticsAreReady(diag) Then GoTo Failed

    gGanttColdState = GANTT_COLD_STATE_READY
    afterSheet = GanttOpenLifecycle_ActiveSheetName()
    GanttOpenLifecycle_Record "Gantt_FinalizeReadyState", _
        "READY|" & reason & "|Routes=" & CStr(diag(1, 5)) & "|Segments=" & CStr(diag(1, 6)), _
        beforeSheet, afterSheet, GanttOpenLifecycle_MsSince(startedAt)
    Gantt_FinalizeReadyState = True
    Exit Function

Failed:
    gGanttColdState = GANTT_COLD_STATE_PENDING
    afterSheet = GanttOpenLifecycle_ActiveSheetName()
    GanttOpenLifecycle_Record "Gantt_FinalizeReadyState", _
        "PENDING|" & reason & "|Err=" & CStr(Err.Number), _
        beforeSheet, afterSheet, GanttOpenLifecycle_MsSince(startedAt)

End Function

Public Function Gantt_IsReadyStateValid() As Boolean

    Dim ws As Worksheet
    Dim diag As Variant

    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)
    diag = GanttDependencySvg_GetPersistentCacheDiagnostics(ws)

    Gantt_IsReadyStateValid = Gantt_DependencyDiagnosticsAreReady(diag)
    Exit Function

Failed:
    Gantt_IsReadyStateValid = False

End Function

Private Function Gantt_DependencyDiagnosticsAreReady(ByVal diag As Variant) As Boolean

    If CBool(diag(1, 7)) Then Exit Function
    If Not CBool(diag(1, 10)) Then Exit Function

    If IsAggregatedScaleMode() Then
        Gantt_DependencyDiagnosticsAreReady = True
        Exit Function
    End If

    Gantt_DependencyDiagnosticsAreReady = _
        CBool(diag(1, 2)) And _
        CBool(diag(1, 3)) And _
        CBool(diag(1, 4)) And _
        (CLng(diag(1, 5)) > 0) And _
        CBool(diag(1, 11))

End Function

Private Function GanttOpenLifecycle_ActiveSheetName() As String

    On Error Resume Next
    If Not Application.ActiveSheet Is Nothing Then
        If Application.ActiveSheet.Parent Is ThisWorkbook Then
            GanttOpenLifecycle_ActiveSheetName = CStr(Application.ActiveSheet.Name)
        End If
    End If
    On Error GoTo 0

End Function

Private Function GanttOpenLifecycle_ElapsedMs() As Double
    GanttOpenLifecycle_ElapsedMs = GanttOpenLifecycle_MsSince(gOpenLifecycleStart)
End Function

Private Function GanttOpenLifecycle_MsSince(ByVal startedAt As Double) As Double

    Dim nowValue As Double

    nowValue = Timer
    If nowValue < startedAt Then nowValue = nowValue + 86400#
    GanttOpenLifecycle_MsSince = (nowValue - startedAt) * 1000#

End Function
'------------------------------------------------------------------------------
' FR:
' Point d'entree public du redraw GANTT complet: lance le workflow de rendu
' depuis l'etat courant WBS/CALC/GanttLive.
'
' EN:
' Public entry point for a full GANTT redraw: starts the rendering workflow
' from the current WBS/CALC/GanttLive state.
'
' Entrees / Inputs:
' - isNewSheet indique si le layout vient d'etre cree.
' - activateGantt controle la restauration du contexte visuel appelant.
'
' Sorties / Outputs:
' - Feuille GANTT reconstruite, timeline/shapes/liens/markers/UI mis a jour.
'
' Appele par / Called by:
' - Outils/harness historiques; le runtime produit doit appeler CommitGanttUpdate.
'
' Notes:
' - Wrapper de compatibilite strict: ne rend pas directement et passe par CommitGanttUpdate.
'------------------------------------------------------------------------------
Public Sub Refresh_Gantt(Optional ByVal isNewSheet As Boolean = False, Optional ByVal activateGantt As Boolean = True)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("Refresh_Gantt", "Workflow")

    CommitGanttUpdate _
        GanttEngine_ActiveDataSource(), _
        GANTT_UPDATE_SCOPE_FULL, _
        GANTT_RENDER_INTENT_SHOW, _
        "Refresh_Gantt_CompatibilityWrapper"

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Public Sub Refresh_Gantt_DisplayOnly()

    CommitGanttUpdate _
        GanttEngine_ActiveDataSource(), _
        GANTT_UPDATE_SCOPE_PARTIAL, _
        GANTT_RENDER_INTENT_OFFSCREEN, _
        "Refresh_Gantt_DisplayOnly_CompatibilityWrapper"

End Sub

'------------------------------------------------------------------------------
' FR: Reprojette la timeline et la geometrie sans reconstruire le moteur.
' EN: Reprojects timeline and geometry without rebuilding the engine.
'------------------------------------------------------------------------------
Public Sub Refresh_Gantt_TimelineOnly()

    CommitGanttUpdate _
        GanttEngine_ActiveDataSource(), _
        GANTT_UPDATE_SCOPE_FULL, _
        GANTT_RENDER_INTENT_SHOW, _
        "Refresh_Gantt_TimelineOnly_CompatibilityWrapper"

End Sub

'------------------------------------------------------------------------------
' FR: Repare une couche SVG persistante dont les routes memoire ne sont pas prouvees.
' EN: Repairs a persisted SVG layer whose in-memory routes are not proven.
'------------------------------------------------------------------------------
Public Sub Gantt_RepairColdStartSvgIfNeeded()

    Dim ws As Worksheet

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)
    If ws Is Nothing Then Exit Sub

    If GanttDependencySvg_HasLayer(ws) And Not GanttDependencySvg_HasRoutes() Then
        If GanttDependencySvg_TryHydratePersistentCache(ws) Then Exit Sub
        GanttColdState_MarkPending "ColdStartSvgLayerWithoutRoutes"
    End If

SafeExit:
    On Error GoTo 0

End Sub
'=====================================================
' Predictive TEST Day registry - phase 1
' Scope: TODAY_LINE, TASK_*, TASK_*_P, MS_* only.
' Delegates to the common engine whenever another visual family could become stale.
'=====================================================
'------------------------------------------------------------------------------
' FR:
' Tente un refresh TEST rapide en mode Day/Detail en calculant le registre attendu
' et en appliquant uniquement les differences de shapes.
'
' EN:
' Attempts a fast TEST refresh in Day/Detail mode by computing the expected
' registry and applying only shape differences.
'
' Entrees / Inputs:
' - GANTT existant, tbl_WBS, maps base/test, demande de rendu TEST one-shot.
'
' Sorties / Outputs:
' - Shapes TASK/MS/TODAY mises a jour sans redraw complet, ou False pour fallback.
'
' Appele par / Called by:
' - GanttLive_ApplyTestRender.
'
' Notes:
' - Fallback obligatoire si timeline, summaries, liens, contraintes ou scale rendent le diff dangereux.
'------------------------------------------------------------------------------
Public Function Gantt_TryApplyTestDayPredictiveRegistry() As Boolean

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("GanttPredictiveRegistry_TryApplyTestDay", "Gantt Registry")

    On Error GoTo Failed

    'The former implementation built a global expected registry and only then
    'rejected dependency shapes. The differential refresh now decides and
    'applies the local transaction before any global shape scan.
    If Not CommitGanttUpdate( _
        GanttEngine_ActiveDataSource(), _
        GANTT_UPDATE_SCOPE_PARTIAL, _
        GANTT_RENDER_INTENT_SHOW, _
        "GanttPredictiveRegistry_Compatibility") Then GoTo Failed
    Profiler_RecordOperation "GanttPredictiveDelegatedToLocalTransaction", 1, 0#
    Gantt_TryApplyTestDayPredictiveRegistry = True
    Exit Function

Failed:
    Profiler_RecordOperation "GanttPredictiveRegistryFallback_LocalError", 1, 0#
    Gantt_TryApplyTestDayPredictiveRegistry = False

End Function

