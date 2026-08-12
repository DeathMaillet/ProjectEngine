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
' CONTRATS / CONTRACTS : Refresh_Gantt, Refresh_Gantt_DisplayOnly, Gantt_TryApplyTestDayPredictiveRegistry
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
' - Workflow Gantt, boutons, refresh production, refresh TEST/SCENARIO via GanttLive.
'
' Notes:
' - Wrapper stable autour de RunGanttRefreshCore; ne contient pas de logique CALC.
'------------------------------------------------------------------------------
Public Sub Refresh_Gantt(Optional ByVal isNewSheet As Boolean = False, Optional ByVal activateGantt As Boolean = True)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("Refresh_Gantt", "Workflow")

    RunGanttRefreshCore False, isNewSheet, activateGantt

End Sub

'------------------------------------------------------------------------------
' FR: Met en forme ou met a jour un element UI/shape du GANTT.
' EN: Formats or updates a GANTT UI/shape element.
'------------------------------------------------------------------------------
Public Sub Refresh_Gantt_DisplayOnly()

    RunGanttRefreshCore True, False, True

End Sub
'=====================================================
' Predictive TEST Day registry - phase 1
' Scope: TODAY_LINE, TASK_*, TASK_*_P, MS_* only.
' Falls back to the classic Refresh_Gantt path whenever another visual family
' could become stale.
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
    Dim wsWBS As Worksheet
    Dim wsGantt As Worksheet
    Dim tblWBS As ListObject
    Dim dataArr As Variant
    Dim mapWBS As Object
    Dim hasChildren As Object
    Dim baseById As Object
    Dim testById As Object
    Dim projectStart As Variant
    Dim projectFinish As Variant
    Dim totalDays As Long
    Dim rowCount As Long
    Dim expected As Object
    Dim existing As Object
    Dim diffCount As Long
    Dim fallbackReason As String

    Set perfScope = Profiler_BeginScope("GanttPredictiveRegistry_TryApplyTestDay", "Gantt Registry")

    On Error GoTo Fallback

    Gantt_TryApplyTestDayPredictiveRegistry = False

    EnsureGanttViewInitialized

    If GetGanttTimelineScaleMode() <> GANTT_SCALE_DAY Then fallbackReason = "ScaleNotDay": GoTo Fallback
    If GetGanttViewMode() <> GANTT_VIEW_DETAIL Then fallbackReason = "ViewNotDetail": GoTo Fallback
    If Not GanttLive_IsTestRenderRequested() Then fallbackReason = "NoTestRenderRequest": GoTo Fallback

    Set wsGantt = ThisWorkbook.Worksheets(GANTT_SHEET)
    If wsGantt Is Nothing Then fallbackReason = "NoGanttSheet": GoTo Fallback
    If IsGanttSheetLayoutEmpty(wsGantt) Then fallbackReason = "EmptyLayout": GoTo Fallback

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    If tblWBS.DataBodyRange Is Nothing Then fallbackReason = "EmptyWBS": GoTo Fallback

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    ValidateGanttSourceColumns mapWBS

    dataArr = tblWBS.DataBodyRange.Value
    rowCount = UBound(dataArr, 1)
    If rowCount < 1 Then fallbackReason = "NoRows": GoTo Fallback
    If GetLastRenderedGanttRow(wsGantt) <> FIRST_TASK_ROW + rowCount - 1 Then fallbackReason = "RowCountMismatch": GoTo Fallback

    Set hasChildren = GanttHierarchy_BuildDirectParentPresenceFromWbs(dataArr, mapWBS)
    Set baseById = GanttLive_BuildBaseByIdMap()
    Set testById = GanttLive_BuildTestByIdMap()

    If GanttPredictive_HasUnsupportedSummaryDelta(dataArr, mapWBS, hasChildren, baseById, testById) Then fallbackReason = "SummaryDelta": GoTo Fallback

    ResolveDisplayedProjectRange dataArr, mapWBS, hasChildren, baseById, testById, True, projectStart, projectFinish
    If Not HasValue(projectStart) Or Not HasValue(projectFinish) Then fallbackReason = "NoProjectRange": GoTo Fallback

    totalDays = GetTimelineSlotCount(projectStart, projectFinish)
    If totalDays < 1 Then fallbackReason = "NoTimelineSlots": GoTo Fallback
    If Not GanttPredictive_CurrentDayTimelineMatches(wsGantt, projectStart, projectFinish, totalDays) Then fallbackReason = "TimelineMismatch": GoTo Fallback

    Set expected = GanttPredictive_BuildExpectedRegistry(wsGantt, dataArr, mapWBS, hasChildren, projectStart, totalDays, baseById, testById)
    Set existing = GanttPredictive_BuildExistingRegistry(wsGantt)

    diffCount = GanttPredictive_CountDiffs(expected, existing)
    Profiler_RecordOperation "GanttPredictiveRegistryDiffs", diffCount, 0#

    If diffCount > 0 Then
        If GanttPredictive_HasExistingPrefix(existing, "DEP_") Then fallbackReason = "DependencyShapesPresent": GoTo Fallback
        If GanttPredictive_HasExistingPrefix(existing, "CSTR_") Then fallbackReason = "ConstraintShapesPresent": GoTo Fallback
        If GetGanttShowConstraints() Then fallbackReason = "ConstraintsEnabled": GoTo Fallback
    End If

    GanttPredictive_ApplyDiff wsGantt, expected, existing
    ApplyGanttUiState wsGantt, False
    Gantt_CapturePendingGeometryRepair wsGantt

    If IsPlanningWorkflowActive() Then
        If GanttDrag_IsWatching() Then GanttDrag_RebuildWatchMaps
    Else
        On Error Resume Next
        GanttDrag_ReconcileWatchState
        On Error GoTo 0
    End If

    Gantt_TryApplyTestDayPredictiveRegistry = True
    Exit Function

Fallback:
    If fallbackReason = "" Then fallbackReason = "ErrorOrUnsupported"
    Profiler_RecordOperation "GanttPredictiveRegistryFallback_" & fallbackReason, 1, 0#
    Gantt_TryApplyTestDayPredictiveRegistry = False

End Function

'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Private Function GanttPredictive_CurrentDayTimelineMatches( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal projectFinish As Variant, _
    ByVal totalDays As Long) As Boolean

    Dim firstVal As Variant
    Dim lastVal As Variant
    Dim lastCol As Long

    On Error GoTo SafeExit

    lastCol = FIRST_TIMELINE_COL + totalDays - 1
    firstVal = ws.Cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Value
    lastVal = ws.Cells(HEADER_ROW_2, lastCol).Value

    If Not HasValue(firstVal) Then GoTo SafeExit
    If Not HasValue(lastVal) Then GoTo SafeExit

    GanttPredictive_CurrentDayTimelineMatches = _
        (CLng(CDbl(firstVal)) = CLng(CDbl(projectStart))) And _
        (CLng(CDbl(lastVal)) = CLng(CDbl(projectFinish)))

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Private Function GanttPredictive_HasUnsupportedSummaryDelta( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal baseById As Object, _
    ByVal testById As Object) As Boolean

    Dim r As Long
    Dim wbsVal As String
    Dim idVal As String

    For r = 1 To UBound(dataArr, 1)
        wbsVal = NormalizeWBS(CStr(dataArr(r, mapWBS("WBS"))))
        If hasChildren.Exists(wbsVal) Then
            idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))
            If GanttLive_HasRenderableTestDelta(idVal, baseById, testById) Then
                GanttPredictive_HasUnsupportedSummaryDelta = True
                Exit Function
            End If
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR:
' Reconstruit en memoire le registre attendu des shapes TEST Day sans toucher la feuille.
'
' EN:
' Rebuilds in memory the expected TEST Day shape registry without touching the sheet.
'
' Entrees / Inputs:
' - Feuille GANTT, data WBS, maps, timeline Day, maps base/test.
'
' Sorties / Outputs:
' - Dictionnaire de records attendus pour TASK/MS/TODAY.
'
' Appele par / Called by:
' - Gantt_TryApplyTestDayPredictiveRegistry.
'
' Notes:
' - Ignore les summaries: un delta summary force le fallback complet.
'------------------------------------------------------------------------------
Private Function GanttPredictive_BuildExpectedRegistry( _
    ByVal ws As Worksheet, _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object) As Object

    Dim perfScope As clsPerfScope
    Dim expected As Object
    Dim parentCompleteMap As Object
    Dim r As Long
    Dim rowCount As Long
    Dim ganttRow As Long
    Dim wbsVal As String
    Dim idVal As String
    Dim rawStartVal As Variant
    Dim rawFinishVal As Variant
    Dim renderStartVal As Variant
    Dim renderFinishVal As Variant
    Dim rawDurationDays As Long
    Dim progressVal As Double
    Dim isCritical As Boolean
    Dim hasDelta As Boolean
    Dim isParent As Boolean
    Dim isLoE As Boolean
    Dim isMilestone As Boolean

    Set perfScope = Profiler_BeginScope("GanttPredictiveRegistry_BuildExpected", "Gantt Registry")
    Set expected = CreateObject("Scripting.Dictionary")
    Set parentCompleteMap = GanttHierarchy_BuildLeafCompletionByAncestor(dataArr, mapWBS)

    rowCount = UBound(dataArr, 1)

    For r = 1 To rowCount
        ganttRow = FIRST_TASK_ROW + r - 1
        wbsVal = NormalizeWBS(CStr(dataArr(r, mapWBS("WBS"))))
        idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))

        rawStartVal = GanttLive_GetDisplayStart(idVal, baseById, testById, True)
        rawFinishVal = GanttLive_GetDisplayFinish(idVal, baseById, testById, True)
        renderStartVal = GetRenderStartForCurrentScale(rawStartVal)
        renderFinishVal = GetRenderFinishForCurrentScale(rawFinishVal)

        isParent = hasChildren.Exists(wbsVal)
        isMilestone = TaskTypeRules_IsMilestoneRow(dataArr, mapWBS, r)
        isLoE = TaskTypeRules_IsLevelOfEffortRow(dataArr, mapWBS, r)

        If isParent Then GoTo NextShape
        If Not ShouldRenderTaskInCurrentView(isParent, renderStartVal, renderFinishVal) Then GoTo NextShape
        If Not HasValue(rawStartVal) Or Not HasValue(rawFinishVal) Then GoTo NextShape
        If Not HasValue(renderStartVal) Or Not HasValue(renderFinishVal) Then GoTo NextShape

        rawDurationDays = CLng(CDbl(rawFinishVal) - CDbl(rawStartVal) + 1)
        progressVal = 0#

        If isLoE Then
            progressVal = GetLoEProgressFromToday(rawStartVal, rawFinishVal)
        ElseIf HasValue(GanttLive_GetDisplayProgress(idVal, baseById, testById, True)) Then
            progressVal = CDbl(GanttLive_GetDisplayProgress(idVal, baseById, testById, True))
        End If

        isCritical = ShouldHighlightGanttAnalyticsPath(dataArr, mapWBS, r, idVal, testById, True)
        hasDelta = GanttLive_HasRenderableTestDelta(idVal, baseById, testById)

        If isMilestone Then
            GanttShapeRegistry_AddMilestoneRecord expected, ws, ganttRow, projectStart, rawStartVal, progressVal, isCritical, "MS_" & CStr(r), hasDelta
        ElseIf ShouldDrawCompactTaskMarker(ws, ganttRow, projectStart, rawStartVal, rawFinishVal, rawDurationDays, isLoE) Then
            GanttShapeRegistry_AddCompactTaskRecords expected, ws, ganttRow, projectStart, rawStartVal, rawFinishVal, progressVal, isCritical, "TASK_" & CStr(r), hasDelta
        Else
            GanttShapeRegistry_AddTaskBarRecords expected, ws, ganttRow, projectStart, rawStartVal, rawFinishVal, progressVal, isCritical, "TASK_" & CStr(r), hasDelta, isLoE
        End If

NextShape:
    Next r

    GanttShapeRegistry_AddTodayLineRecord expected, ws, projectStart, totalDays, rowCount
    Set GanttPredictive_BuildExpectedRegistry = expected

End Function

'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Private Function GanttPredictive_BuildExistingRegistry(ByVal ws As Worksheet) As Object

    Dim perfScope As clsPerfScope
    Dim d As Object
    Dim shp As Shape

    Set perfScope = Profiler_BeginScope("GanttPredictiveRegistry_BuildExisting", "Gantt Registry")
    Set d = CreateObject("Scripting.Dictionary")

    For Each shp In ws.Shapes
        If IsGanttGeometryRepairShape(shp.Name) Then d.Add shp.Name, shp
    Next shp

    Set GanttPredictive_BuildExistingRegistry = d

End Function
'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Private Function GanttPredictive_CountDiffs(ByVal expected As Object, ByVal existing As Object) As Long

    Dim key As Variant
    Dim shp As Shape
    Dim rec As Object
    Dim countVal As Long

    For Each key In expected.Keys
        Set rec = expected(CStr(key))
        If Not existing.Exists(CStr(key)) Then
            countVal = countVal + 1
        Else
            Set shp = existing(CStr(key))
            If GanttShapeRegistry_ShapeDiffers(shp, rec) Then countVal = countVal + 1
        End If
    Next key

    For Each key In existing.Keys
        If GanttPredictive_IsScopedShapeName(CStr(key)) Then
            If Not expected.Exists(CStr(key)) Then countVal = countVal + 1
        End If
    Next key

    GanttPredictive_CountDiffs = countVal

End Function

'------------------------------------------------------------------------------
' FR: Actualise Gantt Predictive Apply Diff sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Gantt Predictive Apply Diff without changing the business rules that produce the data.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub GanttPredictive_ApplyDiff(ByVal ws As Worksheet, ByVal expected As Object, ByVal existing As Object)

    Dim perfScope As clsPerfScope
    Dim key As Variant
    Dim rec As Object
    Dim shp As Shape
    Dim createdCount As Long
    Dim updatedCount As Long
    Dim deletedCount As Long

    Set perfScope = Profiler_BeginScope("GanttPredictiveRegistry_ApplyDiff", "Gantt Registry")

    For Each key In expected.Keys
        Set rec = expected(CStr(key))
        If existing.Exists(CStr(key)) Then
            Set shp = existing(CStr(key))
            If GanttShapeRegistry_ShapeDiffers(shp, rec) Then
                If GanttShapeRegistry_TypeMismatch(shp, rec) Then
                    shp.Delete
                    GanttShapeRegistry_CreateShapeFromRecord ws, rec
                    createdCount = createdCount + 1
                Else
                    GanttShapeRegistry_UpdateShapeFromRecord shp, rec
                    updatedCount = updatedCount + 1
                End If
            End If
        Else
            GanttShapeRegistry_CreateShapeFromRecord ws, rec
            createdCount = createdCount + 1
        End If
    Next key

    For Each key In existing.Keys
        If GanttPredictive_IsScopedShapeName(CStr(key)) Then
            If Not expected.Exists(CStr(key)) Then
                Set shp = existing(CStr(key))
                shp.Delete
                deletedCount = deletedCount + 1
            End If
        End If
    Next key

    Profiler_RecordOperation "GanttPredictiveRegistryCreates", createdCount, 0#
    Profiler_RecordOperation "GanttPredictiveRegistryUpdates", updatedCount, 0#
    Profiler_RecordOperation "GanttPredictiveRegistryDeletes", deletedCount, 0#

End Sub
'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Private Function GanttPredictive_IsScopedShapeName(ByVal shapeName As String) As Boolean

    GanttPredictive_IsScopedShapeName = _
        (shapeName = "TODAY_LINE") Or _
        (Left$(shapeName, 5) = "TASK_") Or _
        (Left$(shapeName, 3) = "MS_")

End Function

'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Private Function GanttPredictive_HasExistingPrefix(ByVal existing As Object, ByVal prefixText As String) As Boolean

    Dim key As Variant

    For Each key In existing.Keys
        If Left$(CStr(key), Len(prefixText)) = prefixText Then
            GanttPredictive_HasExistingPrefix = True
            Exit Function
        End If
    Next key

End Function
