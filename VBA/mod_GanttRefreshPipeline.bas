Attribute VB_Name = "mod_GanttRefreshPipeline"
Option Explicit

'===============================================================================
' MODULE : mod_GanttRefreshPipeline
' DOMAINE / DOMAIN : Gantt
'
' FR
' Orchestre les etapes du pipeline dans leur ordre contractuel.
' Delegue les calculs et rendus a leurs proprietaires.
'
' EN
' Orchestrates pipeline stages in their contractual order.
' Delegates calculations and rendering to their owners.
'
' CONTRATS / CONTRACTS : RunGanttRefreshCore
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"

Private Const FIRST_TIMELINE_COL As Long = 11
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

Private Const TEST_START_HEADER As String = "Test Start"
Private Const TEST_FINISH_HEADER As String = "Test Finish"
Private Const TEST_PROGRESS_HEADER As String = "Test %"

Private Const GANTT_ROW_HEIGHT_HEADER_1 As Double = 21
Private Const GANTT_ROW_HEIGHT_HEADER_2 As Double = 18
Private Const GANTT_ROW_HEIGHT_TASK As Double = 18

Private Const GANTT_VIEW_DETAIL As String = "DETAIL"
Private Const GANTT_VIEW_SUMMARY As String = "SUMMARY"
Private Const GANTT_SCALE_DAY As String = "DAY"
Private Const GANTT_SCALE_WEEK As String = "WEEK"
Private Const GANTT_SCALE_MONTH As String = "MONTH"

Private gLastRenderSignature As String
Private gLastTimelineSignature As String
Private gLastWorksheetShapeCount As Long
Private gLastRefreshSucceeded As Boolean
Private gLastRefreshEffectiveScope As String

Public Function GanttRefresh_LastRunSucceeded() As Boolean

    GanttRefresh_LastRunSucceeded = gLastRefreshSucceeded

End Function

Public Function GanttRefresh_LastEffectiveScope() As String

    GanttRefresh_LastEffectiveScope = gLastRefreshEffectiveScope

End Function


'------------------------------------------------------------------------------
' FR:
' Orchestre le rendu GANTT: charge WBS/CALC, choisit base ou simulation,
' prepare le layout, dessine les shapes, liens, today line, contraintes et etat UI.
'
' EN:
' Orchestrates GANTT rendering: loads WBS/CALC, chooses base or simulation,
' prepares layout, draws shapes, links, today line, constraints, and UI state.
'
' Entrees / Inputs:
' - Mode displayOnly/full, statut nouvelle feuille, option d'activation.
' - tbl_WBS, tbl_CALC, maps GanttLive base/test, etats de toggles.
'
' Sorties / Outputs:
' - GANTT visuel coherent; flags live consommes; watcher Drag reconcilie.
' - Messages console en cas d'erreur VBA de rendu.
'
' Appele par / Called by:
' - CommitGanttUpdate via mod_GanttEngine.
'
' Notes:
' - Procedure la plus sensible du module; elle coordonne rendu, UI, live TEST, contraintes et Drag.
'------------------------------------------------------------------------------
Public Sub RunGanttRefreshCore( _
    ByVal displayOnly As Boolean, _
    ByVal isNewSheet As Boolean, _
    ByVal activateGantt As Boolean, _
    Optional ByVal forceTimelineLayout As Boolean = False)

    Dim perfScope As clsPerfScope

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim wsGantt As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject
    Dim dataArr As Variant
    Dim mapWBS As Object
    Dim hasChildren As Object
    Dim rowById As Object

    Dim projectStart As Variant
    Dim projectFinish As Variant
    Dim totalDays As Long

    Dim rowCount As Long
    Dim renderableRowCount As Long
    Dim ganttRow As Long
    Dim r As Long

    Dim baseById As Object
    Dim testById As Object
    Dim isTestMode As Boolean
    Dim renderMode As String
    Dim renderConstraintMarkers As Boolean
    Dim renderDeadlineMarkers As Boolean
    Dim testInputMap As Object
    Dim calcDrivingMap As Object
    Dim constraintById As Object

    Dim wasGanttSheetCreated As Boolean
    Dim needsVisualLayoutStabilization As Boolean
    Dim consoleMessages As Collection
    Dim wsActiveBeforeRefresh As Worksheet
    Dim selectionAddressBeforeRefresh As String
    Dim shouldRestoreActiveContext As Boolean
    Dim refreshErrNumber As Long
    Dim refreshErrDescription As String
    Dim renderSignature As String
    Dim timelineSignature As String
    Dim timelineNeedsUpdate As Boolean
    Dim timelineCompatibleChange As Boolean
    Dim noOpApplied As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldInternalWrite As Boolean
    Dim applicationStateCaptured As Boolean
    Dim localChangeSet As Object
    Dim localChangedIds As Object
    Dim localChangedRows As Object
    Dim localFallbackReason As String
    Dim localApplied As Boolean
    Dim dependencyLocalApplied As Boolean
    Dim forceCanonicalRebuild As Boolean
    Dim canonicalFallbackReason As String
    Dim physicalGeometryCurrent As Boolean
    Dim physicalDirtyIds As Object
    Dim incidentDirtyCount As Long
    Dim incidentDirtyMarked As Boolean
    Dim topologyChanged As Boolean

    Set perfScope = Profiler_BeginScope("RunGanttRefreshCore", "Gantt")
    gLastRefreshSucceeded = False
    gLastRefreshEffectiveScope = ""

    On Error GoTo SafeExit

    If Not displayOnly Then GanttRefresh_InvalidateDifferentialState "FullBuild"

    Set consoleMessages = New Collection
    shouldRestoreActiveContext = False

    If Not displayOnly Then GanttDependency_ClearExpandedLinksCache
    EnsureGanttViewInitialized
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldInternalWrite = GetGanttInternalWrite()
    applicationStateCaptured = True
    SetGanttInternalWrite True
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then
        Gantt_SafeEmptyState
        GoTo SafeExit
    End If

    'Important:
    '- EnsureGanttSheet can detect whether the GANTT sheet has just been created.
    '- That information must be propagated to PrepareGanttFullLayout.
    '- Otherwise FreezeGanttAfterFinish is never called after deleting/recreating GANTT.
    Set wsGantt = EnsureGanttSheet(wasGanttSheetCreated)
    If wasGanttSheetCreated Then isNewSheet = True
    If displayOnly And Not wasGanttSheetCreated Then
        If Not GanttLocal_HasCommittedSnapshot() Then
            If GanttLocal_PrimeNormalState() Then
                gLastWorksheetShapeCount = wsGantt.Shapes.Count
            End If
        ElseIf gLastWorksheetShapeCount <= 0 Then
            gLastWorksheetShapeCount = wsGantt.Shapes.Count
        End If
    End If

    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)
    If tblCalc.DataBodyRange Is Nothing Then
        Gantt_SafeEmptyState
        GoTo SafeExit
    End If
    needsVisualLayoutStabilization = (Not displayOnly) And (isNewSheet Or IsGanttSheetLayoutEmpty(wsGantt))

    If GetGanttPreserveTestInputs() Then
        Set testInputMap = BuildGanttTestInputMap(wsGantt)
    Else
        Set testInputMap = Nothing
    End If

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    ValidateGanttSourceColumns mapWBS

    dataArr = tblWBS.DataBodyRange.value
    rowCount = UBound(dataArr, 1)
    renderableRowCount = CountRenderableGanttRows(dataArr, mapWBS)
    If renderableRowCount < 1 Then
        Gantt_SafeEmptyState
        GoTo SafeExit
    End If

    Set hasChildren = GanttHierarchy_BuildDirectParentPresenceFromWbs(dataArr, mapWBS)
    Set rowById = GanttRenderer_BuildWbsRowById(dataArr, mapWBS)

    Set baseById = GanttLive_BuildBaseByIdMap()
    Set testById = GanttLive_BuildTestByIdMap()
    renderMode = GanttLive_GetPendingRenderMode()
    isTestMode = GanttLive_IsTestRenderRequested()
    renderConstraintMarkers = (renderMode <> "SCENARIO") And GetGanttShowConstraints()
    renderDeadlineMarkers = (renderMode = "") And GetGanttShowConstraints() And IsAnalyticsEnabled()

    ResolveDisplayedProjectRange dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, projectStart, projectFinish
    If Not HasValue(projectStart) Or Not HasValue(projectFinish) Then
        Gantt_SafeEmptyState
        GoTo SafeExit
    End If

    If IsAggregatedScaleMode() Then
        projectStart = GetScaleProjectStart(projectStart)
        projectFinish = GetScaleProjectFinish(projectFinish)
    End If

    totalDays = GetTimelineSlotCount(projectStart, projectFinish)
    If totalDays < 1 Then
        Gantt_SafeEmptyState
        GoTo SafeExit
    End If

    renderSignature = GanttRefresh_BuildRenderSignature( _
        dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, _
        projectStart, projectFinish, totalDays, renderMode)
    timelineSignature = GanttRefresh_BuildTimelineSignature( _
        projectStart, projectFinish, totalDays, rowCount)
    physicalGeometryCurrent = GanttDependencySvg_IsPhysicalStateCurrent(wsGantt)

    If displayOnly And Not forceTimelineLayout Then
        If Len(gLastTimelineSignature) = 0 Then
            GanttRefresh_PrimeColdTimelineSignature _
                dataArr, mapWBS, hasChildren, baseById, rowCount, _
                projectStart, projectFinish, totalDays, timelineSignature
        End If
    End If

    If displayOnly And Not forceTimelineLayout Then
        If Len(gLastTimelineSignature) = 0 And _
           GanttLocal_HasCommittedSnapshot() And _
           physicalGeometryCurrent Then
            gLastTimelineSignature = timelineSignature
            If gLastWorksheetShapeCount <= 0 Then gLastWorksheetShapeCount = wsGantt.Shapes.Count
            Profiler_RecordOperation "GanttDiffColdTimelineSignatureHydrated", 1, 0#
        End If
    End If

    If displayOnly And Not forceTimelineLayout Then
        If renderSignature = gLastRenderSignature And _
           wsGantt.Shapes.Count = gLastWorksheetShapeCount And _
           physicalGeometryCurrent Then
            Profiler_RecordOperation "GanttDiffNoOpPaths", 1, 0#
            noOpApplied = True
            GoTo RenderComplete
        End If
    End If

    If displayOnly And Not forceTimelineLayout Then
        If Len(gLastTimelineSignature) > 0 And _
           timelineSignature <> gLastTimelineSignature Then
            timelineCompatibleChange = GanttRefresh_IsTimelineSubsetCompatible( _
                gLastTimelineSignature, timelineSignature)
            If timelineCompatibleChange Then
                timelineNeedsUpdate = True
                Profiler_RecordOperation "GanttRenderScope_SUBSET_TimelineCompatible", 1, 0#
                Profiler_RecordOperation "GanttStructuralRenderChange_FALSE", 1, 0#
            Else
                forceCanonicalRebuild = True
                canonicalFallbackReason = "TimelineSignatureChanged"
                Profiler_RecordOperation "GanttFullScopeReason_TimelineSignatureChanged", 1, 0#
                Profiler_RecordOperation "GanttStructuralRenderChange_TRUE", 1, 0#
                GoTo LocalDecisionComplete
            End If
        End If
    End If

    'A local transaction is decided before building a global expected registry
    'or scanning the worksheet Shapes collection.
    If displayOnly And Not forceTimelineLayout And _
       (timelineSignature = gLastTimelineSignature Or timelineCompatibleChange) Then
        Set localChangeSet = GanttLocal_BuildChangeSet( _
            dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, renderMode)

        If CBool(localChangeSet("LocalEligible")) Then
            If wsGantt.Shapes.Count = gLastWorksheetShapeCount Then
                Set localChangedIds = localChangeSet("ChangedIds")
                Set localChangedRows = localChangeSet("ChangedRows")
                topologyChanged = CBool(localChangeSet("TopologyChanged"))

                If localChangedIds.Count = 0 And physicalGeometryCurrent And Not timelineNeedsUpdate Then
                    GanttLocal_CommitChangeSet localChangeSet
                    gLastRenderSignature = renderSignature
                    gLastTimelineSignature = timelineSignature
                    gLastWorksheetShapeCount = wsGantt.Shapes.Count
                    Profiler_RecordOperation "GanttLocalNoOpPaths", 1, 0#
                    noOpApplied = True
                    GoTo RenderComplete
                ElseIf localChangedIds.Count = 0 Then
                    Set physicalDirtyIds = GanttDependencySvg_GetPhysicallyDirtyTaskIds(wsGantt, rowById)
                    If Not physicalDirtyIds Is Nothing Then
                        If physicalDirtyIds.Count > 0 Then
                            Set localChangedIds = physicalDirtyIds
                            Set localChangedRows = GanttRefresh_BuildChangedRowsForIds(localChangedIds, rowById)
                            Profiler_RecordOperation "GanttLocalPhysicalDirtyPaths", 1, 0#
                            Profiler_RecordOperation "GanttLocalPhysicalDirtyIds", localChangedIds.Count, 0#
                        Else
                            forceCanonicalRebuild = True
                            canonicalFallbackReason = "PhysicalGeometryMismatchUnresolved"
                            Profiler_RecordOperation "GanttLocalFallback_PhysicalGeometryMismatchUnresolved", 1, 0#
                            GoTo LocalDecisionComplete
                        End If
                    Else
                        forceCanonicalRebuild = True
                        canonicalFallbackReason = "PhysicalGeometrySnapshotMissing"
                        Profiler_RecordOperation "GanttLocalFallback_PhysicalGeometrySnapshotMissing", 1, 0#
                        GoTo LocalDecisionComplete
                    End If
                End If

                Set calcDrivingMap = CanonicalIdentity_GetDrivingLogicByIdMap()
                GanttLayout_UpdateAffectedLeftPanelRows _
                    wsGantt, dataArr, mapWBS, hasChildren, calcDrivingMap, _
                    localChangedIds, rowById

                If timelineNeedsUpdate Then
                    PrepareGanttDifferentialLayout _
                        wsGantt, rowCount, projectStart, totalDays, True
                End If

                GanttTimelineProjection_Begin wsGantt, projectStart, totalDays

                If topologyChanged Then
                    If Not GanttDependency_PrimeLocalIndex(wsGantt, True) Then
                        localFallbackReason = "DependencyTopologyIndexRefreshFailed"
                        Profiler_RecordOperation "GanttLocalFallback_" & localFallbackReason, 1, 0#
                        forceCanonicalRebuild = True
                        canonicalFallbackReason = localFallbackReason
                        GoTo LocalDecisionComplete
                    End If
                    Profiler_RecordOperation "GanttDependencyTopologyRefreshes", 1, 0#
                End If

                incidentDirtyMarked = GanttDependency_MarkAffectedLinksDirty(localChangedIds, incidentDirtyCount)
                If Not incidentDirtyMarked Then
                    If GanttDependency_PrimeLocalIndex(wsGantt) Then
                        Profiler_RecordOperation "GanttDependencyIncidentIndexPrimedForLocalCommit", 1, 0#
                        incidentDirtyMarked = GanttDependency_MarkAffectedLinksDirty(localChangedIds, incidentDirtyCount)
                    End If
                End If

                If Not incidentDirtyMarked Then
                    localFallbackReason = "IncidentDependencyIndexAbsent"
                    Profiler_RecordOperation "GanttLocalFallback_" & localFallbackReason, 1, 0#
                    forceCanonicalRebuild = True
                    canonicalFallbackReason = localFallbackReason
                    GoTo LocalDecisionComplete
                End If
                Profiler_RecordOperation "GanttLocalActualGeometryChangedIds", localChangedIds.Count, 0#
                Profiler_RecordOperation "GanttLocalIncidentDirtyLinks", incidentDirtyCount, 0#
                Profiler_RecordOperation "GanttRenderScope_SUBSET_ProjectionIds", localChangedIds.Count, 0#

                GanttShapeRegistry_BeginLocalRenderSession wsGantt, localChangedIds, rowById
                GanttRefresh_ProjectTasksStage _
                    wsGantt, dataArr, mapWBS, hasChildren, projectStart, totalDays, _
                    baseById, testById, isTestMode, localChangedIds, localChangedRows
                GanttShapeRegistry_EndRenderSession wsGantt

                dependencyLocalApplied = GanttRefresh_RouteDependenciesStage( _
                    wsGantt, mapWBS, dataArr, hasChildren, rowById, _
                    projectStart, totalDays, baseById, testById, isTestMode, _
                    localChangedIds, False, localFallbackReason)

                If dependencyLocalApplied Then
                    gLastRefreshEffectiveScope = "INCREMENTAL"
                    GanttLocal_CommitChangeSet localChangeSet
                    localApplied = True
                    Profiler_RecordOperation "GanttLocalCommits", 1, 0#
                    GoTo RenderComplete
                End If

                Profiler_RecordOperation "GanttLocalFallback_" & localFallbackReason, 1, 0#
            Else
                Profiler_RecordOperation "GanttLocalFallback_ShapeCountMismatch", 1, 0#
                forceCanonicalRebuild = True
                canonicalFallbackReason = "ShapeCountMismatch"
            End If
        Else
            Select Case CStr(localChangeSet("FallbackReason"))
                Case "DatasetChanged", "StructuralRowChanged", "ChangeSetTooLarge"
                    forceCanonicalRebuild = True
                    canonicalFallbackReason = CStr(localChangeSet("FallbackReason"))
            End Select
        End If
    End If

LocalDecisionComplete:

    If forceCanonicalRebuild Then
        gLastRefreshEffectiveScope = "FULL"
        Profiler_RecordOperation "GanttRenderScope_FULL", 1, 0#
        Profiler_RecordOperation _
            "GanttFullRebuildReason_" & canonicalFallbackReason, 1, 0#
        GanttShapeRegistry_InvalidateCanonical "CanonicalFallback"
        GanttDependency_InvalidateLocalIndex "CanonicalFallback"
        GanttDependency_ClearExpandedLinksCache
    End If

    If displayOnly And Not forceCanonicalRebuild Then
        timelineNeedsUpdate = forceTimelineLayout Or _
            (timelineSignature <> gLastTimelineSignature)
        PrepareGanttDifferentialLayout _
            wsGantt, rowCount, projectStart, totalDays, timelineNeedsUpdate
    Else
        Set calcDrivingMap = CanonicalIdentity_GetDrivingLogicByIdMap()
        PrepareGanttFullLayout wsGantt, dataArr, mapWBS, hasChildren, calcDrivingMap, rowCount, projectStart, totalDays, testInputMap, isNewSheet, activateGantt
    End If

    If Not displayOnly Or forceCanonicalRebuild Then
        EnsureGanttVisualLayoutReadyBeforeDrawing wsGantt, activateGantt, activateGantt
    End If

    If renderConstraintMarkers Or renderDeadlineMarkers Then
        Set constraintById = BuildGanttConstraintMapFromCalc(renderDeadlineMarkers)
    Else
        Set constraintById = CreateObject("Scripting.Dictionary")
    End If

    GanttTimelineProjection_Begin wsGantt, projectStart, totalDays


    GanttShapeRegistry_BeginRenderSession _
        wsGantt, True, False, False, _
        (displayOnly And forceTimelineLayout And Not forceCanonicalRebuild)
    GanttRefresh_ProjectTasksStage wsGantt, dataArr, mapWBS, hasChildren, projectStart, totalDays, baseById, testById, isTestMode
    DrawTodayLine wsGantt, projectStart, totalDays, rowCount
    GanttShapeRegistry_EndRenderSession wsGantt

    dependencyLocalApplied = GanttRefresh_RouteDependenciesStage( _
        wsGantt, mapWBS, dataArr, hasChildren, rowById, projectStart, totalDays, _
        baseById, testById, isTestMode, Nothing, (Not displayOnly Or forceCanonicalRebuild), _
        localFallbackReason)
    If Not dependencyLocalApplied Then Err.Raise 5, "RunGanttRefreshCore", _
        "Dependency routing failed: " & localFallbackReason
    ApplyGanttUiState wsGantt, (Not displayOnly Or forceCanonicalRebuild), _
        (Not displayOnly Or forceCanonicalRebuild)
    If displayOnly And Not forceCanonicalRebuild Then
        GanttConstraint_ClearOverlay wsGantt, rowCount, totalDays
    End If
    If renderConstraintMarkers Or renderDeadlineMarkers Then DrawConstraintMarkers_Gantt wsGantt, dataArr, mapWBS, hasChildren, projectStart, totalDays, constraintById, renderConstraintMarkers, renderDeadlineMarkers

    GanttLocal_CaptureFullSnapshot _
        dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, _
        IIf(displayOnly, "DIFF", "FULL")

    If Not displayOnly Or forceCanonicalRebuild Then Gantt_ClearPendingGeometryRepair

RenderComplete:
    GanttDependencySvg_EnsureFrontLayer wsGantt
    If Not noOpApplied Then
        gLastRenderSignature = renderSignature
        gLastTimelineSignature = timelineSignature
        gLastWorksheetShapeCount = wsGantt.Shapes.Count
        Profiler_RecordOperation "GanttDiffPaths", 1, 0#
    End If
    If Len(gLastRefreshEffectiveScope) = 0 Then
        If displayOnly Then
            gLastRefreshEffectiveScope = "INCREMENTAL"
        Else
            gLastRefreshEffectiveScope = "FULL"
        End If
    End If
    gLastRefreshSucceeded = True

SafeExit:
    refreshErrNumber = Err.Number
    refreshErrDescription = Err.Description
    GanttShapeRegistry_CancelRenderSession
    GanttTimelineProjection_End

    If shouldRestoreActiveContext Then
        RestoreGanttCallerVisualContext wsActiveBeforeRefresh, selectionAddressBeforeRefresh
    End If

    If applicationStateCaptured Then
        Application.EnableEvents = oldEnableEvents
        Application.ScreenUpdating = oldScreenUpdating
        SetGanttInternalWrite oldInternalWrite
    End If

    If Not noOpApplied Then
        If IsPlanningWorkflowActive() Then
            If GanttDrag_IsWatching() Then GanttDrag_RebuildWatchMaps
        Else
            On Error Resume Next
            GanttDrag_ReconcileWatchState
            On Error GoTo 0
        End If
    End If

    If Not GetGanttPreserveTestInputs() Then
        GanttLive_ClearTestRenderRequest
    End If

    If refreshErrNumber <> 0 Then

        If displayOnly Then
            Profiler_RecordOperation "GanttDiffFallback_RenderError", 1, 0#
            Gantt_AddConsoleMessage consoleMessages, "STOP", _
                "Erreur VBA dans Refresh_Gantt_DisplayOnly" & vbCrLf & _
                "-> vérifier le dernier bloc modifié dans mod_Gantt" & vbCrLf & _
                "-> " & refreshErrDescription, _
                "VBA error in Refresh_Gantt_DisplayOnly" & vbCrLf & _
                "-> check the last edited block in mod_Gantt" & vbCrLf & _
                "-> " & refreshErrDescription
        Else
            Gantt_AddConsoleMessage consoleMessages, "STOP", _
                "Erreur VBA dans Refresh_Gantt" & vbCrLf & _
                "-> vérifier le dernier bloc modifié dans mod_Gantt" & vbCrLf & _
                "-> " & refreshErrDescription, _
                "VBA error in Refresh_Gantt" & vbCrLf & _
                "-> check the last edited block in mod_Gantt" & vbCrLf & _
                "-> " & refreshErrDescription
        End If

        CalcBridge_ShowPlanningConsole consoleMessages

    End If

End Sub

Private Sub GanttRefresh_ProjectTasksStage( _
    ByVal wsGantt As Worksheet, _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    Optional ByVal projectionIds As Variant, _
    Optional ByVal projectionRows As Variant)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("GanttRefresh_ProjectTasksStage", "Gantt Engine")
    Profiler_RecordOperation "GanttRefreshStage_ProjectTasks", 1, 0#
    If IsMissing(projectionIds) And IsMissing(projectionRows) Then
        DrawGanttShapes wsGantt, dataArr, mapWBS, hasChildren, _
            projectStart, totalDays, baseById, testById, isTestMode
    Else
        DrawGanttShapes wsGantt, dataArr, mapWBS, hasChildren, _
            projectStart, totalDays, baseById, testById, isTestMode, _
            projectionIds, projectionRows
    End If

End Sub

Private Function GanttRefresh_RouteDependenciesStage( _
    ByVal wsGantt As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal rowById As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByVal affectedIds As Object, _
    ByVal fullScope As Boolean, _
    ByRef fallbackReason As String) As Boolean

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("GanttRefresh_RouteDependenciesStage", "Gantt Engine")
    fallbackReason = ""
    Profiler_RecordOperation "GanttRefreshStage_RouteDependencies", 1, 0#

    If fullScope Then
        Profiler_RecordOperation "GanttRefreshStage_RouteDependencies_FULL", 1, 0#
        DrawDependencyLinks _
            wsGantt, mapWBS, dataArr, hasChildren, rowById, projectStart, totalDays, _
            baseById, testById, isTestMode, True
        GanttRefresh_RouteDependenciesStage = True
        Exit Function
    End If

    Profiler_RecordOperation "GanttRefreshStage_RouteDependencies_PARTIAL", 1, 0#
    GanttRefresh_RouteDependenciesStage = GanttDependency_DrawAffectedLinks( _
        wsGantt, mapWBS, dataArr, hasChildren, rowById, _
        projectStart, totalDays, baseById, testById, isTestMode, _
        affectedIds, fallbackReason)

End Function

'------------------------------------------------------------------------------
' FR: Invalide le cache de post-condition differentielle apres une mutation externe.
' EN: Invalidates the differential post-condition cache after an external mutation.
'------------------------------------------------------------------------------
Public Sub GanttRefresh_InvalidateDifferentialState(Optional ByVal reason As String = "")

    gLastRenderSignature = ""
    gLastTimelineSignature = ""
    gLastWorksheetShapeCount = -1
    GanttLocal_Invalidate reason
    GanttShapeRegistry_InvalidateCanonical reason
    GanttDependency_InvalidateLocalIndex reason
    If Len(reason) > 0 Then
        Profiler_RecordOperation "GanttDiffInvalidation_" & reason, 1, 0#
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Applique uniquement le style analytique des shapes existantes.
' EN: Applies only analytical styling to existing shapes.
'------------------------------------------------------------------------------
Public Sub GanttRefresh_ApplyAnalyticsStyleOnly()

    Dim wsWBS As Worksheet
    Dim wsGantt As Worksheet
    Dim tblWBS As ListObject
    Dim dataArr As Variant
    Dim mapWBS As Object
    Dim hasChildren As Object
    Dim baseById As Object
    Dim testById As Object
    Dim rowById As Object
    Dim changeSet As Object
    Dim changedIds As Object
    Dim changedRows As Object
    Dim projectStart As Variant
    Dim projectFinish As Variant
    Dim totalDays As Long
    Dim isTestMode As Boolean
    Dim canonicalReadyAtStart As Boolean
    Dim styleRecordsAvailable As Boolean
    Dim oldScreenUpdating As Boolean
    Dim screenStateCaptured As Boolean

    On Error GoTo Fallback

    EnsureGanttViewInitialized
    Set wsGantt = ThisWorkbook.Worksheets(GANTT_SHEET)
    If IsGanttSheetLayoutEmpty(wsGantt) Then GoTo Fallback
    If Not GanttDependencySvg_IsPhysicalStateCurrent(wsGantt) Then
        Profiler_RecordOperation "GanttStyleOnlyGeometryInvalidations", 1, 0#
        If Not CommitGanttUpdate( _
            GanttEngine_ActiveDataSource(), _
            GANTT_UPDATE_SCOPE_INCREMENTAL, _
            GANTT_RENDER_INTENT_OFFSCREEN, _
            "GanttRefresh_ApplyAnalyticsStyleOnlyGeometryInvalidation") Then
            Err.Raise 5, "GanttRefresh_ApplyAnalyticsStyleOnly", "Gantt geometry reconciliation failed."
        End If
        Exit Sub
    End If
    If gLastWorksheetShapeCount > 0 Then
        If wsGantt.Shapes.Count <> gLastWorksheetShapeCount Then
            Profiler_RecordOperation "GanttDiffFallback_StyleShapeCountMismatch", 1, 0#
            GoTo Fallback
        End If
    End If

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    If tblWBS.DataBodyRange Is Nothing Then GoTo Fallback

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    ValidateGanttSourceColumns mapWBS
    dataArr = tblWBS.DataBodyRange.value
    Set hasChildren = GanttHierarchy_BuildDirectParentPresenceFromWbs(dataArr, mapWBS)
    Set rowById = GanttRenderer_BuildWbsRowById(dataArr, mapWBS)
    Set baseById = GanttLive_BuildBaseByIdMap()
    Set testById = GanttLive_BuildTestByIdMap()
    isTestMode = (Trim$(GanttLive_GetActiveSimulationMode()) <> "")

    ResolveDisplayedProjectRange _
        dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, _
        projectStart, projectFinish
    If Not HasValue(projectStart) Or Not HasValue(projectFinish) Then GoTo Fallback

    If IsAggregatedScaleMode() Then
        projectStart = GetScaleProjectStart(projectStart)
        projectFinish = GetScaleProjectFinish(projectFinish)
    End If
    totalDays = GetTimelineSlotCount(projectStart, projectFinish)

    oldScreenUpdating = Application.ScreenUpdating
    screenStateCaptured = True
    Application.ScreenUpdating = False

    If Not GanttLocal_HasCommittedSnapshot() Then GanttLocal_PrimeNormalState
    Set changeSet = GanttLocal_BuildChangeSet( _
        dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, "STYLE")

    If CBool(changeSet("LocalEligible")) Then
        Set changedIds = changeSet("ChangedIds")
        Set changedRows = changeSet("ChangedRows")
        canonicalReadyAtStart = GanttShapeRegistry_HasCanonicalRecords()
        Profiler_RecordOperation "GanttStyleOnlyCanonicalReadyAtStart", Abs(CLng(canonicalReadyAtStart)), 0#
        GanttShapeRegistry_BeginLocalRenderSession wsGantt, changedIds, rowById, True
        styleRecordsAvailable = GanttRenderer_ApplyAnalyticsStyleOnlyRecords( _
            wsGantt, dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, changedRows)
        If Not styleRecordsAvailable Then
            GanttShapeRegistry_CancelRenderSession
            GanttTimelineProjection_End
            If canonicalReadyAtStart Then
                Profiler_RecordOperation "GanttDiffFallback_StyleOnlyCanonicalRecordMissing", 1, 0#
            Else
                Profiler_RecordOperation "GanttDiffFallback_StyleOnlyCanonicalCold", 1, 0#
            End If

            ' The common engine correctly treats analytics-only changes as
            ' geometry no-ops. Hydrate and style only the affected task rows
            ' through the existing projection owner instead of forcing FULL.
            GanttTimelineProjection_Begin wsGantt, projectStart, totalDays
            GanttShapeRegistry_BeginLocalRenderSession wsGantt, changedIds, rowById
            GanttRefresh_ProjectTasksStage _
                wsGantt, dataArr, mapWBS, hasChildren, projectStart, totalDays, _
                baseById, testById, isTestMode, changedIds, changedRows
            GanttShapeRegistry_EndRenderSession wsGantt
            GanttTimelineProjection_End
            GanttLocal_CommitChangeSet changeSet
            GanttUiControls_EnsureCanonical wsGantt
            If screenStateCaptured Then Application.ScreenUpdating = oldScreenUpdating
            Profiler_RecordOperation "GanttDiffStyleOnlyHydrationPaths", 1, 0#
            Exit Sub
        End If
        GanttShapeRegistry_EndRenderSession wsGantt
        GanttLocal_CommitChangeSet changeSet
        Profiler_RecordOperation "GanttLocalStyleIds", changedIds.Count, 0#
    Else
        GanttShapeRegistry_CancelRenderSession
        GanttTimelineProjection_End
        If screenStateCaptured Then Application.ScreenUpdating = oldScreenUpdating
        Profiler_RecordOperation "GanttDiffFallback_StyleOnlyEngine", 1, 0#
        CommitGanttUpdate _
            GanttEngine_ActiveDataSource(), _
            GANTT_UPDATE_SCOPE_PARTIAL, _
            GANTT_RENDER_INTENT_OFFSCREEN, _
            "GanttRefresh_ApplyAnalyticsStyleOnlyNonLocal"
        Exit Sub
    End If
    GanttUiControls_EnsureCanonical wsGantt

    Application.ScreenUpdating = oldScreenUpdating
    Profiler_RecordOperation "GanttDiffStyleOnlyPaths", 1, 0#
    Exit Sub

Fallback:
    GanttShapeRegistry_CancelRenderSession
    GanttTimelineProjection_End
    If screenStateCaptured Then Application.ScreenUpdating = oldScreenUpdating
    Profiler_RecordOperation "GanttDiffFallback_StyleOnly", 1, 0#
    CommitGanttUpdate _
        GanttEngine_ActiveDataSource(), _
        GANTT_UPDATE_SCOPE_PARTIAL, _
        GANTT_RENDER_INTENT_OFFSCREEN, _
        "GanttRefresh_ApplyAnalyticsStyleOnlyFallback"

End Sub

Private Function GanttRefresh_BuildChangedRowsForIds( _
    ByVal changedIds As Object, _
    ByVal rowById As Object) As Object

    Dim changedRows As Object
    Dim idKey As Variant
    Dim ganttRow As Long
    Dim dataRow As Long

    Set changedRows = CreateObject("Scripting.Dictionary")
    If changedIds Is Nothing Or rowById Is Nothing Then
        Set GanttRefresh_BuildChangedRowsForIds = changedRows
        Exit Function
    End If

    For Each idKey In changedIds.keys
        If rowById.Exists(CStr(idKey)) Then
            ganttRow = CLng(rowById(CStr(idKey)))
            dataRow = ganttRow - FIRST_TASK_ROW + 1
            If dataRow > 0 Then changedRows(CStr(dataRow)) = True
        End If
    Next idKey

    Set GanttRefresh_BuildChangedRowsForIds = changedRows

End Function

'------------------------------------------------------------------------------
' FR: Construit la signature du modele visuel courant sans lire les shapes.
' EN: Builds the current visual-model signature without reading shapes.
'------------------------------------------------------------------------------
Private Function GanttRefresh_BuildRenderSignature( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByVal projectStart As Variant, _
    ByVal projectFinish As Variant, _
    ByVal totalDays As Long, _
    ByVal renderMode As String) As String

    Dim parts() As String
    Dim r As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim progressVal As Variant
    Dim taskNameVal As String
    Dim taskTypeVal As String
    Dim summaryDisplayVal As String
    Dim criticalPathVal As String
    Dim longestPathVal As String

    ReDim parts(0 To UBound(dataArr, 1))
    parts(0) = _
        GetGanttTimelineScaleMode() & "|" & GetGanttViewMode() & "|" & _
        CStr(GetGanttShowConstraints()) & "|" & UCase$(Trim$(renderMode)) & "|" & _
        GanttRefresh_SignatureValue(projectStart) & "|" & _
        GanttRefresh_SignatureValue(projectFinish) & "|" & CStr(totalDays)

    For r = 1 To UBound(dataArr, 1)
        idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))
        wbsVal = NormalizeWBS(CStr(dataArr(r, mapWBS("WBS"))))
        progressVal = GanttLive_GetDisplayProgress(idVal, baseById, testById, isTestMode)
        taskNameVal = GanttRefresh_ArrayText(dataArr, r, mapWBS, "Task Name")
        taskTypeVal = GanttRefresh_ArrayText(dataArr, r, mapWBS, "Task Type")
        summaryDisplayVal = GanttRefresh_ArrayText(dataArr, r, mapWBS, "S")
        criticalPathVal = GanttRefresh_ArrayText(dataArr, r, mapWBS, "Critical Path")
        longestPathVal = GanttRefresh_ArrayText(dataArr, r, mapWBS, "Longest Path")
        parts(r) = _
            idVal & "|" & wbsVal & "|" & taskNameVal & "|" & taskTypeVal & "|" & _
            summaryDisplayVal & "|" & criticalPathVal & "|" & longestPathVal & "|" & _
            GanttRefresh_SignatureValue(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode)) & "|" & _
            GanttRefresh_SignatureValue(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode)) & "|" & _
            GanttRefresh_SignatureValue(progressVal) & "|" & _
            CStr(hasChildren.Exists(wbsVal))
    Next r

    GanttRefresh_BuildRenderSignature = Join(parts, vbLf)

End Function

'------------------------------------------------------------------------------
' FR: Lit un champ optionnel du buffer WBS pour la signature de rendu.
' EN: Reads an optional WBS buffer field for the render signature.
'------------------------------------------------------------------------------
Private Function GanttRefresh_ArrayText( _
    ByRef dataArr As Variant, _
    ByVal dataRow As Long, _
    ByVal columnMap As Object, _
    ByVal columnName As String) As String

    If columnMap Is Nothing Then Exit Function
    If Not columnMap.Exists(columnName) Then Exit Function

    GanttRefresh_ArrayText = CStr(dataArr(dataRow, CLng(columnMap(columnName))))

End Function

'------------------------------------------------------------------------------
' FR: Construit la signature de la projection horizontale.
' EN: Builds the horizontal projection signature.
'------------------------------------------------------------------------------
Private Function GanttRefresh_BuildTimelineSignature( _
    ByVal projectStart As Variant, _
    ByVal projectFinish As Variant, _
    ByVal totalDays As Long, _
    ByVal rowCount As Long) As String

    GanttRefresh_BuildTimelineSignature = _
        GetGanttTimelineScaleMode() & "|" & _
        GanttRefresh_SignatureValue(projectStart) & "|" & _
        GanttRefresh_SignatureValue(projectFinish) & "|" & _
        CStr(totalDays) & "|" & CStr(rowCount)

End Function

Private Function GanttRefresh_IsTimelineSubsetCompatible( _
    ByVal previousSignature As String, _
    ByVal currentSignature As String) As Boolean

    Dim oldParts As Variant
    Dim newParts As Variant

    If Len(previousSignature) = 0 Or Len(currentSignature) = 0 Then Exit Function
    oldParts = Split(previousSignature, "|")
    newParts = Split(currentSignature, "|")
    If UBound(oldParts) < 4 Or UBound(newParts) < 4 Then Exit Function

    'Same scale, same left origin, and same rendered row count means a project
    'finish extension/contraction changes only the timeline tail. Existing X
    'anchors remain valid, so task/dependency work can stay dirty-set scoped.
    GanttRefresh_IsTimelineSubsetCompatible = _
        (CStr(oldParts(0)) = CStr(newParts(0))) And _
        (CStr(oldParts(1)) = CStr(newParts(1))) And _
        (CStr(oldParts(4)) = CStr(newParts(4)))

End Function

Private Sub GanttRefresh_PrimeColdTimelineSignature( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal baseById As Object, _
    ByVal rowCount As Long, _
    ByVal projectStart As Variant, _
    ByVal projectFinish As Variant, _
    ByVal totalDays As Long, _
    ByVal timelineSignature As String)

    Dim emptySimulation As Object
    Dim baseProjectStart As Variant
    Dim baseProjectFinish As Variant
    Dim baseTotalDays As Long
    Dim baseTimelineSignature As String

    On Error GoTo Failed
    If Len(timelineSignature) = 0 Then Exit Sub
    If Not GanttLocal_HasCommittedSnapshot() Then Exit Sub

    Set emptySimulation = CreateObject("Scripting.Dictionary")
    ResolveDisplayedProjectRange _
        dataArr, mapWBS, hasChildren, baseById, emptySimulation, False, _
        baseProjectStart, baseProjectFinish
    If Not HasValue(baseProjectStart) Or Not HasValue(baseProjectFinish) Then Exit Sub

    If IsAggregatedScaleMode() Then
        baseProjectStart = GetScaleProjectStart(baseProjectStart)
        baseProjectFinish = GetScaleProjectFinish(baseProjectFinish)
    End If

    baseTotalDays = GetTimelineSlotCount(baseProjectStart, baseProjectFinish)
    If baseTotalDays < 1 Then Exit Sub

    baseTimelineSignature = GanttRefresh_BuildTimelineSignature( _
        baseProjectStart, baseProjectFinish, baseTotalDays, rowCount)
    If baseTimelineSignature <> timelineSignature Then
        Profiler_RecordOperation "GanttDiffColdTimelineSignatureNeedsFull", 1, 0#
        Exit Sub
    End If

    gLastTimelineSignature = timelineSignature
    Profiler_RecordOperation "GanttDiffColdTimelineSignatureHydratedFromBaseState", 1, 0#
    Exit Sub

Failed:
    Profiler_RecordOperation "GanttDiffColdTimelineSignatureHydrateError", 1, 0#

End Sub

'------------------------------------------------------------------------------
' FR: Normalise une valeur pour une signature runtime non persistante.
' EN: Normalizes one value for a non-persistent runtime signature.
'------------------------------------------------------------------------------
Private Function GanttRefresh_SignatureValue(ByVal value As Variant) As String

    If HasValue(value) Then
        If IsDate(value) Or IsNumeric(value) Then
            GanttRefresh_SignatureValue = CStr(CDbl(value))
        Else
            GanttRefresh_SignatureValue = CStr(value)
        End If
    Else
        GanttRefresh_SignatureValue = "#"
    End If

End Function



'------------------------------------------------------------------------------
' FR: Nettoie, restaure ou normalise une partie de l'etat visuel GANTT.
' EN: Cleans, restores, or normalizes part of the GANTT visual state.
'------------------------------------------------------------------------------
Private Sub RestoreGanttCallerVisualContext(ByVal ws As Worksheet, ByVal selectionAddress As String)

    'Navigation is owned by the workflow facade, not by the renderer.

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Add Console Message dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Add Console Message helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub Gantt_AddConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

End Sub
