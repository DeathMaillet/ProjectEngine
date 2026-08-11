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
' - Refresh_Gantt et Refresh_Gantt_DisplayOnly.
'
' Notes:
' - Procedure la plus sensible du module; elle coordonne rendu, UI, live TEST, contraintes et Drag.
'------------------------------------------------------------------------------
Public Sub RunGanttRefreshCore( _
    ByVal displayOnly As Boolean, _
    ByVal isNewSheet As Boolean, _
    ByVal activateGantt As Boolean)

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

    Set perfScope = Profiler_BeginScope("RunGanttRefreshCore", "Gantt")

    On Error GoTo SafeExit

    Set consoleMessages = New Collection
    shouldRestoreActiveContext = Not activateGantt
    If shouldRestoreActiveContext Then
        On Error Resume Next
        Set wsActiveBeforeRefresh = Application.ActiveSheet
        selectionAddressBeforeRefresh = CStr(Application.Selection.Address(False, False))
        On Error GoTo SafeExit
    End If

    GanttDependency_ClearExpandedLinksCache
    EnsureGanttViewInitialized
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

    If renderConstraintMarkers Or renderDeadlineMarkers Then
        Set constraintById = BuildGanttConstraintMapFromCalc(renderDeadlineMarkers)
    Else
        Set constraintById = CreateObject("Scripting.Dictionary")
    End If

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

    If displayOnly Then
        PrepareGanttDisplayOnlyLayout wsGantt, rowCount, projectStart, totalDays, testInputMap
    Else
        Set calcDrivingMap = CanonicalIdentity_GetDrivingLogicByIdMap()
        PrepareGanttFullLayout wsGantt, dataArr, mapWBS, hasChildren, calcDrivingMap, rowCount, projectStart, totalDays, testInputMap, isNewSheet, activateGantt
    End If

    If Not displayOnly Then
        EnsureGanttVisualLayoutReadyBeforeDrawing wsGantt, activateGantt, activateGantt
    End If

    DrawGanttShapes wsGantt, dataArr, mapWBS, hasChildren, projectStart, totalDays, baseById, testById, isTestMode
    DrawDependencyLinks wsGantt, mapWBS, dataArr, hasChildren, rowById, projectStart, totalDays, baseById, testById, isTestMode
    DrawTodayLine wsGantt, projectStart, totalDays, rowCount
    ApplyGanttUiState wsGantt
    If renderConstraintMarkers Or renderDeadlineMarkers Then DrawConstraintMarkers_Gantt wsGantt, dataArr, mapWBS, hasChildren, projectStart, totalDays, constraintById, renderConstraintMarkers, renderDeadlineMarkers

    If Not displayOnly Then
        If activateGantt Then
            Gantt_ClearPendingGeometryRepair
        Else
            Gantt_CapturePendingGeometryRepair wsGantt
        End If
    End If

SafeExit:
    refreshErrNumber = Err.Number
    refreshErrDescription = Err.Description

    If shouldRestoreActiveContext Then
        RestoreGanttCallerVisualContext wsActiveBeforeRefresh, selectionAddressBeforeRefresh
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    SetGanttInternalWrite False

    If IsPlanningWorkflowActive() Then
        If GanttDrag_IsWatching() Then GanttDrag_RebuildWatchMaps
    Else
        On Error Resume Next
        GanttDrag_ReconcileWatchState
        On Error GoTo 0
    End If

    If Not GetGanttPreserveTestInputs() Then
        GanttLive_ClearTestRenderRequest
    End If

    If refreshErrNumber <> 0 Then

        If displayOnly Then
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



'------------------------------------------------------------------------------
' FR: Nettoie, restaure ou normalise une partie de l'etat visuel GANTT.
' EN: Cleans, restores, or normalizes part of the GANTT visual state.
'------------------------------------------------------------------------------
Private Sub RestoreGanttCallerVisualContext(ByVal ws As Worksheet, ByVal selectionAddress As String)

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Activate
    If Len(selectionAddress) > 0 Then
        ws.Range(selectionAddress).Select
    End If
    On Error GoTo 0

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

