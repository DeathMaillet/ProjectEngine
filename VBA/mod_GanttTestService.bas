Attribute VB_Name = "mod_GanttTestService"
Option Explicit

'===============================================================================
' MODULE : mod_GanttTestService
' DOMAINE / DOMAIN : Gantt
'
' FR
' Possede le workflow specialise indique par son nom et expose ses contrats stables.
' Ne possede pas les domaines appeles en dependance.
'
' EN
' Owns the named specialized workflow and exposes its stable contracts.
' Does not own the domains it calls as dependencies.
'
' CONTRATS / CONTRACTS : GanttTestService_RunTestEngine
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"

Private Const COL_TEST_START As Long = 5
Private Const COL_TEST_FINISH As Long = 6
Private Const COL_TEST_PROGRESS As Long = 9

'------------------------------------------------------------------------------
' FR: Orchestre Gantt Test Service Run Test Engine en preservant l'ordre contractuel des etapes du domaine.
' EN: Orchestrates Gantt Test Service Run Test Engine while preserving the domain's contractual step order.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Public Sub GanttTestService_RunTestEngine( _
    Optional ByVal silentMode As Boolean = False, _
    Optional ByRef transactionSucceeded As Variant, _
    Optional ByRef transactionMessages As Variant, _
    Optional ByRef transactionGanttRebuilt As Variant, _
    Optional ByVal recordSilentMessages As Boolean = True)

    Dim wsGantt As Worksheet
    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim wsTest As Worksheet

    Dim tblWBS As ListObject
    Dim tblCalc As ListObject
    Dim tblTest As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim mapCalcDriving As Object
    Dim wbsToId As Object
    Dim calcById As Object

    Dim dataWBS As Variant
    Dim outArr() As Variant

    Dim rowCount As Long
    Dim r As Long
    Dim testRow As Long

    Dim idVal As String
    Dim wbsVal As String
    Dim taskTypeVal As String
    Dim ganttRow As Long

    Dim baseStart As Variant
    Dim baseFinish As Variant
    Dim baseDuration As Variant
    Dim baseProgress As Variant
    Dim drivingLogic As String

    Dim baselineStart As Variant
    Dim baselineDuration As Variant
    Dim actualStart As Variant
    Dim actualFinish As Variant
    Dim forecastStart As Variant
    Dim forecastFinish As Variant

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgressRaw As Variant
    Dim testProgressNorm As Variant

    Dim inputStart As Variant
    Dim inputFinish As Variant
    Dim inputDuration As Variant
    Dim inputProgress As Variant

    Dim isSummary As Boolean
    Dim hasActual As Boolean
    Dim anyTestValue As Boolean
    Dim hasRenderableDelta As Boolean
    Dim hasAnyTestInput As Boolean

    Dim consoleMessages As Collection
    Dim workflowStarted As Boolean
    Dim testSucceeded As Boolean
    Dim ganttRebuilt As Boolean

    On Error GoTo SafeExit

    GanttTestAnalyticsSnapshot_Clear
    If Not silentMode Then workflowStarted = EnsurePlanningWorkflowStarted("Run_Gantt_Test_Engine")
    AppEvents_EnsureInitialized
    Set consoleMessages = New Collection

    Set wsGantt = ThisWorkbook.Worksheets(GANTT_SHEET)
    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)

    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then GoTo CleanExit
    If tblCalc.DataBodyRange Is Nothing Then GoTo CleanExit

    Set wsTest = Ensure_CalcGanttTest_Sheet()
    Set tblTest = Ensure_CalcGanttTest_Table(wsTest)

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set mapCalcDriving = CanonicalIdentity_GetDrivingLogicByIdMap()
    Set wbsToId = CanonicalIdentity_GetWbsToIdMap()
    Set calcById = BuildCalcByIdMap_Live(tblCalc, mapCalc)

    ValidateGanttTestSourceColumns mapWBS, mapCalc
    ValidateCalcGanttTestColumns tblTest

    dataWBS = tblWBS.DataBodyRange.value
    rowCount = UBound(dataWBS, 1)

    ResizeTableToRowCount_Generic tblTest, rowCount
    ReDim outArr(1 To rowCount, 1 To tblTest.ListColumns.Count)

    testRow = 0

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataWBS(r, mapWBS("ID"))))
        wbsVal = NormalizeWBS(CStr(dataWBS(r, mapWBS("WBS"))))
        taskTypeVal = Trim$(CStr(dataWBS(r, mapWBS("Task Type"))))

        testRow = testRow + 1

        baseStart = GetCellValue(dataWBS(r, mapWBS("Calculated Start")))
        baseFinish = GetCellValue(dataWBS(r, mapWBS("Calculated Finish")))
        baseDuration = GetCellValue(dataWBS(r, mapWBS("Calculated Duration")))
        baseProgress = GetCellValue(dataWBS(r, mapWBS("% Progress")))

        baselineStart = GetCellValue(dataWBS(r, mapWBS("Baseline Start")))
        baselineDuration = GetCellValue(dataWBS(r, mapWBS("Baseline Duration")))
        actualStart = GetCellValue(dataWBS(r, mapWBS("Actual Start")))
        actualFinish = GetCellValue(dataWBS(r, mapWBS("Actual Finish")))
        forecastStart = GetCellValue(dataWBS(r, mapWBS("Forecast Start")))
        forecastFinish = GetCellValue(dataWBS(r, mapWBS("Forecast Finish")))

        If mapCalcDriving.Exists(idVal) Then
            drivingLogic = CStr(mapCalcDriving(idVal))
        Else
            drivingLogic = ""
        End If

        isSummary = (UCase$(drivingLogic) = "SUMMARY")
        hasActual = (HasValue(actualStart) Or HasValue(actualFinish))

        testStart = Empty
        testFinish = Empty
        testProgressRaw = Empty
        testProgressNorm = Empty
        inputStart = Empty
        inputFinish = Empty
        inputDuration = Empty
        inputProgress = Empty
        anyTestValue = False

        ganttRow = FindGanttRowByWBS(wsGantt, wbsVal)

        If ganttRow > 0 Then
            testStart = GetCellValue(wsGantt.Cells(ganttRow, COL_TEST_START).value)
            testFinish = GetCellValue(wsGantt.Cells(ganttRow, COL_TEST_FINISH).value)
            testProgressRaw = GetCellValue(wsGantt.Cells(ganttRow, COL_TEST_PROGRESS).value)
        End If

        If HasValue(testStart) Or HasValue(testFinish) Or HasValue(testProgressRaw) Then
            anyTestValue = True
        End If

        If HasValue(testProgressRaw) Then
            testProgressNorm = NormalizePercentInput(testProgressRaw)
        Else
            testProgressNorm = Empty
        End If

        BuildTestInputValues_FromWBSSource _
            baselineStart, baselineDuration, actualStart, actualFinish, forecastStart, forecastFinish, _
            baseProgress, testStart, testFinish, testProgressNorm, _
            inputStart, inputFinish, inputDuration, inputProgress

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "ID")) = idVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "WBS")) = wbsVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Task Type")) = taskTypeVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Cal")) = NormalizeCalendarType(dataWBS(r, mapWBS("Cal")))
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Is Summary")) = IIf(isSummary, "YES", "NO")
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Parent ID")) = GetParentIdFromWBS(wbsVal, wbsToId)

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Start")) = baseStart
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Finish")) = baseFinish
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Duration")) = baseDuration
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Base Progress")) = baseProgress
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Driving Logic")) = drivingLogic
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Has Actual")) = IIf(hasActual, "YES", "NO")

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test Start")) = testStart
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test Finish")) = testFinish
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test % Raw")) = testProgressRaw
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Test % Normalized")) = testProgressNorm

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Start")) = inputStart
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Finish")) = inputFinish
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Duration")) = inputDuration
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Input Progress")) = inputProgress

        If calcById.Exists(idVal) Then
            outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Predecessors")) = calcById(idVal)
        End If

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Lag")) = Empty
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Any Test Value")) = IIf(anyTestValue, "YES", "NO")
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Warning Flag")) = ""
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Warning Text")) = ""
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Error Flag")) = ""

    Next r

    FormatCalcGanttTestColumns tblTest
    tblTest.DataBodyRange.value = outArr
    FormatCalcGanttTestColumns tblTest

    hasAnyTestInput = GanttLive_HasAnyTestInput(tblTest)

    If Not Run_Gantt_Test_Backend(tblTest, consoleMessages) Then
        GanttLive_AbortTestEngine wsGantt
        If silentMode Then
            If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Test_Engine"
        Else
            CalcBridge_ShowPlanningConsole consoleMessages
        End If
        GoTo CleanExit
    End If

    hasRenderableDelta = GanttLive_HasAnyRenderableTestDelta(tblTest)

    GanttLive_ApplyTestRender wsGantt
    If Not hasAnyTestInput Then GanttLive_ClearActiveSimulationMode
    testSucceeded = True
    ganttRebuilt = True

    If silentMode Then
        If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Test_Engine"
        GoTo CleanExit
    End If

    If Not hasAnyTestInput Then

        GanttLive_AddBiConsoleMessage consoleMessages, "INFO", _
            "Aucune saisie TEST détectée. L’affichage simulation a été réinitialisé.", _
            "No TEST input detected. Simulation display has been reset."

    ElseIf hasRenderableDelta Then

        GanttLive_AddBiConsoleMessage consoleMessages, "INFO", _
            "Simulation exécutée.", _
            "Simulation executed."

    Else

        GanttLive_AddBiConsoleMessage consoleMessages, "INFO", _
            "Simulation exécutée, mais aucun changement visible n'a été produit." & vbCrLf & _
            "-> cause probable : priorité locale (Actual/Forecast) et/ou contrainte réseau inchangée.", _
            "Simulation executed, but no visible change was produced." & vbCrLf & _
            "-> probable cause: local priority (Actual/Forecast) and/or unchanged network constraint."

    End If

    CalcBridge_ShowPlanningConsole consoleMessages

CleanExit:
    If Not IsMissing(transactionSucceeded) Then transactionSucceeded = testSucceeded
    If Not IsMissing(transactionGanttRebuilt) Then transactionGanttRebuilt = ganttRebuilt
    If Not IsMissing(transactionMessages) Then Set transactionMessages = consoleMessages

    If workflowStarted Then
        EndPlanningWorkflow
        On Error Resume Next
        GanttDrag_ReconcileWatchState
        On Error GoTo 0
    End If
    Exit Sub

SafeExit:
    testSucceeded = False
    GanttLive_SafeExit_TestEngine

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection

        GanttLive_AddVbaOrStructuredError consoleMessages, "Run_Gantt_Test_Engine", Err.Description

        If silentMode Then
            If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Test_Engine"
        Else
            CalcBridge_ShowPlanningConsole consoleMessages
        End If
    End If

    GoTo CleanExit

End Sub



'------------------------------------------------------------------------------
' FR:
' Transforme tbl_CALC_GANTT_TEST en dataset Core, synchronise les contraintes,
' execute Run_Calc_Core, puis reporte les resultats TEST et diagnostics dans la table live.
'
' EN:
' Transforms tbl_CALC_GANTT_TEST into a Core dataset, syncs constraints,
' runs Run_Calc_Core, then writes TEST results and diagnostics back to the live table.
'
' Entrees / Inputs:
' - tbl_CALC_GANTT_TEST preparee par Run_Gantt_Test_Engine.
' - tbl_LOGIC_LINKS/CALC via le bridge des liens Core.
' - Collection console optionnelle.
'
' Sorties / Outputs:
' - Colonnes Calc Test Start/Finish/Duration/Progress.
' - Warning/Error flags et messages utilisateurs.
' - Diagnostics Core ajoutes a la console en cas d'erreur racine.
'
' Appele par / Called by:
' - Run_Gantt_Test_Engine.
'
' Notes:
' - Frontiere principale entre GanttLive et Core Calculation.
' - Filtre les erreurs Core heritees pour ne pas dupliquer les causes racines.
'------------------------------------------------------------------------------
Private Function Run_Gantt_Test_Backend( _
    ByVal tblTest As ListObject, _
    Optional ByVal consoleMessages As Variant) As Boolean

    Dim dataCore As Variant
    Dim mapCore As Object
    Dim linksBySuccId As Object

    Dim outStart() As Variant
    Dim outFinish() As Variant
    Dim outDuration() As Variant
    Dim outProgress() As Variant
    Dim outWarnFlag() As Variant
    Dim outWarnText() As Variant
    Dim outErrFlag() As Variant

    Dim idToRowTest As Object
    Dim idToWbs As Object
    Dim errorIds As Object
    Dim rootErrorIds As Object
    Dim dependencyDiagnostics As Object
    Dim constraintDiagnostics As Object
    Dim cascadeDiagnostics As Object
    Dim warningActualIds As Object
    Dim constraintMessages As Collection
    Dim analyticsById As Object
    Dim resultById As Object

    Dim rowCount As Long
    Dim r As Long
    Dim idVal As String
    Dim errorMsg As String

    On Error GoTo SafeExit

    Run_Gantt_Test_Backend = False

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then
        Run_Gantt_Test_Backend = True
        Exit Function
    End If

    rowCount = tblTest.ListRows.Count

    ReDim outStart(1 To rowCount, 1 To 1)
    ReDim outFinish(1 To rowCount, 1 To 1)
    ReDim outDuration(1 To rowCount, 1 To 1)
    ReDim outProgress(1 To rowCount, 1 To 1)
    ReDim outWarnFlag(1 To rowCount, 1 To 1)
    ReDim outWarnText(1 To rowCount, 1 To 1)
    ReDim outErrFlag(1 To rowCount, 1 To 1)

    Set idToRowTest = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set errorIds = CreateObject("Scripting.Dictionary")
    Set rootErrorIds = CreateObject("Scripting.Dictionary")
    Set dependencyDiagnostics = CreateObject("Scripting.Dictionary")
    Set constraintDiagnostics = CreateObject("Scripting.Dictionary")
    Set cascadeDiagnostics = CreateObject("Scripting.Dictionary")
    Set warningActualIds = CreateObject("Scripting.Dictionary")

    If GanttLive_HasConsoleCollection(consoleMessages) Then
        Set constraintMessages = consoleMessages
    Else
        Set constraintMessages = New Collection
    End If

    If Not Sync_Constraints_To_CALC_ForWorkflow(constraintMessages) Then Exit Function

    BuildGanttTestCoreDataset _
        tblTest, dataCore, mapCore, idToRowTest, idToWbs, _
        outWarnFlag, outWarnText, warningActualIds

    Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded( _
        ThisWorkbook.Worksheets(CALC_SHEET).ListObjects(CALC_TABLE))

    Run_Calc_Core dataCore, mapCore, linksBySuccId, , dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataCore(r, mapCore("ID"))))
        If idVal = "" Then GoTo NextRow

        outStart(r, 1) = dataCore(r, mapCore("Calculated Start"))
        outFinish(r, 1) = dataCore(r, mapCore("Calculated Finish"))
        outDuration(r, 1) = dataCore(r, mapCore("Calculated Duration"))
        outProgress(r, 1) = dataCore(r, mapCore("Input Progress"))

        If UCase$(Trim$(CStr(dataCore(r, mapCore("Error flag"))))) = "ERROR" Then

            outErrFlag(r, 1) = "ERROR"
            errorIds(idVal) = True

            errorMsg = Trim$(CStr(dataCore(r, mapCore("ErrorMsg"))))

            If errorMsg <> "" Then
                outWarnText(r, 1) = errorMsg
            End If

            If Not GanttLive_IsInheritedCoreError(errorMsg) Then
                rootErrorIds(idVal) = True
            End If

        Else
            outErrFlag(r, 1) = ""
        End If

NextRow:
    Next r

    GanttLive_RemoveDerivedLOERootErrors dataCore, mapCore, errorIds, rootErrorIds

    GanttLive_ApplyActualImpactWarnings idToRowTest, warningActualIds, outWarnFlag, outWarnText

    tblTest.ListColumns("Calc Test Start").DataBodyRange.value = outStart
    tblTest.ListColumns("Calc Test Finish").DataBodyRange.value = outFinish
    tblTest.ListColumns("Calc Test Duration").DataBodyRange.value = outDuration
    tblTest.ListColumns("Calc Test Progress").DataBodyRange.value = outProgress
    tblTest.ListColumns("Warning Flag").DataBodyRange.value = outWarnFlag
    tblTest.ListColumns("Warning Text").DataBodyRange.value = outWarnText
    tblTest.ListColumns("Error Flag").DataBodyRange.value = outErrFlag

    FormatCalcGanttTestColumns tblTest

    If errorIds.Count > 0 Then
        If rootErrorIds.Count > 0 Then
            CalcBridge_AppendCoreErrorMessagesFromData consoleMessages, dataCore, mapCore, rootErrorIds, "TEST", dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics
        Else
            ShowGanttLiveGroupedMessage errorIds, idToWbs, _
                "Erreur de calcul dans le moteur live", "corriger les valeurs test ou la logique amont", _
                "Calculation error in live engine", "fix test values or upstream logic", vbCritical, consoleMessages
        End If

        GanttTestAnalyticsSnapshot_Clear
        Exit Function
    End If

    Set analyticsById = CalcBridge_BuildAnalyticsSnapshotById(dataCore, mapCore, linksBySuccId)
    Set resultById = GanttSimulation_BuildResultByIdMap()
    GanttTestAnalyticsSnapshot_Set resultById, analyticsById

    Run_Gantt_Test_Backend = True
    Exit Function

SafeExit:
    GanttTestAnalyticsSnapshot_Clear
    If Err.Number <> 0 Then
        GanttLive_AddVbaOrStructuredError consoleMessages, "Run_Gantt_Test_Backend", Err.Description
    End If

End Function
'------------------------------------------------------------------------------
' FR: Actualise Apply Test Render sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Test Render without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub GanttLive_ApplyTestRender(ByVal wsGantt As Worksheet)

    Dim fastRenderApplied As Boolean

    SetGanttPreserveTestInputs True
    GanttLive_RequestTestRender
    fastRenderApplied = Gantt_TryApplyTestDayPredictiveRegistry()
    If Not fastRenderApplied Then Refresh_Gantt
    GanttLive_SetActiveSimulationMode "TEST"
    SetGanttPreserveTestInputs False

    If Not wsGantt Is Nothing Then
        wsGantt.Activate
    End If

End Sub
'------------------------------------------------------------------------------
' FR:
' Construit les inputs Core d'une ligne TEST selon la priorite metier
' Test > Actual > Forecast > Baseline pour les dates, et TEST > base pour progress.
'
' EN:
' Builds one TEST row's Core inputs using the business priority
' Test > Actual > Forecast > Baseline for dates, and TEST > base for progress.
'
' Entrees / Inputs:
' - Dates baseline/actual/forecast, valeurs TEST et progress base/test.
'
' Sorties / Outputs:
' - inputStart, inputFinish, inputDuration et inputProgress passes par reference.
'
' Appele par / Called by:
' - Run_Gantt_Test_Engine pendant la construction de tbl_CALC_GANTT_TEST.
'
' Notes:
' - Si Test Start et Test Finish existent, la duree est derivee de l'intervalle saisi.
'------------------------------------------------------------------------------
Private Sub BuildTestInputValues_FromWBSSource( _
    ByVal baselineStart As Variant, _
    ByVal baselineDuration As Variant, _
    ByVal actualStart As Variant, _
    ByVal actualFinish As Variant, _
    ByVal forecastStart As Variant, _
    ByVal forecastFinish As Variant, _
    ByVal baseProgress As Variant, _
    ByVal testStart As Variant, _
    ByVal testFinish As Variant, _
    ByVal testProgress As Variant, _
    ByRef inputStart As Variant, _
    ByRef inputFinish As Variant, _
    ByRef inputDuration As Variant, _
    ByRef inputProgress As Variant)

    inputStart = Empty
    inputFinish = Empty
    inputDuration = baselineDuration
    inputProgress = baseProgress

    '----------------------------------------
    ' START candidate
    ' Priorité stricte : Test > Actual > Forecast > Baseline
    '----------------------------------------
    If HasValue(testStart) Then
        inputStart = testStart
    ElseIf HasValue(actualStart) Then
        inputStart = actualStart
    ElseIf HasValue(forecastStart) Then
        inputStart = forecastStart
    ElseIf HasValue(baselineStart) Then
        inputStart = baselineStart
    End If

    '----------------------------------------
    ' FINISH candidate
    ' Priorité stricte : Test > Actual > Forecast
    '----------------------------------------
    If HasValue(testFinish) Then
        inputFinish = testFinish
    ElseIf HasValue(actualFinish) Then
        inputFinish = actualFinish
    ElseIf HasValue(forecastFinish) Then
        inputFinish = forecastFinish
    End If

    '----------------------------------------
    ' Cas explicite Test Start + Test Finish :
    ' on dérive la durée du test saisi
    '----------------------------------------
    If HasValue(testStart) And HasValue(testFinish) Then
        inputDuration = CDbl(testFinish) - CDbl(testStart) + 1
    End If

    '----------------------------------------
    ' Progress
    '----------------------------------------
    If HasValue(testProgress) Then
        inputProgress = testProgress
    End If

End Sub



'------------------------------------------------------------------------------
' FR:
' Convertit tbl_CALC_GANTT_TEST en tableau compatible Run_Calc_Core pour TEST,
' en fusionnant saisies TEST, donnees WBS brutes et contraintes CALC.
'
' EN:
' Converts tbl_CALC_GANTT_TEST into a Run_Calc_Core-compatible TEST array,
' merging TEST inputs, raw WBS data, and CALC constraints.
'
' Entrees / Inputs:
' - tbl_CALC_GANTT_TEST, raw WBS map, contraintes CALC, maps ID/row/WBS.
'
' Sorties / Outputs:
' - dataCore et mapCore pour Run_Calc_Core; warnings actual-task alimentes.
'
' Appele par / Called by:
' - Run_Gantt_Test_Backend.
'
' Notes:
' - C'est le bridge le plus sensible entre UI TEST et Core Calculation.
'------------------------------------------------------------------------------
Private Sub BuildGanttTestCoreDataset( _
    ByVal tblTest As ListObject, _
    ByRef dataCore As Variant, _
    ByRef mapCore As Object, _
    ByVal idToRowTest As Object, _
    ByVal idToWbs As Object, _
    ByRef outWarnFlag() As Variant, _
    ByRef outWarnText() As Variant, _
    ByVal warningActualIds As Object)

    Dim arrTest As Variant
    Dim rowCount As Long
    Dim r As Long

    Dim headers As Variant
    Dim c As Long

    Dim idVal As String
    Dim wbsVal As String
    Dim taskTypeVal As String
    Dim isSummary As String
    Dim parentId As String

    Dim rawById As Object
    Dim calcConstraintById As Object
    Dim calcConstraintValues As Variant

    Dim rawActualStart As Variant
    Dim rawActualFinish As Variant
    Dim rawForecastStart As Variant
    Dim rawForecastFinish As Variant
    Dim rawBaselineStart As Variant
    Dim rawBaselineDuration As Variant
    Dim rawCal As String

    Dim taskName As String
    Dim constraintActive As String
    Dim startConstraintType As String
    Dim startConstraintDate As Variant
    Dim finishConstraintType As String
    Dim finishConstraintDate As Variant

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgress As Variant

    Dim actualStart As Variant
    Dim actualFinish As Variant
    Dim forecastStart As Variant
    Dim forecastFinish As Variant
    Dim baselineStart As Variant
    Dim baselineDuration As Variant
    Dim inputProgress As Variant

    Dim hasActual As Boolean
    Dim isSummaryBool As Boolean
    Dim rowHasMeaningfulTest As Boolean

    headers = Array( _
        "ID", "WBS", "Task Name", "Task Type", "Cal", "ParentID", "IsSummary", _
        "Actual Start", "Actual Finish", "Forecast Start", "Forecast Finish", _
        "Baseline Start", "Baseline Duration", "Constraint Active", _
        "Start Constraint Type", "Start Constraint Date", _
        "Finish Constraint Type", "Finish Constraint Date", _
        "Calculated Start", "Calculated Finish", "Calculated Duration", _
        "Error flag", "ErrorMsg", "Input Progress" _
    )

    Set mapCore = CreateObject("Scripting.Dictionary")
    For c = LBound(headers) To UBound(headers)
        mapCore(CStr(headers(c))) = c + 1
    Next c

    arrTest = tblTest.DataBodyRange.value
    rowCount = UBound(arrTest, 1)

    ReDim dataCore(1 To rowCount, 1 To UBound(headers) + 1)

    Set rawById = BuildRawWbsSourceById_Live()
    Set calcConstraintById = BuildCalcConstraintByIdMap_GanttLive()

    For r = 1 To rowCount

        idVal = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "ID"))))
        wbsVal = NormalizeWBS(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "WBS"))))
        taskTypeVal = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Task Type"))))
        isSummary = UCase$(Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Is Summary")))))
        parentId = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Parent ID"))))

        idToRowTest(idVal) = r
        idToWbs(idVal) = wbsVal

        rawActualStart = Empty
        rawActualFinish = Empty
        rawForecastStart = Empty
        rawForecastFinish = Empty
        rawBaselineStart = Empty
        rawBaselineDuration = Empty
        rawCal = CALENDAR_7D

        taskName = vbNullString
        constraintActive = "No"
        startConstraintType = vbNullString
        startConstraintDate = Empty
        finishConstraintType = vbNullString
        finishConstraintDate = Empty

        If rawById.Exists(idVal) Then
            rawActualStart = rawById(idVal)(0)
            rawActualFinish = rawById(idVal)(1)
            rawForecastStart = rawById(idVal)(2)
            rawForecastFinish = rawById(idVal)(3)
            rawBaselineStart = rawById(idVal)(4)
            rawBaselineDuration = rawById(idVal)(5)
            rawCal = CStr(rawById(idVal)(7))
        End If

        If calcConstraintById.Exists(idVal) Then
            calcConstraintValues = calcConstraintById(idVal)
            taskName = Trim$(CStr(calcConstraintValues(0)))
            constraintActive = Trim$(CStr(calcConstraintValues(1)))
            startConstraintType = Trim$(CStr(calcConstraintValues(2)))
            startConstraintDate = calcConstraintValues(3)
            finishConstraintType = Trim$(CStr(calcConstraintValues(4)))
            finishConstraintDate = calcConstraintValues(5)
        End If

        testStart = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test Start")))
        testFinish = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test Finish")))
        testProgress = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test % Normalized")))

        hasActual = (HasValue(rawActualStart) Or HasValue(rawActualFinish))
        isSummaryBool = (isSummary = "YES")
        rowHasMeaningfulTest = (Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Any Test Value")))) = "YES")

        outWarnFlag(r, 1) = ""
        outWarnText(r, 1) = ""

        actualStart = rawActualStart
        actualFinish = rawActualFinish
        forecastStart = rawForecastStart
        forecastFinish = rawForecastFinish
        baselineStart = rawBaselineStart
        baselineDuration = rawBaselineDuration
        inputProgress = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Base Progress")))

        If HasValue(testProgress) Then
            inputProgress = testProgress
        End If

        If rowHasMeaningfulTest And Not isSummaryBool Then

            If HasValue(testStart) Then forecastStart = testStart
            If HasValue(testFinish) Then forecastFinish = testFinish

            If HasValue(testStart) And HasValue(testFinish) Then
                baselineStart = testStart
                baselineDuration = CDbl(testFinish) - CDbl(testStart) + 1
            End If

            If hasActual Then
                warningActualIds(idVal) = True
                outWarnFlag(r, 1) = "WARNING"
                outWarnText(r, 1) = "TEST INPUT ON ACTUAL TASK - SIMULATION MAY SHOW NO EFFECT AND LOCK MAY NOT MATCH PROD"
            End If

        End If

        dataCore(r, mapCore("ID")) = idVal
        dataCore(r, mapCore("WBS")) = wbsVal
        dataCore(r, mapCore("Task Name")) = taskName
        dataCore(r, mapCore("Task Type")) = taskTypeVal
        dataCore(r, mapCore("Cal")) = NormalizeCalendarType(rawCal)
        dataCore(r, mapCore("ParentID")) = parentId
        dataCore(r, mapCore("IsSummary")) = isSummaryBool
        dataCore(r, mapCore("Actual Start")) = actualStart
        dataCore(r, mapCore("Actual Finish")) = actualFinish
        dataCore(r, mapCore("Forecast Start")) = forecastStart
        dataCore(r, mapCore("Forecast Finish")) = forecastFinish
        dataCore(r, mapCore("Baseline Start")) = baselineStart
        dataCore(r, mapCore("Baseline Duration")) = baselineDuration
        dataCore(r, mapCore("Constraint Active")) = constraintActive
        dataCore(r, mapCore("Start Constraint Type")) = startConstraintType
        dataCore(r, mapCore("Start Constraint Date")) = startConstraintDate
        dataCore(r, mapCore("Finish Constraint Type")) = finishConstraintType
        dataCore(r, mapCore("Finish Constraint Date")) = finishConstraintDate
        dataCore(r, mapCore("Calculated Start")) = Empty
        dataCore(r, mapCore("Calculated Finish")) = Empty
        dataCore(r, mapCore("Calculated Duration")) = Empty
        dataCore(r, mapCore("Error flag")) = ""
        dataCore(r, mapCore("ErrorMsg")) = ""
        dataCore(r, mapCore("Input Progress")) = inputProgress

    Next r

End Sub



'------------------------------------------------------------------------------
' FR: Retourne les IDs non-summary dont les valeurs TEST changent reellement le rendu de base.
' EN: Returns non-summary IDs whose TEST values actually change the base rendering.
'------------------------------------------------------------------------------
Private Function GanttLive_BuildMeaningfulChangedTestIds(ByVal tblTest As ListObject) As Object

    Dim d As Object
    Dim arr As Variant
    Dim mapTest As Object
    Dim r As Long
    Dim idVal As String

    Dim baseStart As Variant
    Dim baseFinish As Variant
    Dim baseProgress As Variant

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgress As Variant

    Set d = CreateObject("Scripting.Dictionary")

    If tblTest Is Nothing Then
        Set GanttLive_BuildMeaningfulChangedTestIds = d
        Exit Function
    End If

    If tblTest.DataBodyRange Is Nothing Then
        Set GanttLive_BuildMeaningfulChangedTestIds = d
        Exit Function
    End If

    Set mapTest = CanonicalIdentity_BuildColumnMap(tblTest)
    arr = tblTest.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapTest("ID"))))
        If idVal = "" Then GoTo NextRow

        If UCase$(Trim$(CStr(arr(r, mapTest("Is Summary"))))) = "YES" Then GoTo NextRow
        If Trim$(CStr(arr(r, mapTest("Any Test Value")))) <> "YES" Then GoTo NextRow

        baseStart = GetCellValue(arr(r, mapTest("Base Start")))
        baseFinish = GetCellValue(arr(r, mapTest("Base Finish")))
        baseProgress = GetCellValue(arr(r, mapTest("Base Progress")))

        testStart = GetCellValue(arr(r, mapTest("Test Start")))
        testFinish = GetCellValue(arr(r, mapTest("Test Finish")))
        testProgress = GetCellValue(arr(r, mapTest("Test % Normalized")))

        If HasMeaningfulLockDelta(baseStart, baseFinish, baseProgress, testStart, testFinish, testProgress) Then
            d(idVal) = True
        End If

NextRow:
    Next r

    Set GanttLive_BuildMeaningfulChangedTestIds = d

End Function



'------------------------------------------------------------------------------
' FR: Parcourt l'aval reseau depuis les TEST modifies pour detecter les taches avec Actual impactees.
' EN: Walks downstream from changed TEST IDs to detect impacted tasks that already have Actuals.
'------------------------------------------------------------------------------
Private Function GanttLive_FindActualTasksImpactedByTestChanges( _
    ByVal tblTest As ListObject, _
    ByVal changedTestIds As Object, _
    ByVal linksBySuccId As Object) As Object

    Dim impactedActualIds As Object
    Dim childrenByPred As Object
    Dim queue As Collection
    Dim visited As Object

    Dim mapTest As Object
    Dim arr As Variant
    Dim hasActualById As Object

    Dim r As Long
    Dim idVal As String
    Dim startId As Variant
    Dim currentId As String
    Dim childId As Variant

    Set impactedActualIds = CreateObject("Scripting.Dictionary")

    If tblTest Is Nothing Then
        Set GanttLive_FindActualTasksImpactedByTestChanges = impactedActualIds
        Exit Function
    End If

    If changedTestIds Is Nothing Then
        Set GanttLive_FindActualTasksImpactedByTestChanges = impactedActualIds
        Exit Function
    End If

    If changedTestIds.Count = 0 Then
        Set GanttLive_FindActualTasksImpactedByTestChanges = impactedActualIds
        Exit Function
    End If

    Set mapTest = CanonicalIdentity_BuildColumnMap(tblTest)
    arr = tblTest.DataBodyRange.value
    Set hasActualById = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapTest("ID"))))
        If idVal <> "" Then
            hasActualById(idVal) = (UCase$(Trim$(CStr(arr(r, mapTest("Has Actual"))))) = "YES")
        End If
    Next r

    Set childrenByPred = GanttLive_BuildChildrenByPred_FromLinks(arr, mapTest, linksBySuccId)

    Set queue = New Collection
    Set visited = CreateObject("Scripting.Dictionary")

    For Each startId In changedTestIds.Keys

        ' IMPORTANT :
        ' la tâche modifiée elle-même doit être incluse si elle a de l’Actual
        If hasActualById.Exists(CStr(startId)) Then
            If hasActualById(CStr(startId)) Then
                impactedActualIds(CStr(startId)) = True
            End If
        End If

        If Not visited.Exists(CStr(startId)) Then
            visited(CStr(startId)) = True
            queue.Add CStr(startId)
        End If

    Next startId

    Do While queue.Count > 0

        currentId = CStr(queue(1))
        queue.Remove 1

        If childrenByPred.Exists(currentId) Then
            For Each childId In childrenByPred(currentId)

                If Not visited.Exists(CStr(childId)) Then
                    visited(CStr(childId)) = True
                    queue.Add CStr(childId)
                End If

                If hasActualById.Exists(CStr(childId)) Then
                    If hasActualById(CStr(childId)) Then
                        impactedActualIds(CStr(childId)) = True
                    End If
                End If

            Next childId
        End If

    Loop

    Set GanttLive_FindActualTasksImpactedByTestChanges = impactedActualIds

End Function



'------------------------------------------------------------------------------
' FR: Construit une adjacency pred -> succ a partir des liens Core pour les warnings d'impact Actual.
' EN: Builds pred -> successor adjacency from Core links for Actual impact warnings.
'------------------------------------------------------------------------------
Private Function GanttLive_BuildChildrenByPred_FromLinks( _
    ByRef arr As Variant, _
    ByVal mapTest As Object, _
    ByVal linksBySuccId As Object) As Object

    Dim childrenByPred As Object
    Dim r As Long
    Dim idVal As String
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String

    Set childrenByPred = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapTest("ID"))))
        If idVal <> "" Then
            Set childrenByPred(idVal) = New Collection
        End If
    Next r

    If linksBySuccId Is Nothing Then
        Set GanttLive_BuildChildrenByPred_FromLinks = childrenByPred
        Exit Function
    End If

    For Each succId In linksBySuccId.Keys
        For Each oneLink In linksBySuccId(CStr(succId))
            predId = Core_GetLinkPredId(oneLink)
            If predId <> "" Then
                If Not childrenByPred.Exists(predId) Then
                    Set childrenByPred(predId) = New Collection
                End If
                childrenByPred(predId).Add CStr(succId)
            End If
        Next oneLink
    Next succId

    Set GanttLive_BuildChildrenByPred_FromLinks = childrenByPred

End Function



'------------------------------------------------------------------------------
' FR: Inscrit les warnings de taches Actual impactees dans les tableaux de sortie TEST.
' EN: Writes impacted-Actual warnings into TEST output arrays.
'------------------------------------------------------------------------------
Private Sub GanttLive_ApplyActualImpactWarnings( _
    ByVal idToRowTest As Object, _
    ByVal warningActualIds As Object, _
    ByRef outWarnFlag() As Variant, _
    ByRef outWarnText() As Variant)

    Dim idVal As Variant
    Dim rowIndex As Long
    Dim warnText As String

    If warningActualIds Is Nothing Then Exit Sub
    If warningActualIds.Count = 0 Then Exit Sub

    warnText = "TEST INPUT IMPACTS ACTUAL TASK - LOCK MAY NOT FULLY PERSIST"

    For Each idVal In warningActualIds.Keys

        If idToRowTest.Exists(CStr(idVal)) Then
            rowIndex = CLng(idToRowTest(CStr(idVal)))
            outWarnFlag(rowIndex, 1) = "WARNING"

            If Trim$(CStr(outWarnText(rowIndex, 1))) = "" Then
                outWarnText(rowIndex, 1) = warnText
            ElseIf InStr(1, CStr(outWarnText(rowIndex, 1)), warnText, vbTextCompare) = 0 Then
                outWarnText(rowIndex, 1) = CStr(outWarnText(rowIndex, 1)) & " | " & warnText
            End If
        End If

    Next idVal

End Sub



'------------------------------------------------------------------------------
' FR: Scanne tbl_CALC_GANTT_TEST pour savoir si au moins une ligne non-summary a un delta visible.
' EN: Scans tbl_CALC_GANTT_TEST to see whether at least one non-summary row has a visible delta.
'------------------------------------------------------------------------------
Private Function GanttLive_HasAnyRenderableTestDelta(ByVal tblTest As ListObject) As Boolean

    Dim baseById As Object
    Dim testById As Object
    Dim arr As Variant
    Dim mapTest As Object
    Dim r As Long
    Dim idVal As String

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then Exit Function

    Set baseById = GanttLive_BuildBaseByIdMap()
    Set testById = GanttLive_BuildTestByIdMap()
    Set mapTest = CanonicalIdentity_BuildColumnMap(tblTest)

    arr = tblTest.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapTest("ID"))))
        If idVal = "" Then GoTo NextRow

        If UCase$(Trim$(CStr(arr(r, mapTest("Is Summary"))))) = "YES" Then GoTo NextRow

        If GanttLive_HasRenderableTestDelta(idVal, baseById, testById) Then
            GanttLive_HasAnyRenderableTestDelta = True
            Exit Function
        End If

NextRow:
    Next r

End Function
