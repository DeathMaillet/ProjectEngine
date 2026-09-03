Attribute VB_Name = "mod_GanttScenarioService"
Option Explicit

'===============================================================================
' MODULE : mod_GanttScenarioService
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
' CONTRATS / CONTRACTS : GanttScenarioService_RunScenarioEngine
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
' FR: Orchestre Gantt Scenario Service Run Scenario Engine en preservant l'ordre contractuel des etapes du domaine.
' EN: Orchestrates Gantt Scenario Service Run Scenario Engine while preserving the domain's contractual step order.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Public Sub GanttScenarioService_RunScenarioEngine( _
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
    Dim calcRowById As Object
    Dim ganttRowByWbs As Object

    Dim dataWBS As Variant
    Dim dataCalc As Variant
    Dim outArr() As Variant

    Dim rowCount As Long
    Dim calcRowCount As Long
    Dim r As Long
    Dim calcRow As Long
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
    Dim consoleMessages As Collection
    Dim scenarioSucceeded As Boolean
    Dim ganttRebuilt As Boolean

    On Error GoTo SafeExit

    AppEvents_EnsureInitialized
    Set consoleMessages = New Collection

    Set wsGantt = ThisWorkbook.Worksheets(GANTT_SHEET)

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)

    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set wsTest = Ensure_CalcGanttTest_Sheet()
    Set tblTest = Ensure_CalcGanttTest_Table(wsTest)

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set mapCalcDriving = CanonicalIdentity_GetDrivingLogicByIdMap()
    Set wbsToId = CanonicalIdentity_GetWbsToIdMap()
    Set calcById = BuildCalcByIdMap_Live(tblCalc, mapCalc)
    Set calcRowById = CreateObject("Scripting.Dictionary")
    Set ganttRowByWbs = BuildGanttRowByWbsMap(wsGantt)

    ValidateGanttTestSourceColumns mapWBS, mapCalc
    ValidateCalcGanttTestColumns tblTest

    If Not mapCalc.Exists("Calculated Start") Then
        Err.Raise vbObjectError + 9101, "Run_Gantt_Scenario_Engine", "Missing CALC column: Calculated Start"
    End If

    If Not mapCalc.Exists("Calculated Finish") Then
        Err.Raise vbObjectError + 9102, "Run_Gantt_Scenario_Engine", "Missing CALC column: Calculated Finish"
    End If

    dataWBS = tblWBS.DataBodyRange.value
    dataCalc = tblCalc.DataBodyRange.value

    rowCount = UBound(dataWBS, 1)
    calcRowCount = UBound(dataCalc, 1)

    For r = 1 To calcRowCount
        idVal = Trim$(CStr(dataCalc(r, mapCalc("ID"))))
        If idVal <> "" Then
            If Not calcRowById.Exists(idVal) Then
                calcRowById(idVal) = r
            End If
        End If
    Next r

    ResizeTableToRowCount_Generic tblTest, rowCount
    ReDim outArr(1 To rowCount, 1 To tblTest.ListColumns.Count)

    testRow = 0

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataWBS(r, mapWBS(VTS_COL_ID))))
        wbsVal = NormalizeWBS(CStr(dataWBS(r, mapWBS(VTS_COL_WBS))))
        taskTypeVal = Trim$(CStr(dataWBS(r, mapWBS(VTS_COL_TASK_TYPE))))

        testRow = testRow + 1

        baseStart = Empty
        baseFinish = Empty
        baseDuration = Empty
        baseProgress = GetCellValue(dataWBS(r, mapWBS(VTS_COL_PROGRESS_PERCENT)))
        drivingLogic = ""

        If calcRowById.Exists(idVal) Then

            calcRow = CLng(calcRowById(idVal))

            baseStart = GetCellValue(dataCalc(calcRow, mapCalc("Calculated Start")))
            baseFinish = GetCellValue(dataCalc(calcRow, mapCalc("Calculated Finish")))

            If HasValue(baseStart) And HasValue(baseFinish) Then
                baseDuration = CDbl(baseFinish) - CDbl(baseStart) + 1
            Else
                baseDuration = Empty
            End If

            If mapCalcDriving.Exists(idVal) Then
                drivingLogic = CStr(mapCalcDriving(idVal))
            ElseIf mapCalc.Exists("Driving Logic") Then
                drivingLogic = CStr(dataCalc(calcRow, mapCalc("Driving Logic")))
            End If

        End If

        isSummary = (UCase$(Trim$(drivingLogic)) = "SUMMARY")
        hasActual = False

        testStart = Empty
        testFinish = Empty
        testProgressRaw = Empty
        testProgressNorm = Empty
        inputStart = Empty
        inputFinish = Empty
        inputDuration = Empty
        inputProgress = Empty
        anyTestValue = False

        ganttRow = 0
        If ganttRowByWbs.Exists(wbsVal) Then ganttRow = CLng(ganttRowByWbs(wbsVal))

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

        inputStart = baseStart
        inputFinish = Empty
        inputDuration = baseDuration
        inputProgress = baseProgress

        If HasValue(testStart) Then inputStart = testStart
        If HasValue(testFinish) Then inputFinish = testFinish

        If HasValue(testStart) And HasValue(testFinish) Then
            inputDuration = CDbl(testFinish) - CDbl(testStart) + 1
        End If

        If HasValue(testProgressNorm) Then
            inputProgress = testProgressNorm
        End If

        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "ID")) = idVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "WBS")) = wbsVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Task Type")) = taskTypeVal
        outArr(testRow, GetColumnIndex_GanttLive(tblTest, "Cal")) = NormalizeCalendarType(dataWBS(r, mapWBS(VTS_COL_CAL)))
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

    If Not Run_Gantt_Scenario_Backend(tblTest, consoleMessages) Then
        GanttLive_AbortTestEngine wsGantt
        If silentMode Then
            If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Scenario_Engine"
        Else
            CalcBridge_ShowPlanningConsole consoleMessages
        End If
        GoTo CleanExit
    End If

    GanttLive_ApplyScenarioRender wsGantt
    scenarioSucceeded = True
    ganttRebuilt = True

    If silentMode Then
        If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Scenario_Engine"
        GoTo CleanExit
    End If

    GanttLive_AddBiConsoleMessage consoleMessages, "INFO", _
        "Scénario mis à jour.", _
        "Scenario updated."

    CalcBridge_ShowPlanningConsole consoleMessages

CleanExit:
    If Not IsMissing(transactionSucceeded) Then transactionSucceeded = scenarioSucceeded
    If Not IsMissing(transactionGanttRebuilt) Then transactionGanttRebuilt = ganttRebuilt
    If Not IsMissing(transactionMessages) Then Set transactionMessages = consoleMessages
    Exit Sub

SafeExit:
    GanttLive_SafeExit_TestEngine

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection

        GanttLive_AddVbaOrStructuredError consoleMessages, "Run_Gantt_Scenario_Engine", Err.Description

        If silentMode Then
            If recordSilentMessages Then CalcBridge_RecordPlanningMessages consoleMessages, "Run_Gantt_Scenario_Engine"
        Else
            CalcBridge_ShowPlanningConsole consoleMessages
        End If
    End If

    Resume CleanExit

End Sub




'------------------------------------------------------------------------------
' FR:
' Execute le Core sur le dataset SCENARIO, reporte les resultats dans
' tbl_CALC_GANTT_TEST et convertit les erreurs Core en messages scenario.
'
' EN:
' Runs Core on the SCENARIO dataset, writes results back to tbl_CALC_GANTT_TEST,
' and converts Core errors into scenario messages.
'
' Entrees / Inputs:
' - tbl_CALC_GANTT_TEST preparee par Run_Gantt_Scenario_Engine.
'
' Sorties / Outputs:
' - Colonnes Calc Test, Warning/Error Flag et messages scenario.
'
' Appele par / Called by:
' - Run_Gantt_Scenario_Engine.
'
' Notes:
' - Frontiere SCENARIO vers Core Calculation; pas d'ecriture durable WBS.
'------------------------------------------------------------------------------
Private Function Run_Gantt_Scenario_Backend( _
    ByVal tblTest As ListObject, _
    Optional ByVal consoleMessages As Variant) As Boolean

    Dim dataCore As Variant
    Dim mapCore As Object
    Dim linksBySuccId As Object
    Dim executionNetwork As clsCompiledExecutionNetwork

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

    Dim rowCount As Long
    Dim r As Long
    Dim idVal As String
    Dim errorMsg As String

    On Error GoTo SafeExit

    Run_Gantt_Scenario_Backend = False

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then
        Run_Gantt_Scenario_Backend = True
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

    BuildGanttScenarioCoreDataset tblTest, dataCore, mapCore, idToRowTest, idToWbs

    Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded( _
        ThisWorkbook.Worksheets(CALC_SHEET).ListObjects(CALC_TABLE))
    Set executionNetwork = CompileExecutionNetwork(dataCore, mapCore, linksBySuccId)

    Run_Calc_Core dataCore, mapCore, linksBySuccId, , dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics, executionNetwork

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
            CalcBridge_AppendCoreErrorMessagesFromData consoleMessages, dataCore, mapCore, rootErrorIds, "SCENARIO", dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics
        Else
            ShowGanttLiveGroupedMessage errorIds, idToWbs, _
                "Erreur de calcul dans le scénario", "corriger les valeurs de test ou la logique amont", _
                "Calculation error in scenario", "fix test values or upstream logic", vbCritical, consoleMessages
        End If

        Exit Function
    End If

    Run_Gantt_Scenario_Backend = True
    Exit Function

SafeExit:
    If Err.Number <> 0 Then
        GanttLive_AddVbaOrStructuredError consoleMessages, "Run_Gantt_Scenario_Backend", Err.Description
    End If

End Function



'------------------------------------------------------------------------------
' FR:
' Convertit la table scenario live en dataset Core sans Actual, Forecast ni contraintes,
' avec les dates calculees courantes comme baseline de simulation.
'
' EN:
' Converts the live scenario table into a Core dataset without Actual, Forecast,
' or constraints, using current calculated dates as the simulation baseline.
'
' Entrees / Inputs:
' - tbl_CALC_GANTT_TEST et noms de taches WBS par ID.
'
' Sorties / Outputs:
' - dataCore/mapCore prets pour Run_Calc_Core et maps ID/row/WBS alimentees.
'
' Appele par / Called by:
' - Run_Gantt_Scenario_Backend.
'
' Notes:
' - Le scenario est volontairement deconnecte des contraintes pour construire une nouvelle baseline possible.
'------------------------------------------------------------------------------
Private Sub BuildGanttScenarioCoreDataset( _
    ByVal tblTest As ListObject, _
    ByRef dataCore As Variant, _
    ByRef mapCore As Object, _
    ByVal idToRowTest As Object, _
    ByVal idToWbs As Object)

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
    Dim taskName As String
    Dim taskNameById As Object

    Dim baseStart As Variant
    Dim baseDuration As Variant
    Dim baseProgress As Variant

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgress As Variant

    Dim inputStart As Variant
    Dim inputDuration As Variant
    Dim inputProgress As Variant

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

    Set taskNameById = BuildTaskNameByIdFromWbs_Live()

    For r = 1 To rowCount

        idVal = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "ID"))))
        wbsVal = NormalizeWBS(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "WBS"))))
        taskTypeVal = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Task Type"))))
        isSummary = UCase$(Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Is Summary")))))
        parentId = Trim$(CStr(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Parent ID"))))
        taskName = vbNullString
        If taskNameById.Exists(idVal) Then taskName = Trim$(CStr(taskNameById(idVal)))

        idToRowTest(idVal) = r
        idToWbs(idVal) = wbsVal

        baseStart = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Base Start")))
        baseDuration = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Base Duration")))
        baseProgress = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Base Progress")))

        testStart = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test Start")))
        testFinish = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test Finish")))
        testProgress = GetCellValue(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Test % Normalized")))

        inputStart = baseStart
        inputDuration = baseDuration
        inputProgress = baseProgress

        If HasValue(testStart) Then
            inputStart = testStart
        End If

        If HasValue(testFinish) Then
            If HasValue(testStart) Then
                inputDuration = CDbl(testFinish) - CDbl(testStart) + 1
            ElseIf HasValue(inputDuration) Then
                inputStart = SubtractWorkingDays(testFinish, inputDuration, arrTest(r, GetColumnIndex_GanttLive(tblTest, "Cal")))
            End If
        End If

        If HasValue(testProgress) Then
            inputProgress = testProgress
        End If

        dataCore(r, mapCore("ID")) = idVal
        dataCore(r, mapCore("WBS")) = wbsVal
        dataCore(r, mapCore("Task Name")) = taskName
        dataCore(r, mapCore("Task Type")) = taskTypeVal
        dataCore(r, mapCore("Cal")) = NormalizeCalendarType(arrTest(r, GetColumnIndex_GanttLive(tblTest, "Cal")))
        dataCore(r, mapCore("ParentID")) = parentId
        dataCore(r, mapCore("IsSummary")) = IIf(isSummary = "YES", True, False)

        dataCore(r, mapCore("Actual Start")) = Empty
        dataCore(r, mapCore("Actual Finish")) = Empty
        dataCore(r, mapCore("Forecast Start")) = Empty
        dataCore(r, mapCore("Forecast Finish")) = Empty

        dataCore(r, mapCore("Baseline Start")) = inputStart
        dataCore(r, mapCore("Baseline Duration")) = inputDuration

        dataCore(r, mapCore("Constraint Active")) = "No"
        dataCore(r, mapCore("Start Constraint Type")) = vbNullString
        dataCore(r, mapCore("Start Constraint Date")) = Empty
        dataCore(r, mapCore("Finish Constraint Type")) = vbNullString
        dataCore(r, mapCore("Finish Constraint Date")) = Empty

        dataCore(r, mapCore("Calculated Start")) = Empty
        dataCore(r, mapCore("Calculated Finish")) = Empty
        dataCore(r, mapCore("Calculated Duration")) = Empty
        dataCore(r, mapCore("Error flag")) = ""
        dataCore(r, mapCore("ErrorMsg")) = ""
        dataCore(r, mapCore("Input Progress")) = inputProgress

    Next r

End Sub


'------------------------------------------------------------------------------
' FR: Actualise Apply Scenario Render sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Scenario Render without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub GanttLive_ApplyScenarioRender(ByVal wsGantt As Worksheet)

    SetGanttPreserveTestInputs True
    GanttLive_RequestScenarioRender
    If Not EnsureGanttForCurrentPlanning(GANTT_ENSURE_LOCAL_UPDATE, "GanttLive_ApplyScenarioRender") Then
        Err.Raise 5, "GanttLive_ApplyScenarioRender", "Gantt SCENARIO render did not reach READY state."
    End If
    GanttLive_SetActiveSimulationMode "SCENARIO"
    SetGanttPreserveTestInputs False

    If Not wsGantt Is Nothing Then
        wsGantt.Activate
    End If

End Sub


