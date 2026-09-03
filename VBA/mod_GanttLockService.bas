Attribute VB_Name = "mod_GanttLockService"
Option Explicit

'===============================================================================
' MODULE : mod_GanttLockService
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
' CONTRATS / CONTRACTS : GanttLockService_RunLockChanges
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"


Private Const GANTT_FIRST_TASK_ROW As Long = 4
Private Const GANTT_COL_WBS As Long = 1

Private Const COL_TEST_START As Long = 5
Private Const COL_TEST_FINISH As Long = 6
Private Const COL_TEST_PROGRESS As Long = 9

'------------------------------------------------------------------------------
' FR: Orchestre Gantt Lock Service Run Lock Changes en preservant l'ordre contractuel des etapes du domaine.
' EN: Orchestrates Gantt Lock Service Run Lock Changes while preserving the domain's contractual step order.
'------------------------------------------------------------------------------

Public Sub GanttLockService_RunLockChanges()

    Dim wsWBS As Worksheet
    Dim wsGantt As Worksheet
    Dim wsCalc As Worksheet

    Dim tblWBS As ListObject
    Dim tblCalc As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object

    Dim wbsBackup As Object
    Dim ganttTestBackup As Object
    Dim appliedChanges As Object
    Dim simulatedById As Object

    Dim hasCalcErrors As Boolean
    Dim lockMatches As Boolean
    Dim consoleMessages As Collection
    Dim workflowStarted As Boolean

    On Error GoTo SafeExit

    GanttLockTrace_Log "Enter Run_Gantt_Lock_Changes"
    workflowStarted = EnsurePlanningWorkflowStarted("Run_Gantt_Lock_Changes")
    GanttLockTrace_Log "Workflow guard started"
    AppEvents_EnsureInitialized
    GanttLockTrace_Log "App events initialized"
    Set consoleMessages = New Collection
    GanttLockTrace_Log "Console message collection created"

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set wsGantt = ThisWorkbook.Worksheets(GANTT_SHEET)
    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)
    GanttLockTrace_Log "Worksheets resolved"

    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)
    GanttLockTrace_Log "WBS and CALC tables resolved"

    If tblWBS.DataBodyRange Is Nothing Then
        GanttLockTrace_Log "Exit: tbl_WBS has no data body"
        GoTo CleanExit
    End If
    If tblCalc.DataBodyRange Is Nothing Then
        GanttLockTrace_Log "Exit: tbl_CALC has no data body"
        GoTo CleanExit
    End If

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    GanttLockTrace_Log "Column maps built"

    GanttLockTrace_Log "ValidateLockSourceColumns start"
    ValidateLockSourceColumns mapWBS, mapCalc
    GanttLockTrace_Log "ValidateLockSourceColumns ok"

    If GanttLive_IsScenarioActive() Then
        GanttLockTrace_Log "Scenario active: CreateScenarioPlanningFromCurrentScenario start"
        CreateScenarioPlanningFromCurrentScenario
        GanttLockTrace_Log "Scenario active: CreateScenarioPlanningFromCurrentScenario returned"
        GoTo CleanExit
    End If

    GanttLockTrace_Log "BuildModifiedTestChangesMap initial start"
    Set appliedChanges = BuildModifiedTestChangesMap(wsGantt, tblWBS, mapWBS)
    GanttLockTrace_Log "BuildModifiedTestChangesMap initial count=" & CStr(appliedChanges.Count)

    If appliedChanges.Count = 0 Then
        GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
            "Aucune modification test à verrouiller.", _
            "No test changes to lock."

        GanttLockTrace_Log "No changes: before CalcBridge_ShowPlanningConsole"
        GanttLockTrace_ShowConsole consoleMessages, "LOCK"
        GanttLockTrace_Log "No changes: after CalcBridge_ShowPlanningConsole"
        GoTo CleanExit
    End If

    GanttLockTrace_Log "Run_Gantt_Test_Engine True start"
    Run_Gantt_Test_Engine True
    GanttLockTrace_Log "Run_Gantt_Test_Engine True returned"
    GanttLockTrace_Log "CALC_GANTT_TEST table resolved"

    GanttLockTrace_Log "GanttLive_CalcGanttTestHasErrors start"
    If GanttSimulation_HasErrors() Then
        GanttLockTrace_Log "GanttLive_CalcGanttTestHasErrors returned True"
        GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
            "Lock annulé : la simulation TEST préalable contient des erreurs." & vbCrLf & _
            "-> corriger les valeurs test ou la logique amont avant de verrouiller.", _
            "Lock cancelled: the preliminary TEST simulation contains errors." & vbCrLf & _
            "-> fix test values or upstream logic before locking."

        GanttLockTrace_Log "TEST errors: before CalcBridge_ShowPlanningConsole"
        GanttLockTrace_ShowConsole consoleMessages, "LOCK"
        GanttLockTrace_Log "TEST errors: after CalcBridge_ShowPlanningConsole"
        GoTo CleanExit
    End If
    GanttLockTrace_Log "GanttLive_CalcGanttTestHasErrors returned False"

    GanttLockTrace_Log "BuildModifiedTestChangesMap post-test start"
    Set appliedChanges = BuildModifiedTestChangesMap(wsGantt, tblWBS, mapWBS)
    GanttLockTrace_Log "BuildModifiedTestChangesMap post-test count=" & CStr(appliedChanges.Count)

    If appliedChanges.Count = 0 Then
        GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
            "Aucune modification test à verrouiller.", _
            "No test changes to lock."

        GanttLockTrace_Log "No changes post-test: before CalcBridge_ShowPlanningConsole"
        GanttLockTrace_ShowConsole consoleMessages, "LOCK"
        GanttLockTrace_Log "No changes post-test: after CalcBridge_ShowPlanningConsole"
        GoTo CleanExit
    End If

    GanttLockTrace_Log "BuildSimulatedLockResultMap start"
    Set simulatedById = BuildSimulatedLockResultMap(appliedChanges)
    GanttLockTrace_Log "BuildSimulatedLockResultMap count=" & CStr(simulatedById.Count)

    If simulatedById.Count = 0 Then
        GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
            "Lock annulé : aucun résultat simulé exploitable n'a été trouvé après le refresh TEST.", _
            "Lock cancelled: no usable simulated result was found after TEST refresh."

        GanttLockTrace_Log "No simulated results: before CalcBridge_ShowPlanningConsole"
        GanttLockTrace_ShowConsole consoleMessages, "LOCK"
        GanttLockTrace_Log "No simulated results: after CalcBridge_ShowPlanningConsole"
        GoTo CleanExit
    End If

    GanttLockTrace_Log "BackupWBSLockColumns start"
    Set wbsBackup = BackupWBSLockColumns(tblWBS, mapWBS)
    GanttLockTrace_Log "BackupWBSLockColumns count=" & CStr(wbsBackup.Count)
    GanttLockTrace_Log "BackupGanttTestInputs start"
    Set ganttTestBackup = BackupGanttTestInputs(wsGantt)
    GanttLockTrace_Log "BackupGanttTestInputs count=" & CStr(ganttTestBackup.Count)

    GanttLockTrace_Log "ApplyModifiedTestChangesToWBS start"
    ApplyModifiedTestChangesToWBS tblWBS, mapWBS, appliedChanges
    GanttLockTrace_Log "ApplyModifiedTestChangesToWBS returned"

    GanttLockTrace_Log "GanttLive_ClearTestRenderRequest before planning update"
    GanttLive_ClearTestRenderRequest
    GanttLockTrace_Log "Run_Planning_Update start"
    Run_Planning_Update
    GanttLockTrace_Log "Run_Planning_Update returned"
    GanttLockTrace_Log "DrainPlanningWorkflowDeferredDisplayMessages start"
    DrainPlanningWorkflowDeferredDisplayMessages consoleMessages
    GanttLockTrace_Log "DrainPlanningWorkflowDeferredDisplayMessages returned"

    GanttLockTrace_Log "CalcTableHasErrors start"
    hasCalcErrors = CalcTableHasErrors(tblCalc, mapCalc)
    GanttLockTrace_Log "CalcTableHasErrors result=" & CStr(hasCalcErrors)
    GanttLockTrace_Log "LockResultsMatchSimulatedResult start"
    lockMatches = LockResultsMatchSimulatedResult(tblWBS, mapWBS, simulatedById)
    GanttLockTrace_Log "LockResultsMatchSimulatedResult result=" & CStr(lockMatches)

    If hasCalcErrors Or (Not lockMatches) Then

        GanttLockTrace_Log "Rollback branch entered"
        GanttLive_RollbackFailedLock _
            tblWBS, mapWBS, wsGantt, wbsBackup, ganttTestBackup
        GanttLockTrace_Log "Rollback returned"

        If hasCalcErrors Then
            GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
                "Lock annulé : le calcul a détecté des erreurs. Les valeurs WBS d'origine ont été restaurées et les colonnes test ont été conservées.", _
                "Lock cancelled: calculation found errors. Original WBS values were restored and test inputs were preserved."
        Else
            GanttLive_AddBiConsoleMessage consoleMessages, "WARNING", _
                "Lock annulé : le recalcul réel ne correspond pas au résultat simulé retenu. Les valeurs WBS d'origine ont été restaurées et les colonnes test ont été conservées.", _
                "Lock cancelled: the real recalculation does not match the retained simulated result. Original WBS values were restored and test inputs were preserved."
        End If

        GanttLockTrace_Log "Rollback branch: before CalcBridge_ShowPlanningConsole"
        GanttLockTrace_ShowConsole consoleMessages, "LOCK"
        GanttLockTrace_Log "Rollback branch: after CalcBridge_ShowPlanningConsole"
        GoTo CleanExit
    End If

    GanttLockTrace_Log "Finalize branch entered"
    GanttLive_FinalizeSuccessfulLock wsGantt, appliedChanges
    GanttLockTrace_Log "Finalize returned"

    GanttLive_AddBiConsoleMessage consoleMessages, "INFO", _
        "Lock appliqué avec succès.", _
        "Lock successfully applied."

    GanttLockTrace_Log "Success branch: before CalcBridge_ShowPlanningConsole"
    GanttLockTrace_ShowConsole consoleMessages, "LOCK"
    GanttLockTrace_Log "Success branch: after CalcBridge_ShowPlanningConsole"

CleanExit:
    GanttLockTrace_Log "CleanExit reached"
    If workflowStarted Then
        GanttLockTrace_Log "EndPlanningWorkflow start"
        EndPlanningWorkflow
        GanttLockTrace_Log "EndPlanningWorkflow returned"
    End If
    GanttLockTrace_Log "Exit Run_Gantt_Lock_Changes"
    Exit Sub

SafeExit:
    GanttLockTrace_Log "SafeExit reached: " & Err.Description
    If consoleMessages Is Nothing Then Set consoleMessages = New Collection

    GanttLive_AddBiConsoleMessage consoleMessages, "STOP", _
        "Erreur VBA dans Run_Gantt_Lock_Changes" & vbCrLf & _
        "-> vérifier le dernier bloc modifié dans mod_GanttLive", _
        "VBA error in Run_Gantt_Lock_Changes" & vbCrLf & _
        "-> check the last edited block in mod_GanttLive"

    GanttLockTrace_Log "SafeExit: before CalcBridge_ShowPlanningConsole"
    GanttLockTrace_ShowConsole consoleMessages, "LOCK"
    GanttLockTrace_Log "SafeExit: after CalcBridge_ShowPlanningConsole"
    Resume CleanExit

End Sub
'=====================================================
' HELPERS - LOCK FLOW
'=====================================================

'------------------------------------------------------------------------------
' FR: Valide Validate Lock Source Columns et applique la politique d'erreur definie par le composant.
' EN: Validates Validate Lock Source Columns and applies the component's defined failure policy.
'------------------------------------------------------------------------------

Private Sub ValidateLockSourceColumns(ByVal mapWBS As Object, ByVal mapCalc As Object)

    Dim reqWBS As Variant
    Dim reqCalc As Variant
    Dim c As Variant

    reqWBS = Array( _
        VTS_COL_ID, _
        VTS_COL_FORECAST_START, _
        VTS_COL_FORECAST_FINISH, _
        VTS_COL_PROGRESS_PERCENT, _
        VTS_COL_CALCULATED_START, _
        VTS_COL_CALCULATED_FINISH _
    )

    reqCalc = Array( _
        "ID", _
        "Error flag" _
    )

    For Each c In reqWBS
        If Not mapWBS.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 980, , "Missing column in tbl_WBS: " & CStr(c)
        End If
    Next c

    For Each c In reqCalc
        If Not mapCalc.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 981, , "Missing column in tbl_CALC: " & CStr(c)
        End If
    Next c

End Sub

'------------------------------------------------------------------------------
' FR: Construit la map Modified Test Changes Map a partir des donnees fournies par l'appelant.
' EN: Builds the Modified Test Changes Map map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildModifiedTestChangesMap( _
    ByVal wsGantt As Worksheet, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object) As Object

    Dim d As Object
    Dim baseById As Object
    Dim wbsToId As Object
    Dim calcDrivingMap As Object

    Dim lastRow As Long
    Dim r As Long

    Dim wbsVal As String
    Dim idVal As String

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgressRaw As Variant
    Dim testProgressNorm As Variant

    Dim hasTestStart As Boolean
    Dim hasTestFinish As Boolean
    Dim hasTestProgress As Boolean

    Dim baseStart As Variant
    Dim baseFinish As Variant
    Dim baseProgress As Variant

    Set d = CreateObject("Scripting.Dictionary")
    Set baseById = GanttLive_BuildBaseByIdMap()
    Set wbsToId = CanonicalIdentity_GetWbsToIdMap()
    Set calcDrivingMap = CanonicalIdentity_GetDrivingLogicByIdMap()

    lastRow = GetLastGanttRow_Live(wsGantt)
    If lastRow < GANTT_FIRST_TASK_ROW Then
        Set BuildModifiedTestChangesMap = d
        Exit Function
    End If

    For r = GANTT_FIRST_TASK_ROW To lastRow

        wbsVal = NormalizeWBS(CStr(wsGantt.Cells(r, GANTT_COL_WBS).value))
        If wbsVal = "" Then GoTo NextRow

        If Not wbsToId.Exists(wbsVal) Then GoTo NextRow
        idVal = CStr(wbsToId(wbsVal))

        If Not calcDrivingMap.Exists(idVal) Then GoTo NextRow
        If UCase$(Trim$(CStr(calcDrivingMap(idVal)))) = "SUMMARY" Then GoTo NextRow

        testStart = GetCellValue(wsGantt.Cells(r, COL_TEST_START).value)
        testFinish = GetCellValue(wsGantt.Cells(r, COL_TEST_FINISH).value)
        testProgressRaw = GetCellValue(wsGantt.Cells(r, COL_TEST_PROGRESS).value)

        hasTestStart = HasValue(testStart)
        hasTestFinish = HasValue(testFinish)
        hasTestProgress = HasValue(testProgressRaw)

        If hasTestProgress Then
            testProgressNorm = NormalizePercentInput(testProgressRaw)
        Else
            testProgressNorm = Empty
        End If

        If Not baseById.Exists(idVal) Then GoTo NextRow

        baseStart = baseById(idVal)(0)
        baseFinish = baseById(idVal)(1)
        baseProgress = baseById(idVal)(3)

        If hasTestStart Or hasTestFinish Or hasTestProgress Then
            If HasMeaningfulLockDelta(baseStart, baseFinish, baseProgress, testStart, testFinish, testProgressNorm) Then
                d(idVal) = Array(testStart, testFinish, testProgressNorm, hasTestStart, hasTestFinish, hasTestProgress)
            End If
        End If

NextRow:
    Next r

    Set BuildModifiedTestChangesMap = d

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Backup WBS Lock Columns sans modifier les donnees d'entree.
' EN: Returns the Backup WBS Lock Columns map without mutating input data.
'------------------------------------------------------------------------------

Private Function BackupWBSLockColumns(ByVal tblWBS As ListObject, ByVal mapWBS As Object) As Object

    Dim d As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")

    If tblWBS.DataBodyRange Is Nothing Then
        Set BackupWBSLockColumns = d
        Exit Function
    End If

    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapWBS(VTS_COL_ID))))
        If idVal <> "" Then
            d(idVal) = Array( _
                GetCellValue(arr(r, mapWBS(VTS_COL_FORECAST_START))), _
                GetCellValue(arr(r, mapWBS(VTS_COL_FORECAST_FINISH))), _
                GetCellValue(arr(r, mapWBS(VTS_COL_PROGRESS_PERCENT))), _
                NormalizeCalendarType(arr(r, mapWBS(VTS_COL_CAL))) _
            )
        End If
    Next r

    Set BackupWBSLockColumns = d

End Function

'------------------------------------------------------------------------------
' FR: Traite la map Restore WBS Lock Columns sans modifier les donnees d'entree.
' EN: Handles the Restore WBS Lock Columns map without mutating input data.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub RestoreWBSLockColumns(ByVal tblWBS As ListObject, ByVal mapWBS As Object, ByVal backup As Object)

    Dim r As Long
    Dim idVal As String
    Dim vals As Variant

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If backup Is Nothing Then Exit Sub

    For r = 1 To tblWBS.ListRows.Count
        idVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_ID)).value))

        If idVal <> "" Then
            If backup.Exists(idVal) Then
                vals = backup(idVal)

                tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_FORECAST_START)).value = vals(0)
                tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_FORECAST_FINISH)).value = vals(1)
                tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_PROGRESS_PERCENT)).value = vals(2)
            End If
        End If
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Modified Test Changes To WBS sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Modified Test Changes To WBS without changing the business rules that produce the data.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub ApplyModifiedTestChangesToWBS( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal changes As Object)

    Dim r As Long
    Dim idVal As String
    Dim vals As Variant

    Dim hasTestStart As Boolean
    Dim hasTestFinish As Boolean
    Dim hasTestProgress As Boolean

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If changes Is Nothing Then Exit Sub
    If changes.Count = 0 Then Exit Sub

    For r = 1 To tblWBS.ListRows.Count

        idVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_ID)).value))

        If idVal <> "" Then
            If changes.Exists(idVal) Then

                vals = changes(idVal)

                hasTestStart = CBool(vals(3))
                hasTestFinish = CBool(vals(4))
                hasTestProgress = CBool(vals(5))

                'Important:
                'A blank TEST Start means "do not lock/change Forecast Start",
                'not "clear Forecast Start".
                If hasTestStart Then
                    tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_FORECAST_START)).value = vals(0)
                End If

                'Important:
                'A blank TEST Finish means "do not lock/change Forecast Finish",
                'not "clear Forecast Finish".
                If hasTestFinish Then
                    tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_FORECAST_FINISH)).value = vals(1)
                End If

                If hasTestProgress Then
                    tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_PROGRESS_PERCENT)).value = vals(2)
                End If

            End If
        End If

    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la map Calc Table Has Errors sans modifier les donnees d'entree.
' EN: Returns the Calc Table Has Errors map without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcTableHasErrors(ByVal tblCalc As ListObject, ByVal mapCalc As Object) As Boolean

    Dim arr As Variant
    Dim r As Long

    If tblCalc.DataBodyRange Is Nothing Then Exit Function

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, mapCalc("Error flag"))))) = "ERROR" Then
            CalcTableHasErrors = True
            Exit Function
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Backup Gantt Test Inputs sans modifier les donnees d'entree.
' EN: Returns the Backup Gantt Test Inputs map without mutating input data.
'------------------------------------------------------------------------------

Private Function BackupGanttTestInputs(ByVal wsGantt As Worksheet) As Object

    Dim d As Object
    Dim lastRow As Long
    Dim r As Long
    Dim wbsVal As String

    Dim testStart As Variant
    Dim testFinish As Variant
    Dim testProgress As Variant

    Set d = CreateObject("Scripting.Dictionary")

    lastRow = GetLastGanttRow_Live(wsGantt)
    If lastRow < GANTT_FIRST_TASK_ROW Then
        Set BackupGanttTestInputs = d
        Exit Function
    End If

    For r = GANTT_FIRST_TASK_ROW To lastRow

        wbsVal = NormalizeWBS(CStr(wsGantt.Cells(r, GANTT_COL_WBS).value))
        If wbsVal = "" Then GoTo NextRow

        testStart = GetCellValue(wsGantt.Cells(r, COL_TEST_START).value)
        testFinish = GetCellValue(wsGantt.Cells(r, COL_TEST_FINISH).value)
        testProgress = GetCellValue(wsGantt.Cells(r, COL_TEST_PROGRESS).value)

        If HasValue(testStart) Or HasValue(testFinish) Or HasValue(testProgress) Then
            d(wbsVal) = Array(testStart, testFinish, testProgress)
        End If

NextRow:
    Next r

    Set BackupGanttTestInputs = d

End Function

'------------------------------------------------------------------------------
' FR: Traite la map Restore Gantt Test Inputs From Backup sans modifier les donnees d'entree.
' EN: Handles the Restore Gantt Test Inputs From Backup map without mutating input data.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub RestoreGanttTestInputsFromBackup(ByVal wsGantt As Worksheet, ByVal backup As Object)

    Dim wbsToRow As Object
    Dim lastRow As Long
    Dim r As Long

    Dim wbsVal As String
    Dim vals As Variant
    Dim rowIndex As Long
    Dim key As Variant
    Dim dateFormat As String

    If backup Is Nothing Then Exit Sub
    If backup.Count = 0 Then Exit Sub

    Set wbsToRow = CreateObject("Scripting.Dictionary")
    dateFormat = Settings_GetDateNumberFormat()

    lastRow = GetLastGanttRow_Live(wsGantt)
    If lastRow < GANTT_FIRST_TASK_ROW Then Exit Sub

    For r = GANTT_FIRST_TASK_ROW To lastRow
        wbsVal = NormalizeWBS(CStr(wsGantt.Cells(r, GANTT_COL_WBS).value))
        If wbsVal <> "" Then
            wbsToRow(wbsVal) = r
        End If
    Next r

    For Each key In backup.Keys

        wbsVal = CStr(key)

        If wbsToRow.Exists(wbsVal) Then

            rowIndex = wbsToRow(wbsVal)
            vals = backup(wbsVal)

            wsGantt.Cells(rowIndex, COL_TEST_START).value = vals(0)
            wsGantt.Cells(rowIndex, COL_TEST_FINISH).value = vals(1)
            wsGantt.Cells(rowIndex, COL_TEST_PROGRESS).value = vals(2)

            wsGantt.Cells(rowIndex, COL_TEST_START).NumberFormat = dateFormat
            wsGantt.Cells(rowIndex, COL_TEST_FINISH).NumberFormat = dateFormat
            wsGantt.Cells(rowIndex, COL_TEST_PROGRESS).NumberFormat = "0%"

        End If

    Next key

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Clear Gantt Test Inputs For Changed Tasks dans le perimetre possede par le composant.
' EN: Resets Clear Gantt Test Inputs For Changed Tasks within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ClearGanttTestInputs_ForChangedTasks( _
    ByVal wsGantt As Worksheet, _
    ByVal appliedChanges As Object)

    Dim wbsToRow As Object
    Dim wbsFromId As Object
    Dim lastRow As Long
    Dim r As Long
    Dim rowIndex As Long
    Dim wbsVal As String
    Dim idVal As Variant

    Set wbsToRow = CreateObject("Scripting.Dictionary")

    lastRow = GetLastGanttRow_Live(wsGantt)
    If lastRow < GANTT_FIRST_TASK_ROW Then Exit Sub

    For r = GANTT_FIRST_TASK_ROW To lastRow
        wbsVal = NormalizeWBS(CStr(wsGantt.Cells(r, GANTT_COL_WBS).value))
        If wbsVal <> "" Then
            wbsToRow(wbsVal) = r
        End If
    Next r

    Set wbsFromId = BuildIdToWbsMapFromWBS_Live()

    For Each idVal In appliedChanges.Keys

        If wbsFromId.Exists(CStr(idVal)) Then

            wbsVal = CStr(wbsFromId(CStr(idVal)))

            If wbsToRow.Exists(wbsVal) Then
                rowIndex = CLng(wbsToRow(wbsVal))

                wsGantt.Cells(rowIndex, COL_TEST_START).ClearContents
                wsGantt.Cells(rowIndex, COL_TEST_FINISH).ClearContents
                wsGantt.Cells(rowIndex, COL_TEST_PROGRESS).ClearContents

                wsGantt.Cells(rowIndex, COL_TEST_START).Interior.Color = RGB(255, 255, 153)
                wsGantt.Cells(rowIndex, COL_TEST_FINISH).Interior.Color = RGB(255, 255, 153)
                wsGantt.Cells(rowIndex, COL_TEST_PROGRESS).Interior.Color = RGB(255, 255, 153)
            End If

        End If

    Next idVal

End Sub


'------------------------------------------------------------------------------
' FR: Traite la map Rollback Failed Lock sans modifier les donnees d'entree.
' EN: Handles the Rollback Failed Lock map without mutating input data.
'------------------------------------------------------------------------------

Private Sub GanttLive_RollbackFailedLock( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal wsGantt As Worksheet, _
    ByVal wbsBackup As Object, _
    ByVal ganttTestBackup As Object)

    GanttLockTrace_Log "Rollback helper: RestoreWBSLockColumns start"
    RestoreWBSLockColumns tblWBS, mapWBS, wbsBackup
    GanttLockTrace_Log "Rollback helper: RestoreWBSLockColumns returned"

    GanttLockTrace_Log "Rollback helper: clear test render request"
    GanttLive_ClearTestRenderRequest
    GanttLockTrace_Log "Rollback helper: Run_Calc_Engine start"
    Run_Calc_Engine
    GanttLockTrace_Log "Rollback helper: Run_Calc_Engine returned"

    GanttLockTrace_Log "Rollback helper: RestoreGanttTestInputsFromBackup start"
    RestoreGanttTestInputsFromBackup wsGantt, ganttTestBackup
    GanttLockTrace_Log "Rollback helper: RestoreGanttTestInputsFromBackup returned"

End Sub

'------------------------------------------------------------------------------
' FR: Traite la map Finalize Successful Lock sans modifier les donnees d'entree.
' EN: Handles the Finalize Successful Lock map without mutating input data.
'------------------------------------------------------------------------------

Private Sub GanttLive_FinalizeSuccessfulLock( _
    ByVal wsGantt As Worksheet, _
    ByVal appliedChanges As Object)

    GanttLockTrace_Log "Finalize helper: ClearGanttTestInputs_ForChangedTasks start"
    ClearGanttTestInputs_ForChangedTasks wsGantt, appliedChanges
    GanttLockTrace_Log "Finalize helper: ClearGanttTestInputs_ForChangedTasks returned"
    GanttLockTrace_Log "Finalize helper: ClearCalcGanttTestResults start"
    ClearCalcGanttTestResults
    GanttLockTrace_Log "Finalize helper: ClearCalcGanttTestResults returned"
    GanttLockTrace_Log "Finalize helper: clear live state start"
    GanttLive_ClearTestRenderRequest
    GanttLive_ClearActiveSimulationMode
    SetGanttPreserveTestInputs False
    GanttLocal_ForgetIds appliedChanges, "LockFinalizedTestState"
    GanttRefresh_MarkRenderSignatureDirty "LockFinalizedTestState"
    GanttLockTrace_Log "Finalize helper: Ensure LOCAL_UPDATE start"
    If Not EnsureGanttForCurrentPlanning(GANTT_ENSURE_LOCAL_UPDATE, "GanttLock_FinalizeAppliedChanges") Then
        Err.Raise 5, "GanttLock_FinalizeAppliedChanges", "Gantt LOCK render did not reach READY state."
    End If
    GanttLockTrace_Log "Finalize helper: Ensure LOCAL_UPDATE returned"

End Sub

'------------------------------------------------------------------------------
' FR: Construit la map Simulated Lock Result Map a partir des donnees fournies par l'appelant.
' EN: Builds the Simulated Lock Result Map map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildSimulatedLockResultMap(ByVal appliedChanges As Object) As Object

    Dim d As Object
    Dim testById As Object
    Dim key As Variant
    Dim simVals As Variant
    Dim changeVals As Variant

    Dim compareStart As Boolean
    Dim compareFinish As Boolean
    Dim compareProgress As Boolean

    Set d = CreateObject("Scripting.Dictionary")
    Set testById = GanttLive_BuildTestByIdMap()

    If appliedChanges Is Nothing Then
        Set BuildSimulatedLockResultMap = d
        Exit Function
    End If

    For Each key In appliedChanges.Keys

        If testById.Exists(CStr(key)) Then

            simVals = testById(CStr(key))
            changeVals = appliedChanges(CStr(key))

            compareStart = CBool(changeVals(3))
            compareFinish = CBool(changeVals(4))
            compareProgress = CBool(changeVals(5))

            d(CStr(key)) = Array( _
                simVals(0), _
                simVals(1), _
                simVals(3), _
                compareStart, _
                compareFinish, _
                compareProgress)

        End If

    Next key

    Set BuildSimulatedLockResultMap = d

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Lock Results Match Simulated Result sans modifier les donnees d'entree.
' EN: Returns the Lock Results Match Simulated Result map without mutating input data.
'------------------------------------------------------------------------------

Private Function LockResultsMatchSimulatedResult( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal simulatedById As Object) As Boolean

    Dim r As Long
    Dim idVal As String
    Dim simVals As Variant

    Dim calcStartVal As Variant
    Dim calcFinishVal As Variant
    Dim progressVal As Variant

    Dim compareStart As Boolean
    Dim compareFinish As Boolean
    Dim compareProgress As Boolean

    If tblWBS.DataBodyRange Is Nothing Then
        LockResultsMatchSimulatedResult = False
        Exit Function
    End If

    For r = 1 To tblWBS.ListRows.Count

        idVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_ID)).value))

        If idVal <> "" Then
            If simulatedById.Exists(idVal) Then

                simVals = simulatedById(idVal)

                compareStart = CBool(simVals(3))
                compareFinish = CBool(simVals(4))
                compareProgress = CBool(simVals(5))

                If compareStart Then
                    calcStartVal = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_CALCULATED_START)).value)
                    If ValuesDiffer(calcStartVal, simVals(0)) Then Exit Function
                End If

                If compareFinish Then
                    calcFinishVal = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_CALCULATED_FINISH)).value)
                    If ValuesDiffer(calcFinishVal, simVals(1)) Then Exit Function
                End If

                If compareProgress Then
                    progressVal = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_PROGRESS_PERCENT)).value)
                    If ValuesDiffer(progressVal, simVals(2)) Then Exit Function
                End If

            End If
        End If

    Next r

    LockResultsMatchSimulatedResult = True

End Function

