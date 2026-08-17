Attribute VB_Name = "mod_CoreBridgeAnalytics"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgeAnalytics
' DOMAINE / DOMAIN : Core Bridge
'
' FR
' Calcule et route les deadlines, warnings parent/task type et analytics de reseau du Bridge.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Computes and routes deadlines, parent/task-type warnings and Bridge network analytics.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : CalcBridge_ComputeDeadlineAnalytics, CalcBridge_ShowParentDateWarnings, CalcBridge_AppendTaskTypeWarnings, CalcBridge_RunAnalyticsAndPush, CalcBridge_AddAnalyticsTopologyWarning
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================



'------------------------------------------------------------------------------
' FR: Traite la collection Compute Deadline Analytics sans modifier les donnees d'entree.
' EN: Handles the Compute Deadline Analytics collection without mutating input data.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Public Sub CalcBridge_ComputeDeadlineAnalytics( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    Optional ByVal consoleMessages As Collection)

    Dim dataArr As Variant
    Dim outDeadlineFloat() As Variant
    Dim exceededIds As Object
    Dim idToWbs As Object
    Dim ackTokensById As Object
    Dim ackTokens As String
    Dim r As Long
    Dim rowCount As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String
    Dim deadlineVal As Variant
    Dim calcFinishVal As Variant
    Dim deadlineFloatVal As Double
    Dim eventHashVal As String

    On Error GoTo SafeExit

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If mapCalc Is Nothing Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("WBS") Then Exit Sub
    If Not mapCalc.Exists("Task Name") Then Exit Sub
    If Not mapCalc.Exists("Deadline") Then Exit Sub
    If Not mapCalc.Exists("Deadline Float") Then Exit Sub
    If Not mapCalc.Exists("Calculated Finish") Then Exit Sub

    dataArr = tblCalc.DataBodyRange.value
    rowCount = UBound(dataArr, 1)
    ReDim outDeadlineFloat(1 To rowCount, 1 To 1)

    Set exceededIds = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set ackTokensById = CreateObject("Scripting.Dictionary")

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataArr(r, mapCalc("ID"))))
        wbsVal = NormalizeWBS(dataArr(r, mapCalc("WBS")))
        taskNameVal = Trim$(CStr(dataArr(r, mapCalc("Task Name"))))
        idToWbs(idVal) = wbsVal

        deadlineVal = dataArr(r, mapCalc("Deadline"))
        calcFinishVal = dataArr(r, mapCalc("Calculated Finish"))

        If HasValue(deadlineVal) And HasValue(calcFinishVal) Then
            If IsDate(deadlineVal) And IsDate(calcFinishVal) Then

                deadlineFloatVal = CDbl(CDate(deadlineVal)) - CDbl(CDate(calcFinishVal))
                outDeadlineFloat(r, 1) = deadlineFloatVal

                If deadlineFloatVal < 0# Then
                    If idVal <> "" Then
                        exceededIds(idVal) = True

                        eventHashVal = BuildPlanningEventHash( _
                            "WARNING", _
                            "DEADLINE_EXCEEDED", _
                            "Deadline depassee", _
                            "Deadline exceeded", _
                            "la date calculee finit apres la deadline", _
                            "calculated finish is after the deadline", _
                            "CalcBridge_ComputeDeadlineAnalytics", _
                            "CALC", _
                            "tbl_CALC", _
                            idVal, _
                            wbsVal, _
                            taskNameVal)

                        ackTokensById(idVal) = BuildPlanningWarningAckToken("DEADLINE_EXCEEDED", eventHashVal)

                        LogPlanningEvent _
                            "WARNING", _
                            "DEADLINE_EXCEEDED", _
                            eventHashVal, _
                            "Deadline depassee", _
                            "Deadline exceeded", _
                            "la date calculee finit apres la deadline", _
                            "calculated finish is after the deadline", _
                            "CalcBridge_ComputeDeadlineAnalytics", _
                            "CALC", _
                            "tbl_CALC", _
                            idVal, _
                            wbsVal, _
                            taskNameVal
                    End If
                End If
            End If
        End If

    Next r

    tblCalc.ListColumns("Deadline Float").DataBodyRange.value = outDeadlineFloat

    If exceededIds.Count > 0 Then
        ackTokens = CalcBridge_BuildWarningAckTokenList(exceededIds, ackTokensById)

        If consoleMessages Is Nothing Then
            Set consoleMessages = New Collection

            CalcBridge_AddGroupedWarningToCollection consoleMessages, exceededIds, idToWbs, _
                "Deadline depassee", _
                "la date calculee finit apres la deadline", _
                "Deadline exceeded", _
                "calculated finish is after the deadline", _
                True, _
                ackTokens

            CalcBridge_ShowPlanningConsole consoleMessages
        Else
            CalcBridge_AddGroupedWarningToCollection consoleMessages, exceededIds, idToWbs, _
                "Deadline depassee", _
                "la date calculee finit apres la deadline", _
                "Deadline exceeded", _
                "calculated finish is after the deadline", _
                True, _
                ackTokens
        End If
    End If

SafeExit:
End Sub

'------------------------------------------------------------------------------
' FR: Projette la collection Parent Date Warnings vers l'interface autorisee par la politique runtime.
' EN: Projects the Parent Date Warnings collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Public Sub CalcBridge_ShowParentDateWarnings( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    Optional ByVal consoleMessages As Collection)

    Dim dataArr As Variant
    Dim rowById As Object
    Dim parentIds As Object
    Dim warnParentDates As Object
    Dim idToWbs As Object
    Dim ackTokensById As Object
    Dim ackTokens As String
    Dim eventHashVal As String

    Dim key As Variant
    Dim rowIdx As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String

    On Error GoTo SafeExit

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Actual Start") Then Exit Sub
    If Not mapCalc.Exists("Actual Finish") Then Exit Sub
    If Not mapCalc.Exists("Forecast Start") Then Exit Sub
    If Not mapCalc.Exists("Forecast Finish") Then Exit Sub
    If Not mapCalc.Exists("Baseline Start") Then Exit Sub
    If Not mapCalc.Exists("Baseline Duration") Then Exit Sub
    If Not mapCalc.Exists("Baseline Finish") Then Exit Sub

    dataArr = tblCalc.DataBodyRange.value

    Set rowById = Core_BuildRowById(dataArr, mapCalc)
    Set parentIds = Core_BuildParentIds(dataArr, mapCalc, rowById)

    Set warnParentDates = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set ackTokensById = CreateObject("Scripting.Dictionary")

    For Each key In rowById.Keys

        idVal = CStr(key)
        rowIdx = CLng(rowById(idVal))

        If mapCalc.Exists("WBS") Then
            wbsVal = NormalizeWBS(dataArr(rowIdx, mapCalc("WBS")))
        Else
            wbsVal = "-"
        End If

        idToWbs(idVal) = wbsVal

        If parentIds.Exists(idVal) Then

            If HasValue(dataArr(rowIdx, mapCalc("Actual Start"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Actual Finish"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Forecast Start"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Forecast Finish"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Baseline Start"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Baseline Duration"))) Or _
               HasValue(dataArr(rowIdx, mapCalc("Baseline Finish"))) Then

                warnParentDates(idVal) = True

                If mapCalc.Exists("Task Name") Then
                    taskNameVal = Trim$(CStr(dataArr(rowIdx, mapCalc("Task Name"))))
                Else
                    taskNameVal = vbNullString
                End If

                eventHashVal = BuildPlanningEventHash( _
                    "WARNING", _
                    "PARENT_DATES_IGNORED", _
                    "Dates saisies sur tache parent", _
                    "Dates entered on summary task", _
                    "les valeurs sont ignorees, calcul par les taches enfants", _
                    "values are ignored and calculated from child tasks", _
                    "CalcBridge_ShowParentDateWarnings", _
                    "CALC", _
                    "tbl_CALC", _
                    idVal, _
                    wbsVal, _
                    taskNameVal)
                ackTokensById(idVal) = BuildPlanningWarningAckToken("PARENT_DATES_IGNORED", eventHashVal)

                On Error Resume Next

                LogPlanningEvent _
                    "WARNING", _
                    "PARENT_DATES_IGNORED", _
                    eventHashVal, _
                    "Dates saisies sur tache parent", _
                    "Dates entered on summary task", _
                    "les valeurs sont ignorees, calcul par les taches enfants", _
                    "values are ignored and calculated from child tasks", _
                    "CalcBridge_ShowParentDateWarnings", _
                    "CALC", _
                    "tbl_CALC", _
                    idVal, _
                    wbsVal, _
                    taskNameVal
                On Error GoTo SafeExit

            End If

        End If

    Next key

    If warnParentDates.Count > 0 Then

        ackTokens = CalcBridge_BuildWarningAckTokenList(warnParentDates, ackTokensById)

        If consoleMessages Is Nothing Then
            Set consoleMessages = New Collection

            CalcBridge_AddGroupedWarningToCollection consoleMessages, warnParentDates, idToWbs, _
                "Dates saisies sur tâche parent", _
                "les valeurs sont ignorées, calcul par les tâches enfants", _
                "Dates entered on summary task", _
                "values are ignored and calculated from child tasks", _
                True, _
                ackTokens

            CalcBridge_ShowPlanningConsole consoleMessages
        Else
            CalcBridge_AddGroupedWarningToCollection consoleMessages, warnParentDates, idToWbs, _
                "Dates saisies sur tâche parent", _
                "les valeurs sont ignorées, calcul par les tâches enfants", _
                "Dates entered on summary task", _
                "values are ignored and calculated from child tasks", _
                True, _
                ackTokens
        End If

    End If

SafeExit:
End Sub


'------------------------------------------------------------------------------
' FR: Indique si la valeur Progress Strictly Between Zero And One satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Progress Strictly Between Zero And One value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function CalcBridge_IsProgressStrictlyBetweenZeroAndOne(ByVal v As Variant) As Boolean

    Dim p As Variant

    If Not HasValue(v) Then Exit Function

    p = NormalizePercentInput(v)
    If Not HasValue(p) Then Exit Function

    CalcBridge_IsProgressStrictlyBetweenZeroAndOne = _
        (CDbl(p) > 0.0000001 And CDbl(p) < 0.9999999)

End Function


'------------------------------------------------------------------------------
' FR: Ajoute la collection Task Type Warnings a la structure cible fournie par l'appelant.
' EN: Adds the Task Type Warnings collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AppendTaskTypeWarnings( _
    ByVal warningMessages As Collection, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object)

    Dim arrWBS As Variant
    Dim arrCalc As Variant
    Dim idToWbs As Object
    Dim calcRowById As Object

    Dim warnLOEProgress As Object
    Dim warnLOEBaseline As Object
    Dim warnMilestoneProgress As Object
    Dim warnMilestoneDuration As Object
    Dim warnIgnoredCal As Object

    Dim r As Long
    Dim calcRow As Long
    Dim idVal As String
    Dim calVal As String

    On Error GoTo SafeExit

    If warningMessages Is Nothing Then Exit Sub
    If tblWBS Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then Exit Sub
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If mapWBS Is Nothing Then Exit Sub
    If mapCalc Is Nothing Then Exit Sub

    If Not mapWBS.Exists("ID") Then Exit Sub
    If Not mapWBS.Exists("WBS") Then Exit Sub
    If Not mapWBS.Exists("Task Type") Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Task Type") Then Exit Sub

    arrWBS = tblWBS.DataBodyRange.value
    arrCalc = tblCalc.DataBodyRange.value

    Set idToWbs = BuildIdToWbsFromWBS(tblWBS)
    Set calcRowById = Core_BuildRowById(arrCalc, mapCalc)

    Set warnLOEProgress = CreateObject("Scripting.Dictionary")
    Set warnLOEBaseline = CreateObject("Scripting.Dictionary")
    Set warnMilestoneProgress = CreateObject("Scripting.Dictionary")
    Set warnMilestoneDuration = CreateObject("Scripting.Dictionary")
    Set warnIgnoredCal = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arrWBS, 1)

        idVal = Trim$(CStr(arrWBS(r, mapWBS("ID"))))
        If idVal = "" Then GoTo NextRow

        If mapWBS.Exists("Cal") Then
            calVal = NormalizeCalendarType(arrWBS(r, mapWBS("Cal")))
            If calVal = CALENDAR_5D Or calVal = CALENDAR_6D Then
                If calcRowById.Exists(idVal) Then
                    calcRow = CLng(calcRowById(idVal))
                    If Core_IsSummaryRow(arrCalc, calcRow, mapCalc) Or _
                       TaskTypeRules_IsLevelOfEffortRow(arrWBS, mapWBS, r) Or _
                       TaskTypeRules_IsMilestoneRow(arrWBS, mapWBS, r) Then
                        warnIgnoredCal(idVal) = True
                    End If
                ElseIf TaskTypeRules_IsLevelOfEffortRow(arrWBS, mapWBS, r) Or _
                       TaskTypeRules_IsMilestoneRow(arrWBS, mapWBS, r) Then
                    warnIgnoredCal(idVal) = True
                End If
            End If
        End If
        If TaskTypeRules_IsLevelOfEffortRow(arrWBS, mapWBS, r) Then

            If mapWBS.Exists("% Progress") Then
                If HasValue(arrWBS(r, mapWBS("% Progress"))) Then warnLOEProgress(idVal) = True
            End If

            If (mapWBS.Exists("Baseline Start") And HasValue(arrWBS(r, mapWBS("Baseline Start")))) Or _
               (mapWBS.Exists("Baseline Duration") And HasValue(arrWBS(r, mapWBS("Baseline Duration")))) Or _
               (mapWBS.Exists("Baseline Finish") And HasValue(arrWBS(r, mapWBS("Baseline Finish")))) Then
                warnLOEBaseline(idVal) = True
            End If

        ElseIf TaskTypeRules_IsMilestoneRow(arrWBS, mapWBS, r) Then

            If mapWBS.Exists("% Progress") Then
                If CalcBridge_IsProgressStrictlyBetweenZeroAndOne(arrWBS(r, mapWBS("% Progress"))) Then
                    warnMilestoneProgress(idVal) = True
                End If
            End If

            If calcRowById.Exists(idVal) Then
                calcRow = CLng(calcRowById(idVal))

                If CalcBridge_RowHasDurationGreaterThanOne(arrCalc, mapCalc, calcRow) Then
                    warnMilestoneDuration(idVal) = True
                End If
            End If

        End If

NextRow:
    Next r

    If warnIgnoredCal.Count > 0 Then
        CalcBridge_AddGroupedWarningToCollection warningMessages, warnIgnoredCal, idToWbs, _
            "Calendrier ignoré sur tâche Summary / LOE / Milestone", _
            "le calendrier n'est utilisé que pour les tâches normales", _
            "Calendar ignored on Summary / LOE / Milestone task", _
            "calendars apply only to normal tasks"
    End If
    If warnLOEProgress.Count > 0 Then
        CalcBridge_AddGroupedWarningToCollection warningMessages, warnLOEProgress, idToWbs, _
            "% Progress renseigné sur LOE", _
            "le progress LOE est calculé automatiquement par date du jour ; la saisie manuelle est ignorée dans le Gantt", _
            "% Progress entered on LOE", _
            "LOE progress is automatically calculated from today's date; manual input is ignored in the Gantt"
    End If

    If warnLOEBaseline.Count > 0 Then
        CalcBridge_AddGroupedWarningToCollection warningMessages, warnLOEBaseline, idToWbs, _
            "Baseline renseignée sur LOE", _
            "une LOE doit être pilotée par ses liens SS/FF ; vérifier que la baseline saisie ne crée pas de confusion", _
            "Baseline entered on LOE", _
            "a LOE must be driven by its SS/FF links; check that entered baseline values are not misleading"
    End If

    If warnMilestoneProgress.Count > 0 Then
        CalcBridge_AddGroupedWarningToCollection warningMessages, warnMilestoneProgress, idToWbs, _
            "% Progress partiel renseigné sur Milestone", _
            "une milestone doit être à 0% ou 100% ; toute valeur intermédiaire doit être corrigée", _
            "Partial % Progress entered on Milestone", _
            "a milestone must be either 0% or 100%; any intermediate value should be corrected"
    End If

    If warnMilestoneDuration.Count > 0 Then
        CalcBridge_AddGroupedWarningToCollection warningMessages, warnMilestoneDuration, idToWbs, _
            "Durée supérieure à 1 jour sur Milestone", _
            "la tâche sera rendue comme milestone mais la durée saisie peut induire en erreur", _
            "Duration greater than 1 day on Milestone", _
            "the task will be rendered as a milestone but the entered duration may be misleading"
    End If

SafeExit:
End Sub


'------------------------------------------------------------------------------
' FR: Retourne la map Row Has Duration Greater Than One sans modifier les donnees d'entree.
' EN: Returns the Row Has Duration Greater Than One map without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcBridge_RowHasDurationGreaterThanOne( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowIdx As Long) As Boolean

    If mapCalc.Exists("Calculated Duration") Then
        If IsNumeric(dataArr(rowIdx, mapCalc("Calculated Duration"))) Then
            If CDbl(dataArr(rowIdx, mapCalc("Calculated Duration"))) > 1# Then
                CalcBridge_RowHasDurationGreaterThanOne = True
                Exit Function
            End If
        End If
    End If

    If mapCalc.Exists("Baseline Duration") Then
        If IsNumeric(dataArr(rowIdx, mapCalc("Baseline Duration"))) Then
            If CDbl(dataArr(rowIdx, mapCalc("Baseline Duration"))) > 1# Then
                CalcBridge_RowHasDurationGreaterThanOne = True
                Exit Function
            End If
        End If
    End If

    If mapCalc.Exists("Actual Duration") Then
        If IsNumeric(dataArr(rowIdx, mapCalc("Actual Duration"))) Then
            If CDbl(dataArr(rowIdx, mapCalc("Actual Duration"))) > 1# Then
                CalcBridge_RowHasDurationGreaterThanOne = True
                Exit Function
            End If
        End If
    End If

    If CalcBridge_PairDurationGreaterThanOne(dataArr, mapCalc, rowIdx, "Baseline Start", "Baseline Finish") Then
        CalcBridge_RowHasDurationGreaterThanOne = True
        Exit Function
    End If

    If CalcBridge_PairDurationGreaterThanOne(dataArr, mapCalc, rowIdx, "Actual Start", "Actual Finish") Then
        CalcBridge_RowHasDurationGreaterThanOne = True
        Exit Function
    End If

    If CalcBridge_PairDurationGreaterThanOne(dataArr, mapCalc, rowIdx, "Forecast Start", "Forecast Finish") Then
        CalcBridge_RowHasDurationGreaterThanOne = True
        Exit Function
    End If

    If CalcBridge_PairDurationGreaterThanOne(dataArr, mapCalc, rowIdx, "Calculated Start", "Calculated Finish") Then
        CalcBridge_RowHasDurationGreaterThanOne = True
        Exit Function
    End If

End Function


'------------------------------------------------------------------------------
' FR: Retourne la map Pair Duration Greater Than One sans modifier les donnees d'entree.
' EN: Returns the Pair Duration Greater Than One map without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcBridge_PairDurationGreaterThanOne( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowIdx As Long, _
    ByVal startColName As String, _
    ByVal finishColName As String) As Boolean

    Dim startVal As Variant
    Dim finishVal As Variant

    If Not mapCalc.Exists(startColName) Then Exit Function
    If Not mapCalc.Exists(finishColName) Then Exit Function

    startVal = dataArr(rowIdx, mapCalc(startColName))
    finishVal = dataArr(rowIdx, mapCalc(finishColName))

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    CalcBridge_PairDurationGreaterThanOne = _
        ((CDbl(finishVal) - CDbl(startVal) + 1) > 1#)

End Function

'------------------------------------------------------------------------------
' FR: Orchestre Run Analytics And Push en preservant l'ordre contractuel des etapes du domaine.
' EN: Orchestrates Run Analytics And Push while preserving the domain's contractual step order.
'------------------------------------------------------------------------------

' FR: Calcule en memoire les floats et chemins du dataset fourni, sans ecriture CALC/WBS.
' EN: Computes floats and paths in memory for the supplied dataset without CALC/WBS writes.
Public Function CalcBridge_BuildAnalyticsSnapshotById( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal linksBySuccId As Object, _
    Optional ByVal executionNetwork As clsCompiledExecutionNetwork) As Object

    Dim perfScope As clsPerfScope
    Dim result As Object
    Dim rowById As Object
    Dim parentIds As Object
    Dim validIds As Object
    Dim directChildrenById As Object
    Dim predsById As Object
    Dim childrenById As Object
    Dim topoOrder As Collection
    Dim analyticsPredLagBySuccPred As Object
    Dim analyticsPredTypeBySuccPred As Object
    Dim sharedNetworkFinishById As Object
    Dim outTotalFloat As Variant
    Dim outFreeFloat As Variant
    Dim outCriticalPath As Variant
    Dim outLongestPath As Variant
    Dim idKey As Variant
    Dim rowIndex As Long

    Set perfScope = Profiler_BeginScope("GanttTestAnalytics_Compute", "Analytics")
    Set result = CreateObject("Scripting.Dictionary")

    If mapCalc Is Nothing Then GoTo SafeExit
    If linksBySuccId Is Nothing Then GoTo SafeExit

    If executionNetwork Is Nothing Then
        Set executionNetwork = CompileExecutionNetwork(dataArr, mapCalc, linksBySuccId)
    End If

    Set rowById = executionNetwork.RowById
    Set parentIds = executionNetwork.ParentIds
    Set validIds = executionNetwork.ValidLeafIds
    Set directChildrenById = executionNetwork.DirectChildrenById
    Set predsById = executionNetwork.AnalyticsPredsById
    Set childrenById = executionNetwork.AnalyticsChildrenById
    Set analyticsPredLagBySuccPred = executionNetwork.PredLagBySuccPred
    Set analyticsPredTypeBySuccPred = executionNetwork.PredTypeBySuccPred
    Set topoOrder = executionNetwork.TopoOrder
    If topoOrder.Count <> validIds.Count Then GoTo SafeExit
    If IsCriticalPathMultiNetworkEnabled() Then
        Set sharedNetworkFinishById = CompiledNetwork_BuildCurrentFinishById(executionNetwork, dataArr, mapCalc)
    End If

    ComputeCurrentFloatAndCritical _
        Nothing, mapCalc, rowById, childrenById, directChildrenById, parentIds, validIds, topoOrder, _
        Nothing, analyticsPredLagBySuccPred, analyticsPredTypeBySuccPred, _
        dataArr, outTotalFloat, outFreeFloat, outCriticalPath, sharedNetworkFinishById

    ComputeLongestPath _
        Nothing, mapCalc, rowById, predsById, childrenById, validIds, _
        analyticsPredLagBySuccPred, analyticsPredTypeBySuccPred, dataArr, outLongestPath, sharedNetworkFinishById

    For Each idKey In rowById.Keys
        rowIndex = CLng(rowById(CStr(idKey)))
        result(CStr(idKey)) = Array( _
            outTotalFloat(rowIndex, 1), _
            outFreeFloat(rowIndex, 1), _
            CStr(outCriticalPath(rowIndex, 1)), _
            CStr(outLongestPath(rowIndex, 1)))
    Next idKey

SafeExit:
    Set CalcBridge_BuildAnalyticsSnapshotById = result

End Function

'------------------------------------------------------------------------------
Public Sub CalcBridge_RunAnalyticsAndPush( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal linksBySuccId As Object, _
    Optional ByVal consoleMessages As Collection, _
    Optional ByVal executionNetwork As clsCompiledExecutionNetwork)

    Dim perfScope As clsPerfScope

    Dim dataArr As Variant
    Dim rowById As Object
    Dim parentIds As Object
    Dim validIds As Object
    Dim directChildrenById As Object
    Dim predsById As Object
    Dim childrenById As Object
    Dim topoOrder As Collection
    Dim idToWbs As Object
    Dim outCriticalREX() As Variant
    Dim errMissingBaselineForREX As Object
    Dim analyticsPredLagBySuccPred As Object
    Dim analyticsPredTypeBySuccPred As Object
    Dim sharedNetworkFinishById As Object
    Dim sharedRexNetworkFinishById As Object
    Dim rexStartById As Object
    Dim rexFinishById As Object
    Dim rexDurationById As Object
    Dim rexDiagnosticsById As Object
    Dim rexDataArr As Variant

    Set perfScope = Profiler_BeginScope("CalcBridge_RunAnalyticsAndPush", "Analytics")

    On Error GoTo ErrHandler

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If mapCalc Is Nothing Then Exit Sub

    dataArr = tblCalc.DataBodyRange.value

    If executionNetwork Is Nothing Then
        Set executionNetwork = CompileExecutionNetwork(dataArr, mapCalc, linksBySuccId)
    End If

    Set rowById = executionNetwork.RowById
    Set parentIds = executionNetwork.ParentIds
    Set validIds = executionNetwork.ValidLeafIds
    Set directChildrenById = executionNetwork.DirectChildrenById
    Set predsById = executionNetwork.AnalyticsPredsById
    Set childrenById = executionNetwork.AnalyticsChildrenById
    Set idToWbs = executionNetwork.IdToWbs
    Set errMissingBaselineForREX = CreateObject("Scripting.Dictionary")
    Set analyticsPredLagBySuccPred = executionNetwork.PredLagBySuccPred
    Set analyticsPredTypeBySuccPred = executionNetwork.PredTypeBySuccPred
    Set topoOrder = executionNetwork.TopoOrder

    If IsCriticalPathMultiNetworkEnabled() Then
        Set sharedNetworkFinishById = CompiledNetwork_BuildCurrentFinishById(executionNetwork, dataArr, mapCalc)
    End If

    If topoOrder.Count <> validIds.Count Then

        CalcBridge_AddAnalyticsTopologyWarning consoleMessages

        Exit Sub
    End If

    ComputeCurrentFloatAndCritical _
        tblCalc, mapCalc, rowById, childrenById, directChildrenById, parentIds, validIds, topoOrder, _
        consoleMessages, analyticsPredLagBySuccPred, analyticsPredTypeBySuccPred, , , , , sharedNetworkFinishById

    ReDim outCriticalREX(1 To tblCalc.ListRows.Count, 1 To 1)

    ComputeLongestPath _
        tblCalc, mapCalc, rowById, predsById, childrenById, validIds, _
        analyticsPredLagBySuccPred, analyticsPredTypeBySuccPred, , , sharedNetworkFinishById

    ComputeCriticalPathREX _
        tblCalc, mapCalc, rowById, predsById, childrenById, directChildrenById, _
        parentIds, validIds, topoOrder, idToWbs, outCriticalREX, errMissingBaselineForREX, _
        consoleMessages, analyticsPredLagBySuccPred, analyticsPredTypeBySuccPred, executionNetwork

    If mapCalc.Exists("Longest Path REX") And errMissingBaselineForREX.Count = 0 Then
        Set rexStartById = CreateObject("Scripting.Dictionary")
        Set rexFinishById = CreateObject("Scripting.Dictionary")
        Set rexDurationById = CreateObject("Scripting.Dictionary")
        Set rexDiagnosticsById = CreateObject("Scripting.Dictionary")
        rexDataArr = dataArr

        If BuildBaselineRexTemporalState( _
            rexDataArr, mapCalc, rowById, predsById, validIds, topoOrder, _
            analyticsPredLagBySuccPred, analyticsPredTypeBySuccPred, _
            rexStartById, rexFinishById, rexDurationById, rexDiagnosticsById) Then

            ApplyBaselineRexTemporalStateToArray _
                rexDataArr, mapCalc, rowById, rexStartById, rexFinishById, rexDurationById

            If IsCriticalPathMultiNetworkEnabled() Then
                Set sharedRexNetworkFinishById = CalcBridge_BuildFinishByIdFromColumn(rexDataArr, mapCalc, rowById, validIds, "Baseline Finish", False, executionNetwork)
            End If

            ComputeLongestPath _
                tblCalc, mapCalc, rowById, predsById, childrenById, validIds, _
                analyticsPredLagBySuccPred, analyticsPredTypeBySuccPred, rexDataArr, , sharedRexNetworkFinishById, _
                "Baseline Start", "Baseline Finish", "Longest Path REX", False
        End If
    End If

    If errMissingBaselineForREX.Count > 0 Then
        If consoleMessages Is Nothing Then
            Set consoleMessages = New Collection
            CalcBridge_AddConsoleMessage consoleMessages, "WARNING", _
                CalcBridge_BuildMissingBaselineRexMessage(errMissingBaselineForREX, idToWbs)
            CalcBridge_ShowPlanningConsole consoleMessages
        Else
            CalcBridge_AddConsoleMessage consoleMessages, "WARNING", _
                CalcBridge_BuildMissingBaselineRexMessage(errMissingBaselineForREX, idToWbs)
        End If
    End If

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "CalcBridge_RunAnalyticsAndPush", Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Construit une map finish par ID pour l'etat temporel fourni, avec projection par composante compilee si disponible.
' EN: Builds a finish-by-ID map for the supplied temporal state, using compiled component projection when available.
'------------------------------------------------------------------------------
Private Function CalcBridge_BuildFinishByIdFromColumn( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object, _
    ByVal validIds As Object, _
    ByVal finishColumnName As String, _
    ByVal excludeActualFinished As Boolean, _
    Optional ByVal executionNetwork As clsCompiledExecutionNetwork) As Object

    Dim rawFinishById As Object
    Dim idKey As Variant
    Dim taskId As String
    Dim rowIndex As Long
    Dim finishVal As Variant

    Set rawFinishById = CreateObject("Scripting.Dictionary")

    If mapCalc Is Nothing Then GoTo SafeExit
    If rowById Is Nothing Then GoTo SafeExit
    If validIds Is Nothing Then GoTo SafeExit
    If Not mapCalc.Exists(finishColumnName) Then GoTo SafeExit

    For Each idKey In validIds.Keys
        taskId = CStr(idKey)
        If rowById.Exists(taskId) Then
            rowIndex = CLng(rowById(taskId))
            If (Not excludeActualFinished) Or Not CalcBridge_IsActualFinishedInCalcArray(dataArr, rowIndex, mapCalc) Then
                finishVal = GetCellValue(dataArr(rowIndex, mapCalc(finishColumnName)))
                If HasValue(finishVal) Then rawFinishById(taskId) = finishVal
            End If
        End If
    Next idKey

SafeExit:
    If Not executionNetwork Is Nothing Then
        Set CalcBridge_BuildFinishByIdFromColumn = CompiledNetwork_BuildFinishByIdFromValues(executionNetwork, rawFinishById)
    Else
        Set CalcBridge_BuildFinishByIdFromColumn = rawFinishById
    End If

End Function
'------------------------------------------------------------------------------
' FR: Indique si la ligne CALC porte un Actual Finish exploitable.
' EN: Returns whether the CALC row has a usable Actual Finish.
'------------------------------------------------------------------------------
Private Function CalcBridge_IsActualFinishedInCalcArray( _
    ByRef dataArr As Variant, _
    ByVal rowIndex As Long, _
    ByVal mapCalc As Object) As Boolean

    If mapCalc Is Nothing Then Exit Function
    If Not mapCalc.Exists("Actual Finish") Then Exit Function

    CalcBridge_IsActualFinishedInCalcArray = HasValue(GetCellValue(dataArr(rowIndex, mapCalc("Actual Finish"))))

End Function
'------------------------------------------------------------------------------
' FR: Construit la collection Leaf IDs a partir des donnees fournies par l'appelant.
' EN: Builds the Leaf IDs collection from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CalcBridge_BuildLeafIds( _
    ByVal rowById As Object, _
    ByVal parentIds As Object, _
    Optional ByRef dataArr As Variant, _
    Optional ByVal mapCalc As Object) As Object

    Dim d As Object
    Dim key As Variant
    Dim rowIdx As Long

    Set d = CreateObject("Scripting.Dictionary")

    If rowById Is Nothing Then
        Set CalcBridge_BuildLeafIds = d
        Exit Function
    End If

    For Each key In rowById.Keys

        If Not parentIds Is Nothing Then
            If parentIds.Exists(CStr(key)) Then GoTo NextKey
        End If

        rowIdx = CLng(rowById(CStr(key)))

        'LOE is not part of analytics / critical path / floats / REX.
        'This exclusion happens before REX baseline-duration validation.
        If TaskTypeRules_IsLevelOfEffortRow(dataArr, mapCalc, rowIdx) Then
            GoTo NextKey
        End If

        d(CStr(key)) = True

NextKey:
    Next key

    Set CalcBridge_BuildLeafIds = d

End Function

'------------------------------------------------------------------------------
' FR: Construit la map Empty Collections a partir des donnees fournies par l'appelant.
' EN: Builds the Empty Collections map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CalcBridge_BuildEmptyCollections(ByVal ids As Object) As Object

    Dim d As Object
    Dim key As Variant

    Set d = CreateObject("Scripting.Dictionary")

    For Each key In ids.Keys
        Set d(CStr(key)) = New Collection
    Next key

    Set CalcBridge_BuildEmptyCollections = d

End Function

'------------------------------------------------------------------------------
' FR: Construit la map parent ID -> enfants directs a partir des donnees fournies par l'appelant.
' EN: Builds the parent-ID-to-direct-children map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CalcBridge_BuildDirectChildrenById( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object) As Object

    Dim d As Object
    Dim key As Variant
    Dim r As Long
    Dim idVal As String
    Dim parentId As String

    Set d = CreateObject("Scripting.Dictionary")

    For Each key In rowById.Keys
        Set d(CStr(key)) = New Collection
    Next key

    If Not mapCalc.Exists("ID") Then
        Set CalcBridge_BuildDirectChildrenById = d
        Exit Function
    End If

    If Not mapCalc.Exists("ParentID") Then
        Set CalcBridge_BuildDirectChildrenById = d
        Exit Function
    End If

    For r = 1 To UBound(dataArr, 1)
        idVal = Trim$(CStr(dataArr(r, mapCalc("ID"))))
        parentId = Trim$(CStr(dataArr(r, mapCalc("ParentID"))))

        If idVal <> "" And parentId <> "" Then
            If d.Exists(parentId) Then
                d(parentId).Add idVal
            End If
        End If
    Next r

    Set CalcBridge_BuildDirectChildrenById = d

End Function

'------------------------------------------------------------------------------
' FR: Construit la map Analytics Network From Expanded Links a partir des donnees fournies par l'appelant.
' EN: Builds the Analytics Network From Expanded Links map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Sub CalcBridge_BuildAnalyticsNetworkFromExpandedLinks( _
    ByVal linksBySuccId As Object, _
    ByVal validIds As Object, _
    ByVal predsById As Object, _
    ByVal childrenById As Object, _
    ByRef predLagBySuccPred As Object, _
    ByRef predTypeBySuccPred As Object)

    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String
    Dim linkType As String
    Dim linkLag As Double
    Dim linkKey As String

    Set predLagBySuccPred = CreateObject("Scripting.Dictionary")
    Set predTypeBySuccPred = CreateObject("Scripting.Dictionary")

    If linksBySuccId Is Nothing Then Exit Sub
    If validIds Is Nothing Then Exit Sub
    If predsById Is Nothing Then Exit Sub
    If childrenById Is Nothing Then Exit Sub

    For Each succId In linksBySuccId.Keys
        If validIds.Exists(CStr(succId)) Then
            For Each oneLink In linksBySuccId(CStr(succId))
                predId = Core_GetLinkPredId(oneLink)

                If predId <> "" Then
                    If validIds.Exists(predId) Then
                        On Error Resume Next
                        linkType = Core_GetLinkType(oneLink)
                        linkLag = Core_GetLinkLag(oneLink)
                        If Err.Number <> 0 Then
                            Err.Clear
                            On Error GoTo 0
                            GoTo NextLink
                        End If
                        On Error GoTo 0

                        linkType = UCase$(Trim$(linkType))
                        If linkType <> "FS" And linkType <> "SS" And linkType <> "FF" Then GoTo NextLink

                        linkKey = CStr(succId) & "|" & predId
                        predLagBySuccPred(linkKey) = linkLag
                        predTypeBySuccPred(linkKey) = linkType
                        predsById(CStr(succId)).Add predId
                        childrenById(predId).Add CStr(succId)
                    End If
                End If

NextLink:
            Next oneLink
        End If
    Next succId

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la collection Topological Order sans modifier les donnees d'entree.
' EN: Returns the Topological Order collection without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcBridge_TopologicalOrder( _
    ByVal validIds As Object, _
    ByVal predsById As Object, _
    ByVal childrenById As Object) As Collection

    Dim q As Collection
    Dim topo As Collection
    Dim indegree As Object
    Dim key As Variant
    Dim childId As Variant
    Dim currentId As String

    Set q = New Collection
    Set topo = New Collection
    Set indegree = CreateObject("Scripting.Dictionary")

    For Each key In validIds.Keys
        indegree(CStr(key)) = predsById(CStr(key)).Count
        If CLng(indegree(CStr(key))) = 0 Then q.Add CStr(key)
    Next key

    Do While q.Count > 0
        currentId = CStr(q(1))
        q.Remove 1
        topo.Add currentId

        If childrenById.Exists(currentId) Then
            For Each childId In childrenById(currentId)
                indegree(CStr(childId)) = CLng(indegree(CStr(childId))) - 1
                If CLng(indegree(CStr(childId))) = 0 Then q.Add CStr(childId)
            Next childId
        End If
    Loop

    Set CalcBridge_TopologicalOrder = topo

End Function

'------------------------------------------------------------------------------
' FR: Construit la map Id To WBS From Data a partir des donnees fournies par l'appelant.
' EN: Builds the Id To WBS From Data map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CalcBridge_BuildIdToWbsFromData( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object) As Object

    Dim d As Object
    Dim key As Variant
    Dim rowIdx As Long

    Set d = CreateObject("Scripting.Dictionary")

    For Each key In rowById.Keys
        rowIdx = CLng(rowById(CStr(key)))

        If mapCalc.Exists("WBS") Then
            d(CStr(key)) = Trim$(CStr(dataArr(rowIdx, mapCalc("WBS"))))
        Else
            d(CStr(key)) = "-"
        End If
    Next key

    Set CalcBridge_BuildIdToWbsFromData = d

End Function


'------------------------------------------------------------------------------
' FR: Ajoute le warning Analytics lorsque l'ordre topologique est incomplet.
' EN: Adds the Analytics warning when the topological order is incomplete.
'------------------------------------------------------------------------------
Public Sub CalcBridge_AddAnalyticsTopologyWarning(ByVal consoleMessages As Collection)

    CalcBridge_AddOrShowConsoleMessage consoleMessages, "WARNING", _
        "Analytics non calculées : ordre topologique incomplet." & vbCrLf & _
        "-> vérifier les cycles ou la reconstruction tbl_LOGIC_LINKS.", _
        "Analytics not calculated: incomplete topological order." & vbCrLf & _
        "-> check cycles or tbl_LOGIC_LINKS rebuild."

End Sub


'------------------------------------------------------------------------------
' FR: Construit la map Id To WBS From WBS a partir des donnees fournies par l'appelant.
' EN: Builds the Id To WBS From WBS map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildIdToWbsFromWBS(ByVal tblWBS As ListObject) As Object

    Dim mapWBS As Object
    Dim arrWBS As Variant
    Dim idToWbs As Object
    Dim r As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    Set idToWbs = CreateObject("Scripting.Dictionary")

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildIdToWbsFromWBS = idToWbs
        Exit Function
    End If

    arrWBS = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arrWBS, 1)
        idVal = Trim$(CStr(arrWBS(r, mapWBS("ID"))))
        wbsVal = NormalizeWBS(arrWBS(r, mapWBS("WBS")))

        If idVal <> "" Then
            idToWbs(idVal) = wbsVal
        End If
    Next r

    Set BuildIdToWbsFromWBS = idToWbs

End Function



