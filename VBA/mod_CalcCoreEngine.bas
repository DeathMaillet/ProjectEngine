Attribute VB_Name = "mod_CalcCoreEngine"
Option Explicit

'===============================================================================
' MODULE : mod_CalcCoreEngine
' DOMAINE / DOMAIN : Core Calculation
'
' FR
' Execute le calcul planning unique sur les taches feuilles, les contraintes, les dependances et le post-process LOE.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Runs the single planning calculation over leaf tasks, constraints, dependencies and LOE post-processing.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Run_Calc_Core
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

'------------------------------------------------------------------------------
' FR:
' Orchestre le calcul Core en memoire: prepare les index, nettoie le scope de
' recalcul, calcule les taches feuilles en ordre topologique, traite les LOE,
' propage les erreurs bloquantes puis consolide les summaries.
'
' EN:
' Orchestrates the in-memory Core calculation: builds indexes, clears the
' recalculation scope, computes leaf tasks in topological order, processes LOE
' tasks, propagates blocking errors, and rolls up summary dates.
'
' Entrees / Inputs:
' - dataArr avec les colonnes Core deja mappees dans mapCol.
' - linksBySuccId prepare par le Bridge, indexe par successeur.
' - recalcScope optionnel pour un recalcul partiel.
' - dictionnaires optionnels de diagnostics dependency/constraint/cascade.
'
' Sorties / Outputs:
' - Met a jour dans dataArr les dates calculees, durees, flags et messages.
' - Renseigne les diagnostics structures passes par le caller.
' - Preserve les sorties hors scope lors d'un recalcul partiel.
'
' Appele par / Called by:
' - Les adaptateurs Bridge/production qui preparent les donnees Core.
' - Les workflows TEST/SCENARIO qui executent le moteur en memoire.
'
' Notes:
' - Point d'entree stable du moteur Core; aucune lecture/ecriture feuille.
' - Les taches LOE sont exclues du topo standard puis calculees en post-process.
'------------------------------------------------------------------------------
Public Sub Run_Calc_Core( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal linksBySuccId As Object, _
    Optional ByVal recalcScope As Object, _
    Optional ByVal dependencyDiagnostics As Object, _
    Optional ByVal constraintDiagnostics As Object, _
    Optional ByVal cascadeDiagnostics As Object)

    Dim perfScope As clsPerfScope

    Dim requiredCols As Variant

    Dim rowById As Object
    Dim parentIds As Object
    Dim directChildrenById As Object
    Dim childrenByPred As Object
    Dim validIds As Object
    Dim loeIds As Object
    Dim indegree As Object
    Dim topoOrder As Collection

    Dim calcStartById As Object
    Dim calcFinishById As Object
    Dim blockingErrors As Object

    Dim currentId As Variant
    Dim rowIdx As Long
    Dim shouldCompute As Boolean
    Dim isPartialMode As Boolean

    Set perfScope = Profiler_BeginScope("Run_Calc_Core", "Core Calculation")

    requiredCols = Array( _
        "ID", _
        "WBS", _
        "Task Name", _
        "ParentID", _
        "IsSummary", _
        "Task Type", _
        "Cal", _
        "Actual Start", _
        "Actual Finish", _
        "Forecast Start", _
        "Forecast Finish", _
        "Baseline Start", _
        "Baseline Duration", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration", _
        "Error flag", _
        "ErrorMsg" _
    )

    Core_RequireColumns mapCol, requiredCols, "Run_Calc_Core"

    isPartialMode = Not (recalcScope Is Nothing)

    Set rowById = Core_BuildRowById(dataArr, mapCol)
    Set parentIds = Core_BuildParentIds(dataArr, mapCol, rowById)
    Set directChildrenById = Core_BuildDirectChildrenById(dataArr, mapCol, rowById)
    Set childrenByPred = Core_BuildChildrenByPred(rowById, linksBySuccId)

    Set loeIds = Core_BuildLOEIds(dataArr, mapCol, rowById)

    Set validIds = Core_BuildValidLeafIds(rowById, parentIds)
    Core_RemoveIdsFromDictionary validIds, loeIds

    Set indegree = Core_BuildIndegree(validIds, linksBySuccId)
    Set topoOrder = Core_TopoSortLeafNetwork(validIds, childrenByPred, indegree)

    Set calcStartById = CreateObject("Scripting.Dictionary")
    Set calcFinishById = CreateObject("Scripting.Dictionary")
    Set blockingErrors = CreateObject("Scripting.Dictionary")

    If isPartialMode Then
        Core_LoadExistingCalcOutputs dataArr, mapCol, rowById, calcStartById, calcFinishById
        Core_ClearCalcOutputs_ForScope dataArr, mapCol, rowById, recalcScope, calcStartById, calcFinishById
    Else
        Core_ClearAllCalcOutputs dataArr, mapCol
    End If

    Core_ValidateLOEAsNonPredecessor dataArr, mapCol, rowById, linksBySuccId, loeIds, blockingErrors

    If Core_HasTopoFailure(topoOrder, validIds) Then
        Core_MarkTopoFailure dataArr, mapCol, validIds, rowById, linksBySuccId, blockingErrors
        GoTo SafeExit
    End If

    For Each currentId In topoOrder

        shouldCompute = False

        If isPartialMode Then
            If recalcScope.Exists(CStr(currentId)) Then
                shouldCompute = True
            End If
        Else
            shouldCompute = True
        End If

        If shouldCompute Then
            Core_ComputeOneLeafTask _
                dataArr, mapCol, CStr(currentId), rowById, linksBySuccId, _
                calcStartById, calcFinishById, blockingErrors, dependencyDiagnostics, constraintDiagnostics, cascadeDiagnostics
        End If

    Next currentId

    Core_ApplyLOEPostProcess dataArr, mapCol, rowById, linksBySuccId, loeIds, _
                             calcStartById, calcFinishById, blockingErrors

    If blockingErrors.Count > 0 Then
        Core_PropagateBlockingErrors dataArr, mapCol, blockingErrors, childrenByPred, rowById, constraintDiagnostics, dependencyDiagnostics, cascadeDiagnostics
    End If

    For Each currentId In calcStartById.Keys
        If rowById.Exists(CStr(currentId)) Then
            rowIdx = CLng(rowById(CStr(currentId)))
            Core_SetCalcTriplet dataArr, rowIdx, mapCol, _
                calcStartById(CStr(currentId)), _
                calcFinishById(CStr(currentId)), _
                NormalizeCalendarType(Core_GetVal(dataArr, rowIdx, mapCol, "Cal"))
        End If
    Next currentId

    Core_RollupSummaryDates dataArr, mapCol, rowById, directChildrenById, parentIds

SafeExit:
End Sub


'------------------------------------------------------------------------------
' FR:
' Recharge les dates calculees deja presentes dans dataArr afin qu'un recalcul
' partiel puisse reutiliser les predecesseurs et successeurs hors scope.
'
' EN:
' Reloads calculated dates already present in dataArr so a partial recalculation
' can reuse predecessors and successors that are outside the scope.
'
' Entrees / Inputs:
' - dataArr, mapCol et rowById du run Core courant.
' - Dictionnaires calcStartById et calcFinishById a alimenter.
'
' Sorties / Outputs:
' - Ajoute les dates Calculated Start/Finish existantes dans les caches Core.
'
' Appele par / Called by:
' - Run_Calc_Core en mode recalcul partiel.
'
' Notes:
' - Ne calcule rien; cette procedure ne fait que rehydrater l'etat existant.
'------------------------------------------------------------------------------
Private Sub Core_LoadExistingCalcOutputs( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal calcStartById As Object, _
    ByVal calcFinishById As Object)

    Dim perfScope As clsPerfScope

    Dim idVal As Variant
    Dim rowIdx As Long
    Dim calcStart As Variant
    Dim calcFinish As Variant

    Set perfScope = Profiler_BeginScope("Core_LoadExistingCalcOutputs", "Array")

    For Each idVal In rowById.Keys

        rowIdx = CLng(rowById(CStr(idVal)))

        calcStart = Core_GetVal(dataArr, rowIdx, mapCol, "Calculated Start")
        calcFinish = Core_GetVal(dataArr, rowIdx, mapCol, "Calculated Finish")

        If HasValue(calcStart) Then
            calcStartById(CStr(idVal)) = calcStart
        End If

        If HasValue(calcFinish) Then
            calcFinishById(CStr(idVal)) = calcFinish
        End If

    Next idVal

End Sub


'------------------------------------------------------------------------------
' FR:
' Calcule une tache feuille standard a partir de ses dates terrain, de ses
' contraintes et de ses liens predecesseurs; enregistre les blocages metier
' des que la tache devient non calculable.
'
' EN:
' Computes one standard leaf task from actual/forecast dates, constraints, and
' predecessor links; records business blocking errors as soon as the task cannot
' be scheduled consistently.
'
' Entrees / Inputs:
' - currentId, dataArr, mapCol et rowById pour lire la ligne de tache.
' - linksBySuccId, directChildrenById et calcStart/FinishById pour les dependances.
' - Dictionnaires de diagnostics dependency/constraint et erreurs bloquantes.
'
' Sorties / Outputs:
' - Ajoute start/finish calcules dans calcStartById et calcFinishById.
' - Ajoute Error flag/ErrorMsg et diagnostics en cas de violation.
'
' Appele par / Called by:
' - Run_Calc_Core pendant le parcours topologique des taches feuilles.
'
' Notes:
' - Gere FS/SS/FF, lag, sources summary, contraintes start/finish et dates
'   actual/forecast.
' - Les LOE ne passent pas par ce calcul; elles sont traitees ensuite.
'------------------------------------------------------------------------------
Private Sub Core_ComputeOneLeafTask( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal taskId As String, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal calcStartById As Object, _
    ByVal calcFinishById As Object, _
    ByVal blockingErrors As Object, _
    Optional ByVal dependencyDiagnostics As Object, _
    Optional ByVal constraintDiagnostics As Object, _
    Optional ByVal cascadeDiagnostics As Object)

    Dim perfScope As clsPerfScope

    Dim rowIdx As Long

    Dim baselineStart As Variant
    Dim baselineDuration As Variant
    Dim actualStart As Variant
    Dim actualFinish As Variant
    Dim forecastStart As Variant
    Dim forecastFinish As Variant

    Dim predAllowedStart As Variant
    Dim predAllowedFinish As Variant
    Dim constraintAllowedStart As Variant
    Dim constraintLatestStart As Variant
    Dim constraintMustStart As Variant
    Dim constraintAllowedFinish As Variant
    Dim constraintLatestFinish As Variant
    Dim constraintMustFinish As Variant
    Dim allowedStart As Variant
    Dim allowedFinish As Variant
    Dim constraintActive As String
    Dim startConstraintType As String
    Dim finishConstraintType As String
    Dim startConstraintDate As Variant
    Dim finishConstraintDate As Variant
    Dim hasExplicitStart As Boolean

    Dim normalAllowedStart As Variant
    Dim summaryAllowedStart As Variant
    Dim summaryStartBySource As Object
    Dim summaryStartDiagBySource As Object
    Dim normalAllowedStartDiag As Object
    Dim summaryAllowedStartDiag As Object
    Dim predAllowedStartDiag As Object
    Dim candidateDiag As Object

    Dim sourceStart As Variant
    Dim sourceFinish As Variant
    Dim calcStart As Variant
    Dim calcFinish As Variant
    Dim mustFinishStart As Variant

    Dim oneLink As Variant
    Dim predId As String
    Dim linkType As String
    Dim lagVal As Double
    Dim summarySourceId As String

    Dim candidateStart As Variant
    Dim candidateFinish As Variant
    Dim parentKey As Variant

    Dim effectiveDuration As Variant
    Dim taskTypeVal As String
    Dim calType As String

    Set perfScope = Profiler_BeginScope("Core_ComputeOneLeafTask", "Core Leaf")

    If Not rowById.Exists(taskId) Then Exit Sub
    rowIdx = CLng(rowById(taskId))

    calType = NormalizeCalendarType(Core_GetVal(dataArr, rowIdx, mapCol, "Cal"))
    baselineStart = Core_GetVal(dataArr, rowIdx, mapCol, "Baseline Start")
    baselineDuration = Core_GetVal(dataArr, rowIdx, mapCol, "Baseline Duration")
    actualStart = Core_GetVal(dataArr, rowIdx, mapCol, "Actual Start")
    actualFinish = Core_GetVal(dataArr, rowIdx, mapCol, "Actual Finish")
    forecastStart = Core_GetVal(dataArr, rowIdx, mapCol, "Forecast Start")
    forecastFinish = Core_GetVal(dataArr, rowIdx, mapCol, "Forecast Finish")
    constraintActive = UCase$(Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Constraint Active"))))
    startConstraintType = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Start Constraint Type")))
    startConstraintDate = Core_GetVal(dataArr, rowIdx, mapCol, "Start Constraint Date")
    finishConstraintType = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Finish Constraint Type")))
    finishConstraintDate = Core_GetVal(dataArr, rowIdx, mapCol, "Finish Constraint Date")

    effectiveDuration = baselineDuration

    'Milestone rule:
    'A milestone without explicit duration is treated as 1 day by the core.
    'This is a calculation default only; it does not write 1 back to WBS/CALC input fields.
    If Not HasValue(effectiveDuration) Then
        taskTypeVal = Core_NormalizeTaskType(Core_GetVal(dataArr, rowIdx, mapCol, "Task Type"))

        If taskTypeVal = "MILESTONE" Then
            effectiveDuration = 1
        End If
    End If

    predAllowedStart = Empty
    predAllowedFinish = Empty
    constraintAllowedStart = Empty
    constraintLatestStart = Empty
    constraintMustStart = Empty
    constraintAllowedFinish = Empty
    constraintLatestFinish = Empty
    constraintMustFinish = Empty
    mustFinishStart = Empty
    allowedStart = Empty
    allowedFinish = Empty

    normalAllowedStart = Empty
    summaryAllowedStart = Empty
    Set summaryStartBySource = CreateObject("Scripting.Dictionary")
    Set summaryStartDiagBySource = CreateObject("Scripting.Dictionary")

    If Not linksBySuccId Is Nothing Then
        If linksBySuccId.Exists(taskId) Then
            For Each oneLink In linksBySuccId(taskId)

                predId = Core_GetLinkPredId(oneLink)
                linkType = Core_GetLinkType(oneLink)
                lagVal = Core_GetLinkLag(oneLink)
                summarySourceId = Core_GetLinkSummarySourceId(oneLink)

                If predId = "" Then
                    Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                        "Missing predecessor"
                    Exit Sub
                End If

                If Not rowById.Exists(predId) Then
                    Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                        "Missing predecessor: ID " & predId
                    Exit Sub
                End If

                If blockingErrors.Exists(predId) Then
                    Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                        "Blocked by predecessor error: ID " & predId
                    Exit Sub
                End If

                Select Case linkType

                    Case "SS"
                        If calcStartById.Exists(predId) Then
                            candidateStart = ApplyLag(calcStartById(predId), lagVal, calType, "SS")

                            Set candidateDiag = Core_CreateStartDependencyDiagnostic( _
                                taskId, predId, linkType, lagVal, candidateStart, calcStartById(predId), "START", summarySourceId)

                            If summarySourceId <> "" Then
                                If summaryStartBySource.Exists(summarySourceId) Then
                                    If CDbl(candidateStart) < CDbl(summaryStartBySource(summarySourceId)) Then
                                        summaryStartBySource(summarySourceId) = candidateStart
                                        Set summaryStartDiagBySource(summarySourceId) = candidateDiag
                                    End If
                                Else
                                    summaryStartBySource(summarySourceId) = candidateStart
                                    Set summaryStartDiagBySource(summarySourceId) = candidateDiag
                                End If
                            Else
                                If Not HasValue(normalAllowedStart) Or CDbl(candidateStart) > CDbl(normalAllowedStart) Then
                                    Set normalAllowedStartDiag = candidateDiag
                                End If
                                normalAllowedStart = Core_MaxDateIfBoth(normalAllowedStart, candidateStart)
                            End If
                        End If

                    Case "FF"
                        If calcFinishById.Exists(predId) Then
                            candidateFinish = ApplyLag(calcFinishById(predId), lagVal, calType, "FF")
                            predAllowedFinish = Core_MaxDateIfBoth(predAllowedFinish, candidateFinish)
                        End If

                    Case "FS", ""
                        If calcFinishById.Exists(predId) Then
                            candidateStart = ApplyLag(calcFinishById(predId), lagVal, calType, "FS")
                            Set candidateDiag = Core_CreateStartDependencyDiagnostic( _
                                taskId, predId, IIf(linkType = "", "FS", linkType), lagVal, candidateStart, calcFinishById(predId), "FINISH", summarySourceId)
                            If Not HasValue(normalAllowedStart) Or CDbl(candidateStart) > CDbl(normalAllowedStart) Then
                                Set normalAllowedStartDiag = candidateDiag
                            End If
                            normalAllowedStart = Core_MaxDateIfBoth(normalAllowedStart, candidateStart)
                        End If

                    Case Else
                        Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                            "Unsupported link type: " & linkType
                        Exit Sub

                End Select

            Next oneLink
        End If
    End If

    For Each parentKey In summaryStartBySource.Keys
        If Not HasValue(summaryAllowedStart) Or CDbl(summaryStartBySource(CStr(parentKey))) > CDbl(summaryAllowedStart) Then
            If summaryStartDiagBySource.Exists(CStr(parentKey)) Then
                Set summaryAllowedStartDiag = summaryStartDiagBySource(CStr(parentKey))
            End If
        End If
        summaryAllowedStart = Core_MaxDateIfBoth(summaryAllowedStart, summaryStartBySource(CStr(parentKey)))
    Next parentKey

    predAllowedStart = Core_MaxDateIfBoth(normalAllowedStart, summaryAllowedStart)
    If HasValue(predAllowedStart) Then
        If HasValue(summaryAllowedStart) And CDbl(predAllowedStart) = CDbl(summaryAllowedStart) Then
            Set predAllowedStartDiag = summaryAllowedStartDiag
        Else
            Set predAllowedStartDiag = normalAllowedStartDiag
        End If
    End If

    If constraintActive = "YES" Then
        If startConstraintType = "Start No Earlier Than" Then
            constraintAllowedStart = startConstraintDate
        ElseIf startConstraintType = "Start No Later Than" Then
            constraintLatestStart = startConstraintDate
        ElseIf startConstraintType = "Must Start On" Then
            constraintMustStart = startConstraintDate
        ElseIf startConstraintType <> "" Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Type de contrainte debut non reconnu", _
                    "Unknown start constraint type")
            Exit Sub
        End If

        If finishConstraintType = "Finish No Earlier Than" Then
            constraintAllowedFinish = finishConstraintDate
        ElseIf finishConstraintType = "Finish No Later Than" Then
            constraintLatestFinish = finishConstraintDate
        ElseIf finishConstraintType = "Must Finish On" Then
            constraintMustFinish = finishConstraintDate
        ElseIf finishConstraintType <> "" Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Type de contrainte fin non reconnu", _
                    "Unknown finish constraint type")
            Exit Sub
        End If
    End If

    allowedStart = Core_MaxDateIfBoth(predAllowedStart, constraintAllowedStart)
    allowedFinish = Core_MaxDateIfBoth(predAllowedFinish, constraintAllowedFinish)
    hasExplicitStart = HasValue(actualStart) Or HasValue(forecastStart) Or HasValue(constraintAllowedStart) Or HasValue(constraintMustStart)

    If HasValue(actualStart) And HasValue(predAllowedStart) Then
        If CDbl(actualStart) < CDbl(predAllowedStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                "Actual Start violates dependencies"
            Exit Sub
        End If
    End If

    If HasValue(actualStart) And HasValue(constraintAllowedStart) Then
        If CDbl(actualStart) < CDbl(constraintAllowedStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Actual Start avant contrainte debut", _
                    "Actual Start is before start constraint")
            Exit Sub
        End If
    End If

    If HasValue(actualStart) And HasValue(constraintLatestStart) Then
        If CDbl(actualStart) > CDbl(constraintLatestStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Actual Start apres contrainte debut max", _
                    "Actual Start is after latest start constraint")
            Exit Sub
        End If
    End If

    If HasValue(actualStart) And HasValue(constraintMustStart) Then
        If CDbl(actualStart) <> CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Actual Start different de contrainte Must Start On", _
                    "Actual Start differs from Must Start On constraint")
            Exit Sub
        End If
    End If

    If HasValue(actualFinish) And HasValue(predAllowedFinish) Then
        If CDbl(actualFinish) < CDbl(predAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                "Actual Finish violates finish constraints"
            Exit Sub
        End If
    End If

    If HasValue(actualFinish) And HasValue(constraintAllowedFinish) Then
        If CDbl(actualFinish) < CDbl(constraintAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Actual Finish avant contrainte fin", _
                    "Actual Finish is before finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(actualFinish) And HasValue(constraintLatestFinish) Then
        If CDbl(actualFinish) > CDbl(constraintLatestFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Actual Finish apres contrainte fin max", _
                    "Actual Finish is after latest finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(actualFinish) And HasValue(constraintMustFinish) Then
        If CDbl(actualFinish) <> CDbl(constraintMustFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Actual Finish different de contrainte Must Finish On", _
                    "Actual Finish differs from Must Finish On constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastStart) And HasValue(predAllowedStart) Then
        If CDbl(forecastStart) < CDbl(predAllowedStart) Then
            Core_RecordForecastStartDependencyDiagnostic dependencyDiagnostics, taskId, forecastStart, predAllowedStart, predAllowedStartDiag
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                "Forecast Start violates dependencies"
            Exit Sub
        End If
    End If

    If HasValue(forecastStart) And HasValue(constraintAllowedStart) Then
        If CDbl(forecastStart) < CDbl(constraintAllowedStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Forecast Start avant contrainte debut", _
                    "Forecast Start is before start constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastStart) And HasValue(constraintLatestStart) Then
        If CDbl(forecastStart) > CDbl(constraintLatestStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Forecast Start apres contrainte debut max", _
                    "Forecast Start is after latest start constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastStart) And HasValue(constraintMustStart) Then
        If CDbl(forecastStart) <> CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Forecast Start different de contrainte Must Start On", _
                    "Forecast Start differs from Must Start On constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastFinish) And HasValue(predAllowedFinish) Then
        If CDbl(forecastFinish) < CDbl(predAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Forecast Finish avant contrainte fin amont", _
                    "Forecast Finish is before upstream finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastFinish) And HasValue(constraintAllowedFinish) Then
        If CDbl(forecastFinish) < CDbl(constraintAllowedFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Forecast Finish avant contrainte fin", _
                    "Forecast Finish is before finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastFinish) And HasValue(constraintLatestFinish) Then
        If CDbl(forecastFinish) > CDbl(constraintLatestFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Forecast Finish apres contrainte fin max", _
                    "Forecast Finish is after latest finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(forecastFinish) And HasValue(constraintMustFinish) Then
        If CDbl(forecastFinish) <> CDbl(constraintMustFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Forecast Finish different de contrainte Must Finish On", _
                    "Forecast Finish differs from Must Finish On constraint")
            Exit Sub
        End If
    End If

    If HasValue(constraintMustFinish) Then
        If Not HasValue(effectiveDuration) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                "Baseline Duration missing"
            Exit Sub
        End If

        mustFinishStart = SubtractWorkingDays(constraintMustFinish, effectiveDuration, calType)

        If HasValue(allowedFinish) Then
            If CDbl(allowedFinish) > CDbl(constraintMustFinish) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                        "Calculated Finish different de contrainte Must Finish On", _
                        "Calculated Finish differs from Must Finish On constraint")
                Exit Sub
            End If
        End If

        If HasValue(allowedStart) Then
            If CDbl(allowedStart) > CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                        "Calculated Start avant reseau impose par contrainte Must Finish On", _
                        "Calculated Start is before network allowed start due to Must Finish On constraint")
                Exit Sub
            End If
        End If

        If HasValue(constraintMustStart) Then
            If CDbl(constraintMustStart) <> CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                        "Duree incompatible avec contraintes Must Start On / Must Finish On", _
                        "Duration is incompatible with Must Start On / Must Finish On constraints")
                Exit Sub
            End If
        End If

        If HasValue(actualStart) Then
            If CDbl(actualStart) <> CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                        "Actual Start different du debut impose par Must Finish On", _
                        "Actual Start differs from start implied by Must Finish On constraint")
                Exit Sub
            End If
        End If

        If HasValue(forecastStart) Then
            If CDbl(forecastStart) <> CDbl(mustFinishStart) Then
                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                    Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                        "Forecast Start different du debut impose par Must Finish On", _
                        "Forecast Start differs from start implied by Must Finish On constraint")
                Exit Sub
            End If
        End If
    End If

    If HasValue(constraintMustStart) And HasValue(allowedStart) Then
        If CDbl(allowedStart) > CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Calculated Start different de contrainte Must Start On", _
                    "Calculated Start differs from Must Start On constraint")
            Exit Sub
        End If
    End If

    sourceStart = Core_GetSourceStart(dataArr, rowIdx, mapCol)

    If HasValue(constraintMustFinish) Then
        calcStart = mustFinishStart
    ElseIf HasValue(constraintMustStart) Then
        calcStart = constraintMustStart
    ElseIf HasValue(sourceStart) Then
        calcStart = Core_MaxDateIfBoth(sourceStart, allowedStart)
    Else
        calcStart = Empty

        If HasValue(allowedFinish) And HasValue(effectiveDuration) Then
            calcStart = SubtractWorkingDays(allowedFinish, effectiveDuration, calType)
        End If

        If HasValue(allowedStart) Then
            calcStart = Core_MaxDateIfBoth(calcStart, allowedStart)
        End If
    End If

    If Not HasValue(calcStart) Then
        Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
            "Start date not computable"
        Exit Sub
    End If

    If HasValue(constraintLatestStart) Then
        If CDbl(calcStart) > CDbl(constraintLatestStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Calculated Start apres contrainte debut max", _
                    "Calculated Start is after latest start constraint")
            Exit Sub
        End If
    End If

    If HasValue(constraintMustStart) Then
        If CDbl(calcStart) <> CDbl(constraintMustStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Calculated Start different de contrainte Must Start On", _
                    "Calculated Start differs from Must Start On constraint")
            Exit Sub
        End If
    End If

    sourceFinish = Core_GetSourceFinish(dataArr, rowIdx, mapCol)

    If HasValue(constraintMustFinish) Then
        calcFinish = constraintMustFinish
    ElseIf HasValue(sourceFinish) Then
        calcFinish = sourceFinish
    Else
        If Not HasValue(effectiveDuration) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                "Baseline Duration missing"
            Exit Sub
        End If

        calcFinish = AddWorkingDays(calcStart, effectiveDuration, calType)
    End If

    If HasValue(allowedFinish) Then
        If CDbl(calcFinish) < CDbl(allowedFinish) Then
            calcFinish = allowedFinish

            If Not hasExplicitStart Then
                If HasValue(effectiveDuration) Then
                    calcStart = SubtractWorkingDays(calcFinish, effectiveDuration, calType)

                    If HasValue(allowedStart) Then
                        If CDbl(calcStart) < CDbl(allowedStart) Then calcStart = allowedStart
                    End If
                End If
            End If
        End If
    End If

    If HasValue(constraintLatestFinish) Then
        If CDbl(calcFinish) > CDbl(constraintLatestFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Calculated Finish apres contrainte fin max", _
                    "Calculated Finish is after latest finish constraint")
            Exit Sub
        End If
    End If

    If HasValue(constraintMustFinish) Then
        If CDbl(calcFinish) <> CDbl(constraintMustFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
                Core_BuildConstraintDiagnosticMessage(dataArr, rowIdx, mapCol, constraintDiagnostics, taskId, startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, mustFinishStart, _
                    "Calculated Finish different de contrainte Must Finish On", _
                    "Calculated Finish differs from Must Finish On constraint")
            Exit Sub
        End If
    End If

    If CDbl(calcFinish) < CDbl(calcStart) Then
        Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, taskId, _
            "Finish before start"
        Exit Sub
    End If

    calcStartById(taskId) = calcStart
    calcFinishById(taskId) = calcFinish

End Sub

'------------------------------------------------------------------------------
' FR: Enregistre le diagnostic structure d'une contrainte puis renvoie le message bilingue stocke dans ErrorMsg.
' EN: Records the structured constraint diagnostic and returns the bilingual message stored in ErrorMsg.
'------------------------------------------------------------------------------
Private Function Core_BuildConstraintDiagnosticMessage( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal constraintDiagnostics As Object, _
    ByVal taskId As String, _
    ByVal startConstraintType As String, _
    ByVal startConstraintDate As Variant, _
    ByVal finishConstraintType As String, _
    ByVal finishConstraintDate As Variant, _
    ByVal actualStart As Variant, _
    ByVal actualFinish As Variant, _
    ByVal forecastStart As Variant, _
    ByVal forecastFinish As Variant, _
    ByVal calcStart As Variant, _
    ByVal calcFinish As Variant, _
    ByVal predAllowedStart As Variant, _
    ByVal predAllowedFinish As Variant, _
    ByVal allowedStart As Variant, _
    ByVal allowedFinish As Variant, _
    ByVal effectiveDuration As Variant, _
    ByVal mustFinishStart As Variant, _
    ByVal frText As String, _
    ByVal enText As String) As String

    Core_RecordConstraintDiagnostic constraintDiagnostics, dataArr, rowIdx, mapCol, taskId, _
        startConstraintType, startConstraintDate, finishConstraintType, finishConstraintDate, _
        actualStart, actualFinish, forecastStart, forecastFinish, calcStart, calcFinish, _
        predAllowedStart, predAllowedFinish, allowedStart, allowedFinish, effectiveDuration, _
        mustFinishStart, frText, enText

    Core_BuildConstraintDiagnosticMessage = Core_BuildConstraintCoreMessage(dataArr, rowIdx, mapCol, frText, enText)

End Function


'------------------------------------------------------------------------------
' FR: Construit l'objet diagnostic detaille d'une violation de contrainte a partir du contexte de calcul de la tache.
' EN: Builds the detailed diagnostic object for a constraint violation from the task calculation context.
'------------------------------------------------------------------------------
Private Sub Core_RecordConstraintDiagnostic( _
    ByVal constraintDiagnostics As Object, _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal taskId As String, _
    ByVal startConstraintType As String, _
    ByVal startConstraintDate As Variant, _
    ByVal finishConstraintType As String, _
    ByVal finishConstraintDate As Variant, _
    ByVal actualStart As Variant, _
    ByVal actualFinish As Variant, _
    ByVal forecastStart As Variant, _
    ByVal forecastFinish As Variant, _
    ByVal calcStart As Variant, _
    ByVal calcFinish As Variant, _
    ByVal predAllowedStart As Variant, _
    ByVal predAllowedFinish As Variant, _
    ByVal allowedStart As Variant, _
    ByVal allowedFinish As Variant, _
    ByVal effectiveDuration As Variant, _
    ByVal mustFinishStart As Variant, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim diag As Object
    Dim txt As String
    Dim sideVal As String
    Dim typeVal As String
    Dim dateVal As Variant
    Dim checkedField As String
    Dim checkedValue As Variant
    Dim relationVal As String
    Dim allowedValue As Variant

    If constraintDiagnostics Is Nothing Then Exit Sub

    txt = UCase$(frText & " " & enText)

    If InStr(1, txt, "FINISH", vbTextCompare) > 0 Or _
       InStr(1, txt, "FIN", vbTextCompare) > 0 Then
        sideVal = "FINISH"
        typeVal = Trim$(finishConstraintType)
        dateVal = finishConstraintDate
    Else
        sideVal = "START"
        typeVal = Trim$(startConstraintType)
        dateVal = startConstraintDate
    End If

    If InStr(1, txt, "MUST FINISH ON", vbTextCompare) > 0 Then
        sideVal = "FINISH"
        typeVal = "Must Finish On"
        dateVal = finishConstraintDate
    ElseIf InStr(1, txt, "MUST START ON", vbTextCompare) > 0 Then
        sideVal = "START"
        typeVal = "Must Start On"
        dateVal = startConstraintDate
    ElseIf InStr(1, txt, "UPSTREAM FINISH", vbTextCompare) > 0 Or _
           InStr(1, txt, "FIN AMONT", vbTextCompare) > 0 Then
        sideVal = "FINISH"
        typeVal = "Upstream finish constraint"
        dateVal = predAllowedFinish
    End If

    If InStr(1, txt, "ACTUAL START", vbTextCompare) > 0 Then
        checkedField = "Actual Start"
        checkedValue = actualStart
    ElseIf InStr(1, txt, "ACTUAL FINISH", vbTextCompare) > 0 Then
        checkedField = "Actual Finish"
        checkedValue = actualFinish
    ElseIf InStr(1, txt, "FORECAST START", vbTextCompare) > 0 Then
        checkedField = "Forecast Start"
        checkedValue = forecastStart
    ElseIf InStr(1, txt, "FORECAST FINISH", vbTextCompare) > 0 Then
        checkedField = "Forecast Finish"
        checkedValue = forecastFinish
    ElseIf InStr(1, txt, "CALCULATED START", vbTextCompare) > 0 Then
        checkedField = "Calculated Start"
        checkedValue = calcStart
    ElseIf InStr(1, txt, "CALCULATED FINISH", vbTextCompare) > 0 Then
        checkedField = "Calculated Finish"
        checkedValue = calcFinish
    ElseIf InStr(1, txt, "DUREE", vbTextCompare) > 0 Or _
           InStr(1, txt, "DURATION", vbTextCompare) > 0 Then
        checkedField = "Effective Duration"
        checkedValue = effectiveDuration
    Else
        checkedField = sideVal
        If sideVal = "FINISH" Then checkedValue = calcFinish Else checkedValue = calcStart
    End If

    If InStr(1, txt, "BEFORE", vbTextCompare) > 0 Or _
       InStr(1, txt, "AVANT", vbTextCompare) > 0 Then
        relationVal = ">="
    ElseIf InStr(1, txt, "AFTER", vbTextCompare) > 0 Or _
           InStr(1, txt, "APRES", vbTextCompare) > 0 Then
        relationVal = "<="
    Else
        relationVal = "="
    End If

    If typeVal = "Upstream finish constraint" Then
        allowedValue = predAllowedFinish
    ElseIf sideVal = "FINISH" Then
        If InStr(1, txt, "NO EARLIER", vbTextCompare) > 0 Or InStr(1, txt, "BEFORE", vbTextCompare) > 0 Or InStr(1, txt, "AVANT", vbTextCompare) > 0 Then
            allowedValue = Core_FirstValue(allowedFinish, finishConstraintDate)
        ElseIf InStr(1, txt, "NO LATER", vbTextCompare) > 0 Or InStr(1, txt, "LATEST", vbTextCompare) > 0 Or InStr(1, txt, "MAX", vbTextCompare) > 0 Or InStr(1, txt, "AFTER", vbTextCompare) > 0 Or InStr(1, txt, "APRES", vbTextCompare) > 0 Then
            allowedValue = finishConstraintDate
        Else
            allowedValue = finishConstraintDate
        End If
    Else
        If InStr(1, txt, "NO EARLIER", vbTextCompare) > 0 Or InStr(1, txt, "BEFORE", vbTextCompare) > 0 Or InStr(1, txt, "AVANT", vbTextCompare) > 0 Then
            allowedValue = Core_FirstValue(allowedStart, startConstraintDate)
        ElseIf InStr(1, txt, "NO LATER", vbTextCompare) > 0 Or InStr(1, txt, "LATEST", vbTextCompare) > 0 Or InStr(1, txt, "MAX", vbTextCompare) > 0 Or InStr(1, txt, "AFTER", vbTextCompare) > 0 Or InStr(1, txt, "APRES", vbTextCompare) > 0 Then
            allowedValue = startConstraintDate
        ElseIf InStr(1, txt, "MUST FINISH", vbTextCompare) > 0 Then
            allowedValue = mustFinishStart
        Else
            allowedValue = startConstraintDate
        End If
    End If

    Set diag = CreateObject("Scripting.Dictionary")
    diag("TaskID") = CStr(taskId)
    diag("WBS") = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "WBS")))
    diag("TaskName") = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Task Name")))
    diag("ConstraintSide") = sideVal
    diag("ConstraintType") = typeVal
    diag("ConstraintDate") = dateVal
    diag("CheckedField") = checkedField
    diag("CheckedValue") = checkedValue
    diag("ExpectedOperator") = relationVal
    diag("AllowedValue") = allowedValue
    diag("ActualStart") = actualStart
    diag("ActualFinish") = actualFinish
    diag("ForecastStart") = forecastStart
    diag("ForecastFinish") = forecastFinish
    diag("CalculatedStart") = calcStart
    diag("CalculatedFinish") = calcFinish
    diag("PredAllowedStart") = predAllowedStart
    diag("PredAllowedFinish") = predAllowedFinish
    diag("AllowedStart") = allowedStart
    diag("AllowedFinish") = allowedFinish
    diag("EffectiveDuration") = effectiveDuration
    diag("MustFinishImpliedStart") = mustFinishStart
    diag("RawFR") = frText
    diag("RawEN") = enText

    Set constraintDiagnostics.Item(CStr(taskId)) = diag

End Sub


'------------------------------------------------------------------------------
' FR: Retourne la premiere valeur renseignee entre une valeur prioritaire et une valeur de repli.
' EN: Returns the first populated value between a primary value and a fallback value.
'------------------------------------------------------------------------------
Private Function Core_FirstValue(ByVal primaryValue As Variant, ByVal fallbackValue As Variant) As Variant

    If HasValue(primaryValue) Then
        Core_FirstValue = primaryValue
    Else
        Core_FirstValue = fallbackValue
    End If

End Function

'------------------------------------------------------------------------------
' FR: Construit le message Core FR/EN minimal associe a une violation de contrainte, avec ID, WBS et nom de tache.
' EN: Builds the minimal FR/EN Core message for a constraint violation, including task ID, WBS, and name.
'------------------------------------------------------------------------------
Private Function Core_BuildConstraintCoreMessage( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal frText As String, _
    ByVal enText As String) As String

    Dim idVal As String
    Dim wbsVal As String
    Dim taskName As String

    idVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "ID")))
    wbsVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "WBS")))
    taskName = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Task Name")))

    Core_BuildConstraintCoreMessage = _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "ID : " & idVal & vbCrLf & _
        "WBS : " & wbsVal & vbCrLf & _
        "Task : " & taskName & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText & vbCrLf & vbCrLf & _
        "ID: " & idVal & vbCrLf & _
        "WBS: " & wbsVal & vbCrLf & _
        "Task: " & taskName

End Function

'------------------------------------------------------------------------------
' FR:
' Enregistre une erreur bloquante sur une tache et la reporte immediatement
' dans les colonnes d'erreur de la ligne Core.
'
' EN:
' Registers a blocking error on a task and immediately mirrors it to the Core
' row error columns.
'
' Entrees / Inputs:
' - currentId, rowIdx et message metier deja formule.
' - dataArr/mapCol pour mettre a jour la ligne.
' - blockingErrors pour memoriser la racine de blocage.
'
' Sorties / Outputs:
' - blockingErrors(currentId) = message.
' - Error flag/ErrorMsg mis a jour dans dataArr.
'
' Appele par / Called by:
' - Core_ComputeOneLeafTask, Core_MarkTopoFailure, Core_ValidateLOEAsNonPredecessor, Core_ApplyLOEPostProcess.
'
' Notes:
' - Point commun de creation des erreurs dures qui seront ensuite propagees.
'------------------------------------------------------------------------------
Private Sub Core_AddBlockingError( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal blockingErrors As Object, _
    ByVal taskId As String, _
    ByVal errText As String)

    blockingErrors(CStr(taskId)) = True
    Core_AddErrorMessage_Row dataArr, rowIdx, mapCol, errText, True

End Sub

'------------------------------------------------------------------------------
' FR:
' Transforme un echec de tri topologique en erreur bloquante sur toutes les
' taches feuilles valides, avec un diagnostic de cycle quand il est retrouvable.
'
' EN:
' Converts a topological sort failure into a blocking error on every valid leaf
' task, including a cycle diagnostic when one can be recovered.
'
' Entrees / Inputs:
' - validIds, linksBySuccId, rowById, dataArr et mapCol du run Core.
' - blockingErrors a completer.
'
' Sorties / Outputs:
' - Marque les taches impactees en erreur bloquante.
'
' Appele par / Called by:
' - Run_Calc_Core lorsque l'ordre topologique ne couvre pas tous les IDs valides.
'
' Notes:
' - Cette erreur stoppe le calcul normal car l'ordre de dependance n'est plus fiable.
'------------------------------------------------------------------------------
Private Sub Core_MarkTopoFailure( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal validIds As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal blockingErrors As Object)

    Dim taskId As Variant
    Dim rowIdx As Long
    Dim cycleMessage As String

    cycleMessage = Core_BuildTopoCycleDiagnosticMessage(dataArr, mapCol, validIds, rowById, linksBySuccId)
    If Trim$(cycleMessage) = "" Then cycleMessage = "Cycle detected"

    For Each taskId In validIds.Keys
        If rowById.Exists(CStr(taskId)) Then
            rowIdx = CLng(rowById(CStr(taskId)))
            blockingErrors(CStr(taskId)) = True
            Core_AddErrorMessage_Row dataArr, rowIdx, mapCol, cycleMessage, True
        End If
    Next taskId

End Sub

'------------------------------------------------------------------------------
' FR: Construit le message utilisateur pour un cycle logique detecte par le tri topologique.
' EN: Builds the user-facing message for a logical cycle detected by the topological sort.
'------------------------------------------------------------------------------
Private Function Core_BuildTopoCycleDiagnosticMessage( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal validIds As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object) As String

    Dim cycleEdges As Collection
    Dim cycleText As String

    On Error GoTo SafeFallback

    Set cycleEdges = Core_FindFirstTopoCycleEdges(validIds, linksBySuccId)

    If cycleEdges Is Nothing Then GoTo SafeFallback
    If cycleEdges.Count = 0 Then GoTo SafeFallback

    cycleText = Core_FormatTopoCycleEdges(dataArr, mapCol, rowById, cycleEdges)
    If Trim$(cycleText) = "" Then GoTo SafeFallback

    Core_BuildTopoCycleDiagnosticMessage = _
        "Cycle detected" & vbCrLf & _
        BiMsg( _
            "Cycle logique détecté :" & vbCrLf & vbCrLf & cycleText & vbCrLf & vbCrLf & _
            "-> corriger un des liens de cette boucle dans Predecessors WBS.", _
            "Logic cycle detected:" & vbCrLf & vbCrLf & cycleText & vbCrLf & vbCrLf & _
            "-> fix one of the links in this loop in Predecessors WBS.")
    Exit Function

SafeFallback:
    Core_BuildTopoCycleDiagnosticMessage = "Cycle detected"

End Function

'------------------------------------------------------------------------------
' FR: Recherche dans le graphe des liens le premier cycle entre taches feuilles valides.
' EN: Searches the dependency graph for the first cycle between valid leaf tasks.
'------------------------------------------------------------------------------
Private Function Core_FindFirstTopoCycleEdges( _
    ByVal validIds As Object, _
    ByVal linksBySuccId As Object) As Collection

    Dim childrenByPred As Object
    Dim state As Object
    Dim stackIndex As Object
    Dim stackIds As Collection
    Dim stackEdges As Collection
    Dim cycleEdges As Collection
    Dim idKey As Variant
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String
    Dim edgeInfo As Object

    Set childrenByPred = CreateObject("Scripting.Dictionary")
    Set state = CreateObject("Scripting.Dictionary")
    Set stackIndex = CreateObject("Scripting.Dictionary")
    Set stackIds = New Collection
    Set stackEdges = New Collection
    Set cycleEdges = New Collection

    If validIds Is Nothing Then
        Set Core_FindFirstTopoCycleEdges = cycleEdges
        Exit Function
    End If

    For Each idKey In validIds.Keys
        Set childrenByPred(CStr(idKey)) = New Collection
        state(CStr(idKey)) = 0
    Next idKey

    If Not linksBySuccId Is Nothing Then
        For Each succId In linksBySuccId.Keys
            If validIds.Exists(CStr(succId)) Then
                For Each oneLink In linksBySuccId(CStr(succId))
                    predId = Core_GetLinkPredId(oneLink)
                    If predId <> "" Then
                        If validIds.Exists(predId) Then
                            Set edgeInfo = CreateObject("Scripting.Dictionary")
                            edgeInfo("FromId") = predId
                            edgeInfo("ToId") = CStr(succId)
                            edgeInfo("LinkType") = Core_GetLinkType(oneLink)
                            edgeInfo("Lag") = Core_GetLinkLag(oneLink)
                            childrenByPred(predId).Add edgeInfo
                        End If
                    End If
                Next oneLink
            End If
        Next succId
    End If

    For Each idKey In validIds.Keys
        If CLng(state(CStr(idKey))) = 0 Then
            Core_DFS_FindTopoCycleEdges _
                CStr(idKey), childrenByPred, state, stackIndex, stackIds, stackEdges, cycleEdges
            If cycleEdges.Count > 0 Then Exit For
        End If
    Next idKey

    Set Core_FindFirstTopoCycleEdges = cycleEdges

End Function

'------------------------------------------------------------------------------
' FR: Parcourt le graphe en profondeur et extrait les aretes du premier cycle rencontre.
' EN: Traverses the graph depth-first and extracts the edges of the first encountered cycle.
'------------------------------------------------------------------------------
Private Sub Core_DFS_FindTopoCycleEdges( _
    ByVal currentId As String, _
    ByVal childrenByPred As Object, _
    ByVal state As Object, _
    ByVal stackIndex As Object, _
    ByVal stackIds As Collection, _
    ByVal stackEdges As Collection, _
    ByVal cycleEdges As Collection)

    Dim edgeInfo As Variant
    Dim childId As String
    Dim startIdx As Long
    Dim i As Long

    If cycleEdges.Count > 0 Then Exit Sub

    state(currentId) = 1
    stackIds.Add currentId
    stackIndex(currentId) = stackIds.Count

    If childrenByPred.Exists(currentId) Then
        For Each edgeInfo In childrenByPred(currentId)
            If cycleEdges.Count > 0 Then Exit For

            childId = CStr(edgeInfo("ToId"))
            If Not state.Exists(childId) Then state(childId) = 0

            If CLng(state(childId)) = 0 Then
                stackEdges.Add edgeInfo
                Core_DFS_FindTopoCycleEdges _
                    childId, childrenByPred, state, stackIndex, stackIds, stackEdges, cycleEdges
                If cycleEdges.Count > 0 Then Exit For
                If stackEdges.Count > 0 Then stackEdges.Remove stackEdges.Count

            ElseIf CLng(state(childId)) = 1 Then
                If stackIndex.Exists(childId) Then
                    startIdx = CLng(stackIndex(childId))
                    For i = startIdx To stackEdges.Count
                        cycleEdges.Add stackEdges(i)
                    Next i
                    cycleEdges.Add edgeInfo
                    Exit For
                End If
            End If
        Next edgeInfo
    End If

    state(currentId) = 2

    If stackIds.Count > 0 Then
        If CStr(stackIds(stackIds.Count)) = currentId Then
            stackIds.Remove stackIds.Count
        End If
    End If

    If stackIndex.Exists(currentId) Then stackIndex.Remove currentId

End Sub

'------------------------------------------------------------------------------
' FR: Transforme les aretes d'un cycle en texte lisible avec taches, type de lien et lag.
' EN: Formats cycle edges as readable text with tasks, link type, and lag.
'------------------------------------------------------------------------------
Private Function Core_FormatTopoCycleEdges( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal cycleEdges As Collection) As String

    Dim result As String
    Dim edgeInfo As Variant
    Dim fromId As String
    Dim toId As String
    Dim linkType As String
    Dim lagVal As Double

    result = ""

    For Each edgeInfo In cycleEdges
        fromId = CStr(edgeInfo("FromId"))
        toId = CStr(edgeInfo("ToId"))
        linkType = UCase$(Trim$(CStr(edgeInfo("LinkType"))))
        lagVal = CDbl(edgeInfo("Lag"))

        If result <> "" Then result = result & vbCrLf & vbCrLf

        result = result & _
            Core_CycleTaskLabel(dataArr, mapCol, rowById, fromId) & vbCrLf & _
            "--" & linkType & Core_FormatCycleLag(lagVal) & "-->" & vbCrLf & _
            Core_CycleTaskLabel(dataArr, mapCol, rowById, toId)
    Next edgeInfo

    Core_FormatTopoCycleEdges = result

End Function

'------------------------------------------------------------------------------
' FR:
' Cree le diagnostic source qui explique quelle dependance impose une date de
' debut minimale plus tardive que la date demandee.
'
' EN:
' Creates the source diagnostic explaining which dependency imposes a minimum
' start date later than the requested date.
'
' Entrees / Inputs:
' - Task cible, predecesseur bloquant, type de lien, lag et dates candidates.
' - Source eventuelle quand le predecesseur bloquant provient d'un summary.
'
' Sorties / Outputs:
' - Dictionnaire diagnostic pret a etre stocke ou enrichi.
'
' Appele par / Called by:
' - Core_ComputeOneLeafTask pendant l'analyse des liens de debut.
'
' Notes:
' - Ce diagnostic est la base des messages dependency et des cascades aval.
'------------------------------------------------------------------------------
Private Function Core_CreateStartDependencyDiagnostic( _
    ByVal taskId As String, _
    ByVal predId As String, _
    ByVal linkType As String, _
    ByVal lagVal As Double, _
    ByVal candidateStart As Variant, _
    ByVal predecessorDate As Variant, _
    ByVal predecessorDateKind As String, _
    Optional ByVal summarySourceId As String = "") As Object

    Dim diag As Object

    Set diag = CreateObject("Scripting.Dictionary")

    diag("TaskID") = CStr(taskId)
    diag("BlockingPredecessorID") = CStr(predId)
    diag("BlockingLinkType") = UCase$(Trim$(linkType))
    diag("BlockingLag") = CDbl(lagVal)
    diag("BlockingCandidateDate") = candidateStart
    diag("BlockingPredecessorDate") = predecessorDate
    diag("BlockingPredecessorDateKind") = UCase$(Trim$(predecessorDateKind))
    diag("ExpandedFrom") = CStr(summarySourceId)

    Set Core_CreateStartDependencyDiagnostic = diag

End Function


'------------------------------------------------------------------------------
' FR:
' Stocke le diagnostic structure lorsqu'une Forecast Start est anterieure a la
' date minimale imposee par les dependances.
'
' EN:
' Stores the structured diagnostic when a Forecast Start is earlier than the
' minimum start date imposed by dependencies.
'
' Entrees / Inputs:
' - Task cible, Forecast Start demandee, date minimale autorisee.
' - Diagnostic source de la dependance bloquante.
'
' Sorties / Outputs:
' - dependencyDiagnostics(taskId) avec les champs de cause et de date.
'
' Appele par / Called by:
' - Core_ComputeOneLeafTask au moment de valider Forecast Start.
'
' Notes:
' - Ne formule pas le message utilisateur; elle preserve les donnees auditables.
'------------------------------------------------------------------------------
Private Sub Core_RecordForecastStartDependencyDiagnostic( _
    ByVal dependencyDiagnostics As Object, _
    ByVal taskId As String, _
    ByVal requestedStart As Variant, _
    ByVal minimumAllowedStart As Variant, _
    ByVal sourceDiagnostic As Object)

    Dim diag As Object

    If dependencyDiagnostics Is Nothing Then Exit Sub
    If sourceDiagnostic Is Nothing Then Exit Sub

    Set diag = CreateObject("Scripting.Dictionary")

    diag("TaskID") = CStr(taskId)
    diag("RequestedStart") = requestedStart
    diag("MinimumAllowedStart") = minimumAllowedStart

    If sourceDiagnostic.Exists("BlockingPredecessorID") Then diag("BlockingPredecessorID") = sourceDiagnostic("BlockingPredecessorID")
    If sourceDiagnostic.Exists("BlockingLinkType") Then diag("BlockingLinkType") = sourceDiagnostic("BlockingLinkType")
    If sourceDiagnostic.Exists("BlockingLag") Then diag("BlockingLag") = sourceDiagnostic("BlockingLag")
    If sourceDiagnostic.Exists("BlockingCandidateDate") Then diag("BlockingCandidateDate") = sourceDiagnostic("BlockingCandidateDate")
    If sourceDiagnostic.Exists("BlockingPredecessorDate") Then diag("BlockingPredecessorDate") = sourceDiagnostic("BlockingPredecessorDate")
    If sourceDiagnostic.Exists("BlockingPredecessorDateKind") Then diag("BlockingPredecessorDateKind") = sourceDiagnostic("BlockingPredecessorDateKind")
    If sourceDiagnostic.Exists("ExpandedFrom") Then diag("ExpandedFrom") = sourceDiagnostic("ExpandedFrom")

    Set dependencyDiagnostics.Item(CStr(taskId)) = diag

End Sub

'------------------------------------------------------------------------------
' FR: Construit le libelle humain d'une tache de cycle a partir du WBS, du nom ou de l'ID.
' EN: Builds the human-readable label for a cycle task from WBS, task name, or ID.
'------------------------------------------------------------------------------
Private Function Core_CycleTaskLabel( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal taskId As String) As String

    Dim rowIdx As Long
    Dim wbsVal As String
    Dim taskNameVal As String

    If rowById Is Nothing Then GoTo Fallback
    If Not rowById.Exists(taskId) Then GoTo Fallback

    rowIdx = CLng(rowById(taskId))
    wbsVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "WBS")))
    taskNameVal = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Task Name")))

    If wbsVal <> "" And taskNameVal <> "" Then
        Core_CycleTaskLabel = wbsVal & " " & taskNameVal
    ElseIf wbsVal <> "" Then
        Core_CycleTaskLabel = wbsVal
    ElseIf taskNameVal <> "" Then
        Core_CycleTaskLabel = taskNameVal
    Else
        Core_CycleTaskLabel = "ID " & taskId
    End If
    Exit Function

Fallback:
    Core_CycleTaskLabel = "ID " & taskId

End Function

'------------------------------------------------------------------------------
' FR: Formate le lag d'un lien de cycle avec signe explicite et decimal standardise.
' EN: Formats a cycle link lag with an explicit sign and standardized decimal separator.
'------------------------------------------------------------------------------
Private Function Core_FormatCycleLag(ByVal lagVal As Double) As String

    Dim lagText As String

    If Abs(lagVal) < 0.0000001 Then
        Core_FormatCycleLag = "+0"
        Exit Function
    End If

    lagText = Replace$(Format$(Abs(lagVal), "0.##"), ",", ".")

    If lagVal > 0 Then
        Core_FormatCycleLag = "+" & lagText
    Else
        Core_FormatCycleLag = "-" & lagText
    End If

End Function

'------------------------------------------------------------------------------
' FR:
' Propage les erreurs bloquantes racines vers les successeurs afin que toute la
' chaine aval soit marquee non calculable.
'
' EN:
' Propagates root blocking errors to successors so the full downstream chain is
' marked as non-computable.
'
' Entrees / Inputs:
' - blockingErrors racines, childrenByPred, rowById, dataArr et mapCol.
' - cascadeDiagnostics optionnel pour tracer la propagation.
'
' Sorties / Outputs:
' - Error flag/ErrorMsg sur les descendants impactes.
' - cascadeDiagnostics rempli avec la racine et le parent de propagation.
'
' Appele par / Called by:
' - Run_Calc_Core apres le calcul standard et le post-process LOE.
'
' Notes:
' - Les descendants recoivent un message de chaine bloquee s'ils n'ont pas deja
'   leur propre erreur racine.
'------------------------------------------------------------------------------
Private Sub Core_PropagateBlockingErrors( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal blockingErrors As Object, _
    ByVal childrenByPred As Object, _
    ByVal rowById As Object, _
    Optional ByVal constraintDiagnostics As Object, _
    Optional ByVal dependencyDiagnostics As Object, _
    Optional ByVal cascadeDiagnostics As Object)

    Dim allErrorIds As Object
    Dim taskId As Variant
    Dim rowIdx As Long

    Set allErrorIds = CreateObject("Scripting.Dictionary")

    For Each taskId In blockingErrors.Keys
        allErrorIds(CStr(taskId)) = True
    Next taskId

    For Each taskId In blockingErrors.Keys
        Core_PropagateErrorToChildren CStr(taskId), childrenByPred, allErrorIds
        Core_RecordCascadeDiagnosticsForRoot CStr(taskId), childrenByPred, constraintDiagnostics, dependencyDiagnostics, cascadeDiagnostics
    Next taskId

    For Each taskId In allErrorIds.Keys
        If rowById.Exists(CStr(taskId)) Then
            rowIdx = CLng(rowById(CStr(taskId)))
            Core_SetErrorFlag_Row dataArr, rowIdx, mapCol, True

            If Not blockingErrors.Exists(CStr(taskId)) Then
                Core_AddErrorMessage_Row dataArr, rowIdx, mapCol, _
                    "Blocked by predecessor chain", True
            End If
        End If
    Next taskId

End Sub


'------------------------------------------------------------------------------
' FR:
' Cree les diagnostics de cascade pour tous les descendants d'une erreur racine,
' en conservant la cause initiale et le parent qui propage le blocage.
'
' EN:
' Creates cascade diagnostics for all descendants of one root error, preserving
' the initial cause and the parent that propagated the block.
'
' Entrees / Inputs:
' - rootId, childrenByPred, allErrorIds et diagnostics racines disponibles.
'
' Sorties / Outputs:
' - cascadeDiagnostics enrichi pour chaque descendant impacte.
'
' Appele par / Called by:
' - Core_PropagateBlockingErrors, une fois par erreur racine.
'
' Notes:
' - Parcours en largeur pour produire une chaine de responsabilite lisible.
'------------------------------------------------------------------------------
Private Sub Core_RecordCascadeDiagnosticsForRoot( _
    ByVal rootId As String, _
    ByVal childrenByPred As Object, _
    ByVal constraintDiagnostics As Object, _
    ByVal dependencyDiagnostics As Object, _
    ByVal cascadeDiagnostics As Object)

    Dim q As Collection
    Dim currentId As Variant
    Dim childId As Variant
    Dim parentById As Object
    Dim diag As Object
    Dim rootType As String

    If cascadeDiagnostics Is Nothing Then Exit Sub
    If childrenByPred Is Nothing Then Exit Sub

    If Not constraintDiagnostics Is Nothing Then
        If constraintDiagnostics.Exists(CStr(rootId)) Then rootType = "CONSTRAINT"
    End If

    If rootType = "" Then
        If Not dependencyDiagnostics Is Nothing Then
            If dependencyDiagnostics.Exists(CStr(rootId)) Then rootType = "DEPENDENCY"
        End If
    End If

    If rootType = "" Then rootType = "CORE"

    Set q = New Collection
    Set parentById = CreateObject("Scripting.Dictionary")
    q.Add CStr(rootId)
    parentById(CStr(rootId)) = ""

    Do While q.Count > 0
        currentId = q(1)
        q.Remove 1

        If childrenByPred.Exists(CStr(currentId)) Then
            For Each childId In childrenByPred(CStr(currentId))
                If Not parentById.Exists(CStr(childId)) Then
                    parentById(CStr(childId)) = CStr(currentId)
                    q.Add CStr(childId)

                    Set diag = CreateObject("Scripting.Dictionary")
                    diag("TaskID") = CStr(childId)
                    diag("RootErrorID") = CStr(rootId)
                    diag("RootErrorType") = rootType
                    diag("ParentPropagatedFrom") = CStr(currentId)

                    If rootType = "CONSTRAINT" Then
                        If Not constraintDiagnostics Is Nothing Then
                            If constraintDiagnostics.Exists(CStr(rootId)) Then Set diag("RootConstraintDiagnostic") = constraintDiagnostics(CStr(rootId))
                        End If
                    ElseIf rootType = "DEPENDENCY" Then
                        If Not dependencyDiagnostics Is Nothing Then
                            If dependencyDiagnostics.Exists(CStr(rootId)) Then Set diag("RootDependencyDiagnostic") = dependencyDiagnostics(CStr(rootId))
                        End If
                    End If

                    Set cascadeDiagnostics.Item(CStr(childId)) = diag
                End If
            Next childId
        End If
    Loop

End Sub

'------------------------------------------------------------------------------
' FR: Efface les sorties calculees uniquement pour les IDs du recalcul partiel et retire leurs caches en memoire.
' EN: Clears calculated outputs only for partial recalculation IDs and removes their in-memory cache entries.
'------------------------------------------------------------------------------
Private Sub Core_ClearCalcOutputs_ForScope( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal recalcScope As Object, _
    Optional ByVal calcStartById As Object, _
    Optional ByVal calcFinishById As Object)

    Dim perfScope As clsPerfScope

    Dim idKey As Variant
    Dim taskId As String
    Dim rowIdx As Long

    Set perfScope = Profiler_BeginScope("Core_ClearCalcOutputs_ForScope", "Array")

    If recalcScope Is Nothing Then Exit Sub
    If rowById Is Nothing Then Exit Sub

    For Each idKey In recalcScope.Keys

        taskId = CStr(idKey)

        If rowById.Exists(taskId) Then
            rowIdx = CLng(rowById(taskId))
            Core_ClearCalcOutputs_Row dataArr, rowIdx, mapCol
        End If

        If Not calcStartById Is Nothing Then
            If calcStartById.Exists(taskId) Then calcStartById.Remove taskId
        End If

        If Not calcFinishById Is Nothing Then
            If calcFinishById.Exists(taskId) Then calcFinishById.Remove taskId
        End If

    Next idKey

End Sub


'------------------------------------------------------------------------------
' FR: Identifie les taches Level of Effort afin de les exclure du calcul topo standard et de les traiter a part.
' EN: Identifies Level of Effort tasks so they can be excluded from standard topological calculation and handled separately.
'------------------------------------------------------------------------------
Private Function Core_BuildLOEIds( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object) As Object

    Dim perfScope As clsPerfScope

    Dim d As Object
    Dim idVal As Variant
    Dim rowIdx As Long
    Dim taskType As String

    Set perfScope = Profiler_BeginScope("Core_BuildLOEIds", "Dictionary")

    Set d = CreateObject("Scripting.Dictionary")

    If rowById Is Nothing Then
        Set Core_BuildLOEIds = d
        Exit Function
    End If

    For Each idVal In rowById.Keys

        rowIdx = CLng(rowById(CStr(idVal)))
        taskType = Core_NormalizeTaskType(Core_GetVal(dataArr, rowIdx, mapCol, "Task Type"))

        If taskType = "LEVEL OF EFFORT" Then
            d(CStr(idVal)) = True
        End If

    Next idVal

    Set Core_BuildLOEIds = d

End Function

'------------------------------------------------------------------------------
' FR: Normalise les libelles de type de tache consommes par le Core.
' EN: Normalizes task type labels consumed by the Core.
'------------------------------------------------------------------------------
Private Function Core_NormalizeTaskType(ByVal rawValue As Variant) As String

    Dim s As String

    s = UCase$(Trim$(CStr(rawValue)))

    Select Case s

        Case "", "TASK", "STANDARD", "NORMAL"
            Core_NormalizeTaskType = "TASK"

        Case "MILESTONE", "MS", "JALON"
            Core_NormalizeTaskType = "MILESTONE"

        Case "LEVEL OF EFFORT", "LOE", "LEVEL-OF-EFFORT", "LEVEL_OF_EFFORT"
            Core_NormalizeTaskType = "LEVEL OF EFFORT"

        Case Else
            Core_NormalizeTaskType = s

    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retire d'un dictionnaire de travail tous les IDs presents dans un second dictionnaire.
' EN: Removes from a working dictionary every ID present in a second dictionary.
'------------------------------------------------------------------------------
Private Sub Core_RemoveIdsFromDictionary( _
    ByVal targetDict As Object, _
    ByVal idsToRemove As Object)

    Dim idVal As Variant

    If targetDict Is Nothing Then Exit Sub
    If idsToRemove Is Nothing Then Exit Sub

    For Each idVal In idsToRemove.Keys
        If targetDict.Exists(CStr(idVal)) Then
            targetDict.Remove CStr(idVal)
        End If
    Next idVal

End Sub

'------------------------------------------------------------------------------
' FR: Valide Core Validate LOE As Non Predecessor et applique la politique d'erreur definie par le composant.
' EN: Validates Core Validate LOE As Non Predecessor and applies the component's defined failure policy.
'------------------------------------------------------------------------------

Private Sub Core_ValidateLOEAsNonPredecessor( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal loeIds As Object, _
    ByVal blockingErrors As Object)

    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String
    Dim loeRow As Long
    Dim succRow As Long

    If linksBySuccId Is Nothing Then Exit Sub
    If loeIds Is Nothing Then Exit Sub
    If blockingErrors Is Nothing Then Exit Sub

    For Each succId In linksBySuccId.Keys

        If linksBySuccId.Exists(CStr(succId)) Then

            For Each oneLink In linksBySuccId(CStr(succId))

                predId = Core_GetLinkPredId(oneLink)

                If predId <> "" Then
                    If loeIds.Exists(predId) Then

                        If rowById.Exists(predId) Then
                            loeRow = CLng(rowById(predId))
                            Core_AddBlockingError dataArr, loeRow, mapCol, blockingErrors, predId, _
                                "LOE cannot be used as predecessor"
                        End If

                        If rowById.Exists(CStr(succId)) Then
                            succRow = CLng(rowById(CStr(succId)))
                            Core_AddBlockingError dataArr, succRow, mapCol, blockingErrors, CStr(succId), _
                                "Blocked by invalid LOE predecessor: ID " & predId
                        End If

                    End If
                End If

            Next oneLink

        End If

    Next succId

End Sub


'------------------------------------------------------------------------------
' FR:
' Calcule les taches Level of Effort apres les taches standard, a partir de
' liens SS pour le debut et FF pour la fin, puis neutralise les champs analytiques
' qui ne s'appliquent pas a une LOE.
'
' EN:
' Computes Level of Effort tasks after standard tasks, using SS links for start
' and FF links for finish, then clears analytics fields that do not apply to LOE.
'
' Entrees / Inputs:
' - loeIds, dataArr, mapCol, rowById, linksBySuccId et directChildrenById.
' - calcStartById/calcFinishById issus du calcul standard.
' - blockingErrors pour signaler les LOE invalides ou non calculables.
'
' Sorties / Outputs:
' - Dates calculees LOE ajoutees aux caches Core.
' - Erreurs bloquantes sur LOE si les liens SS/FF requis sont absents ou invalides.
' - Champs Driving Logic, Critical Path, Total/Free Float et REX nettoyes si presents.
'
' Appele par / Called by:
' - Run_Calc_Core apres le parcours topologique des taches non-LOE.
'
' Notes:
' - Les LOE doivent avoir au moins un SS et un FF; FS et autres types sont refuses.
' - Les sources summary SS sont regroupees pour choisir le debut le plus tot du groupe.
'------------------------------------------------------------------------------
Private Sub Core_ApplyLOEPostProcess( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object, _
    ByVal loeIds As Object, _
    ByVal calcStartById As Object, _
    ByVal calcFinishById As Object, _
    ByVal blockingErrors As Object)

    Dim perfScope As clsPerfScope

    Dim loeId As Variant
    Dim rowIdx As Long

    Dim ssCount As Long
    Dim ffCount As Long
    Dim invalidCount As Long

    Dim oneLink As Variant
    Dim linkType As String
    Dim predId As String
    Dim lagVal As Double
    Dim summarySourceId As String

    Dim loeStart As Variant
    Dim loeFinish As Variant
    Dim normalSSStart As Variant
    Dim summarySSStart As Variant
    Dim summarySSStartBySource As Object
    Dim parentKey As Variant

    Dim candidateStart As Variant
    Dim candidateFinish As Variant
    Dim calType As String

    Set perfScope = Profiler_BeginScope("Core_ApplyLOEPostProcess", "Core LOE")

    If loeIds Is Nothing Then Exit Sub
    If loeIds.Count = 0 Then Exit Sub

    For Each loeId In loeIds.Keys

        If Not rowById.Exists(CStr(loeId)) Then GoTo NextLOE
        rowIdx = CLng(rowById(CStr(loeId)))
        calType = NormalizeCalendarType(Core_GetVal(dataArr, rowIdx, mapCol, "Cal"))

        ssCount = 0
        ffCount = 0
        invalidCount = 0

        loeStart = Empty
        loeFinish = Empty
        normalSSStart = Empty
        summarySSStart = Empty
        Set summarySSStartBySource = CreateObject("Scripting.Dictionary")

        If Not linksBySuccId Is Nothing Then
            If linksBySuccId.Exists(CStr(loeId)) Then

                For Each oneLink In linksBySuccId(CStr(loeId))

                    predId = Core_GetLinkPredId(oneLink)
                    linkType = Core_GetLinkType(oneLink)
                    lagVal = Core_GetLinkLag(oneLink)
                    summarySourceId = Core_GetLinkSummarySourceId(oneLink)

                    Select Case linkType

                        Case "SS"
                            ssCount = ssCount + 1

                            If predId = "" Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE SS predecessor missing"
                                GoTo NextLOE
                            End If

                            If Not rowById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE SS predecessor not found: ID " & predId
                                GoTo NextLOE
                            End If

                            If blockingErrors.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE blocked by SS predecessor error: ID " & predId
                                GoTo NextLOE
                            End If

                            If Not calcStartById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE SS predecessor start not available: ID " & predId
                                GoTo NextLOE
                            End If

                            candidateStart = ApplyLag(calcStartById(predId), lagVal, calType, "SS")

                            If summarySourceId <> "" Then
                                If summarySSStartBySource.Exists(summarySourceId) Then
                                    summarySSStartBySource(summarySourceId) = _
                                        Core_MinDateIfBoth(summarySSStartBySource(summarySourceId), candidateStart)
                                Else
                                    summarySSStartBySource(summarySourceId) = candidateStart
                                End If
                            Else
                                normalSSStart = Core_MaxDateIfBoth(normalSSStart, candidateStart)
                            End If

                        Case "FF"
                            ffCount = ffCount + 1

                            If predId = "" Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE FF predecessor missing"
                                GoTo NextLOE
                            End If

                            If Not rowById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE FF predecessor not found: ID " & predId
                                GoTo NextLOE
                            End If

                            If blockingErrors.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE blocked by FF predecessor error: ID " & predId
                                GoTo NextLOE
                            End If

                            If Not calcFinishById.Exists(predId) Then
                                Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                                    "LOE FF predecessor finish not available: ID " & predId
                                GoTo NextLOE
                            End If

                            candidateFinish = ApplyLag(calcFinishById(predId), lagVal, calType, "FF")
                            loeFinish = Core_MaxDateIfBoth(loeFinish, candidateFinish)

                        Case Else
                            invalidCount = invalidCount + 1

                    End Select

                Next oneLink

            End If
        End If

        If ssCount = 0 Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE must have at least one SS predecessor"
            GoTo NextLOE
        End If

        If ffCount = 0 Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE must have at least one FF predecessor"
            GoTo NextLOE
        End If

        If invalidCount > 0 Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE only supports SS and FF predecessors"
            GoTo NextLOE
        End If

        For Each parentKey In summarySSStartBySource.Keys
            summarySSStart = Core_MaxDateIfBoth(summarySSStart, summarySSStartBySource(CStr(parentKey)))
        Next parentKey

        loeStart = Core_MaxDateIfBoth(normalSSStart, summarySSStart)

        If Not HasValue(loeStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE start not computable"
            GoTo NextLOE
        End If

        If Not HasValue(loeFinish) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE finish not computable"
            GoTo NextLOE
        End If

        If CDbl(loeFinish) < CDbl(loeStart) Then
            Core_AddBlockingError dataArr, rowIdx, mapCol, blockingErrors, CStr(loeId), _
                "LOE finish before start"
            GoTo NextLOE
        End If

        calcStartById(CStr(loeId)) = loeStart
        calcFinishById(CStr(loeId)) = loeFinish

        If mapCol.Exists("Driving Logic") Then
            dataArr(rowIdx, mapCol("Driving Logic")) = "LOE"
        End If

        If mapCol.Exists("Critical Path") Then
            dataArr(rowIdx, mapCol("Critical Path")) = vbNullString
        End If

        If mapCol.Exists("Total Float") Then
            dataArr(rowIdx, mapCol("Total Float")) = vbNullString
        End If

        If mapCol.Exists("Free Float") Then
            dataArr(rowIdx, mapCol("Free Float")) = vbNullString
        End If

        If mapCol.Exists("Critical Path REX") Then
            dataArr(rowIdx, mapCol("Critical Path REX")) = vbNullString
        End If

        If mapCol.Exists("Total Float REX") Then
            dataArr(rowIdx, mapCol("Total Float REX")) = vbNullString
        End If

        If mapCol.Exists("Free Float REX") Then
            dataArr(rowIdx, mapCol("Free Float REX")) = vbNullString
        End If

NextLOE:
    Next loeId

End Sub

