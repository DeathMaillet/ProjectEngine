Attribute VB_Name = "mod_GanttSimulationResetService"
Option Explicit

'===============================================================================
' MODULE : mod_GanttSimulationResetService
' DOMAINE / DOMAIN : Gantt Simulation
'
' FR
' Propriétaire du contrat atomique qui quitte TEST ou SCENARIO et restaure la
' projection normale du dernier planning calculé. Ne lance aucun moteur.
'
' EN
' Owns the atomic contract that exits TEST or SCENARIO and restores the normal
' projection of the latest calculated planning state. Runs no engine.
'
' CONTRATS / CONTRACTS : GanttSimulation_ResetToNormal
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : GanttSimulation_ResetToNormal
'===============================================================================

'------------------------------------------------------------------------------
' FR:
' Efface toutes les données temporaires TEST/SCENARIO puis, si une simulation
' était effectivement rendue, restaure le Gantt normal sans recalcul.
' L'argument False permet aux orchestrateurs qui possèdent déjà leur refresh de
' réutiliser exactement le même nettoyage.
'
' EN:
' Clears all temporary TEST/SCENARIO data and, when a simulation was actually
' rendered, restores the normal Gantt without recalculation.
' Passing False lets orchestrators that already own a refresh reuse the exact
' same cleanup.
'------------------------------------------------------------------------------
Public Sub GanttSimulation_ResetToNormal( _
    Optional ByVal refreshDisplayIfNeeded As Boolean = True)

    Dim perfScope As clsPerfScope
    Dim workflowStarted As Boolean
    Dim dragPaused As Boolean
    Dim visualSimulationWasActive As Boolean
    Dim mutationCount As Long
    Dim refreshCount As Long
    Dim resetErrNumber As Long
    Dim resetErrDescription As String

    Set perfScope = Profiler_BeginScope("GanttSimulation_ResetToNormal", "Gantt Simulation")
    On Error GoTo SafeExit

    workflowStarted = EnsurePlanningWorkflowStarted("GanttSimulation_ResetToNormal")
    visualSimulationWasActive = GanttSimulation_HasRenderedSimulationState()

    If refreshDisplayIfNeeded And visualSimulationWasActive Then
        GanttDrag_PauseForLifecycle
        dragPaused = True
    End If

    mutationCount = mutationCount + GanttSheetLayout_ClearTestInputs()

    If GanttSimulation_ClearResultsIfAny() Then
        mutationCount = mutationCount + 1
    End If

    If Trim$(GanttLive_GetPendingRenderMode()) <> "" Then
        GanttLive_ClearTestRenderRequest
        mutationCount = mutationCount + 1
    End If

    If Trim$(GanttLive_GetActiveSimulationMode()) <> "" Or _
       GanttTestAnalyticsSnapshot_IsCurrent() Then
        GanttLive_ClearActiveSimulationMode
        mutationCount = mutationCount + 1
    End If

    If GetGanttPreserveTestInputs() Then
        SetGanttPreserveTestInputs False
        mutationCount = mutationCount + 1
    End If

    If refreshDisplayIfNeeded And visualSimulationWasActive Then
        Refresh_Gantt_DisplayOnly
        refreshCount = 1
    End If

SafeExit:
    resetErrNumber = Err.Number
    resetErrDescription = Err.Description

    Profiler_RecordOperation "GanttSimulationResetMutations", mutationCount, 0#
    Profiler_RecordOperation "GanttSimulationResetRefreshes", refreshCount, 0#

    If workflowStarted Then EndPlanningWorkflow

    If dragPaused Then
        On Error Resume Next
        GanttDrag_ReconcileWatchState
        On Error GoTo 0
    End If

    If resetErrNumber <> 0 Then
        CalcBridge_ShowSingleConsoleMessage _
            "STOP", _
            "Erreur pendant la réinitialisation du Gantt : " & resetErrDescription, _
            "Error while resetting the Gantt: " & resetErrDescription
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Indique si le renderer consomme encore une projection TEST ou SCENARIO.
' EN: Returns whether the renderer still consumes a TEST or SCENARIO projection.
'------------------------------------------------------------------------------
Private Function GanttSimulation_HasRenderedSimulationState() As Boolean

    GanttSimulation_HasRenderedSimulationState = _
        (Trim$(GanttLive_GetPendingRenderMode()) <> "") Or _
        (Trim$(GanttLive_GetActiveSimulationMode()) <> "") Or _
        GanttTestAnalyticsSnapshot_IsCurrent()

End Function
