Attribute VB_Name = "mod_GanttSimulationRenderFlow"
Option Explicit

'===============================================================================
' MODULE : mod_GanttSimulationRenderFlow
' DOMAINE / DOMAIN : Gantt
'
' FR
' Relie les demandes de rendu simulation au refresh Gantt sans posseder le renderer.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Connects simulation render requests to Gantt refresh without owning the renderer.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : GanttLive_AbortTestEngine, GanttLive_SafeExit_TestEngine
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================




'=====================================================
' HELPERS – TEST RENDER FLOW
'=====================================================

'------------------------------------------------------------------------------
' FR: Sort proprement d'une simulation echouee en relachant la preservation des inputs TEST.
' EN: Cleanly aborts a failed simulation by releasing TEST input preservation.
'------------------------------------------------------------------------------
Public Sub GanttLive_AbortTestEngine(ByVal wsGantt As Worksheet)

    SetGanttPreserveTestInputs False

    If Not wsGantt Is Nothing Then
        wsGantt.Activate
    End If

End Sub




'------------------------------------------------------------------------------
' FR: Nettoie les flags de preservation et de rendu live apres erreur TEST/SCENARIO.
' EN: Cleans preservation and live-render flags after a TEST/SCENARIO error.
'------------------------------------------------------------------------------
Public Sub GanttLive_SafeExit_TestEngine()

    SetGanttPreserveTestInputs False
    GanttLive_ClearTestRenderRequest

End Sub
