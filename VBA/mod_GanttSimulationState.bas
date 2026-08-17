Attribute VB_Name = "mod_GanttSimulationState"
Option Explicit

'===============================================================================
' MODULE : mod_GanttSimulationState
' DOMAINE / DOMAIN : Gantt
'
' FR
' Possede l'etat runtime du domaine et ses transitions explicites.
' Ne persiste ni ne recalcule les donnees metier sauf mention contraire.
'
' EN
' Owns domain runtime state and its explicit transitions.
' Does not persist or recalculate business data unless stated otherwise.
'
' CONTRATS / CONTRACTS : GanttLive_RequestTestRender, GanttLive_RequestScenarioRender, GanttLive_ClearTestRenderRequest, GanttLive_IsTestRenderRequested, GanttLive_GetPendingRenderMode, GanttLive_SetActiveSimulationMode, GanttLive_ClearActiveSimulationMode, GanttLive_GetActiveSimulationMode
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private gPendingRenderMode As String
Private gActiveSimulationMode As String


'=====================================================
' One-shot render request.
' The test engine raises this flag just before calling
' Refresh_Gantt. Refresh_Gantt consumes it and clears it.
' This avoids stale simulation staying active on later
' normal refreshes.
'=====================================================
'------------------------------------------------------------------------------
' FR: Demande au prochain refresh GANTT d'utiliser le rendu TEST une seule fois.
' EN: Requests the next GANTT refresh to use TEST rendering once.
'------------------------------------------------------------------------------
Public Sub GanttLive_RequestTestRender()
    gPendingRenderMode = "TEST"
End Sub

'------------------------------------------------------------------------------
' FR: Demande au prochain refresh GANTT d'utiliser le rendu SCENARIO une seule fois.
' EN: Requests the next GANTT refresh to use SCENARIO rendering once.
'------------------------------------------------------------------------------
Public Sub GanttLive_RequestScenarioRender()
    gPendingRenderMode = "SCENARIO"
End Sub

'------------------------------------------------------------------------------
' FR: Efface la demande de rendu live en attente pour eviter un overlay stale.
' EN: Clears the pending live render request to avoid a stale overlay.
'------------------------------------------------------------------------------
Public Sub GanttLive_ClearTestRenderRequest()
    gPendingRenderMode = ""
End Sub

'------------------------------------------------------------------------------
' FR: Indique si un refresh GANTT doit consommer un mode de rendu live en attente.
' EN: Returns whether a GANTT refresh must consume a pending live render mode.
'------------------------------------------------------------------------------
Public Function GanttLive_IsTestRenderRequested() As Boolean
    GanttLive_IsTestRenderRequested = (Trim$(gPendingRenderMode) <> "")
End Function

'------------------------------------------------------------------------------
' FR: Retourne le mode de rendu live en attente normalise pour le renderer GANTT.
' EN: Returns the normalized pending live render mode for the GANTT renderer.
'------------------------------------------------------------------------------
Public Function GanttLive_GetPendingRenderMode() As String
    GanttLive_GetPendingRenderMode = UCase$(Trim$(gPendingRenderMode))
End Function



'------------------------------------------------------------------------------
' FR: Memorise le mode de simulation actif consomme par le rendu et le workflow lock.
' EN: Stores the active simulation mode consumed by rendering and lock workflow.
'------------------------------------------------------------------------------
Public Sub GanttLive_SetActiveSimulationMode(ByVal modeName As String)
    gActiveSimulationMode = UCase$(Trim$(modeName))
End Sub

'------------------------------------------------------------------------------
' FR: Efface le mode de simulation actif.
' EN: Clears the active simulation mode.
'------------------------------------------------------------------------------
Public Sub GanttLive_ClearActiveSimulationMode()
    gActiveSimulationMode = ""
    GanttTestAnalyticsSnapshot_Clear
End Sub

'------------------------------------------------------------------------------
' FR: Retourne le mode de simulation actif normalise.
' EN: Returns the normalized active simulation mode.
'------------------------------------------------------------------------------
Public Function GanttLive_GetActiveSimulationMode() As String
    GanttLive_GetActiveSimulationMode = UCase$(Trim$(gActiveSimulationMode))
End Function

'------------------------------------------------------------------------------
' FR: Indique si le mode SCENARIO est actuellement actif.
' EN: Returns whether SCENARIO mode is currently active.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' FR: Indique si une projection de simulation runtime doit etre abandonnee.
' EN: Returns whether a runtime simulation projection must be abandoned.
'------------------------------------------------------------------------------
Public Function GanttLive_HasActiveSimulationRuntimeState() As Boolean

    GanttLive_HasActiveSimulationRuntimeState = _
        (Trim$(GanttLive_GetPendingRenderMode()) <> "") Or _
        (Trim$(GanttLive_GetActiveSimulationMode()) <> "") Or _
        GanttTestAnalyticsSnapshot_IsCurrent() Or _
        GetGanttPreserveTestInputs()

End Function

Public Function GanttLive_IsScenarioActive() As Boolean
    GanttLive_IsScenarioActive = (UCase$(Trim$(gActiveSimulationMode)) = "SCENARIO")
End Function

'------------------------------------------------------------------------------
' FR: Indique si le mode TEST est actuellement actif.
' EN: Returns whether TEST mode is currently active.
'------------------------------------------------------------------------------
Public Function GanttLive_IsLiveTestActive() As Boolean
    GanttLive_IsLiveTestActive = (UCase$(Trim$(gActiveSimulationMode)) = "TEST")
End Function

