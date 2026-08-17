Attribute VB_Name = "mod_GanttEngine"
Option Explicit

'===============================================================================
' MODULE : mod_GanttEngine
' DOMAIN : Gantt
'
' Owns the single product path that turns the current planning projection into a
' READY Gantt. User workflows choose only a data source and a scope; they do not
' own rendering, SVG commit, or READY.
'===============================================================================

Public Const GANTT_DATA_SOURCE_NORMAL As String = "NORMAL"
Public Const GANTT_DATA_SOURCE_TEST As String = "TEST"
Public Const GANTT_DATA_SOURCE_SCENARIO As String = "SCENARIO"

Public Const GANTT_UPDATE_SCOPE_NOOP As String = "NOOP"
Public Const GANTT_UPDATE_SCOPE_PARTIAL As String = "PARTIAL"
Public Const GANTT_UPDATE_SCOPE_INCREMENTAL As String = "INCREMENTAL"
Public Const GANTT_UPDATE_SCOPE_FULL As String = "FULL"
Public Const GANTT_UPDATE_SCOPE_DEFER As String = "DEFER"

Public Const GANTT_RENDER_INTENT_OFFSCREEN As String = "OFFSCREEN"
Public Const GANTT_RENDER_INTENT_SHOW As String = "SHOW"
Public Const GANTT_RENDER_INTENT_DEFER As String = "DEFER"

Public Function CommitGanttUpdate( _
    ByVal dataSource As String, _
    ByVal updateScope As String, _
    ByVal renderIntent As String, _
    Optional ByVal reason As String = "") As Boolean

    Dim perfScope As clsPerfScope
    Dim ws As Worksheet
    Dim normalizedSource As String
    Dim normalizedScope As String
    Dim normalizedIntent As String
    Dim requestedScope As String
    Dim escalationReason As String
    Dim coreEffectiveScope As String
    Dim finalizeReason As String
    Dim scopeAlreadyRendered As Boolean
    Dim ready As Boolean

    Set perfScope = Profiler_BeginScope("CommitGanttUpdate", "Gantt Engine")
    On Error GoTo Failed

    normalizedSource = UCase$(Trim$(dataSource))
    normalizedScope = UCase$(Trim$(updateScope))
    normalizedIntent = UCase$(Trim$(renderIntent))
    requestedScope = normalizedScope
    If normalizedScope = GANTT_UPDATE_SCOPE_PARTIAL Then normalizedScope = GANTT_UPDATE_SCOPE_INCREMENTAL

    Set ws = ThisWorkbook.Worksheets("GANTT")

    Select Case normalizedIntent
        Case GANTT_RENDER_INTENT_DEFER
            GanttDependencySvg_InvalidatePersistentCache ws, "CommitGanttUpdate|DEFER|" & reason
            GanttDeferredRender_MarkPending "CommitGanttUpdate|DEFER|" & reason
            Profiler_RecordOperation "GanttEngineDeferredCommits", 1, 0#
            CommitGanttUpdate = True
            Exit Function
    End Select

    Select Case normalizedScope
        Case GANTT_UPDATE_SCOPE_NOOP
            ready = Gantt_IsReadyStateValid()
            If Not ready Then
                normalizedScope = GANTT_UPDATE_SCOPE_INCREMENTAL
                escalationReason = "NoopReadyValidationFailed"
            End If

        Case GANTT_UPDATE_SCOPE_INCREMENTAL
            GanttColdState_MarkPending "CommitGanttUpdate|INCREMENTAL|" & normalizedSource & "|" & reason
            RunGanttRefreshCore True, False, False
            coreEffectiveScope = UCase$(Trim$(GanttRefresh_LastEffectiveScope()))
            If coreEffectiveScope = GANTT_UPDATE_SCOPE_FULL Then
                normalizedScope = GANTT_UPDATE_SCOPE_FULL
                escalationReason = "CoreFullScope"
                GanttColdState_MarkPending "CommitGanttUpdate|FULL|" & normalizedSource & "|" & reason
                finalizeReason = "CommitGanttUpdate|FULL|" & normalizedSource & "|" & reason
            Else
                finalizeReason = "CommitGanttUpdate|INCREMENTAL|" & normalizedSource & "|" & reason
            End If
            ready = GanttRefresh_LastRunSucceeded() And _
                Gantt_FinalizeReadyState(finalizeReason)
            If ready Then scopeAlreadyRendered = True
            If Not ready And normalizedScope <> GANTT_UPDATE_SCOPE_FULL Then
                Profiler_RecordOperation "GanttEngineIncrementalFallbacks", 1, 0#
                normalizedScope = GANTT_UPDATE_SCOPE_FULL
                escalationReason = "IncrementalReadyValidationFailed"
            End If

        Case GANTT_UPDATE_SCOPE_FULL
            ' handled below

        Case Else
            Err.Raise 5, "CommitGanttUpdate", "Unknown Gantt update scope: " & updateScope
    End Select

    If normalizedScope = GANTT_UPDATE_SCOPE_FULL And Not scopeAlreadyRendered Then
        GanttColdState_MarkPending "CommitGanttUpdate|FULL|" & normalizedSource & "|" & reason
        RunGanttRefreshCore False, False, False
        coreEffectiveScope = UCase$(Trim$(GanttRefresh_LastEffectiveScope()))
        If Len(coreEffectiveScope) > 0 Then normalizedScope = coreEffectiveScope
        ready = GanttRefresh_LastRunSucceeded() And _
            Gantt_FinalizeReadyState("CommitGanttUpdate|FULL|" & normalizedSource & "|" & reason)
    End If

    If Not ready Then GoTo Failed

    If normalizedIntent = GANTT_RENDER_INTENT_SHOW Then ws.Activate

    GanttEngine_RearmDragWatcherIfReady normalizedIntent
    GanttEngine_RecordScopeDecision requestedScope, normalizedScope, escalationReason
    Profiler_RecordOperation "GanttEngineCommits_" & normalizedScope, 1, 0#
    CommitGanttUpdate = True
    Exit Function

Failed:
    If normalizedScope <> GANTT_UPDATE_SCOPE_FULL And normalizedIntent <> GANTT_RENDER_INTENT_DEFER Then
        On Error Resume Next
        Err.Clear
        escalationReason = "EngineErrorFallback"
        GanttColdState_MarkPending "CommitGanttUpdateErrorFallback|FULL|" & normalizedSource & "|" & reason
        RunGanttRefreshCore False, False, False
        coreEffectiveScope = UCase$(Trim$(GanttRefresh_LastEffectiveScope()))
        ready = GanttRefresh_LastRunSucceeded() And _
            Gantt_FinalizeReadyState("CommitGanttUpdate|ERROR_FALLBACK_FULL|" & normalizedSource & "|" & reason)
        If ready Then
            If normalizedIntent = GANTT_RENDER_INTENT_SHOW Then ws.Activate
            GanttEngine_RearmDragWatcherIfReady normalizedIntent
            GanttEngine_RecordScopeDecision requestedScope, GANTT_UPDATE_SCOPE_FULL, escalationReason
            Profiler_RecordOperation "GanttEngineErrorFallbackFull", 1, 0#
            CommitGanttUpdate = True
            Exit Function
        End If
        On Error GoTo 0
    End If

    GanttColdState_MarkPending "CommitGanttUpdateFailed|" & reason
    CommitGanttUpdate = False

End Function

Private Sub GanttEngine_RecordScopeDecision( _
    ByVal requestedScope As String, _
    ByVal effectiveScope As String, _
    ByVal escalationReason As String)

    If Len(requestedScope) > 0 Then
        Profiler_RecordOperation "GanttEngineRequestedScope_" & requestedScope, 1, 0#
    End If
    If Len(effectiveScope) > 0 Then
        Profiler_RecordOperation "GanttEngineEffectiveScope_" & effectiveScope, 1, 0#
    End If
    If Len(escalationReason) > 0 Then
        Profiler_RecordOperation "GanttEngineEscalationReason_" & escalationReason, 1, 0#
    End If

End Sub

Public Function GanttEngine_IsPhysicalGeometryDirty() As Boolean

    Dim ws As Worksheet

    On Error GoTo Dirty
    Set ws = ThisWorkbook.Worksheets("GANTT")
    GanttEngine_IsPhysicalGeometryDirty = Not GanttDependencySvg_IsTaskGeometryCurrent(ws)
    Exit Function

Dirty:
    GanttEngine_IsPhysicalGeometryDirty = True

End Function

Private Sub GanttEngine_RearmDragWatcherIfReady(ByVal renderIntent As String)

    On Error Resume Next
    If UCase$(Trim$(renderIntent)) = GANTT_RENDER_INTENT_SHOW Then
        GanttDrag_EnsureArmed
    End If
    On Error GoTo 0

End Sub

Public Function GanttEngine_ActiveDataSource() As String

    Dim modeValue As String

    modeValue = UCase$(Trim$(GanttLive_GetPendingRenderMode()))
    If modeValue = "" Then modeValue = UCase$(Trim$(GanttLive_GetActiveSimulationMode()))

    Select Case modeValue
        Case GANTT_DATA_SOURCE_TEST
            GanttEngine_ActiveDataSource = GANTT_DATA_SOURCE_TEST
        Case GANTT_DATA_SOURCE_SCENARIO
            GanttEngine_ActiveDataSource = GANTT_DATA_SOURCE_SCENARIO
        Case Else
            GanttEngine_ActiveDataSource = GANTT_DATA_SOURCE_NORMAL
    End Select

End Function
