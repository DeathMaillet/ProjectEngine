Attribute VB_Name = "mod_GanttDependencyRenderer"
Option Explicit

'===============================================================================
' MODULE : mod_GanttDependencyRenderer
' DOMAINE / DOMAIN : Gantt
'
' FR
' Construit le rendu Excel du composant a partir de donnees deja preparees.
' Ne decide pas les donnees metier a calculer.
'
' EN
' Builds the component's Excel rendering from prepared data.
' Does not decide business data to calculate.
'
' CONTRATS / CONTRACTS : GanttDependency_ClearExpandedLinksCache, DrawDependencyLinks, GetTaskMidX
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"

Private Const FIRST_TIMELINE_COL As Long = 11
Private Const HEADER_ROW_1 As Long = 3
Private Const HEADER_ROW_2 As Long = 4
Private Const FIRST_TASK_ROW As Long = 5
Private Const COL_WBS As Long = 1
Private Const COL_TASK As Long = 2

Private Const COLOR_TASK_BLUE As Long = 12874308
Private Const COLOR_TASK_CRITICAL As Long = 192
Private Const COLOR_PROGRESS_GREEN As Long = 4699504
Private Const COLOR_PROGRESS_ORANGE As Long = 3248093

Private Const GANTT_VIEW_DETAIL As String = "DETAIL"
Private Const GANTT_VIEW_SUMMARY As String = "SUMMARY"
Private Const GANTT_ANALYTICS_PATH_NONE As String = "NONE"
Private Const GANTT_ANALYTICS_PATH_CP As String = "CP"
Private Const GANTT_ANALYTICS_PATH_LP As String = "LP"

Private Const LINK_STUB As Double = 6
Private Const LINK_MIN_CHANNEL_GAP As Double = 4
Private Const LINK_EDGE_PADDING As Double = 8
Private Const LINK_ANCHOR_START As String = "START"
Private Const LINK_ANCHOR_FINISH As String = "FINISH"

Private gExpectedDependencySegments As Object
Private gExistingDependencySegments As Object
Private gDependencyForceCreate As Boolean
Private gDependencyCreates As Long
Private gDependencyUpdates As Long
Private gDependencyDeletes As Long
Private gDependencyInspections As Long

Private gExpandedLinks As Object
Private gLinkSpecsByPrefix As Object
Private gIncomingLinksByTask As Object
Private gOutgoingLinksByTask As Object
Private gSegmentsByLink As Object
Private gDependencyArrowTransfers As Long
Private gDependencyVisibilityKnown As Boolean
Private gDependenciesVisible As Boolean
Private gDependencyHarnessSequence As Long
Private gActiveLogicalRoute As clsGanttDependencyRoute

'------------------------------------------------------------------------------
' FR: Reinitialise Gantt Dependency Clear Expanded Links Cache dans le perimetre possede par le composant.
' EN: Resets Gantt Dependency Clear Expanded Links Cache within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub GanttDependency_ClearExpandedLinksCache()
    Set gExpandedLinks = Nothing
    GanttDependencySvg_Invalidate "ExpandedLinksCache", True
End Sub

'------------------------------------------------------------------------------
' Reroutes only links adjacent to affected task IDs. The index is built during
' a canonical Day render and remains workbook-local until explicit invalidation.
'------------------------------------------------------------------------------
Public Function GanttDependency_DrawAffectedLinks( _
    ByVal wsGantt As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal rowById As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByVal affectedIds As Object, _
    ByRef fallbackReason As String) As Boolean

    Dim perfScope As clsPerfScope
    Dim affectedLinks As Object
    Dim anchorCache As Object
    Dim prefix As Variant
    Dim segmentName As Variant
    Dim spec As Object
    Dim oldSegments As Object
    Dim newSegments As Object
    Dim shp As Shape
    Dim inspectedLinks As Long

    Set perfScope = Profiler_BeginScope("GanttDependency_DrawAffectedLinks", "Gantt Local")
    On Error GoTo Failed

    fallbackReason = ""
    If gLinkSpecsByPrefix Is Nothing Or _
       gIncomingLinksByTask Is Nothing Or _
       gOutgoingLinksByTask Is Nothing Or _
       gSegmentsByLink Is Nothing Then
        fallbackReason = "DependencyIndexAbsent"
        Exit Function
    End If

    Set affectedLinks = CreateObject("Scripting.Dictionary")
    GanttDependency_CollectAffectedLinkPrefixes affectedIds, gIncomingLinksByTask, affectedLinks
    GanttDependency_CollectAffectedLinkPrefixes affectedIds, gOutgoingLinksByTask, affectedLinks

    If IsAggregatedScaleMode() Then
        If GanttDependencySvg_IsRequested() Then
            GanttDependencySvg_SetVisible wsGantt, False
            GanttDependencySvg_AcceptAggregatedScaleHidden wsGantt
        End If
        GanttDependency_DrawAffectedLinks = True
        Exit Function
    End If

    If affectedLinks.Count = 0 Then
        Profiler_RecordOperation "GanttLocalLinksInspected", 0, 0#
        GanttDependency_DrawAffectedLinks = True
        Exit Function
    End If

    If GanttDependencySvg_IsRequested() And Not GanttDependencySvg_HasRoutes() Then
        EnsureExpandedLinksCacheFromCalc
        If HasExpandedLinksAvailable() Then
            GanttDependencySvg_Invalidate "LocalSvgRoutesAbsent", True
            If GanttDependency_TryDrawSvgFull( _
                wsGantt, mapWBS, dataArr, hasChildren, rowById, _
                projectStart, totalDays, baseById, testById, isTestMode) Then
                Profiler_RecordOperation "GanttDependencySvgLocalColdRebuild", 1, 0#
                GanttDependency_DrawAffectedLinks = True
                Exit Function
            End If
        End If

        fallbackReason = GanttDependencySvg_GetFallbackReason()
        If Len(fallbackReason) = 0 Then fallbackReason = "DependencySvgRoutesAbsent"
        Exit Function
    End If

    If GanttDependencySvg_IsRequested() And GanttDependencySvg_HasRoutes() Then
        Set anchorCache = CreateObject("Scripting.Dictionary")

        For Each prefix In affectedLinks.Keys
            If Not gLinkSpecsByPrefix.Exists(CStr(prefix)) Then
                fallbackReason = "DependencySvgSpecMissing"
                Exit Function
            End If

            Set spec = gLinkSpecsByPrefix(CStr(prefix))
            Set gActiveLogicalRoute = GanttDependency_CreateLogicalRoute(spec, CStr(prefix))
            DrawSingleDependencyLink _
                wsGantt, mapWBS, dataArr, hasChildren, rowById, _
                projectStart, totalDays, _
                CStr(spec("PredID")), CStr(spec("SuccID")), _
                baseById, testById, isTestMode, _
                anchorCache, CStr(prefix), _
                CStr(spec("LinkType")), CDbl(spec("Lag"))

            If Not gActiveLogicalRoute Is Nothing Then
                GanttDependencySvg_StoreRoute gActiveLogicalRoute
            End If
            Set gActiveLogicalRoute = Nothing
            inspectedLinks = inspectedLinks + 1
        Next prefix

        If GanttDependencySvg_TryCommit(wsGantt) Then
            Profiler_RecordOperation "GanttLocalLinksInspected", inspectedLinks, 0#
            Profiler_RecordOperation "GanttDependencySvgLocalCommits", 1, 0#
            GanttDependency_DrawAffectedLinks = True
            Exit Function
        End If

        fallbackReason = GanttDependencySvg_GetFallbackReason()
        If Len(fallbackReason) = 0 Then fallbackReason = "DependencySvgLocalCommitFailed"
        Exit Function
    End If

    Set gExpectedDependencySegments = CreateObject("Scripting.Dictionary")
    Set gExistingDependencySegments = CreateObject("Scripting.Dictionary")
    gDependencyForceCreate = False
    gDependencyCreates = 0
    gDependencyUpdates = 0
    gDependencyDeletes = 0
    gDependencyInspections = 0
    gDependencyArrowTransfers = 0

    'Resolve only the stable names owned by affected links.
    For Each prefix In affectedLinks.Keys
        If gSegmentsByLink.Exists(CStr(prefix)) Then
            Set oldSegments = gSegmentsByLink(CStr(prefix))
            For Each segmentName In oldSegments.Keys
                Set shp = Nothing
                On Error Resume Next
                Set shp = wsGantt.Shapes(CStr(segmentName))
                On Error GoTo Failed
                If Not shp Is Nothing Then
                    gExistingDependencySegments.Add CStr(segmentName), shp
                    gDependencyInspections = gDependencyInspections + 1
                End If
            Next segmentName
        End If
    Next prefix

    Set anchorCache = CreateObject("Scripting.Dictionary")

    For Each prefix In affectedLinks.Keys
        If Not gLinkSpecsByPrefix.Exists(CStr(prefix)) Then
            fallbackReason = "DependencySpecMissing"
            GoTo FailedWithoutError
        End If

        Set spec = gLinkSpecsByPrefix(CStr(prefix))
        inspectedLinks = inspectedLinks + 1

        DrawSingleDependencyLink _
            wsGantt, mapWBS, dataArr, hasChildren, rowById, _
            projectStart, totalDays, _
            CStr(spec("PredID")), CStr(spec("SuccID")), _
            baseById, testById, isTestMode, _
            anchorCache, CStr(prefix), _
            CStr(spec("LinkType")), CDbl(spec("Lag"))
    Next prefix

    'Delete only stale segments belonging to the affected link prefixes, then
    'commit the new per-link segment index.
    For Each prefix In affectedLinks.Keys
        If gSegmentsByLink.Exists(CStr(prefix)) Then
            Set oldSegments = gSegmentsByLink(CStr(prefix))
            For Each segmentName In oldSegments.Keys
                If Not gExpectedDependencySegments.Exists(CStr(segmentName)) Then
                    On Error Resume Next
                    wsGantt.Shapes(CStr(segmentName)).Delete
                    If Err.Number = 0 Then gDependencyDeletes = gDependencyDeletes + 1
                    Err.Clear
                    On Error GoTo Failed
                End If
            Next segmentName
        End If

        Set newSegments = CreateObject("Scripting.Dictionary")
        For Each segmentName In gExpectedDependencySegments.Keys
            If Left$(CStr(segmentName), Len(CStr(prefix)) + 1) = CStr(prefix) & "_" Then
                newSegments(CStr(segmentName)) = True
            End If
        Next segmentName

        If gSegmentsByLink.Exists(CStr(prefix)) Then
            Set gSegmentsByLink(CStr(prefix)) = newSegments
        Else
            gSegmentsByLink.Add CStr(prefix), newSegments
        End If
    Next prefix

    Profiler_RecordOperation "GanttLocalLinksInspected", inspectedLinks, 0#
    Profiler_RecordOperation "GanttLocalDependencySegmentsInspected", gDependencyInspections, 0#
    Profiler_RecordOperation "GanttLocalDependencySegmentsCreated", gDependencyCreates, 0#
    Profiler_RecordOperation "GanttLocalDependencySegmentsUpdated", gDependencyUpdates, 0#
    Profiler_RecordOperation "GanttLocalDependencySegmentsDeleted", gDependencyDeletes, 0#
    Profiler_RecordOperation "GanttDependencyArrowTransfers", gDependencyArrowTransfers, 0#

    GanttDependency_DrawAffectedLinks = True

CleanExit:
    Set gActiveLogicalRoute = Nothing
    Set gExpectedDependencySegments = Nothing
    Set gExistingDependencySegments = Nothing
    gDependencyForceCreate = False
    Exit Function

Failed:
    fallbackReason = "DependencyLocalError_" & CStr(Err.Number)
FailedWithoutError:
    GanttDependency_DrawAffectedLinks = False
    GoTo CleanExit

End Function

Private Sub GanttDependency_ResetLocalIndex()

    Set gLinkSpecsByPrefix = CreateObject("Scripting.Dictionary")
    Set gIncomingLinksByTask = CreateObject("Scripting.Dictionary")
    Set gOutgoingLinksByTask = CreateObject("Scripting.Dictionary")
    Set gSegmentsByLink = CreateObject("Scripting.Dictionary")

End Sub

Private Sub GanttDependency_RegisterLink( _
    ByVal shapePrefix As String, _
    ByVal predId As String, _
    ByVal succId As String, _
    ByVal linkType As String, _
    ByVal linkLag As Double)

    Dim spec As Object

    If gLinkSpecsByPrefix Is Nothing Then GanttDependency_ResetLocalIndex

    Set spec = CreateObject("Scripting.Dictionary")
    spec("PredID") = predId
    spec("SuccID") = succId
    spec("LinkType") = linkType
    spec("Lag") = linkLag
    gLinkSpecsByPrefix.Add shapePrefix, spec

    GanttDependency_IndexLink gOutgoingLinksByTask, predId, shapePrefix
    GanttDependency_IndexLink gIncomingLinksByTask, succId, shapePrefix

End Sub

Private Sub GanttDependency_IndexLink( _
    ByVal indexByTask As Object, _
    ByVal taskId As String, _
    ByVal shapePrefix As String)

    Dim linksForTask As Object

    If Not indexByTask.Exists(taskId) Then
        Set linksForTask = CreateObject("Scripting.Dictionary")
        indexByTask.Add taskId, linksForTask
    Else
        Set linksForTask = indexByTask(taskId)
    End If

    linksForTask(shapePrefix) = True

End Sub

Private Sub GanttDependency_RegisterSegment(ByVal shapeName As String)

    Dim prefix As String
    Dim splitPos As Long
    Dim segments As Object

    splitPos = InStrRev(shapeName, "_")
    If splitPos < 1 Then Exit Sub
    prefix = Left$(shapeName, splitPos - 1)

    If gSegmentsByLink Is Nothing Then Set gSegmentsByLink = CreateObject("Scripting.Dictionary")

    If Not gSegmentsByLink.Exists(prefix) Then
        Set segments = CreateObject("Scripting.Dictionary")
        gSegmentsByLink.Add prefix, segments
    Else
        Set segments = gSegmentsByLink(prefix)
    End If

    segments(shapeName) = True

End Sub

Private Sub GanttDependency_CollectAffectedLinkPrefixes( _
    ByVal affectedIds As Object, _
    ByVal sourceIndex As Object, _
    ByVal target As Object)

    Dim idVal As Variant
    Dim prefix As Variant
    Dim linksForTask As Object

    If affectedIds Is Nothing Then Exit Sub

    For Each idVal In affectedIds.Keys
        If sourceIndex.Exists(CStr(idVal)) Then
            Set linksForTask = sourceIndex(CStr(idVal))
            For Each prefix In linksForTask.Keys
                target(CStr(prefix)) = True
            Next prefix
        End If
    Next idVal

End Sub

Public Function GanttDependency_MarkAffectedLinksDirty( _
    ByVal affectedIds As Object, _
    Optional ByRef dirtyCount As Long = 0) As Boolean

    Dim affectedLinks As Object

    On Error GoTo Failed
    dirtyCount = 0
    If affectedIds Is Nothing Then
        GanttDependency_MarkAffectedLinksDirty = True
        Exit Function
    End If

    If gIncomingLinksByTask Is Nothing Or gOutgoingLinksByTask Is Nothing Then
        GanttDependency_MarkAffectedLinksDirty = False
        Exit Function
    End If

    Set affectedLinks = CreateObject("Scripting.Dictionary")
    GanttDependency_CollectAffectedLinkPrefixes affectedIds, gIncomingLinksByTask, affectedLinks
    GanttDependency_CollectAffectedLinkPrefixes affectedIds, gOutgoingLinksByTask, affectedLinks

    dirtyCount = affectedLinks.Count
    If dirtyCount > 0 And GanttDependencySvg_IsRequested() Then
        GanttDependencySvg_MarkLinksDirty affectedLinks
    End If

    Profiler_RecordOperation "GanttDependencyIncidentDirtyLinks", dirtyCount, 0#
    GanttDependency_MarkAffectedLinksDirty = True
    Exit Function

Failed:
    GanttDependency_MarkAffectedLinksDirty = False

End Function

Public Sub GanttDependency_InvalidateLocalIndex(Optional ByVal reason As String = "")

    Set gLinkSpecsByPrefix = Nothing
    Set gIncomingLinksByTask = Nothing
    Set gOutgoingLinksByTask = Nothing
    Set gSegmentsByLink = Nothing
    gDependencyVisibilityKnown = False
    gDependenciesVisible = False
    GanttDependencySvg_Invalidate reason, True

    If Len(Trim$(reason)) > 0 Then
        Profiler_RecordOperation "GanttDependencyIndexInvalidation_" & reason, 1, 0#
    End If

End Sub

Public Function GanttDependency_PrimeLocalIndex( _
    ByVal wsGantt As Worksheet, _
    Optional ByVal refreshFromCalc As Boolean = False) As Boolean

    Dim perfScope As clsPerfScope
    Dim succId As Variant
    Dim linkItem As Variant
    Dim predId As String
    Dim shapePrefix As String
    Dim linkIndex As Long
    Dim i As Long
    Dim shapeName As String
    Dim oldSpecs As Object
    Dim changedPrefixes As Object
    Dim prefix As Variant

    Set perfScope = Profiler_BeginScope("GanttDependency_PrimeLocalIndex", "Gantt Local")
    On Error GoTo Failed

    If wsGantt Is Nothing Then Exit Function
    Set oldSpecs = gLinkSpecsByPrefix
    Set changedPrefixes = CreateObject("Scripting.Dictionary")
    If GanttDependencySvg_HasLayer(wsGantt) And Not GanttDependencySvg_HasRoutes() Then
        If GanttDependencySvg_TryHydratePersistentCache(wsGantt, True) Then
            Profiler_RecordOperation "GanttDependencySvgPrimeHydratedRoutes", 1, 0#
        Else
            Profiler_RecordOperation "GanttDependencySvgPrimeRequiresCanonical", 1, 0#
            GanttDependencySvg_Invalidate "ColdSvgLayerWithoutRoutes", True
        End If
    End If

    If refreshFromCalc Then Set gExpandedLinks = Nothing
    EnsureExpandedLinksCacheFromCalc
    If Not HasExpandedLinksAvailable() Then Exit Function

    GanttDependency_ResetLocalIndex

    For Each succId In gExpandedLinks.Keys
        linkIndex = 0
        For Each linkItem In gExpandedLinks(CStr(succId))
            predId = Trim$(CStr(linkItem("PredID")))
            If predId <> "" Then
                linkIndex = linkIndex + 1
                shapePrefix = "DEP_" & predId & "_" & CStr(succId) & "_" & CStr(linkIndex)
                GanttDependency_RegisterLink _
                    shapePrefix, predId, CStr(succId), _
                    GetLinkTypeFromItem(linkItem), GetLinkLagFromItem(linkItem)
            End If
        Next linkItem
    Next succId

    If Not oldSpecs Is Nothing Then
        For Each prefix In oldSpecs.Keys
            If Not gLinkSpecsByPrefix.Exists(CStr(prefix)) Then
                changedPrefixes(CStr(prefix)) = True
            ElseIf Not GanttDependency_LinkSpecsEqual( _
                oldSpecs(CStr(prefix)), gLinkSpecsByPrefix(CStr(prefix))) Then
                changedPrefixes(CStr(prefix)) = True
            End If
        Next prefix
        For Each prefix In gLinkSpecsByPrefix.Keys
            If Not oldSpecs.Exists(CStr(prefix)) Then changedPrefixes(CStr(prefix)) = True
        Next prefix
    End If

    If refreshFromCalc And GanttDependencySvg_IsRequested() Then
        GanttDependencySvg_ReconcileRouteIds gLinkSpecsByPrefix, changedPrefixes
    End If

    'One bootstrap scan is allowed after opening. Subsequent local transactions
    'resolve only exact names through the per-link segment index.
    For i = 1 To wsGantt.Shapes.Count
        shapeName = CStr(wsGantt.Shapes(i).Name)
        If Left$(shapeName, 4) = "DEP_" Then GanttDependency_RegisterSegment shapeName
    Next i

    gDependencyVisibilityKnown = True
    gDependenciesVisible = Not IsAggregatedScaleMode()

    Profiler_RecordOperation "GanttDependencyPrimeShapeScans", wsGantt.Shapes.Count, 0#
    Profiler_RecordOperation "GanttDependencyPrimeLinks", gLinkSpecsByPrefix.Count, 0#
    GanttDependency_PrimeLocalIndex = True
    Exit Function

Failed:
    GanttDependency_InvalidateLocalIndex "PrimeError"

End Function

Private Function GanttDependency_LinkSpecsEqual( _
    ByVal leftSpec As Object, _
    ByVal rightSpec As Object) As Boolean

    If leftSpec Is Nothing Or rightSpec Is Nothing Then Exit Function
    GanttDependency_LinkSpecsEqual = _
        (CStr(leftSpec("PredID")) = CStr(rightSpec("PredID"))) And _
        (CStr(leftSpec("SuccID")) = CStr(rightSpec("SuccID"))) And _
        (CStr(leftSpec("LinkType")) = CStr(rightSpec("LinkType"))) And _
        (Abs(CDbl(leftSpec("Lag")) - CDbl(rightSpec("Lag"))) < 0.000001)

End Function



'------------------------------------------------------------------------------
' FR:
' Dessine les liens de dependance depuis le cache tbl_LOGIC_LINKS expanse,
' en sautant les modes Week/Month.
'
' EN:
' Draws dependency links from the expanded tbl_LOGIC_LINKS cache,
' skipping Week/Month aggregated modes.
'
' Entrees / Inputs:
' - data WBS, rowById, hasChildren, maps base/test, mode TEST et timeline Day.
'
' Sorties / Outputs:
' - Shapes DEP_* composees de segments avec fleches.
'
' Appele par / Called by:
' - RunGanttRefreshCore.
'
' Notes:
' - Fortement couple a CALC/tbl_LOGIC_LINKS et a la geometrie des task bars.
'------------------------------------------------------------------------------
Public Sub DrawDependencyLinks( _
    ByVal wsGantt As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal rowById As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    Optional ByVal forceCreate As Boolean = False)

    Dim perfScope As clsPerfScope

    Dim succId As Variant
    Dim linkItem As Variant
    Dim predId As String
    Dim shapePrefix As String
    Dim linkIndex As Long
    Dim anchorCache As Object
    Dim aggregatedScale As Boolean
    Dim i As Long
    Dim existingName As String
    Dim svgRendered As Boolean

    Set perfScope = Profiler_BeginScope("DrawDependencyLinks", "Dependency Render")

    On Error GoTo SafeExit

    Set gExpectedDependencySegments = CreateObject("Scripting.Dictionary")
    Set gExistingDependencySegments = CreateObject("Scripting.Dictionary")
    gDependencyForceCreate = forceCreate
    gDependencyCreates = 0
    gDependencyUpdates = 0
    gDependencyDeletes = 0
    gDependencyInspections = 0
    gDependencyArrowTransfers = 0

    aggregatedScale = IsAggregatedScaleMode()
    If aggregatedScale Then
        GanttDependencySvg_SetVisible wsGantt, False
        GanttDependencySvg_AcceptAggregatedScaleHidden wsGantt
        If Not gDependencyVisibilityKnown Or gDependenciesVisible Then
            GanttDependency_SetIndexedVisibility wsGantt, False
        End If
        GoTo SafeExit
    End If

    If GanttDependencySvg_IsRequested() Then
        If GanttDependencySvg_CanReuseVisibleDayLayer(wsGantt) Then
            svgRendered = True
            GoTo SafeExit
        End If

        EnsureExpandedLinksCacheFromCalc
        If HasExpandedLinksAvailable() Then
            svgRendered = GanttDependency_TryDrawSvgFull( _
                wsGantt, mapWBS, dataArr, hasChildren, rowById, _
                projectStart, totalDays, baseById, testById, isTestMode)
            If svgRendered Then GoTo SafeExit
        End If

        GanttDependencySvg_ActivateHistorical wsGantt, _
            GanttDependencySvg_GetFallbackReason()
    End If

    If Not forceCreate Then
        If Not GanttDependency_LoadIndexedExistingSegments(wsGantt) Then
            For i = 1 To wsGantt.Shapes.Count
                existingName = CStr(wsGantt.Shapes(i).Name)
                If Left$(existingName, 4) = "DEP_" Then
                    gExistingDependencySegments.Add existingName, wsGantt.Shapes(i)
                    gDependencyInspections = gDependencyInspections + 1
                End If
            Next i
        End If
    End If

    EnsureExpandedLinksCacheFromCalc
    If Not HasExpandedLinksAvailable() Then GoTo SafeExit

    Set anchorCache = CreateObject("Scripting.Dictionary")
    GanttDependency_ResetLocalIndex

    For Each succId In gExpandedLinks.Keys

        linkIndex = 0

        For Each linkItem In gExpandedLinks(CStr(succId))

            predId = Trim$(CStr(linkItem("PredID")))

            If predId <> "" Then
                If rowById.Exists(predId) And rowById.Exists(CStr(succId)) Then

                    '==================================================
                    ' NOUVELLE RÈGLE :
                    ' on accepte aussi l'affichage des liens parent -> parent
                    ' si le lien existe dans CALC, on le dessine
                    '==================================================
                    linkIndex = linkIndex + 1
                    shapePrefix = "DEP_" & predId & "_" & CStr(succId) & "_" & CStr(linkIndex)
                    GanttDependency_RegisterLink _
                        shapePrefix, predId, CStr(succId), _
                        GetLinkTypeFromItem(linkItem), GetLinkLagFromItem(linkItem)

                    DrawSingleDependencyLink _
                        wsGantt, mapWBS, dataArr, hasChildren, rowById, _
                        projectStart, totalDays, _
                        predId, CStr(succId), _
                        baseById, testById, isTestMode, _
                        anchorCache, shapePrefix, _
                        GetLinkTypeFromItem(linkItem), _
                        GetLinkLagFromItem(linkItem)

                End If
            End If

NextLink:
        Next linkItem

    Next succId

SafeExit:
    If Not aggregatedScale And Not svgRendered Then
        GanttDependency_DeleteStaleSegments wsGantt
        gDependencyVisibilityKnown = True
        gDependenciesVisible = True
    End If
    Profiler_RecordOperation "GanttDiffDependencySegmentsInspected", gDependencyInspections, 0#
    Profiler_RecordOperation "GanttDiffDependencySegmentsCreated", gDependencyCreates, 0#
    Profiler_RecordOperation "GanttDiffDependencySegmentsUpdated", gDependencyUpdates, 0#
    Profiler_RecordOperation "GanttDiffDependencySegmentsDeleted", gDependencyDeletes, 0#
    Profiler_RecordOperation "GanttDependencyArrowTransfers", gDependencyArrowTransfers, 0#
    Set gExpectedDependencySegments = Nothing
    Set gExistingDependencySegments = Nothing
    gDependencyForceCreate = False
    Set gActiveLogicalRoute = Nothing
End Sub

Private Function GanttDependency_TryDrawSvgFull( _
    ByVal wsGantt As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal rowById As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean) As Boolean

    Dim perfScope As clsPerfScope
    Dim succId As Variant
    Dim linkItem As Variant
    Dim predId As String
    Dim shapePrefix As String
    Dim linkIndex As Long
    Dim anchorCache As Object
    Dim spec As Object
    Dim rebuildAll As Boolean
    Dim routesBuilt As Long
    Dim routesReused As Long

    Set perfScope = Profiler_BeginScope("GanttDependency_TryDrawSvgFull", "Dependency SVG")
    On Error GoTo Failed

    rebuildAll = Not GanttDependencySvg_HasRoutes()
    If rebuildAll Then GanttDependencySvg_BeginFullModel

    Set anchorCache = CreateObject("Scripting.Dictionary")
    GanttDependency_ResetLocalIndex

    For Each succId In gExpandedLinks.Keys
        linkIndex = 0
        For Each linkItem In gExpandedLinks(CStr(succId))
            predId = Trim$(CStr(linkItem("PredID")))
            If predId <> "" Then
                If rowById.Exists(predId) And rowById.Exists(CStr(succId)) Then
                    linkIndex = linkIndex + 1
                    shapePrefix = "DEP_" & predId & "_" & CStr(succId) & "_" & CStr(linkIndex)
                    GanttDependency_RegisterLink _
                        shapePrefix, predId, CStr(succId), _
                        GetLinkTypeFromItem(linkItem), GetLinkLagFromItem(linkItem)

                    If rebuildAll Or GanttDependencySvg_ShouldRebuildRoute(shapePrefix) Then
                        Set spec = gLinkSpecsByPrefix(shapePrefix)
                        Set gActiveLogicalRoute = GanttDependency_CreateLogicalRoute(spec, shapePrefix)
                        DrawSingleDependencyLink _
                            wsGantt, mapWBS, dataArr, hasChildren, rowById, _
                            projectStart, totalDays, predId, CStr(succId), _
                            baseById, testById, isTestMode, anchorCache, shapePrefix, _
                            CStr(spec("LinkType")), CDbl(spec("Lag"))
                        GanttDependencySvg_StoreRoute gActiveLogicalRoute
                        Set gActiveLogicalRoute = Nothing
                        routesBuilt = routesBuilt + 1
                    Else
                        routesReused = routesReused + 1
                    End If
                End If
            End If
        Next linkItem
    Next succId

    Profiler_RecordOperation "GanttDependencySvgRoutesBuilt", routesBuilt, 0#
    Profiler_RecordOperation "GanttDependencySvgRoutesReused", routesReused, 0#
    GanttDependency_TryDrawSvgFull = GanttDependencySvg_TryCommit(wsGantt)
    Exit Function

Failed:
    Set gActiveLogicalRoute = Nothing
    Profiler_RecordOperation "GanttDependencySvgRouteBuildError", 1, 0#

End Function

Private Function GanttDependency_CreateLogicalRoute( _
    ByVal spec As Object, _
    ByVal shapePrefix As String) As clsGanttDependencyRoute

    Dim route As clsGanttDependencyRoute

    Set route = New clsGanttDependencyRoute
    route.Initialize _
        shapePrefix, CStr(spec("PredID")), CStr(spec("SuccID")), _
        CStr(spec("LinkType")), CDbl(spec("Lag"))
    Set GanttDependency_CreateLogicalRoute = route

End Function

Private Function GanttDependency_LoadIndexedExistingSegments(ByVal ws As Worksheet) As Boolean

    Dim prefix As Variant
    Dim segmentName As Variant
    Dim segments As Object
    Dim shp As Shape

    If gSegmentsByLink Is Nothing Then Exit Function
    If gSegmentsByLink.Count = 0 Then Exit Function

    For Each prefix In gSegmentsByLink.Keys
        Set segments = gSegmentsByLink(CStr(prefix))
        For Each segmentName In segments.Keys
            Set shp = Nothing
            On Error Resume Next
            Set shp = ws.Shapes(CStr(segmentName))
            On Error GoTo 0

            If shp Is Nothing Then
                Set gExistingDependencySegments = CreateObject("Scripting.Dictionary")
                Exit Function
            End If
            gExistingDependencySegments.Add CStr(segmentName), shp
            gDependencyInspections = gDependencyInspections + 1
        Next segmentName
    Next prefix

    GanttDependency_LoadIndexedExistingSegments = True

End Function

Private Sub GanttDependency_SetIndexedVisibility(ByVal ws As Worksheet, ByVal makeVisible As Boolean)

    Dim prefix As Variant
    Dim segmentName As Variant
    Dim segments As Object
    Dim shp As Shape
    Dim targetVisibility As MsoTriState
    Dim changedCount As Long
    Dim visibilityTransitionKnown As Boolean

    targetVisibility = IIf(makeVisible, msoTrue, msoFalse)
    visibilityTransitionKnown = _
        gDependencyVisibilityKnown And (gDependenciesVisible <> makeVisible)

    If gSegmentsByLink Is Nothing Then
        GanttDependency_SetExistingVisibility ws, makeVisible
        gDependencyVisibilityKnown = True
        gDependenciesVisible = makeVisible
        Exit Sub
    End If

    For Each prefix In gSegmentsByLink.Keys
        Set segments = gSegmentsByLink(CStr(prefix))
        For Each segmentName In segments.Keys
            Set shp = Nothing
            On Error Resume Next
            Set shp = ws.Shapes(CStr(segmentName))
            On Error GoTo 0
            If Not shp Is Nothing Then
                If visibilityTransitionKnown Then
                    shp.Visible = targetVisibility
                    changedCount = changedCount + 1
                ElseIf shp.Visible <> targetVisibility Then
                    shp.Visible = targetVisibility
                    changedCount = changedCount + 1
                End If
            End If
        Next segmentName
    Next prefix

    gDependencyVisibilityKnown = True
    gDependenciesVisible = makeVisible
    Profiler_RecordOperation "GanttDiffDependencyVisibilityUpdated", changedCount, 0#

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Function GetLinkTypeFromItem(ByVal linkItem As Variant) As String

    On Error GoTo SafeExit

    If IsObject(linkItem) Then
        If linkItem.Exists("LinkType") Then
            GetLinkTypeFromItem = UCase$(Trim$(CStr(linkItem("LinkType"))))
            If GetLinkTypeFromItem = "" Then GetLinkTypeFromItem = "FS"
            Exit Function
        End If
    End If

SafeExit:
    GetLinkTypeFromItem = "FS"

End Function
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Function GetLinkLagFromItem(ByVal linkItem As Variant) As Double

    On Error GoTo SafeExit

    If IsObject(linkItem) Then
        If linkItem.Exists("Lag") Then
            If IsNumeric(linkItem("Lag")) Then
                GetLinkLagFromItem = CDbl(linkItem("Lag"))
                Exit Function
            End If
        End If
    End If

SafeExit:
    GetLinkLagFromItem = 0#

End Function
'------------------------------------------------------------------------------
' FR:
' Route un lien FS/SS/FF entre deux taches en choisissant les ancres,
' les points d'entree et le chemin visuel adapte au lag.
'
' EN:
' Routes one FS/SS/FF link between two tasks by choosing anchors,
' entry points, and a visual path suited to lag.
'
' Entrees / Inputs:
' - IDs pred/succ, type de lien, lag, cache d'ancres, maps base/test.
'
' Sorties / Outputs:
' - Segments DEP_* horizontaux/verticaux formant le connecteur.
'
' Appele par / Called by:
' - DrawDependencyLinks.
'
' Notes:
' - Zone a ne pas refactoriser brutalement: nombreuses regles visuelles FS same-day/negative.
'------------------------------------------------------------------------------
Private Sub DrawSingleDependencyLink( _
    ByVal wsGantt As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal rowById As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal predId As String, _
    ByVal succId As String, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByVal anchorCache As Object, _
    ByVal shapePrefix As String, _
    ByVal linkType As String, _
    ByVal linkLag As Double)

    Dim perfScope As clsPerfScope
    Dim stageScope As clsPerfScope

    Dim predRow As Long
    Dim succRow As Long
    Dim predDataRow As Long
    Dim succDataRow As Long

    Dim predX As Double
    Dim predY As Double
    Dim succX As Double
    Dim succY As Double

    Dim succTopX As Double
    Dim succTopY As Double
    Dim succMidLeftX As Double
    Dim succMidLeftY As Double

    Dim predAnchorType As String
    Dim succAnchorType As String
    Dim predDate As Variant
    Dim succDate As Variant
    Dim gapDays As Long
    Dim useMidLeftEntry As Boolean
    Dim cellWidth As Double

    Set perfScope = Profiler_BeginScope("DrawSingleDependencyLink", "Dependency Render")

    If Not rowById.Exists(predId) Then Exit Sub
    If Not rowById.Exists(succId) Then Exit Sub

    predRow = CLng(rowById(predId))
    succRow = CLng(rowById(succId))

    predDataRow = predRow - FIRST_TASK_ROW + 1
    succDataRow = succRow - FIRST_TASK_ROW + 1

    If predDataRow < 1 Or predDataRow > UBound(dataArr, 1) Then Exit Sub
    If succDataRow < 1 Or succDataRow > UBound(dataArr, 1) Then Exit Sub

    linkType = UCase$(Trim$(linkType))
    If linkType = "" Then linkType = "FS"

    GetLinkAnchorTypes linkType, predAnchorType, succAnchorType
    Set stageScope = Profiler_BeginScope("DependencyRoute_CellWidthRead", "Dependency Render")
    cellWidth = wsGantt.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width
    Set stageScope = Nothing

    Set stageScope = Profiler_BeginScope("DependencyRoute_ReferenceDates", "Dependency Render")
    predDate = GetLinkReferenceDate(predId, predAnchorType, baseById, testById, isTestMode)
    succDate = GetLinkReferenceDate(succId, succAnchorType, baseById, testById, isTestMode)
    Set stageScope = Nothing

    If Not HasValue(predDate) Then Exit Sub
    If Not HasValue(succDate) Then Exit Sub

    Set stageScope = Profiler_BeginScope("DependencyRoute_AnchorLookup", "Dependency Render")
    GetCachedTaskAnchorPointByType anchorCache, wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, predDataRow, _
        predAnchorType, predX, predY, baseById, testById, isTestMode
    Set stageScope = Nothing

    Select Case linkType

        Case "SS"
            Set stageScope = Profiler_BeginScope("DependencyRoute_AnchorLookup", "Dependency Render")
            GetCachedTaskAnchorPointByType anchorCache, wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, succDataRow, _
                succAnchorType, succX, succY, baseById, testById, isTestMode
            Set stageScope = Nothing

            gapDays = CLng(CDbl(succDate) - CDbl(predDate) - linkLag)
            Set stageScope = Profiler_BeginScope("DependencyRoute_SegmentConstruction", "Dependency Render")
            RouteDependencyLink_SS wsGantt, shapePrefix, predX, predY, succX, succY, gapDays, cellWidth
            Set stageScope = Nothing

        Case "FF"
            Set stageScope = Profiler_BeginScope("DependencyRoute_AnchorLookup", "Dependency Render")
            If GanttDependency_IsPointMarkerTarget( _
                dataArr, mapWBS, succDataRow, baseById, testById, isTestMode) Then
                GetCachedTaskStartMidEntryPoint _
                    anchorCache, wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, succDataRow, _
                    succX, succY, baseById, testById, isTestMode
                Profiler_RecordOperation "GanttDependencyFfPointMarkerLeftEntry", 1, 0#
            Else
                GetCachedTaskAnchorPointByType anchorCache, wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, succDataRow, _
                    succAnchorType, succX, succY, baseById, testById, isTestMode
            End If
            Set stageScope = Nothing

            gapDays = CLng(CDbl(succDate) - CDbl(predDate) - linkLag)
            Set stageScope = Profiler_BeginScope("DependencyRoute_SegmentConstruction", "Dependency Render")
            RouteDependencyLink_FF wsGantt, shapePrefix, predX, predY, succX, succY, gapDays, cellWidth
            Set stageScope = Nothing

        Case Else   ' FS
            gapDays = CLng(CDbl(succDate) - CDbl(predDate) - 1 - linkLag)

            ' On calcule les 2 points candidats côté successeur
            Set stageScope = Profiler_BeginScope("DependencyRoute_AnchorLookup", "Dependency Render")
            GetCachedTaskTopEntryPoint anchorCache, wsGantt, mapWBS, dataArr, projectStart, totalDays, succDataRow, _
                succTopX, succTopY, baseById, testById, isTestMode

            GetCachedTaskStartMidEntryPoint anchorCache, wsGantt, mapWBS, dataArr, hasChildren, projectStart, totalDays, succDataRow, _
                succMidLeftX, succMidLeftY, baseById, testById, isTestMode
            Set stageScope = Nothing

            ' Règle corrigée :
            ' on ne décide PAS avec gapDays=0/1
            ' on décide avec la place horizontale réelle entre pred et l'entrée gauche du successeur
            useMidLeftEntry = HasRoomForFsMidLeftEntry(predX, succMidLeftX, cellWidth)

            If gapDays < 0 Then
                succX = succMidLeftX
                succY = succMidLeftY
                Set stageScope = Profiler_BeginScope("DependencyRoute_SegmentConstruction", "Dependency Render")
                RouteDependencyLink_FS_Negative wsGantt, shapePrefix, predX, predY, succX, succY, gapDays, cellWidth
                Set stageScope = Nothing

            ElseIf useMidLeftEntry Then
                succX = succMidLeftX
                succY = succMidLeftY
                Set stageScope = Profiler_BeginScope("DependencyRoute_SegmentConstruction", "Dependency Render")
                RouteDependencyLink_FS_Normal wsGantt, shapePrefix, predX, predY, succX, succY, gapDays, cellWidth
                Set stageScope = Nothing

            Else
                succX = succTopX
                succY = succTopY
                Set stageScope = Profiler_BeginScope("DependencyRoute_SegmentConstruction", "Dependency Render")
                RouteDependencyLink_FS_SameDay wsGantt, shapePrefix, predX, predY, succX, succY
                Set stageScope = Nothing
            End If

    End Select

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Cached Task Anchor Point By Type dans le workflow de rendu GANTT.
' EN: Runs the Get Cached Task Anchor Point By Type helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetCachedTaskAnchorPointByType( _
    ByVal anchorCache As Object, _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal anchorType As String, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim cacheKey As String
    Dim cachedValue As Variant

    cacheKey = "TYPE|" & CStr(dataRow) & "|" & UCase$(Trim$(anchorType))

    If Not anchorCache Is Nothing Then
        If anchorCache.Exists(cacheKey) Then
            cachedValue = anchorCache(cacheKey)
            xOut = CDbl(cachedValue(0))
            yOut = CDbl(cachedValue(1))
            Exit Sub
        End If
    End If

    GetTaskAnchorPointByType ws, mapWBS, dataArr, hasChildren, projectStart, totalDays, dataRow, _
        anchorType, xOut, yOut, baseById, testById, isTestMode

    If Not anchorCache Is Nothing Then anchorCache(cacheKey) = Array(xOut, yOut)

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Cached Task Top Entry Point dans le workflow de rendu GANTT.
' EN: Runs the Get Cached Task Top Entry Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetCachedTaskTopEntryPoint( _
    ByVal anchorCache As Object, _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim cacheKey As String
    Dim cachedValue As Variant

    cacheKey = "TOP|" & CStr(dataRow)

    If Not anchorCache Is Nothing Then
        If anchorCache.Exists(cacheKey) Then
            cachedValue = anchorCache(cacheKey)
            xOut = CDbl(cachedValue(0))
            yOut = CDbl(cachedValue(1))
            Exit Sub
        End If
    End If

    GetTaskTopEntryPoint ws, mapWBS, dataArr, projectStart, totalDays, dataRow, _
        xOut, yOut, baseById, testById, isTestMode

    If Not anchorCache Is Nothing Then anchorCache(cacheKey) = Array(xOut, yOut)

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Cached Task Start Mid Entry Point dans le workflow de rendu GANTT.
' EN: Runs the Get Cached Task Start Mid Entry Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetCachedTaskStartMidEntryPoint( _
    ByVal anchorCache As Object, _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim cacheKey As String
    Dim cachedValue As Variant

    cacheKey = "MIDLEFT|" & CStr(dataRow)

    If Not anchorCache Is Nothing Then
        If anchorCache.Exists(cacheKey) Then
            cachedValue = anchorCache(cacheKey)
            xOut = CDbl(cachedValue(0))
            yOut = CDbl(cachedValue(1))
            Exit Sub
        End If
    End If

    GetTaskStartMidEntryPoint ws, mapWBS, dataArr, hasChildren, projectStart, totalDays, dataRow, _
        xOut, yOut, baseById, testById, isTestMode

    If Not anchorCache Is Nothing Then anchorCache(cacheKey) = Array(xOut, yOut)

End Sub
'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function HasRoomForFsMidLeftEntry( _
    ByVal predX As Double, _
    ByVal succMidLeftX As Double, _
    ByVal cellWidth As Double) As Boolean

    Dim minNeeded As Double

    ' Il faut une vraie place visuelle pour arriver par la gauche.
    ' gapDays = 0 peut quand même avoir assez de place à l’écran.
    minNeeded = MaxDouble(10, cellWidth * 0.55)

    HasRoomForFsMidLeftEntry = ((succMidLeftX - predX) >= minNeeded)

End Function
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub RouteDependencyLink_FS_SameDay( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double)

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, succX, predY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", succX, predY, succX, succY, True

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Start Mid Entry Point dans le workflow de rendu GANTT.
' EN: Runs the Get Task Start Mid Entry Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskStartMidEntryPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    ' Entrée milieu gauche = vrai point milieu côté gauche
    GetTaskAnchorPointBySide ws, mapWBS, dataArr, hasChildren, projectStart, totalDays, dataRow, _
        "LEFT", xOut, yOut, baseById, testById, isTestMode

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub RouteDependencyLink_FS_Normal( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long, _
    ByVal cellWidth As Double)

    Dim endX As Double
    Dim directEnough As Boolean
    Dim routeAbove As Boolean
    Dim laneY As Double
    Dim bendX As Double
    Dim entryGap As Double
    Dim finalX As Double
    ' IMPORTANT : plus aucun cas spécial ici basé sur gapDays = 0.
    ' Le choix top / milieu-gauche a déjà été fait en amont.

    If gapDays <= 1 Then
        entryGap = 8
    Else
        entryGap = 8 + (cellWidth / 2)
    End If

    endX = succX - entryGap
    If endX <= predX + 4 Then endX = succX - 4
    If endX <= predX + 2 Then endX = succX

    directEnough = (endX - predX >= LINK_STUB * 2)

    If directEnough Then
        If gapDays > 1 Then
            bendX = predX + LINK_STUB + (cellWidth / 3)
        Else
            bendX = predX + LINK_STUB
        End If

        If bendX > endX - LINK_STUB Then
            bendX = predX + ((endX - predX) / 2)
        End If

        finalX = succX - 2

        If finalX <= endX + 2 Then
            DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
            DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, succY, False
            DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, succY, succX, succY, True
        Else
            DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
            DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, succY, False
            DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, succY, endX, succY, False
            DrawLinkSegment wsGantt, shapePrefix & "_4", endX, succY, succX, succY, True
        End If

        Exit Sub
    End If

    routeAbove = (succY <= predY)

    If routeAbove Then
        laneY = MinDouble(predY, succY) - LINK_MIN_CHANNEL_GAP
    Else
        laneY = MaxDouble(predY, succY) + LINK_MIN_CHANNEL_GAP
    End If

    If gapDays > 1 Then
        bendX = predX + LINK_STUB + (cellWidth / 3)
    Else
        bendX = predX + LINK_STUB
    End If

    If bendX >= succX - 6 Then
        bendX = predX + ((succX - predX) / 2)
    End If

    finalX = succX - 2

    If finalX <= endX + 2 Then
        DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
        DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, laneY, succX - 2, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_4", succX - 2, laneY, succX - 2, succY, False
        DrawLinkSegment wsGantt, shapePrefix & "_5", succX - 2, succY, succX, succY, True
    Else
        DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, bendX, predY, False
        DrawLinkSegment wsGantt, shapePrefix & "_2", bendX, predY, bendX, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_3", bendX, laneY, endX, laneY, False
        DrawLinkSegment wsGantt, shapePrefix & "_4", endX, laneY, endX, succY, False
        DrawLinkSegment wsGantt, shapePrefix & "_5", endX, succY, succX, succY, True
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub RouteDependencyLink_FS_Negative( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long, _
    ByVal cellWidth As Double)

    Dim laneY As Double
    Dim leftX As Double

    If succY <= predY Then
        laneY = MinDouble(predY, succY) - LINK_MIN_CHANNEL_GAP
    Else
        laneY = MaxDouble(predY, succY) + LINK_MIN_CHANNEL_GAP
    End If

    leftX = MinDouble(predX, succX) - MaxDouble(8, cellWidth / 2)
    leftX = leftX - MaxDouble(6, Abs(gapDays) * (cellWidth / 2))

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, predX, laneY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", predX, laneY, leftX, laneY, False
    DrawLinkSegment wsGantt, shapePrefix & "_3", leftX, laneY, leftX, succY, False
    DrawLinkSegment wsGantt, shapePrefix & "_4", leftX, succY, succX, succY, True

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub RouteDependencyLink_SS( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long, _
    ByVal cellWidth As Double)

    Dim busX As Double

    busX = MinDouble(predX, succX) - MaxDouble(8, cellWidth / 2)

    If gapDays < 0 Then
        busX = busX - MaxDouble(6, Abs(gapDays) * (cellWidth / 2))
    End If

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, busX, predY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", busX, predY, busX, succY, False
    DrawLinkSegment wsGantt, shapePrefix & "_3", busX, succY, succX, succY, True

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub RouteDependencyLink_FF( _
    ByVal wsGantt As Worksheet, _
    ByVal shapePrefix As String, _
    ByVal predX As Double, _
    ByVal predY As Double, _
    ByVal succX As Double, _
    ByVal succY As Double, _
    ByVal gapDays As Long, _
    ByVal cellWidth As Double)

    Dim busX As Double

    busX = MinDouble(predX, succX) - MaxDouble(8, cellWidth / 2)

    If gapDays < 0 Then
        busX = busX - MaxDouble(6, Abs(gapDays) * (cellWidth / 2))
    End If

    DrawLinkSegment wsGantt, shapePrefix & "_1", predX, predY, busX, predY, False
    DrawLinkSegment wsGantt, shapePrefix & "_2", busX, predY, busX, succY, False
    DrawLinkSegment wsGantt, shapePrefix & "_3", busX, succY, succX, succY, True

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Top Entry Point dans le workflow de rendu GANTT.
' EN: Runs the Get Task Top Entry Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskTopEntryPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double
    Dim topEntryOffset As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS(VTS_COL_ID))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    If durationVal <= 1 Then
        xOut = GetTaskMidX(ws, projectStart, startVal)
        yOut = GetGanttRowTop(ws, ganttRow) + 3
    Else
        xOut = TimelineLeft(ws, projectStart, startVal) + 4
        yOut = GetGanttBarTop(ws, ganttRow)
    End If

    topEntryOffset = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Width * 0.15
    xOut = xOut + topEntryOffset

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub FormatDependencyLine(ByVal shp As Shape, ByVal withArrow As Boolean)

    With shp.Line
        .ForeColor.RGB = RGB(120, 120, 120)
        .Weight = 1
        .DashStyle = msoLineSolid
        If withArrow Then
            .EndArrowheadStyle = msoArrowheadTriangle
        Else
            .EndArrowheadStyle = msoArrowheadNone
        End If
    End With

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub DrawLinkSegment( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal x1 As Double, _
    ByVal y1 As Double, _
    ByVal x2 As Double, _
    ByVal y2 As Double, _
    ByVal withArrow As Boolean)

    Dim perfScope As clsPerfScope

    Dim shp As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim widthVal As Double
    Dim heightVal As Double
    Dim geometryDiffers As Boolean
    Dim styleDiffers As Boolean
    Dim visibilityDiffers As Boolean

    If Not gActiveLogicalRoute Is Nothing Then
        gActiveLogicalRoute.AddSegment x1, y1, x2, y2, withArrow
        Exit Sub
    End If

    Set perfScope = Profiler_BeginScope("DrawLinkSegment", "Shape Create")

    If Abs(x2 - x1) < 0.1 And Abs(y2 - y1) < 0.1 Then
        If withArrow Then GanttDependency_TransferTerminalArrow ws, shapeName
        Exit Sub
    End If

    If Not gExpectedDependencySegments Is Nothing Then
        gExpectedDependencySegments(shapeName) = True
    End If
    GanttDependency_RegisterSegment shapeName

    leftPos = MinDouble(x1, x2)
    topPos = MinDouble(y1, y2)
    widthVal = Abs(x2 - x1)
    heightVal = Abs(y2 - y1)

    If Not gDependencyForceCreate Then
        If Not gExistingDependencySegments Is Nothing Then
            If gExistingDependencySegments.Exists(shapeName) Then
                Set shp = gExistingDependencySegments(shapeName)
            End If
        End If
    End If

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddLine(x1, y1, x2, y2)
        shp.Name = shapeName
        ApplyGanttRenderLinePlacement shp
        FormatDependencyLine shp, withArrow
        If Not gExistingDependencySegments Is Nothing Then
            gExistingDependencySegments.Add shapeName, shp
        End If
        gDependencyCreates = gDependencyCreates + 1
        Exit Sub
    End If

    gDependencyInspections = gDependencyInspections + 1

    geometryDiffers = _
        Abs(shp.Left - leftPos) > 0.1 Or _
        Abs(shp.Top - topPos) > 0.1 Or _
        Abs(shp.Width - widthVal) > 0.1 Or _
        Abs(shp.Height - heightVal) > 0.1

    styleDiffers = _
        shp.Line.ForeColor.RGB <> RGB(120, 120, 120) Or _
        Abs(shp.Line.Weight - 1) > 0.01 Or _
        shp.Line.DashStyle <> msoLineSolid Or _
        ((shp.Line.EndArrowheadStyle = msoArrowheadNone) = withArrow)
    visibilityDiffers = (shp.Visible <> msoTrue)

    If geometryDiffers Then
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthVal
        shp.Height = heightVal
    End If

    If styleDiffers Then FormatDependencyLine shp, withArrow
    If visibilityDiffers Then shp.Visible = msoTrue

    If geometryDiffers Or styleDiffers Or visibilityDiffers Then
        gDependencyUpdates = gDependencyUpdates + 1
    End If

End Sub

Private Function MinDouble(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If firstValue < secondValue Then
        MinDouble = firstValue
    Else
        MinDouble = secondValue
    End If
End Function

Private Function MaxDouble(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If firstValue > secondValue Then
        MaxDouble = firstValue
    Else
        MaxDouble = secondValue
    End If
End Function

Private Sub GanttDependency_TransferTerminalArrow( _
    ByVal ws As Worksheet, _
    ByVal terminalShapeName As String)

    Dim splitPos As Long
    Dim segmentIndex As Long
    Dim candidateName As String
    Dim shp As Shape

    splitPos = InStrRev(terminalShapeName, "_")
    If splitPos < 1 Then Exit Sub
    If Not IsNumeric(Mid$(terminalShapeName, splitPos + 1)) Then Exit Sub

    segmentIndex = CLng(Mid$(terminalShapeName, splitPos + 1)) - 1

    Do While segmentIndex > 0
        candidateName = Left$(terminalShapeName, splitPos) & CStr(segmentIndex)
        Set shp = Nothing

        If Not gExistingDependencySegments Is Nothing Then
            If gExistingDependencySegments.Exists(candidateName) Then
                Set shp = gExistingDependencySegments(candidateName)
            End If
        End If

        If shp Is Nothing Then
            On Error Resume Next
            Set shp = ws.Shapes(candidateName)
            On Error GoTo 0
        End If

        If Not shp Is Nothing Then
            FormatDependencyLine shp, True
            If Not gExpectedDependencySegments Is Nothing Then
                gExpectedDependencySegments(candidateName) = True
            End If
            GanttDependency_RegisterSegment candidateName
            gDependencyArrowTransfers = gDependencyArrowTransfers + 1
            Exit Sub
        End If

        segmentIndex = segmentIndex - 1
    Loop

End Sub

'------------------------------------------------------------------------------
' Deterministic proof hook for terminal-arrow ownership. It creates and removes
' only two exact harness shapes on GANTT and does not touch the dependency index.
'------------------------------------------------------------------------------
Public Function GanttDependencyHarness_TerminalArrowCase( _
    ByVal terminalDeltaX As Double, _
    ByVal terminalDeltaY As Double) As Variant

    Dim ws As Worksheet
    Dim prefix As String
    Dim shapeName As String
    Dim shp As Shape
    Dim result(1 To 1, 1 To 6) As Variant
    Dim arrowCount As Long
    Dim visibleSegmentCount As Long
    Dim terminalSegmentExists As Boolean
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)
    gDependencyHarnessSequence = gDependencyHarnessSequence + 1
    prefix = "HARNESS_DEP_TERMINAL_" & CStr(gDependencyHarnessSequence)

    Set gExpectedDependencySegments = CreateObject("Scripting.Dictionary")
    Set gExistingDependencySegments = CreateObject("Scripting.Dictionary")
    gDependencyForceCreate = True
    gDependencyCreates = 0
    gDependencyUpdates = 0
    gDependencyDeletes = 0
    gDependencyInspections = 0
    gDependencyArrowTransfers = 0

    On Error GoTo CleanExit

    DrawLinkSegment ws, prefix & "_1", 20#, 20#, 40#, 20#, False
    DrawLinkSegment _
        ws, prefix & "_2", 40#, 20#, _
        40# + terminalDeltaX, 20# + terminalDeltaY, True

    For i = 1 To 2
        shapeName = prefix & "_" & CStr(i)
        Set shp = Nothing
        On Error Resume Next
        Set shp = ws.Shapes(shapeName)
        On Error GoTo CleanExit

        If Not shp Is Nothing Then
            visibleSegmentCount = visibleSegmentCount + 1
            If shp.Line.EndArrowheadStyle <> msoArrowheadNone Then
                arrowCount = arrowCount + 1
            End If
            If i = 2 Then terminalSegmentExists = True
        End If
    Next i

    result(1, 1) = terminalDeltaX
    result(1, 2) = terminalDeltaY
    result(1, 3) = visibleSegmentCount
    result(1, 4) = arrowCount
    result(1, 5) = terminalSegmentExists
    result(1, 6) = gDependencyArrowTransfers

CleanExit:
    On Error Resume Next
    ws.Shapes(prefix & "_1").Delete
    ws.Shapes(prefix & "_2").Delete
    On Error GoTo 0

    Set gExpectedDependencySegments = Nothing
    Set gExistingDependencySegments = Nothing
    gDependencyForceCreate = False

    GanttDependencyHarness_TerminalArrowCase = result

End Function

'------------------------------------------------------------------------------
' FR: Masque ou affiche les segments existants sans detruire leur identite.
' EN: Hides or shows existing segments without destroying their identity.
'------------------------------------------------------------------------------
Private Sub GanttDependency_SetExistingVisibility( _
    ByVal ws As Worksheet, _
    ByVal visibleValue As Boolean)

    Dim shapeName As Variant
    Dim expectedVisibility As MsoTriState
    Dim changedCount As Long

    If ws Is Nothing Then Exit Sub

    expectedVisibility = IIf(visibleValue, msoTrue, msoFalse)
    If gExistingDependencySegments Is Nothing Then Exit Sub

    For Each shapeName In gExistingDependencySegments.Keys
        If gExistingDependencySegments(shapeName).Visible <> expectedVisibility Then
            gExistingDependencySegments(shapeName).Visible = expectedVisibility
            changedCount = changedCount + 1
        End If
    Next shapeName

    gDependencyUpdates = gDependencyUpdates + changedCount
    Profiler_RecordOperation "GanttDiffDependencyVisibilityUpdated", changedCount, 0#

End Sub

'------------------------------------------------------------------------------
' FR: Supprime les segments DEP absents de la topologie graphique attendue.
' EN: Deletes DEP segments absent from the expected graphical topology.
'------------------------------------------------------------------------------
Private Sub GanttDependency_DeleteStaleSegments(ByVal ws As Worksheet)

    Dim shapeName As Variant
    Dim shp As Shape

    If ws Is Nothing Then Exit Sub
    If gExpectedDependencySegments Is Nothing Then Exit Sub

    If gExistingDependencySegments Is Nothing Then Exit Sub

    For Each shapeName In gExistingDependencySegments.Keys
        If Not gExpectedDependencySegments.Exists(CStr(shapeName)) Then
            Set shp = gExistingDependencySegments(CStr(shapeName))
            shp.Delete
            gDependencyDeletes = gDependencyDeletes + 1
        End If
    Next shapeName

End Sub

'------------------------------------------------------------------------------
' FR:
' Recharge le cache global des liens GANTT depuis tbl_LOGIC_LINKS avant dessin des dependances.
'
' EN:
' Reloads the global GANTT link cache from tbl_LOGIC_LINKS before dependency rendering.
'
' Entrees / Inputs:
' - tbl_LOGIC_LINKS via BuildExpandedLinksCacheFromLogicLinksTable.
'
' Sorties / Outputs:
' - gExpandedLinks remplace par le cache courant.
'
' Appele par / Called by:
' - DrawDependencyLinks.
'
' Notes:
' - Couplage direct CALC -> renderer; candidat a extraction Bridge/LinkProvider.
'------------------------------------------------------------------------------
Private Sub EnsureExpandedLinksCacheFromCalc()

    If gExpandedLinks Is Nothing Then
        Set gExpandedLinks = BuildExpandedLinksCacheFromLogicLinksTable()
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Function BuildExpandedLinksCacheFromLogicLinksTable() As Object

    Dim perfScope As clsPerfScope
    Dim network As clsParsedPlanningNetwork
    Dim link As clsParsedPlanningLink
    Dim d As Object
    Dim linkCol As Collection
    Dim tokenInfo As Object
    Dim r As Long
    Dim succId As String
    Dim predId As String
    Dim linkType As String

    Set perfScope = Profiler_BeginScope("BuildExpandedLinksCacheFromLogicLinksTable", "Excel Read")
    Set d = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Set network = ParsedPlanningNetwork_LoadCanonical()

    If Not network.HasColumn("Succ ID") Then GoTo SafeExit
    If Not network.HasColumn("Pred ID") Then GoTo SafeExit
    If Not network.HasColumn("Link Type") Then GoTo SafeExit
    If Not network.HasColumn("Lag") Then GoTo SafeExit

    For r = 1 To network.Count

        Set link = network.Item(r)
        succId = link.SuccId
        predId = link.PredId
        linkType = link.LinkType

        If succId = "" Then GoTo NextRow
        If predId = "" Then GoTo NextRow
        If linkType <> "FS" And linkType <> "SS" And linkType <> "FF" Then GoTo NextRow

        If Not d.Exists(succId) Then
            Set linkCol = New Collection
            d.Add succId, linkCol
        End If

        Set tokenInfo = CreateObject("Scripting.Dictionary")
        tokenInfo("PredID") = predId
        tokenInfo("LinkType") = linkType
        tokenInfo("Lag") = link.Lag
        tokenInfo("RawToken") = link.RawToken

        d(succId).Add tokenInfo

NextRow:
    Next r

SafeExit:
    Set BuildExpandedLinksCacheFromLogicLinksTable = d

End Function
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Function HasExpandedLinksAvailable() As Boolean

    On Error GoTo SafeExit

    If gExpandedLinks Is Nothing Then Exit Function
    If gExpandedLinks.Count <= 0 Then Exit Function

    HasExpandedLinksAvailable = True
    Exit Function

SafeExit:
    HasExpandedLinksAvailable = False

End Function
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Sub GetLinkAnchorTypes( _
    ByVal linkType As String, _
    ByRef predAnchorType As String, _
    ByRef succAnchorType As String)

    Select Case UCase$(Trim$(linkType))
        Case "SS"
            predAnchorType = LINK_ANCHOR_START
            succAnchorType = LINK_ANCHOR_START

        Case "FF"
            predAnchorType = LINK_ANCHOR_FINISH
            succAnchorType = LINK_ANCHOR_FINISH

        Case Else
            predAnchorType = LINK_ANCHOR_FINISH
            succAnchorType = LINK_ANCHOR_START
    End Select

End Sub
'------------------------------------------------------------------------------
' FR: Calcule ou dessine une partie des liens de dependance visibles dans le GANTT.
' EN: Computes or draws part of the dependency links visible in GANTT.
'------------------------------------------------------------------------------
Private Function GetLinkReferenceDate( _
    ByVal taskId As String, _
    ByVal anchorType As String, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean) As Variant

    Select Case UCase$(Trim$(anchorType))
        Case LINK_ANCHOR_START
            GetLinkReferenceDate = GanttLive_GetDisplayStart(taskId, baseById, testById, isTestMode)

        Case Else
            GetLinkReferenceDate = GanttLive_GetDisplayFinish(taskId, baseById, testById, isTestMode)
    End Select

End Function

Private Function GanttDependency_IsPointMarkerTarget( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal dataRow As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean) As Boolean

    Dim taskId As String
    Dim startValue As Variant
    Dim finishValue As Variant

    If TaskTypeRules_IsMilestoneRow(dataArr, mapWBS, dataRow, VTS_COL_TASK_TYPE) Then
        GanttDependency_IsPointMarkerTarget = True
        Exit Function
    End If

    taskId = Trim$(CStr(dataArr(dataRow, mapWBS(VTS_COL_ID))))
    startValue = GetRenderStartForCurrentScale( _
        GanttLive_GetDisplayStart(taskId, baseById, testById, isTestMode))
    finishValue = GetRenderFinishForCurrentScale( _
        GanttLive_GetDisplayFinish(taskId, baseById, testById, isTestMode))

    If Not HasValue(startValue) Or Not HasValue(finishValue) Then Exit Function
    GanttDependency_IsPointMarkerTarget = (CDbl(finishValue) - CDbl(startValue) + 1 <= 1)

End Function
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Anchor Point dans le workflow de rendu GANTT.
' EN: Runs the Get Task Anchor Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskAnchorPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal isFinishSide As Boolean, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim wbs As String
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim sizeVal As Double

    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    wbs = NormalizeWBS(CStr(dataArr(dataRow, mapWBS(VTS_COL_WBS))))
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS(VTS_COL_ID))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    yOut = ws.cells(ganttRow, FIRST_TIMELINE_COL).Top + (ws.rows(ganttRow).Height / 2)

    If hasChildren.Exists(wbs) Then
        If isFinishSide Then
            xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
        Else
            xOut = TimelineLeft(ws, projectStart, startVal)
        End If

    ElseIf durationVal <= 1 Then
        sizeVal = ws.rows(ganttRow).Height - 6

        If isFinishSide Then
            xOut = GetTaskMidX(ws, projectStart, startVal) + (sizeVal / 2)
        Else
            xOut = GetTaskMidX(ws, projectStart, startVal) - (sizeVal / 2)
        End If

    Else
        If isFinishSide Then
            xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
        Else
            xOut = TimelineLeft(ws, projectStart, startVal)
        End If
    End If

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Anchor Point By Type dans le workflow de rendu GANTT.
' EN: Runs the Get Task Anchor Point By Type helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskAnchorPointByType( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal anchorType As String, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    GetTaskAnchorPointBySide ws, mapWBS, dataArr, hasChildren, projectStart, totalDays, dataRow, _
                             IIf(UCase$(Trim$(anchorType)) = LINK_ANCHOR_START, "LEFT", "RIGHT"), _
                             xOut, yOut, baseById, testById, isTestMode

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Finish Entry Point dans le workflow de rendu GANTT.
' EN: Runs the Get Task Finish Entry Point helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskFinishEntryPoint( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim sizeVal As Double
    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS(VTS_COL_ID))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    yOut = ws.cells(ganttRow, FIRST_TIMELINE_COL).Top + (ws.rows(ganttRow).Height / 2)

    If durationVal <= 1 Then
        sizeVal = ws.rows(ganttRow).Height - 6
        xOut = GetTaskMidX(ws, projectStart, startVal) + (sizeVal / 2)
    Else
        xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
    End If

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Left X dans le workflow de rendu GANTT.
' EN: Runs the Get Task Left X helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function GetTaskLeftX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    GetTaskLeftX = TimelineLeft(ws, projectStart, taskDate)

End Function
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Right X dans le workflow de rendu GANTT.
' EN: Runs the Get Task Right X helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function GetTaskRightX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    GetTaskRightX = TimelineRightAfterFinish(ws, projectStart, taskDate)

End Function
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Mid X dans le workflow de rendu GANTT.
' EN: Runs the Get Task Mid X helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetTaskMidX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    GetTaskMidX = TimelineDateRangeMidX(ws, projectStart, taskDate, taskDate)

End Function
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Anchor Point By Side dans le workflow de rendu GANTT.
' EN: Runs the Get Task Anchor Point By Side helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetTaskAnchorPointBySide( _
    ByVal ws As Worksheet, _
    ByVal mapWBS As Object, _
    ByRef dataArr As Variant, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal dataRow As Long, _
    ByVal anchorSide As String, _
    ByRef xOut As Double, _
    ByRef yOut As Double, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim ganttRow As Long
    Dim wbs As String
    Dim idVal As String
    Dim startVal As Variant
    Dim finishVal As Variant
    Dim durationVal As Double
    Dim sizeVal As Double
    Dim timelineLeftBound As Double
    Dim timelineRightBound As Double

    ganttRow = FIRST_TASK_ROW + dataRow - 1
    wbs = NormalizeWBS(CStr(dataArr(dataRow, mapWBS(VTS_COL_WBS))))
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS(VTS_COL_ID))))

    startVal = GetRenderStartForCurrentScale(GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode))
    finishVal = GetRenderFinishForCurrentScale(GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode))

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1

    timelineLeftBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).Left + LINK_EDGE_PADDING
    timelineRightBound = ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Left + _
                         ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + totalDays - 1).Width - LINK_EDGE_PADDING

    yOut = ws.cells(ganttRow, FIRST_TIMELINE_COL).Top + (ws.rows(ganttRow).Height / 2)

    If hasChildren.Exists(wbs) Then
        Select Case UCase$(Trim$(anchorSide))
            Case "LEFT"
                xOut = TimelineLeft(ws, projectStart, startVal)
            Case "RIGHT"
                xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
            Case Else
                xOut = GetTaskMidX(ws, projectStart, startVal)
        End Select

    ElseIf durationVal <= 1 Then
        sizeVal = ws.rows(ganttRow).Height - 6

        Select Case UCase$(Trim$(anchorSide))
            Case "LEFT"
                xOut = GetTaskMidX(ws, projectStart, startVal) - (sizeVal / 2)
            Case "RIGHT"
                xOut = GetTaskMidX(ws, projectStart, startVal) + (sizeVal / 2)
            Case Else
                xOut = GetTaskMidX(ws, projectStart, startVal)
        End Select

    Else
        Select Case UCase$(Trim$(anchorSide))
            Case "LEFT"
                xOut = TimelineLeft(ws, projectStart, startVal)
            Case "RIGHT"
                xOut = TimelineRightAfterFinish(ws, projectStart, finishVal)
            Case Else
                xOut = GetTaskMidX(ws, projectStart, startVal)
        End Select
    End If

    If xOut < timelineLeftBound Then xOut = timelineLeftBound
    If xOut > timelineRightBound Then xOut = timelineRightBound

End Sub
