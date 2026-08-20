Attribute VB_Name = "mod_GanttDependencySvg"
Option Explicit

'===============================================================================
' MODULE : mod_GanttDependencySvg
' DOMAIN : Gantt dependency projection
'
' Owns the SVG materialization of canonical logical routes. The historical
' segment renderer remains the explicit fallback and retains all routing rules.
' No cache is shared between workbooks: module state belongs to ThisWorkbook.
'===============================================================================

Private Const SVG_MODE As String = "SVG"
Private Const HISTORICAL_MODE As String = "HISTORICAL"
Private Const SVG_LAYER_NAME As String = "GANTT_DEPENDENCY_SVG"
Private Const SVG_STAGING_PREFIX As String = "GANTT_DEPENDENCY_SVG_STAGING_"
Private Const SVG_MARGIN As Double = 5#
Private Const SVG_PHYSICAL_TOLERANCE As Double = 0.1
Private Const SVG_LINE_COLOR As String = "#787878"
Private Const SVG_LINE_WEIGHT As Double = 1#
Private Const SVG_ARROW_LENGTH As Double = 8#
Private Const SVG_ARROW_HALF_WIDTH As Double = 3.6
Private Const SVG_CACHE_SHEET_NAME As String = "__PE_SVG_CACHE"
Private Const SVG_CACHE_VERSION As String = "SVG_ROUTE_CACHE_V2"
Private Const SVG_CACHE_ALT_PREFIX As String = "ProjectEngineSvgCache|"
Private Const SVG_CACHE_TAG_SIGNATURE As String = "PE_SVG_CACHE_SIGNATURE"

Private gModeInitialized As Boolean
Private gRequestedMode As String
Private gActiveMode As String
Private gRoutesById As Object
Private gDirtyLinkIds As Object
Private gFullModelDirty As Boolean
Private gLastContentHash As String
Private gLastTaskGeometrySignature As String
Private gLastTaskGeometryByShape As Object
Private gLastFallbackReason As String
Private gLayerVisible As Boolean
Private gGenerationCount As Long
Private gInsertionCount As Long
Private gFallbackCount As Long
Private gLastDirtyLinkCount As Long

Public Function GanttDependencySvg_IsLayerName(ByVal shapeName As String) As Boolean

    GanttDependencySvg_IsLayerName = _
        (shapeName = SVG_LAYER_NAME) Or _
        (Left$(shapeName, Len(SVG_STAGING_PREFIX)) = SVG_STAGING_PREFIX)

End Function

Public Sub GanttDependencySvg_SetRenderMode(ByVal renderMode As String)

    Dim ws As Worksheet

    EnsureModeInitialized
    renderMode = UCase$(Trim$(renderMode))
    If renderMode <> SVG_MODE And renderMode <> HISTORICAL_MODE Then
        Err.Raise 5, "GanttDependencySvg_SetRenderMode", "Unsupported dependency render mode: " & renderMode
    End If

    gRequestedMode = renderMode
    If renderMode = HISTORICAL_MODE Then
        gActiveMode = HISTORICAL_MODE
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("GANTT")
        On Error GoTo 0
        If Not ws Is Nothing Then GanttDependencySvg_SetVisible ws, False
    End If
    Profiler_RecordOperation "GanttDependencyMode_" & renderMode, 1, 0#

End Sub

Public Function GanttDependencySvg_IsRequested() As Boolean

    EnsureModeInitialized
    GanttDependencySvg_IsRequested = (gRequestedMode = SVG_MODE)

End Function

Public Function GanttDependencySvg_IsActive() As Boolean

    EnsureModeInitialized
    GanttDependencySvg_IsActive = (gActiveMode = SVG_MODE)

End Function

Public Function GanttDependencySvg_HasLayer(ByVal ws As Worksheet) As Boolean

    Dim shp As Shape

    Set shp = GetLayerShape(ws)
    GanttDependencySvg_HasLayer = Not shp Is Nothing

End Function

Public Function GanttDependencySvg_GetActiveMode() As String

    EnsureModeInitialized
    GanttDependencySvg_GetActiveMode = gActiveMode

End Function

Public Function GanttDependencySvg_GetFallbackReason() As String
    GanttDependencySvg_GetFallbackReason = gLastFallbackReason
End Function

Private Sub EnsureModeInitialized()

    If gModeInitialized Then Exit Sub
    gRequestedMode = SVG_MODE
    gActiveMode = HISTORICAL_MODE
    Set gDirtyLinkIds = CreateObject("Scripting.Dictionary")
    gFullModelDirty = True
    gModeInitialized = True

End Sub

Public Sub GanttDependencySvg_Invalidate( _
    Optional ByVal reason As String = "", _
    Optional ByVal discardRoutes As Boolean = True)

    EnsureModeInitialized

    If discardRoutes Then Set gRoutesById = Nothing
    gLastTaskGeometrySignature = vbNullString
    Set gLastTaskGeometryByShape = Nothing
    gFullModelDirty = True
    Set gDirtyLinkIds = CreateObject("Scripting.Dictionary")
    gLastDirtyLinkCount = 0

    If Len(Trim$(reason)) > 0 Then
        Profiler_RecordOperation "GanttDependencySvgInvalidation_" & reason, 1, 0#
    End If

End Sub

Public Sub GanttDependencySvg_InvalidatePersistentCache( _
    ByVal ws As Worksheet, _
    Optional ByVal reason As String = "")

    Dim layer As Shape
    Dim cacheWs As Worksheet

    EnsureModeInitialized

    On Error Resume Next
    Set layer = GetLayerShape(ws)
    If Not layer Is Nothing Then
        layer.Tags.Delete SVG_CACHE_TAG_SIGNATURE
        layer.AlternativeText = ""
    End If

    Set cacheWs = GetSvgCacheSheet(False)
    If Not cacheWs Is Nothing Then
        cacheWs.Cells.Clear
        cacheWs.Range("A1").Value = "RecordType"
        cacheWs.Range("A2").Value = "INVALIDATED"
        cacheWs.Range("B2").Value = SVG_CACHE_VERSION
        cacheWs.Range("C2").Value = reason
        cacheWs.Range("D2").Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        cacheWs.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0

    GanttDependencySvg_Invalidate reason, True
    If Len(Trim$(reason)) > 0 Then
        Profiler_RecordOperation "GanttDependencySvgPersistentInvalidation_" & reason, 1, 0#
    End If

End Sub

Public Sub GanttDependencySvg_BeginFullModel()

    EnsureModeInitialized
    Set gRoutesById = CreateObject("Scripting.Dictionary")
    Set gDirtyLinkIds = CreateObject("Scripting.Dictionary")
    gLastTaskGeometrySignature = vbNullString
    Set gLastTaskGeometryByShape = Nothing
    gLastDirtyLinkCount = 0
    gFullModelDirty = True

End Sub

Public Sub GanttDependencySvg_AcceptAggregatedScaleHidden(ByVal ws As Worksheet)

    EnsureModeInitialized
    Set gRoutesById = CreateObject("Scripting.Dictionary")
    Set gDirtyLinkIds = CreateObject("Scripting.Dictionary")
    gLastTaskGeometrySignature = GanttDependencySvg_CurrentTaskGeometrySignature(ws)
    Set gLastTaskGeometryByShape = GanttDependencySvg_BuildTaskGeometryMap(ws)
    gLastFallbackReason = vbNullString
    gFullModelDirty = False
    gLayerVisible = False
    gActiveMode = SVG_MODE
    Profiler_RecordOperation "GanttDependencySvgAggregatedScaleHidden", 1, 0#

End Sub

Public Sub GanttDependencySvg_StoreRoute(ByVal route As clsGanttDependencyRoute)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("DependencySvg_RouteRegistryUpdate", "Dependency SVG")
    EnsureModeInitialized
    If route Is Nothing Then Exit Sub
    If gRoutesById Is Nothing Then Set gRoutesById = CreateObject("Scripting.Dictionary")

    If route.SegmentCount = 0 Then
        If gRoutesById.Exists(route.LinkId) Then gRoutesById.Remove route.LinkId
        Exit Sub
    End If

    If gRoutesById.Exists(route.LinkId) Then
        Set gRoutesById(route.LinkId) = route
    Else
        gRoutesById.Add route.LinkId, route
    End If

    If Not gDirtyLinkIds Is Nothing Then
        If gDirtyLinkIds.Exists(route.LinkId) Then gDirtyLinkIds.Remove route.LinkId
    End If

End Sub

Public Function GanttDependencySvg_HasRoutes() As Boolean

    GanttDependencySvg_HasRoutes = Not gRoutesById Is Nothing
    If GanttDependencySvg_HasRoutes Then GanttDependencySvg_HasRoutes = (gRoutesById.Count > 0)

End Function

Public Function GanttDependencySvg_HasRoute(ByVal linkId As String) As Boolean

    If gRoutesById Is Nothing Then Exit Function
    GanttDependencySvg_HasRoute = gRoutesById.Exists(linkId)

End Function

Public Function GanttDependencySvg_ShouldRebuildRoute(ByVal linkId As String) As Boolean

    If gRoutesById Is Nothing Then
        GanttDependencySvg_ShouldRebuildRoute = True
    ElseIf gFullModelDirty Then
        GanttDependencySvg_ShouldRebuildRoute = True
    ElseIf gDirtyLinkIds Is Nothing Then
        GanttDependencySvg_ShouldRebuildRoute = False
    Else
        GanttDependencySvg_ShouldRebuildRoute = gDirtyLinkIds.Exists(linkId)
    End If

End Function

Public Sub GanttDependencySvg_ReconcileRouteIds( _
    ByVal expectedRouteIds As Object, _
    ByVal changedRouteIds As Object)

    Dim routeId As Variant
    Dim staleIds As Collection
    Dim staleId As Variant

    EnsureModeInitialized
    If gRoutesById Is Nothing Then Exit Sub
    Set staleIds = New Collection

    For Each routeId In gRoutesById.Keys
        If expectedRouteIds Is Nothing Then
            staleIds.Add CStr(routeId)
        ElseIf Not expectedRouteIds.Exists(CStr(routeId)) Then
            staleIds.Add CStr(routeId)
        End If
    Next routeId

    For Each staleId In staleIds
        gRoutesById.Remove CStr(staleId)
    Next staleId

    If Not changedRouteIds Is Nothing Then GanttDependencySvg_MarkLinksDirty changedRouteIds
    Profiler_RecordOperation "GanttDependencySvgRoutesRemoved", staleIds.Count, 0#

End Sub

Public Sub GanttDependencySvg_MarkLinksDirty(ByVal linkIds As Object)

    Dim linkId As Variant

    EnsureModeInitialized
    If gDirtyLinkIds Is Nothing Then Set gDirtyLinkIds = CreateObject("Scripting.Dictionary")
    If linkIds Is Nothing Then Exit Sub

    For Each linkId In linkIds.Keys
        gDirtyLinkIds(CStr(linkId)) = True
    Next linkId
    gLastDirtyLinkCount = linkIds.Count

    Profiler_RecordOperation "GanttDependencySvgLinksDirty", linkIds.Count, 0#

End Sub

Public Function GanttDependencySvg_CanReuseVisibleDayLayer(ByVal ws As Worksheet) As Boolean

    Dim shp As Shape

    EnsureModeInitialized
    If gRequestedMode <> SVG_MODE Then Exit Function
    If gFullModelDirty Then Exit Function
    If Not gDirtyLinkIds Is Nothing Then
        If gDirtyLinkIds.Count > 0 Then Exit Function
    End If
    If Not GanttDependencySvg_HasRoutes() Then Exit Function
    If Not GanttDependencySvg_IsTaskGeometryCurrent(ws) Then
        gFullModelDirty = True
        gLastFallbackReason = "TaskGeometryChangedRequiresRouteRebuild"
        Profiler_RecordOperation "GanttDependencySvgTaskGeometryInvalidations", 1, 0#
        Exit Function
    End If
    If Not GanttDependencySvg_AllRoutesHaveTerminalArrow() Then Exit Function

    Set shp = GetLayerShape(ws)
    If shp Is Nothing Then Exit Function
    If Not GanttDependencySvg_IsLayerPhysicalStateCurrent(ws) Then
        gLastFallbackReason = "SvgLayerPhysicalStateMismatch"
        Profiler_RecordOperation "GanttDependencySvgPhysicalInvalidations", 1, 0#
        Exit Function
    End If

    If shp.Visible <> msoTrue Then shp.Visible = msoTrue
    gLayerVisible = True
    gActiveMode = SVG_MODE
    Profiler_RecordOperation "GanttDependencySvgDayReuse", 1, 0#
    GanttDependencySvg_CanReuseVisibleDayLayer = True

End Function

Public Function GanttDependencySvg_ArrowheadDimensions() As Variant

    Dim outArr(1 To 1, 1 To 2) As Variant

    outArr(1, 1) = SVG_ARROW_LENGTH
    outArr(1, 2) = SVG_ARROW_HALF_WIDTH
    GanttDependencySvg_ArrowheadDimensions = outArr

End Function

Public Sub GanttDependencySvg_EnsureFrontLayer(ByVal ws As Worksheet)

    Dim shp As Shape

    If ws Is Nothing Then Exit Sub
    Set shp = GetLayerShape(ws)
    If shp Is Nothing Then Exit Sub
    If shp.Visible <> msoTrue Then Exit Sub
    If shp.ZOrderPosition <> ws.Shapes.Count Then shp.ZOrder msoBringToFront

End Sub

Public Sub GanttDependencySvg_SetVisible(ByVal ws As Worksheet, ByVal makeVisible As Boolean)

    Dim shp As Shape
    Dim target As MsoTriState

    Set shp = GetLayerShape(ws)
    If shp Is Nothing Then Exit Sub

    target = IIf(makeVisible, msoTrue, msoFalse)
    If shp.Visible <> target Then shp.Visible = target
    gLayerVisible = makeVisible

End Sub

Public Function GanttDependencySvg_TryHydratePersistentCache( _
    ByVal ws As Worksheet, _
    Optional ByVal allowStaleTaskGeometry As Boolean = False) As Boolean

    Dim perfScope As clsPerfScope
    Dim layer As Shape
    Dim cacheWs As Worksheet
    Dim arr As Variant
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim signature As String
    Dim layerSignature As String
    Dim contentHash As String
    Dim routeById As Object
    Dim route As clsGanttDependencyRoute
    Dim linkId As String
    Dim segmentCount As Long
    Dim taskGeometrySignature As String
    Dim currentTaskGeometrySignature As String
    Dim staleTaskGeometry As Boolean

    Set perfScope = Profiler_BeginScope("GanttDependencySvg_HydrateCache", "Dependency SVG")
    EnsureModeInitialized
    On Error GoTo Failed

    If gRequestedMode <> SVG_MODE Then Exit Function
    Set layer = GetLayerShape(ws)
    If layer Is Nothing Then Exit Function

    layerSignature = GetLayerCacheSignature(layer)
    If Len(layerSignature) = 0 Then Exit Function

    Set cacheWs = GetSvgCacheSheet(False)
    If cacheWs Is Nothing Then Exit Function
    If cacheWs.UsedRange.Rows.Count < 3 Then Exit Function

    arr = cacheWs.UsedRange.Value
    If Not IsArray(arr) Then Exit Function
    rowCount = UBound(arr, 1)
    If CStr(arr(2, 1)) <> "META" Then Exit Function
    If CStr(arr(2, 2)) <> SVG_CACHE_VERSION Then Exit Function

    signature = CStr(arr(2, 3))
    contentHash = CStr(arr(2, 4))
    If Len(signature) = 0 Or signature <> layerSignature Then Exit Function
    If UBound(arr, 2) >= 9 Then taskGeometrySignature = CStr(arr(2, 9))
    If Len(taskGeometrySignature) = 0 Then
        gLastFallbackReason = "SvgCacheMissingTaskGeometrySignature"
        Exit Function
    End If
    currentTaskGeometrySignature = GanttDependencySvg_CurrentTaskGeometrySignature(ws)
    If taskGeometrySignature <> currentTaskGeometrySignature Then
        If allowStaleTaskGeometry Then
            staleTaskGeometry = True
            Profiler_RecordOperation "GanttDependencySvgCacheHydratedWithStaleTaskGeometry", 1, 0#
        Else
            gLastFallbackReason = "SvgCacheTaskGeometrySignatureMismatch"
            Exit Function
        End If
    End If
    If Len(currentTaskGeometrySignature) = 0 Then
        gLastFallbackReason = "SvgCacheTaskGeometrySignatureMismatch"
        Exit Function
    End If

    Set routeById = CreateObject("Scripting.Dictionary")
    For rowIndex = 3 To rowCount
        If CStr(arr(rowIndex, 1)) = "ROUTE" Then
            linkId = CStr(arr(rowIndex, 3))
            If Len(linkId) > 0 Then
                If routeById.Exists(linkId) Then
                    Set route = routeById(linkId)
                Else
                    Set route = New clsGanttDependencyRoute
                    route.Initialize linkId, CStr(arr(rowIndex, 4)), CStr(arr(rowIndex, 5)), _
                        CStr(arr(rowIndex, 6)), CDbl(Val(CStr(arr(rowIndex, 7))))
                    routeById.Add linkId, route
                End If
                route.AddSegment _
                    CDbl(Val(CStr(arr(rowIndex, 9)))), _
                    CDbl(Val(CStr(arr(rowIndex, 10)))), _
                    CDbl(Val(CStr(arr(rowIndex, 11)))), _
                    CDbl(Val(CStr(arr(rowIndex, 12)))), _
                    SvgParseBoolean(arr(rowIndex, 13))
                segmentCount = segmentCount + 1
            End If
        End If
    Next rowIndex

    If routeById.Count = 0 Or segmentCount = 0 Then Exit Function

    Set gRoutesById = routeById
    Set gDirtyLinkIds = CreateObject("Scripting.Dictionary")
    gFullModelDirty = staleTaskGeometry
    gLastContentHash = contentHash
    gLastTaskGeometrySignature = taskGeometrySignature
    Set gLastTaskGeometryByShape = GanttDependencySvg_BuildTaskGeometryMap(ws)
    gLastFallbackReason = ""
    gLayerVisible = True
    gActiveMode = SVG_MODE
    If layer.Visible <> msoTrue Then layer.Visible = msoTrue

    Profiler_RecordOperation "GanttDependencySvgCacheHydrated", routeById.Count, 0#
    Profiler_RecordOperation "GanttDependencySvgCacheHydratedSegments", segmentCount, 0#
    GanttDependencySvg_TryHydratePersistentCache = True
    Exit Function

Failed:
    Profiler_RecordOperation "GanttDependencySvgCacheHydrateFailed", 1, 0#

End Function

Public Function GanttDependencySvg_TryCommit(ByVal ws As Worksheet) As Boolean

    Dim perfScope As clsPerfScope
    Dim stageScope As clsPerfScope
    Dim finalizeScope As clsPerfScope
    Dim svgText As String
    Dim tempPath As String
    Dim contentHash As String
    Dim minX As Double
    Dim minY As Double
    Dim maxX As Double
    Dim maxY As Double
    Dim layerWidth As Double
    Dim layerHeight As Double
    Dim oldLayer As Shape
    Dim newLayer As Shape
    Dim stagingName As String
    Dim cacheSignature As String
    Dim taskGeometrySignature As String
    Dim taskGeometryByShape As Object
    Dim dirtyRouteCount As Long
    Dim cleanRouteCount As Long

    Set perfScope = Profiler_BeginScope("GanttDependencySvg_TryCommit", "Dependency SVG")
    EnsureModeInitialized
    On Error GoTo Failed

    If gRequestedMode <> SVG_MODE Then
        gLastFallbackReason = "HistoricalModeRequested"
        Exit Function
    End If
    If Not GanttDependencySvg_HasRoutes() Then
        gLastFallbackReason = "NoLogicalRoutes"
        Exit Function
    End If

    dirtyRouteCount = gLastDirtyLinkCount
    cleanRouteCount = gRoutesById.Count - dirtyRouteCount
    If cleanRouteCount < 0 Then cleanRouteCount = 0
    Profiler_RecordOperation "DependencySvgRoutesDirtyAtCommit", dirtyRouteCount, 0#
    Profiler_RecordOperation "DependencySvgRoutesCleanAtCommit", cleanRouteCount, 0#

    Set stageScope = Profiler_BeginScope("DependencySvg_SerializeDocument", "Dependency SVG")
    svgText = BuildSvgDocument(minX, minY, maxX, maxY)
    Set stageScope = Nothing
    If Len(svgText) = 0 Then
        gLastFallbackReason = "EmptySvg"
        Exit Function
    End If

    Set finalizeScope = Profiler_BeginScope("DependencySvg_FinalizeState", "Dependency SVG")
    Set stageScope = Profiler_BeginScope("DependencySvg_HashSvgText", "Dependency SVG")
    contentHash = StableTextHash(svgText)
    Set stageScope = Nothing
    Set stageScope = Profiler_BeginScope("DependencySvg_BuildRouteSignature", "Dependency SVG")
    cacheSignature = BuildRouteSnapshotSignature(contentHash)
    Set stageScope = Nothing
    Set stageScope = Profiler_BeginScope("DependencySvg_TaskGeometrySignature", "Dependency SVG")
    taskGeometrySignature = GanttDependencySvg_CurrentTaskGeometrySnapshot(ws, taskGeometryByShape)
    Set stageScope = Nothing
    Set finalizeScope = Nothing
    Set stageScope = Profiler_BeginScope("DependencySvg_FindExistingPicture", "Dependency SVG")
    Set oldLayer = GetLayerShape(ws)
    Set stageScope = Nothing
    If Not oldLayer Is Nothing Then
        If contentHash = gLastContentHash Then
            If Len(gLastTaskGeometrySignature) > 0 Then
                If gLastTaskGeometrySignature <> taskGeometrySignature Then
                    gFullModelDirty = True
                    gLastFallbackReason = "TaskGeometryChangedWithoutRouteRebuild"
                    Profiler_RecordOperation "GanttDependencySvgStaleGeometryBlocked", 1, 0#
                    Exit Function
                End If
            End If
            Set stageScope = Profiler_BeginScope("DependencySvg_PositionPicture", "Dependency SVG")
            ApplyLayerPhysicalState oldLayer, minX, minY, maxX - minX, maxY - minY
            Set stageScope = Nothing
            gFullModelDirty = False
            Set gDirtyLinkIds = CreateObject("Scripting.Dictionary")
            gLastDirtyLinkCount = 0
            gLayerVisible = True
            gActiveMode = SVG_MODE
            gLastTaskGeometrySignature = taskGeometrySignature
            Set gLastTaskGeometryByShape = taskGeometryByShape
            ApplyLayerCacheSignature oldLayer, cacheSignature
            Set stageScope = Profiler_BeginScope("DependencySvg_PersistCache", "Dependency SVG")
            SavePersistentRouteCache ws, contentHash, cacheSignature
            Set stageScope = Nothing
            Profiler_RecordOperation "GanttDependencySvgUnchanged", 1, 0#
            GanttDependencySvg_TryCommit = True
            Exit Function
        End If
    End If

    Set stageScope = Profiler_BeginScope("DependencySvg_WriteTempFile", "Dependency SVG")
    tempPath = WriteSvgTempFile(svgText)
    Set stageScope = Nothing
    layerWidth = maxX - minX
    layerHeight = maxY - minY
    If layerWidth < 1# Then layerWidth = 1#
    If layerHeight < 1# Then layerHeight = 1#

    stagingName = SVG_STAGING_PREFIX & CStr(gGenerationCount + 1)
    Set stageScope = Profiler_BeginScope("DependencySvg_AddPicture", "Dependency SVG")
    Set newLayer = ws.Shapes.AddPicture( _
        tempPath, msoFalse, msoTrue, _
        minX, minY, layerWidth, layerHeight)
    Set stageScope = Nothing
    Set stageScope = Profiler_BeginScope("DependencySvg_PositionPicture", "Dependency SVG")
    newLayer.Name = stagingName
    newLayer.LockAspectRatio = msoFalse
    newLayer.Placement = xlFreeFloating
    newLayer.Visible = msoFalse
    Set stageScope = Nothing

    Set stageScope = Profiler_BeginScope("DependencySvg_DeleteExistingPicture", "Dependency SVG")
    If Not oldLayer Is Nothing Then oldLayer.Delete
    Set stageScope = Nothing
    Set stageScope = Profiler_BeginScope("DependencySvg_PositionPicture", "Dependency SVG")
    newLayer.Name = SVG_LAYER_NAME
    ApplyLayerCacheSignature newLayer, cacheSignature
    newLayer.Visible = msoTrue
    Set stageScope = Nothing

    Set stageScope = Profiler_BeginScope("DependencySvg_CleanupTempResources", "Dependency SVG")
    DeleteHistoricalSegments ws
    Set stageScope = Nothing
    Set stageScope = Profiler_BeginScope("DependencySvg_FinalizeState", "Dependency SVG")
    gGenerationCount = gGenerationCount + 1
    gInsertionCount = gInsertionCount + 1
    gLastContentHash = contentHash
    gLastTaskGeometrySignature = taskGeometrySignature
    Set gLastTaskGeometryByShape = taskGeometryByShape
    gLastFallbackReason = ""
    gFullModelDirty = False
    Set gDirtyLinkIds = CreateObject("Scripting.Dictionary")
    gLastDirtyLinkCount = 0
    gLayerVisible = True
    gActiveMode = SVG_MODE
    Set stageScope = Nothing
    Set stageScope = Profiler_BeginScope("DependencySvg_PersistCache", "Dependency SVG")
    SavePersistentRouteCache ws, contentHash, cacheSignature
    Set stageScope = Nothing

    Profiler_RecordOperation "GanttDependencySvgLayersGenerated", 1, 0#
    Profiler_RecordOperation "GanttDependencySvgBytes", Len(svgText), 0#
    Profiler_RecordOperation "GanttDependencySvgRoutes", gRoutesById.Count, 0#
    GanttDependencySvg_TryCommit = True

CleanExit:
    Set stageScope = Profiler_BeginScope("DependencySvg_CleanupTempResources", "Dependency SVG")
    DeleteTempFile tempPath
    Set stageScope = Nothing
    Exit Function

Failed:
    gFallbackCount = gFallbackCount + 1
    gLastFallbackReason = "SvgCommitError_" & CStr(Err.Number)
    gActiveMode = HISTORICAL_MODE
    Profiler_RecordOperation "GanttDependencySvgFallback_" & gLastFallbackReason, 1, 0#
    On Error Resume Next
    If Not newLayer Is Nothing Then newLayer.Delete
    If Not oldLayer Is Nothing Then oldLayer.Visible = msoFalse
    On Error GoTo 0
    GoTo CleanExit

End Function

Public Sub GanttDependencySvg_ActivateHistorical(ByVal ws As Worksheet, Optional ByVal reason As String = "")

    EnsureModeInitialized
    GanttDependencySvg_SetVisible ws, False
    gActiveMode = HISTORICAL_MODE
    If Len(reason) > 0 Then
        gLastFallbackReason = reason
        Profiler_RecordOperation "GanttDependencySvgFallback_" & reason, 1, 0#
    End If

End Sub

Public Function GanttDependencySvg_RouteCount() As Long
    If Not gRoutesById Is Nothing Then GanttDependencySvg_RouteCount = gRoutesById.Count
End Function

Public Function GanttDependencySvg_SegmentCount() As Long

    Dim routeId As Variant
    Dim route As clsGanttDependencyRoute

    If gRoutesById Is Nothing Then Exit Function
    For Each routeId In gRoutesById.Keys
        Set route = gRoutesById(CStr(routeId))
        GanttDependencySvg_SegmentCount = GanttDependencySvg_SegmentCount + route.SegmentCount
    Next routeId

End Function

Public Function GanttDependencySvg_CurrentTaskGeometrySignature(ByVal ws As Worksheet) As String

    Dim shapeValues As Object

    GanttDependencySvg_CurrentTaskGeometrySignature = GanttDependencySvg_CurrentTaskGeometrySnapshot(ws, shapeValues)

End Function

Private Function GanttDependencySvg_CurrentTaskGeometrySnapshot( _
    ByVal ws As Worksheet, _
    ByRef shapeValues As Object) As String

    Dim keys As Variant
    Dim key As Variant
    Dim textValue As String

    If ws Is Nothing Then Exit Function

    On Error GoTo Failed
    Set shapeValues = GanttDependencySvg_BuildTaskGeometryMap(ws)
    If shapeValues Is Nothing Then GoTo Failed

    If shapeValues.Count = 0 Then
        GanttDependencySvg_CurrentTaskGeometrySnapshot = "TASK_GEOMETRY|EMPTY"
        Exit Function
    End If

    keys = SortedDictionaryKeys(shapeValues)
    For Each key In keys
        textValue = textValue & "|" & CStr(shapeValues(CStr(key)))
    Next key

    GanttDependencySvg_CurrentTaskGeometrySnapshot = "TASK_GEOMETRY|" & StableTextHash(textValue)
    Exit Function

Failed:
    GanttDependencySvg_CurrentTaskGeometrySnapshot = vbNullString

End Function

Private Function GanttDependencySvg_BuildTaskGeometryMap(ByVal ws As Worksheet) As Object

    Dim shapeValues As Object
    Dim shp As Shape
    Dim shapeName As String
    Dim i As Long

    If ws Is Nothing Then Exit Function

    Set shapeValues = CreateObject("Scripting.Dictionary")
    On Error GoTo Failed
    For i = 1 To ws.Shapes.Count
        Set shp = ws.Shapes(i)
        shapeName = CStr(shp.Name)
        If IsTaskGeometrySignatureShape(shapeName) Then
            shapeValues(shapeName) = shapeName & "|" & _
                SvgNumber(CDbl(shp.Left)) & "|" & _
                SvgNumber(CDbl(shp.Top)) & "|" & _
                SvgNumber(CDbl(shp.Width)) & "|" & _
                SvgNumber(CDbl(shp.Height)) & "|" & _
                SvgBoolToken(shp.Visible = msoTrue)
        End If
    Next i

    Set GanttDependencySvg_BuildTaskGeometryMap = shapeValues
    Exit Function

Failed:
    Set GanttDependencySvg_BuildTaskGeometryMap = Nothing

End Function

Private Function GanttDependencySvg_ShapeGeometryChanged( _
    ByVal shapeName As String, _
    ByVal currentByShape As Object) As Boolean

    If Len(shapeName) = 0 Then Exit Function
    If currentByShape Is Nothing Then Exit Function
    If gLastTaskGeometryByShape Is Nothing Then Exit Function

    If Not gLastTaskGeometryByShape.Exists(shapeName) Then Exit Function
    If Not currentByShape.Exists(shapeName) Then
        GanttDependencySvg_ShapeGeometryChanged = True
        Exit Function
    End If

    GanttDependencySvg_ShapeGeometryChanged = _
        (CStr(gLastTaskGeometryByShape(shapeName)) <> CStr(currentByShape(shapeName)))

End Function

Public Function GanttDependencySvg_IsTaskGeometryCurrent(ByVal ws As Worksheet) As Boolean

    Dim currentSignature As String

    If Len(gLastTaskGeometrySignature) = 0 Then Exit Function
    currentSignature = GanttDependencySvg_CurrentTaskGeometrySignature(ws)
    If Len(currentSignature) = 0 Then Exit Function
    GanttDependencySvg_IsTaskGeometryCurrent = (currentSignature = gLastTaskGeometrySignature)

End Function

Public Function GanttDependencySvg_IsLayerPhysicalStateCurrent(ByVal ws As Worksheet) As Boolean

    Dim layer As Shape
    Dim expectedLeft As Double
    Dim expectedTop As Double
    Dim expectedWidth As Double
    Dim expectedHeight As Double

    EnsureModeInitialized
    If ws Is Nothing Then Exit Function
    If IsAggregatedScaleMode() Then
        GanttDependencySvg_IsLayerPhysicalStateCurrent = True
        Exit Function
    End If
    If Not GanttDependencySvg_TryGetExpectedLayerBounds( _
        expectedLeft, expectedTop, expectedWidth, expectedHeight) Then Exit Function

    Set layer = GetLayerShape(ws)
    If layer Is Nothing Then Exit Function
    If layer.Visible <> msoTrue Then Exit Function
    If Abs(CDbl(layer.Left) - expectedLeft) > SVG_PHYSICAL_TOLERANCE Then Exit Function
    If Abs(CDbl(layer.Top) - expectedTop) > SVG_PHYSICAL_TOLERANCE Then Exit Function
    If Abs(CDbl(layer.Width) - expectedWidth) > SVG_PHYSICAL_TOLERANCE Then Exit Function
    If Abs(CDbl(layer.Height) - expectedHeight) > SVG_PHYSICAL_TOLERANCE Then Exit Function

    GanttDependencySvg_IsLayerPhysicalStateCurrent = True

End Function

Public Function GanttDependencySvg_IsPhysicalStateCurrent(ByVal ws As Worksheet) As Boolean

    If Not GanttDependencySvg_IsTaskGeometryCurrent(ws) Then Exit Function
    GanttDependencySvg_IsPhysicalStateCurrent = _
        GanttDependencySvg_IsLayerPhysicalStateCurrent(ws)

End Function

Public Function GanttDependencySvg_GetPhysicallyDirtyTaskIds( _
    ByVal ws As Worksheet, _
    ByVal rowById As Object) As Object

    Dim dirtyIds As Object
    Dim currentByShape As Object
    Dim idKey As Variant
    Dim taskId As String
    Dim ganttRow As Long
    Dim dataRow As Long
    Dim shapePrefix As Variant
    Dim shapeName As String

    Set dirtyIds = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set GanttDependencySvg_GetPhysicallyDirtyTaskIds = dirtyIds
        Exit Function
    End If
    If rowById Is Nothing Then
        Set GanttDependencySvg_GetPhysicallyDirtyTaskIds = dirtyIds
        Exit Function
    End If
    If gLastTaskGeometryByShape Is Nothing Then
        Set GanttDependencySvg_GetPhysicallyDirtyTaskIds = Nothing
        Exit Function
    End If

    Set currentByShape = GanttDependencySvg_BuildTaskGeometryMap(ws)
    If currentByShape Is Nothing Then
        Set GanttDependencySvg_GetPhysicallyDirtyTaskIds = Nothing
        Exit Function
    End If

    For Each idKey In rowById.Keys
        taskId = CStr(idKey)
        ganttRow = CLng(rowById(taskId))
        dataRow = ganttRow - 5 + 1
        If dataRow > 0 Then
            For Each shapePrefix In Array("TASK_", "MS_", "SUM_")
                shapeName = CStr(shapePrefix) & CStr(dataRow)
                If GanttDependencySvg_ShapeGeometryChanged(shapeName, currentByShape) Then
                    dirtyIds(taskId) = True
                    Exit For
                End If
                shapeName = CStr(shapePrefix) & taskId
                If GanttDependencySvg_ShapeGeometryChanged(shapeName, currentByShape) Then
                    dirtyIds(taskId) = True
                    Exit For
                End If
            Next shapePrefix
        End If
    Next idKey

    Profiler_RecordOperation "GanttPhysicalDirtyIds", dirtyIds.Count, 0#
    Set GanttDependencySvg_GetPhysicallyDirtyTaskIds = dirtyIds

End Function

Public Function GanttDependencySvg_AllRoutesHaveTerminalArrow() As Boolean

    Dim routeId As Variant
    Dim route As clsGanttDependencyRoute
    Dim segment As Object
    Dim hasArrow As Boolean

    If gRoutesById Is Nothing Then Exit Function
    If gRoutesById.Count = 0 Then Exit Function

    For Each routeId In gRoutesById.Keys
        Set route = gRoutesById(CStr(routeId))
        hasArrow = False
        For Each segment In route.Segments
            If SvgParseBoolean(segment("Arrow")) Then
                hasArrow = True
                Exit For
            End If
        Next segment
        If Not hasArrow Then Exit Function
    Next routeId

    GanttDependencySvg_AllRoutesHaveTerminalArrow = True

End Function

Public Function GanttDependencySvg_GetDiagnostics() As Variant

    Dim result(1 To 1, 1 To 10) As Variant
    Dim dirtyCount As Long

    EnsureModeInitialized
    If Not gDirtyLinkIds Is Nothing Then dirtyCount = gDirtyLinkIds.Count
    result(1, 1) = gRequestedMode
    result(1, 2) = gActiveMode
    result(1, 3) = GanttDependencySvg_RouteCount()
    result(1, 4) = GanttDependencySvg_SegmentCount()
    result(1, 5) = dirtyCount
    result(1, 6) = gFullModelDirty
    result(1, 7) = gGenerationCount
    result(1, 8) = gInsertionCount
    result(1, 9) = gFallbackCount
    result(1, 10) = gLastFallbackReason
    GanttDependencySvg_GetDiagnostics = result

End Function

Public Function GanttDependencySvg_GetPersistentCacheDiagnostics(ByVal ws As Worksheet) As Variant

    Dim result(1 To 1, 1 To 12) As Variant
    Dim cacheWs As Worksheet
    Dim layer As Shape
    Dim layerSignature As String

    EnsureModeInitialized
    Set cacheWs = GetSvgCacheSheet(False)
    Set layer = GetLayerShape(ws)
    If Not layer Is Nothing Then layerSignature = GetLayerCacheSignature(layer)

    result(1, 1) = Not cacheWs Is Nothing
    result(1, 2) = Not layer Is Nothing
    result(1, 3) = Len(layerSignature) > 0
    result(1, 4) = GanttDependencySvg_HasRoutes()
    result(1, 5) = GanttDependencySvg_RouteCount()
    result(1, 6) = GanttDependencySvg_SegmentCount()
    result(1, 7) = gFullModelDirty
    result(1, 8) = gLastContentHash
    result(1, 9) = gLastFallbackReason
    result(1, 10) = GanttDependencySvg_IsTaskGeometryCurrent(ws)
    result(1, 11) = GanttDependencySvg_AllRoutesHaveTerminalArrow()
    result(1, 12) = GanttDependencySvg_IsLayerPhysicalStateCurrent(ws)
    GanttDependencySvg_GetPersistentCacheDiagnostics = result

End Function

Public Function GanttDependencySvg_ExportRoutes(ByVal outputPath As String) As Boolean

    Dim fso As Object
    Dim stream As Object
    Dim routeKeys As Variant
    Dim routeKey As Variant
    Dim route As clsGanttDependencyRoute
    Dim segment As Object
    Dim i As Long

    On Error GoTo Failed
    If gRoutesById Is Nothing Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(outputPath, True, False)
    stream.WriteLine "LinkID" & vbTab & "PredID" & vbTab & "SuccID" & vbTab & _
        "LinkType" & vbTab & "Lag" & vbTab & "SegmentIndex" & vbTab & _
        "X1" & vbTab & "Y1" & vbTab & "X2" & vbTab & "Y2" & vbTab & "Arrow"

    routeKeys = SortedDictionaryKeys(gRoutesById)
    For Each routeKey In routeKeys
        Set route = gRoutesById(CStr(routeKey))
        For i = 1 To route.Segments.Count
            Set segment = route.Segments(i)
            stream.WriteLine _
                route.LinkId & vbTab & route.PredId & vbTab & route.SuccId & vbTab & _
                route.LinkType & vbTab & SvgNumber(route.Lag) & vbTab & CStr(i) & vbTab & _
                SvgNumber(CDbl(segment("X1"))) & vbTab & SvgNumber(CDbl(segment("Y1"))) & vbTab & _
                SvgNumber(CDbl(segment("X2"))) & vbTab & SvgNumber(CDbl(segment("Y2"))) & vbTab & _
                SvgBoolToken(CBool(segment("Arrow")))
        Next i
    Next routeKey

    stream.Close
    GanttDependencySvg_ExportRoutes = True
    Exit Function

Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0

End Function

Public Function GanttDependencySvg_ExportCurrentSvg(ByVal outputPath As String) As Boolean

    Dim svgText As String
    Dim minX As Double
    Dim minY As Double
    Dim maxX As Double
    Dim maxY As Double
    Dim fso As Object
    Dim stream As Object

    On Error GoTo Failed
    If gRoutesById Is Nothing Then Exit Function
    svgText = BuildSvgDocument(minX, minY, maxX, maxY)
    If Len(svgText) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(outputPath, True, False)
    stream.Write svgText
    stream.Close
    GanttDependencySvg_ExportCurrentSvg = True
    Exit Function

Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0

End Function

Private Function BuildSvgDocument( _
    ByRef minX As Double, _
    ByRef minY As Double, _
    ByRef maxX As Double, _
    ByRef maxY As Double) As String

    Dim perfScope As clsPerfScope
    Dim routeKeys As Variant
    Dim routeKey As Variant
    Dim route As clsGanttDependencyRoute
    Dim segment As Object
    Dim body As String
    Dim arrowBody As String
    Dim routePath As String
    Dim i As Long
    Dim firstBounds As Boolean
    Dim routesSerialized As Long
    Dim segmentsSerialized As Long
    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double

    Set perfScope = Profiler_BeginScope("GanttDependencySvg_BuildDocument", "Dependency SVG")
    firstBounds = True
    routeKeys = SortedDictionaryKeys(gRoutesById)

    For Each routeKey In routeKeys
        Set route = gRoutesById(CStr(routeKey))
        If route.Visible And route.SegmentCount > 0 Then
            If firstBounds Then
                minX = route.MinX
                minY = route.MinY
                maxX = route.MaxX
                maxY = route.MaxY
                firstBounds = False
            Else
                If route.MinX < minX Then minX = route.MinX
                If route.MinY < minY Then minY = route.MinY
                If route.MaxX > maxX Then maxX = route.MaxX
                If route.MaxY > maxY Then maxY = route.MaxY
            End If
        End If
    Next routeKey

    If firstBounds Then Exit Function
    minX = minX - SVG_MARGIN
    minY = minY - SVG_MARGIN
    maxX = maxX + SVG_MARGIN
    maxY = maxY + SVG_MARGIN

    For Each routeKey In routeKeys
        Set route = gRoutesById(CStr(routeKey))
        If route.Visible And route.SegmentCount > 0 Then
            routesSerialized = routesSerialized + 1
            routePath = ""
            For i = 1 To route.Segments.Count
                Set segment = route.Segments(i)
                segmentsSerialized = segmentsSerialized + 1
                x1 = CDbl(segment("X1")) - minX
                y1 = CDbl(segment("Y1")) - minY
                x2 = CDbl(segment("X2")) - minX
                y2 = CDbl(segment("Y2")) - minY

                routePath = routePath & "M " & SvgNumber(x1) & " " & SvgNumber(y1) & _
                    " L " & SvgNumber(x2) & " " & SvgNumber(y2) & " "
                If CBool(segment("Arrow")) Then
                    arrowBody = arrowBody & BuildArrowPolygon(x1, y1, x2, y2)
                End If
            Next i

            body = body & "<path id=""" & XmlEscape(CStr(routeKey)) & """ d=""" & _
                Trim$(routePath) & """/>" & vbCrLf
        End If
    Next routeKey

    BuildSvgDocument = _
        "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
        "<svg xmlns=""http://www.w3.org/2000/svg"" viewBox=""0 0 " & _
        SvgNumber(maxX - minX) & " " & SvgNumber(maxY - minY) & _
        """ width=""" & SvgNumber(maxX - minX) & "pt"" height=""" & _
        SvgNumber(maxY - minY) & "pt"">" & vbCrLf & _
        "<g fill=""none"" stroke=""" & SVG_LINE_COLOR & """ stroke-width=""" & _
        SvgNumber(SVG_LINE_WEIGHT) & """ stroke-linecap=""butt"" stroke-linejoin=""miter"">" & vbCrLf & _
        body & "</g>" & vbCrLf & _
        "<g fill=""" & SVG_LINE_COLOR & """ stroke=""none"">" & vbCrLf & _
        arrowBody & "</g>" & vbCrLf & "</svg>"

    Profiler_RecordOperation "DependencySvgRoutesSerialized", routesSerialized, 0#
    Profiler_RecordOperation "DependencySvgSegmentsSerialized", segmentsSerialized, 0#

End Function

Private Function BuildArrowPolygon( _
    ByVal x1 As Double, _
    ByVal y1 As Double, _
    ByVal x2 As Double, _
    ByVal y2 As Double) As String

    Dim dx As Double
    Dim dy As Double
    Dim lengthValue As Double
    Dim ux As Double
    Dim uy As Double
    Dim baseX As Double
    Dim baseY As Double
    Dim perpX As Double
    Dim perpY As Double

    dx = x2 - x1
    dy = y2 - y1
    lengthValue = Sqr((dx * dx) + (dy * dy))
    If lengthValue < 0.1 Then Exit Function

    ux = dx / lengthValue
    uy = dy / lengthValue
    baseX = x2 - (ux * SVG_ARROW_LENGTH)
    baseY = y2 - (uy * SVG_ARROW_LENGTH)
    perpX = -uy * SVG_ARROW_HALF_WIDTH
    perpY = ux * SVG_ARROW_HALF_WIDTH

    BuildArrowPolygon = "<polygon points=""" & _
        SvgNumber(x2) & "," & SvgNumber(y2) & " " & _
        SvgNumber(baseX + perpX) & "," & SvgNumber(baseY + perpY) & " " & _
        SvgNumber(baseX - perpX) & "," & SvgNumber(baseY - perpY) & """/>" & vbCrLf

End Function

Private Function WriteSvgTempFile(ByVal svgText As String) As String

    Dim perfScope As clsPerfScope
    Dim fso As Object
    Dim folderPath As String
    Dim filePath As String
    Dim stream As Object

    Set perfScope = Profiler_BeginScope("GanttDependencySvg_WriteFile", "Dependency SVG")
    folderPath = Environ$("TEMP") & "\ProjectEngineSvg"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

    filePath = folderPath & "\dependencies_" & Format$(Now, "yyyymmdd_hhnnss") & _
        "_" & CStr(gGenerationCount + 1) & ".svg"
    Set stream = fso.CreateTextFile(filePath, True, False)
    stream.Write svgText
    stream.Close
    WriteSvgTempFile = filePath

End Function

Private Sub DeleteTempFile(ByVal filePath As String)

    Dim fso As Object

    If Len(filePath) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then fso.DeleteFile filePath, True

End Sub

Private Sub DeleteHistoricalSegments(ByVal ws As Worksheet)

    Dim i As Long
    Dim deletedCount As Long
    Dim shapeName As String

    For i = ws.Shapes.Count To 1 Step -1
        shapeName = CStr(ws.Shapes(i).Name)
        If Left$(shapeName, 4) = "DEP_" Then
            ws.Shapes(i).Delete
            deletedCount = deletedCount + 1
        End If
    Next i

    Profiler_RecordOperation "GanttDependencySvgHistoricalSegmentsDeleted", deletedCount, 0#

End Sub

Private Sub SavePersistentRouteCache( _
    ByVal ws As Worksheet, _
    ByVal contentHash As String, _
    ByVal cacheSignature As String)

    Dim cacheWs As Worksheet
    Dim routeKeys As Variant
    Dim routeKey As Variant
    Dim route As clsGanttDependencyRoute
    Dim segment As Object
    Dim arr As Variant
    Dim rowIndex As Long
    Dim i As Long
    Dim segmentCount As Long

    On Error GoTo Failed
    If gRoutesById Is Nothing Then Exit Sub
    If gRoutesById.Count = 0 Then Exit Sub
    If Len(cacheSignature) = 0 Then Exit Sub

    segmentCount = GanttDependencySvg_SegmentCount()
    If segmentCount = 0 Then Exit Sub

    Set cacheWs = GetSvgCacheSheet(True)
    If cacheWs Is Nothing Then Exit Sub

    ReDim arr(1 To segmentCount + 2, 1 To 13)
    arr(1, 1) = "RecordType"
    arr(1, 2) = "CacheVersion"
    arr(1, 3) = "SignatureOrLinkID"
    arr(1, 4) = "ContentHashOrPredID"
    arr(1, 5) = "GeneratedAtOrSuccID"
    arr(1, 6) = "ScaleOrLinkType"
    arr(1, 7) = "RouteCountOrLag"
    arr(1, 8) = "SegmentIndex"
    arr(1, 9) = "X1"
    arr(1, 10) = "Y1"
    arr(1, 11) = "X2"
    arr(1, 12) = "Y2"
    arr(1, 13) = "Arrow"

    arr(2, 1) = "META"
    arr(2, 2) = SVG_CACHE_VERSION
    arr(2, 3) = cacheSignature
    arr(2, 4) = contentHash
    arr(2, 5) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    arr(2, 6) = GetGanttTimelineScaleMode()
    arr(2, 7) = gRoutesById.Count
    arr(2, 8) = segmentCount
    arr(2, 9) = gLastTaskGeometrySignature

    rowIndex = 3
    routeKeys = SortedDictionaryKeys(gRoutesById)
    For Each routeKey In routeKeys
        Set route = gRoutesById(CStr(routeKey))
        For i = 1 To route.Segments.Count
            Set segment = route.Segments(i)
            arr(rowIndex, 1) = "ROUTE"
            arr(rowIndex, 2) = SVG_CACHE_VERSION
            arr(rowIndex, 3) = route.LinkId
            arr(rowIndex, 4) = route.PredId
            arr(rowIndex, 5) = route.SuccId
            arr(rowIndex, 6) = route.LinkType
            arr(rowIndex, 7) = SvgNumber(route.Lag)
            arr(rowIndex, 8) = i
            arr(rowIndex, 9) = SvgNumber(CDbl(segment("X1")))
            arr(rowIndex, 10) = SvgNumber(CDbl(segment("Y1")))
            arr(rowIndex, 11) = SvgNumber(CDbl(segment("X2")))
            arr(rowIndex, 12) = SvgNumber(CDbl(segment("Y2")))
            arr(rowIndex, 13) = SvgBoolToken(CBool(segment("Arrow")))
            rowIndex = rowIndex + 1
        Next i
    Next routeKey

    cacheWs.Cells.Clear
    cacheWs.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
    cacheWs.Visible = xlSheetVeryHidden

    Profiler_RecordOperation "GanttDependencySvgCacheSaved", gRoutesById.Count, 0#
    Profiler_RecordOperation "GanttDependencySvgCacheSavedSegments", segmentCount, 0#
    Profiler_RecordOperation "DependencySvgRoutesPersisted", gRoutesById.Count, 0#
    Profiler_RecordOperation "DependencySvgSegmentsPersisted", segmentCount, 0#
    Exit Sub

Failed:
    Profiler_RecordOperation "GanttDependencySvgCacheSaveFailed", 1, 0#

End Sub

Private Function BuildRouteSnapshotSignature(ByVal contentHash As String) As String

    Dim routeKeys As Variant
    Dim routeKey As Variant
    Dim route As clsGanttDependencyRoute
    Dim segment As Object
    Dim i As Long
    Dim segmentCount As Long
    Dim textValue As String

    If gRoutesById Is Nothing Then Exit Function

    textValue = SVG_CACHE_VERSION & "|" & GetGanttTimelineScaleMode() & "|" & _
        CStr(gRoutesById.Count) & "|" & contentHash

    routeKeys = SortedDictionaryKeys(gRoutesById)
    For Each routeKey In routeKeys
        Set route = gRoutesById(CStr(routeKey))
        textValue = textValue & "|" & route.LinkId & "," & route.PredId & "," & _
            route.SuccId & "," & route.LinkType & "," & SvgNumber(route.Lag)
        For i = 1 To route.Segments.Count
            Set segment = route.Segments(i)
            textValue = textValue & "," & CStr(i) & ":" & _
                SvgNumber(CDbl(segment("X1"))) & "," & _
                SvgNumber(CDbl(segment("Y1"))) & "," & _
                SvgNumber(CDbl(segment("X2"))) & "," & _
                SvgNumber(CDbl(segment("Y2"))) & "," & _
                SvgBoolToken(CBool(segment("Arrow")))
            segmentCount = segmentCount + 1
        Next i
    Next routeKey

    BuildRouteSnapshotSignature = SVG_CACHE_VERSION & "|" & CStr(segmentCount) & "|" & StableTextHash(textValue)

End Function

Private Sub ApplyLayerCacheSignature(ByVal shp As Shape, ByVal cacheSignature As String)

    If shp Is Nothing Then Exit Sub
    If Len(cacheSignature) = 0 Then Exit Sub

    On Error Resume Next
    shp.AlternativeText = SVG_CACHE_ALT_PREFIX & cacheSignature
    shp.Tags.Delete SVG_CACHE_TAG_SIGNATURE
    shp.Tags.Add SVG_CACHE_TAG_SIGNATURE, cacheSignature
    On Error GoTo 0

End Sub

Private Sub ApplyLayerPhysicalState( _
    ByVal shp As Shape, _
    ByVal layerLeft As Double, _
    ByVal layerTop As Double, _
    ByVal layerWidth As Double, _
    ByVal layerHeight As Double)

    If shp Is Nothing Then Exit Sub
    If layerWidth < 1# Then layerWidth = 1#
    If layerHeight < 1# Then layerHeight = 1#

    shp.LockAspectRatio = msoFalse
    shp.Placement = xlFreeFloating
    If Abs(CDbl(shp.Left) - layerLeft) > SVG_PHYSICAL_TOLERANCE Then shp.Left = layerLeft
    If Abs(CDbl(shp.Top) - layerTop) > SVG_PHYSICAL_TOLERANCE Then shp.Top = layerTop
    If Abs(CDbl(shp.Width) - layerWidth) > SVG_PHYSICAL_TOLERANCE Then shp.Width = layerWidth
    If Abs(CDbl(shp.Height) - layerHeight) > SVG_PHYSICAL_TOLERANCE Then shp.Height = layerHeight
    If shp.Visible <> msoTrue Then shp.Visible = msoTrue

End Sub

Private Function GanttDependencySvg_TryGetExpectedLayerBounds( _
    ByRef layerLeft As Double, _
    ByRef layerTop As Double, _
    ByRef layerWidth As Double, _
    ByRef layerHeight As Double) As Boolean

    Dim routeKey As Variant
    Dim route As clsGanttDependencyRoute
    Dim firstBounds As Boolean
    Dim maxX As Double
    Dim maxY As Double

    If gRoutesById Is Nothing Then Exit Function
    If gRoutesById.Count = 0 Then Exit Function

    firstBounds = True
    For Each routeKey In gRoutesById.Keys
        Set route = gRoutesById(CStr(routeKey))
        If route.Visible And route.SegmentCount > 0 Then
            If firstBounds Then
                layerLeft = route.MinX
                layerTop = route.MinY
                maxX = route.MaxX
                maxY = route.MaxY
                firstBounds = False
            Else
                If route.MinX < layerLeft Then layerLeft = route.MinX
                If route.MinY < layerTop Then layerTop = route.MinY
                If route.MaxX > maxX Then maxX = route.MaxX
                If route.MaxY > maxY Then maxY = route.MaxY
            End If
        End If
    Next routeKey
    If firstBounds Then Exit Function

    layerLeft = layerLeft - SVG_MARGIN
    layerTop = layerTop - SVG_MARGIN
    maxX = maxX + SVG_MARGIN
    maxY = maxY + SVG_MARGIN
    layerWidth = maxX - layerLeft
    layerHeight = maxY - layerTop
    If layerWidth < 1# Then layerWidth = 1#
    If layerHeight < 1# Then layerHeight = 1#
    GanttDependencySvg_TryGetExpectedLayerBounds = True

End Function

Private Function GetLayerCacheSignature(ByVal shp As Shape) As String

    Dim rawValue As String

    If shp Is Nothing Then Exit Function

    On Error Resume Next
    rawValue = CStr(shp.Tags(SVG_CACHE_TAG_SIGNATURE))
    On Error GoTo 0
    If Len(rawValue) > 0 Then
        GetLayerCacheSignature = rawValue
        Exit Function
    End If

    rawValue = CStr(shp.AlternativeText)
    If Left$(rawValue, Len(SVG_CACHE_ALT_PREFIX)) = SVG_CACHE_ALT_PREFIX Then
        GetLayerCacheSignature = Mid$(rawValue, Len(SVG_CACHE_ALT_PREFIX) + 1)
    End If

End Function

Private Function GetSvgCacheSheet(ByVal createIfMissing As Boolean) As Worksheet

    On Error Resume Next
    Set GetSvgCacheSheet = ThisWorkbook.Worksheets(SVG_CACHE_SHEET_NAME)
    On Error GoTo 0

    If GetSvgCacheSheet Is Nothing And createIfMissing Then
        Set GetSvgCacheSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetSvgCacheSheet.Name = SVG_CACHE_SHEET_NAME
        GetSvgCacheSheet.Visible = xlSheetVeryHidden
    End If

End Function

Private Function GetLayerShape(ByVal ws As Worksheet) As Shape

    On Error Resume Next
    Set GetLayerShape = ws.Shapes(SVG_LAYER_NAME)
    On Error GoTo 0

End Function

Private Function StableTextHash(ByVal textValue As String) As String

    Dim hashValue As Double
    Dim i As Long

    hashValue = 5381#
    For i = 1 To Len(textValue)
        hashValue = (hashValue * 33#) + AscW(Mid$(textValue, i, 1))
        hashValue = hashValue - (Fix(hashValue / 2147483647#) * 2147483647#)
    Next i
    StableTextHash = CStr(CLng(hashValue))

End Function

Private Function SvgNumber(ByVal value As Double) As String

    SvgNumber = Replace$(Format$(value, "0.###"), ",", ".")
    If Right$(SvgNumber, 1) = "." Then SvgNumber = Left$(SvgNumber, Len(SvgNumber) - 1)
    If SvgNumber = "-0" Then SvgNumber = "0"

End Function

Private Function SvgBoolToken(ByVal value As Boolean) As String

    If value Then
        SvgBoolToken = "TRUE"
    Else
        SvgBoolToken = "FALSE"
    End If

End Function

Private Function SvgParseBoolean(ByVal value As Variant) As Boolean

    Dim rawValue As String

    rawValue = UCase$(Trim$(CStr(value)))
    Select Case rawValue
        Case "TRUE", "VRAI", "YES", "OUI", "1", "-1", "Y", "T"
            SvgParseBoolean = True
    End Select

End Function

Private Function IsTaskGeometrySignatureShape(ByVal shapeName As String) As Boolean

    IsTaskGeometrySignatureShape = _
        (Left$(shapeName, 5) = "TASK_") Or _
        (Left$(shapeName, 3) = "MS_") Or _
        (Left$(shapeName, 4) = "SUM_")

End Function

Private Function XmlEscape(ByVal value As String) As String

    value = Replace$(value, "&", "&amp;")
    value = Replace$(value, """", "&quot;")
    value = Replace$(value, "<", "&lt;")
    value = Replace$(value, ">", "&gt;")
    XmlEscape = value

End Function

Private Function SortedDictionaryKeys(ByVal source As Object) As Variant

    Dim values As Variant

    values = source.Keys
    If source.Count > 1 Then QuickSortStrings values, LBound(values), UBound(values)
    SortedDictionaryKeys = values

End Function

Private Sub QuickSortStrings(ByRef values As Variant, ByVal first As Long, ByVal last As Long)

    Dim low As Long
    Dim high As Long
    Dim pivot As String
    Dim temp As Variant

    low = first
    high = last
    pivot = CStr(values((first + last) \ 2))

    Do While low <= high
        Do While StrComp(CStr(values(low)), pivot, vbBinaryCompare) < 0
            low = low + 1
        Loop
        Do While StrComp(CStr(values(high)), pivot, vbBinaryCompare) > 0
            high = high - 1
        Loop
        If low <= high Then
            temp = values(low)
            values(low) = values(high)
            values(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortStrings values, first, high
    If low < last Then QuickSortStrings values, low, last

End Sub
