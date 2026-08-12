Attribute VB_Name = "mod_GanttRenderer"
Option Explicit

'===============================================================================
' MODULE : mod_GanttRenderer
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
' CONTRATS / CONTRACTS : ShouldRenderTaskInCurrentView, DrawGanttShapes, ShouldDrawCompactTaskMarker, DrawTodayLine, ShouldHighlightGanttAnalyticsPath, GetTaskBaseColor, GetProgressFillColor, GanttHierarchy_BuildLeafCompletionByAncestor
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


'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function IsValidDisplayRange(ByVal startVal As Variant, ByVal finishVal As Variant) As Boolean

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    IsValidDisplayRange = (CDbl(finishVal) >= CDbl(startVal))

End Function

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Public Function ShouldRenderTaskInCurrentView( _
    ByVal isParent As Boolean, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant) As Boolean

    ShouldRenderTaskInCurrentView = IsValidDisplayRange(startVal, finishVal)

End Function



'------------------------------------------------------------------------------
' FR:
' Parcourt les lignes WBS et dessine le visuel principal: summaries, jalons,
' barres longues, marqueurs compacts, progress et delta TEST.
'
' EN:
' Walks WBS rows and draws the main visual layer: summaries, milestones,
' long bars, compact markers, progress, and TEST deltas.
'
' Entrees / Inputs:
' - dataArr WBS, maps, hasChildren, timeline, maps base/test GanttLive et mode TEST.
'
' Sorties / Outputs:
' - Shapes TASK/MS/SUM et progress crees via le registre de shapes.
'
' Appele par / Called by:
' - RunGanttRefreshCore apres preparation du layout.
'
' Notes:
' - Ne calcule pas les dates; consomme uniquement les dates d'affichage GanttLive.
'------------------------------------------------------------------------------
Public Sub DrawGanttShapes( _
    ByVal ws As Worksheet, _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean)

    Dim perfScope As clsPerfScope

    Dim r As Long
    Dim rowCount As Long
    Dim renderableRowCount As Long
    Dim ganttRow As Long

    Dim wbs As String
    Dim idVal As String

    Dim rawStartVal As Variant
    Dim rawFinishVal As Variant
    Dim renderStartVal As Variant
    Dim renderFinishVal As Variant

    Dim rawDurationDays As Long
    Dim progressVal As Double
    Dim isCritical As Boolean
    Dim hasDelta As Boolean
    Dim isParent As Boolean
    Dim isLoE As Boolean
    Dim isMilestone As Boolean
    Dim parentCompleteMap As Object

    Set perfScope = Profiler_BeginScope("DrawGanttShapes", "Shape Render")

    rowCount = UBound(dataArr, 1)
    Set parentCompleteMap = GanttHierarchy_BuildLeafCompletionByAncestor(dataArr, mapWBS)

    For r = 1 To rowCount
        ganttRow = FIRST_TASK_ROW + r - 1

        wbs = NormalizeWBS(CStr(dataArr(r, mapWBS("WBS"))))
        idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))

        rawStartVal = GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode)
        rawFinishVal = GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode)

        renderStartVal = GetRenderStartForCurrentScale(rawStartVal)
        renderFinishVal = GetRenderFinishForCurrentScale(rawFinishVal)

        isParent = hasChildren.Exists(wbs)
        isMilestone = TaskTypeRules_IsMilestoneRow(dataArr, mapWBS, r)
        isLoE = TaskTypeRules_IsLevelOfEffortRow(dataArr, mapWBS, r)

        If Not ShouldRenderTaskInCurrentView(isParent, renderStartVal, renderFinishVal) Then GoTo NextShape
        If Not HasValue(rawStartVal) Or Not HasValue(rawFinishVal) Then GoTo NextShape
        If Not HasValue(renderStartVal) Or Not HasValue(renderFinishVal) Then GoTo NextShape

        rawDurationDays = CLng(CDbl(rawFinishVal) - CDbl(rawStartVal) + 1)

        progressVal = 0

        If isLoE Then

            'LOE Primavera-style progress:
            'The visual progress must be derived from TODAY and LOE calculated dates,
            'not from the manual % Progress stored in WBS/CALC.
            progressVal = GetLoEProgressFromToday(rawStartVal, rawFinishVal)

        Else

            If HasValue(GanttLive_GetDisplayProgress(idVal, baseById, testById, isTestMode)) Then
                progressVal = CDbl(GanttLive_GetDisplayProgress(idVal, baseById, testById, isTestMode))
            End If

        End If

        isCritical = ShouldHighlightGanttAnalyticsPath(dataArr, mapWBS, r, idVal, testById, isTestMode)

        hasDelta = False
        If isTestMode Then
            hasDelta = GanttLive_HasRenderableTestDelta(idVal, baseById, testById)
        End If

        If isParent Then

            DrawSummaryBar ws, ganttRow, projectStart, rawStartVal, rawFinishVal, isCritical, totalDays, _
                "SUM_" & CStr(r), Trim$(CStr(dataArr(r, mapWBS("Task Name")))), hasDelta, _
                ParentIsCompleteFromMap(wbs, parentCompleteMap)

        ElseIf isMilestone Then

            DrawMilestone ws, ganttRow, projectStart, rawStartVal, progressVal, isCritical, totalDays, "MS_" & CStr(r), hasDelta

        ElseIf ShouldDrawCompactTaskMarker(ws, ganttRow, projectStart, rawStartVal, rawFinishVal, rawDurationDays, isLoE) Then

            DrawSingleWeekTask ws, ganttRow, projectStart, rawStartVal, progressVal, isCritical, totalDays, _
                "TASK_" & CStr(r), hasDelta, rawFinishVal

        Else

            DrawTaskBar ws, ganttRow, projectStart, rawStartVal, rawFinishVal, progressVal, isCritical, totalDays, _
                "TASK_" & CStr(r), hasDelta, rawStartVal, rawFinishVal, isLoE

        End If

NextShape:
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Public Function ShouldDrawCompactTaskMarker( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    ByVal rawDurationDays As Long, _
    ByVal isLoE As Boolean) As Boolean

    Dim perfScope As clsPerfScope

    Dim visualWidth As Double
    Dim minimumReadableWidth As Double

    Set perfScope = Profiler_BeginScope("ShouldDrawCompactTaskMarker", "Shape Decision")

    If isLoE Then Exit Function
    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    If rawDurationDays <= 1 Then
        ShouldDrawCompactTaskMarker = True
        Exit Function
    End If

    visualWidth = TimelineWidth(ws, projectStart, startVal, finishVal)
    minimumReadableWidth = GetGanttBarHeight(ws, ganttRow)

    ShouldDrawCompactTaskMarker = (visualWidth > 0 And visualWidth < minimumReadableWidth)

End Function

'------------------------------------------------------------------------------
' FR:
' Cree les records puis les shapes d'une barre de tache standard ou LOE avec progress.
'
' EN:
' Creates records then shapes for a standard or LOE task bar with progress.
'
' Entrees / Inputs:
' - Ligne GANTT, dates de rendu, progress, criticite, delta TEST, type LOE.
'
' Sorties / Outputs:
' - Shape TASK_* et eventuel TASK_*_P dans la feuille GANTT.
'
' Appele par / Called by:
' - DrawGanttShapes.
'
' Notes:
' - Delegue la geometrie fine a GanttShapeRegistry_AddTaskBarRecords.
'------------------------------------------------------------------------------
Private Sub DrawTaskBar( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    ByVal progressVal As Double, _
    ByVal isCritical As Boolean, _
    ByVal totalDays As Long, _
    ByVal shapeKey As String, _
    ByVal hasDelta As Boolean, _
    ByVal rawStartVal As Variant, _
    ByVal rawFinishVal As Variant, _
    Optional ByVal isLoE As Boolean = False)

    Dim perfScope As clsPerfScope
    Dim expected As Object

    Set perfScope = Profiler_BeginScope("DrawTaskBar", "Shape Create")
    Set expected = CreateObject("Scripting.Dictionary")

    GanttShapeRegistry_AddTaskBarRecords expected, ws, ganttRow, projectStart, startVal, finishVal, progressVal, isCritical, shapeKey, hasDelta, isLoE
    GanttShapeRegistry_CreateAllFromRecords ws, expected

End Sub

'------------------------------------------------------------------------------
' FR:
' Cree le record puis la shape diamant d'un jalon positionne sur sa date d'affichage.
'
' EN:
' Creates the record then diamond shape for a milestone positioned on its display date.
'
' Entrees / Inputs:
' - Ligne GANTT, date jalon, progress, criticite et delta TEST.
'
' Sorties / Outputs:
' - Shape MS_* dans la feuille GANTT.
'
' Appele par / Called by:
' - DrawGanttShapes.
'
' Notes:
' - Utilise le registre pour garder le rendu compatible avec le diff predictif.
'------------------------------------------------------------------------------
Private Sub DrawMilestone( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal progressVal As Double, _
    ByVal isCritical As Boolean, _
    ByVal totalDays As Long, _
    ByVal shapeKey As String, _
    ByVal hasDelta As Boolean)

    Dim perfScope As clsPerfScope
    Dim expected As Object

    Set perfScope = Profiler_BeginScope("DrawMilestone", "Shape Create")
    Set expected = CreateObject("Scripting.Dictionary")

    GanttShapeRegistry_AddMilestoneRecord expected, ws, ganttRow, projectStart, startVal, progressVal, isCritical, shapeKey, hasDelta
    GanttShapeRegistry_CreateAllFromRecords ws, expected

End Sub


'------------------------------------------------------------------------------
' FR: Retourne une valeur d'etat ou une reference utilisee par le rendu GANTT.
' EN: Returns a state value or reference used by GANTT rendering.
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' FR: Dessine une famille visuelle du GANTT a partir des dates et positions deja calculees.
' EN: Draws one GANTT visual family from already computed dates and positions.
'------------------------------------------------------------------------------
Public Sub DrawTodayLine(ByVal ws As Worksheet, ByVal projectStart As Variant, ByVal totalDays As Long, ByVal rowCount As Long)

    Dim perfScope As clsPerfScope
    Dim expected As Object

    Set perfScope = Profiler_BeginScope("DrawTodayLine", "Shape Create")
    Set expected = CreateObject("Scripting.Dictionary")

    GanttShapeRegistry_AddTodayLineRecord expected, ws, projectStart, totalDays, rowCount
    GanttShapeRegistry_CreateAllFromRecords ws, expected

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Gantt Analytics Path Column Name sans exposer de mutateur sur l'etat source.
' EN: Returns the Gantt Analytics Path Column Name value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function GetGanttAnalyticsPathColumnName() As String

    EnsureGanttViewInitialized

    Select Case GetGanttAnalyticsPathMode()
        Case GANTT_ANALYTICS_PATH_CP
            GetGanttAnalyticsPathColumnName = "Critical Path"
        Case GANTT_ANALYTICS_PATH_LP
            GetGanttAnalyticsPathColumnName = "Longest Path"
        Case Else
            GetGanttAnalyticsPathColumnName = vbNullString
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Public Function ShouldHighlightGanttAnalyticsPath( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal dataRow As Long, _
    Optional ByVal idVal As String = "", _
    Optional ByVal testById As Object = Nothing, _
    Optional ByVal isTestMode As Boolean = False) As Boolean

    Dim pathColumnName As String
    Dim pathValue As String

    pathColumnName = GetGanttAnalyticsPathColumnName()
    If Len(pathColumnName) = 0 Then Exit Function

    If isTestMode And GanttLive_IsTestAnalyticsProjectionActive() Then
        pathValue = UCase$(Trim$(GanttLive_GetDisplayAnalyticsPath(idVal, testById, True, pathColumnName)))
    Else
        If mapWBS Is Nothing Then Exit Function
        If Not mapWBS.Exists(pathColumnName) Then Exit Function
        pathValue = UCase$(Trim$(CStr(dataArr(dataRow, mapWBS(pathColumnName)))))
    End If

    Select Case pathColumnName
        Case "Critical Path"
            ShouldHighlightGanttAnalyticsPath = (pathValue = "CRITICAL")
        Case "Longest Path"
            ShouldHighlightGanttAnalyticsPath = (pathValue = "LONGEST")
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function IsGanttAnalyticsPathHighlightEnabled() As Boolean

    IsGanttAnalyticsPathHighlightEnabled = (Len(GetGanttAnalyticsPathColumnName()) > 0)

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Base Color dans le workflow de rendu GANTT.
' EN: Runs the Get Task Base Color helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetTaskBaseColor(ByVal isCritical As Boolean) As Long

    If IsGanttAnalyticsPathHighlightEnabled() And isCritical Then
        GetTaskBaseColor = COLOR_TASK_CRITICAL
    Else
        GetTaskBaseColor = COLOR_TASK_BLUE
    End If

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Summary Line Color dans le workflow de rendu GANTT.
' EN: Runs the Get Summary Line Color helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function GetSummaryLineColor( _
    ByVal isCritical As Boolean, _
    ByVal isComplete As Boolean) As Long

    If isComplete Then
        GetSummaryLineColor = COLOR_PROGRESS_GREEN
    ElseIf IsGanttAnalyticsPathHighlightEnabled() And isCritical Then
        GetSummaryLineColor = COLOR_TASK_CRITICAL
    Else
        GetSummaryLineColor = COLOR_TASK_BLUE
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function IsParentWBS(ByVal wbs As String, ByRef dataArr As Variant, ByVal mapWBS As Object) As Boolean

    Dim i As Long
    Dim otherWBS As String
    Dim prefix As String

    If Trim$(wbs) = "" Then Exit Function

    prefix = wbs & "."

    For i = 1 To UBound(dataArr, 1)
        otherWBS = NormalizeWBS(CStr(dataArr(i, mapWBS("WBS"))))
        If Left$(otherWBS, Len(prefix)) = prefix Then
            IsParentWBS = True
            Exit Function
        End If
    Next i

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Progress Fill Color dans le workflow de rendu GANTT.
' EN: Runs the Get Progress Fill Color helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetProgressFillColor( _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    ByVal progressVal As Double) As Variant

    Dim durationVal As Double
    Dim coveredDays As Long
    Dim progressFinish As Double
    Dim todayVal As Double

    If progressVal <= 0 Then Exit Function

    If progressVal >= 1 Then
        GetProgressFillColor = COLOR_PROGRESS_GREEN
        Exit Function
    End If

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1
    If durationVal <= 0 Then Exit Function

    coveredDays = Int(durationVal * progressVal + 0.999999)
    If coveredDays < 1 Then coveredDays = 1

    progressFinish = CDbl(startVal) + coveredDays - 1
    todayVal = CDbl(Date)

    ' Tolérance : on considère "à jour" si le progress couvre aujourd'hui - 1
    If progressFinish >= todayVal - 1 Then
        GetProgressFillColor = COLOR_PROGRESS_GREEN
    Else
        GetProgressFillColor = COLOR_PROGRESS_ORANGE
    End If

End Function

'------------------------------------------------------------------------------
' FR: Construit une map, un index ou une structure intermediaire consommee par le rendu GANTT.
' EN: Builds a map, index, or intermediate structure consumed by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GanttHierarchy_BuildLeafCompletionByAncestor(ByRef dataArr As Variant, ByVal mapWBS As Object) As Object

    Dim parentMap As Object
    Dim i As Long
    Dim wbsVal As String
    Dim parentWbs As String
    Dim progressVal As Double

    Set parentMap = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(dataArr, 1)

        wbsVal = NormalizeWBS(CStr(dataArr(i, mapWBS("WBS"))))
        If wbsVal = "" Then GoTo NextI

        If IsParentWBS(wbsVal, dataArr, mapWBS) Then GoTo NextI

        If HasValue(dataArr(i, mapWBS("% Progress"))) Then
            progressVal = CDbl(dataArr(i, mapWBS("% Progress")))
        Else
            progressVal = 0
        End If

        parentWbs = GetParentWBS(wbsVal)

        Do While parentWbs <> ""
            If Not parentMap.Exists(parentWbs) Then
                parentMap(parentWbs) = True
            End If

            If progressVal < 1 Then
                parentMap(parentWbs) = False
            End If

            parentWbs = GetParentWBS(parentWbs)
        Loop

NextI:
    Next i

    Set GanttHierarchy_BuildLeafCompletionByAncestor = parentMap

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Parent Is Complete From Map dans le workflow de rendu GANTT.
' EN: Runs the Parent Is Complete From Map helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function ParentIsCompleteFromMap(ByVal parentWbs As String, ByVal parentCompleteMap As Object) As Boolean

    If parentCompleteMap Is Nothing Then Exit Function
    If parentWbs = "" Then Exit Function

    If parentCompleteMap.Exists(parentWbs) Then
        ParentIsCompleteFromMap = CBool(parentCompleteMap(parentWbs))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Dessine une famille visuelle du GANTT a partir des dates et positions deja calculees.
' EN: Draws one GANTT visual family from already computed dates and positions.
'------------------------------------------------------------------------------
Private Sub DrawSingleWeekTask( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal progressVal As Double, _
    ByVal isCritical As Boolean, _
    ByVal totalDays As Long, _
    ByVal shapeKey As String, _
    ByVal hasDelta As Boolean, _
    Optional ByVal finishVal As Variant)

    Dim perfScope As clsPerfScope
    Dim expected As Object
    Dim markerFinishVal As Variant

    Set perfScope = Profiler_BeginScope("DrawSingleWeekTask", "Shape Create")
    Set expected = CreateObject("Scripting.Dictionary")

    markerFinishVal = finishVal
    If Not HasValue(markerFinishVal) Then markerFinishVal = startVal

    GanttShapeRegistry_AddCompactTaskRecords expected, ws, ganttRow, projectStart, startVal, markerFinishVal, progressVal, isCritical, shapeKey, hasDelta
    GanttShapeRegistry_CreateAllFromRecords ws, expected

End Sub


'------------------------------------------------------------------------------
' FR: Execute le helper Get Render Progress Bounds dans le workflow de rendu GANTT.
' EN: Runs the Get Render Progress Bounds helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub GetRenderProgressBounds( _
    ByVal rawStartVal As Variant, _
    ByVal rawFinishVal As Variant, _
    ByVal progressVal As Double, _
    ByRef renderStartOut As Variant, _
    ByRef renderFinishOut As Variant)

    Dim rawDuration As Long
    Dim coveredDays As Long
    Dim rawProgressFinish As Date

    renderStartOut = Empty
    renderFinishOut = Empty

    If progressVal <= 0 Then Exit Sub
    If Not HasValue(rawStartVal) Then Exit Sub
    If Not HasValue(rawFinishVal) Then Exit Sub

    rawDuration = CLng(CDbl(rawFinishVal) - CDbl(rawStartVal) + 1)
    If rawDuration <= 0 Then Exit Sub

    coveredDays = Int(rawDuration * WorksheetFunction.Min(progressVal, 1) + 0.999999)
    If coveredDays < 1 Then coveredDays = 1

    rawProgressFinish = CDate(rawStartVal) + coveredDays - 1

    If IsAggregatedScaleMode() Then
        renderStartOut = GetScalePeriodStart(CDate(rawStartVal))
        renderFinishOut = GetScalePeriodFinish(rawProgressFinish)
    Else
        renderStartOut = rawStartVal
        renderFinishOut = rawProgressFinish
    End If

End Sub

'=====================================================
' Helpers for factorized refresh core
'=====================================================


'------------------------------------------------------------------------------
' FR: Construit une map, un index ou une structure intermediaire consommee par le rendu GANTT.
' EN: Builds a map, index, or intermediate structure consumed by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GanttHierarchy_BuildDirectParentPresenceFromWbs( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object) As Object

    Dim perfScope As clsPerfScope

    Dim d As Object
    Dim r As Long
    Dim wbsVal As String
    Dim parentWbs As String

    Set perfScope = Profiler_BeginScope("GanttHierarchy_BuildDirectParentPresenceFromWbs", "Dictionary")

    Set d = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(dataArr, 1)
        wbsVal = NormalizeWBS(CStr(dataArr(r, mapWBS("WBS"))))
        parentWbs = GetParentWBS(wbsVal)

        If parentWbs <> "" Then
            d(parentWbs) = True
        End If
    Next r

    Set GanttHierarchy_BuildDirectParentPresenceFromWbs = d

End Function

'------------------------------------------------------------------------------
' FR: Construit une map, un index ou une structure intermediaire consommee par le rendu GANTT.
' EN: Builds a map, index, or intermediate structure consumed by GANTT rendering.
'------------------------------------------------------------------------------
Public Function GanttRenderer_BuildWbsRowById( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object) As Object

    Dim perfScope As clsPerfScope

    Dim d As Object
    Dim r As Long
    Dim idVal As String

    Set perfScope = Profiler_BeginScope("GanttRenderer_BuildWbsRowById", "Dictionary")

    Set d = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(dataArr, 1)
        idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))
        If idVal <> "" Then
            d(idVal) = FIRST_TASK_ROW + r - 1
        End If
    Next r

    Set GanttRenderer_BuildWbsRowById = d

End Function



'------------------------------------------------------------------------------
' FR: Dessine une famille visuelle du GANTT a partir des dates et positions deja calculees.
' EN: Draws one GANTT visual family from already computed dates and positions.
'------------------------------------------------------------------------------
Private Sub DrawSummaryBar( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    ByVal isCritical As Boolean, _
    ByVal totalDays As Long, _
    ByVal shapeKey As String, _
    ByVal labelText As String, _
    ByVal hasDelta As Boolean, _
    ByVal isComplete As Boolean)

    Dim perfScope As clsPerfScope

    Dim x1 As Double
    Dim x2 As Double
    Dim yTop As Double
    Dim yMid As Double
    Dim yBottom As Double
    Dim shp As Shape

    Dim displayText As String
    Dim maxLen As Long
    Dim txtShape As Shape
    Dim textLeft As Double
    Dim textTop As Double
    Dim lineColor As Long

    Set perfScope = Profiler_BeginScope("DrawSummaryBar", "Shape Create")

    If Not HasValue(startVal) Then Exit Sub
    If Not HasValue(finishVal) Then Exit Sub

    'Strict date-column anchoring.
    x1 = TimelineLeft(ws, projectStart, startVal)
    x2 = TimelineRightAfterFinish(ws, projectStart, finishVal)

    If x2 <= x1 Then Exit Sub

    yTop = GetGanttSummaryTop(ws, ganttRow)
    yMid = GetGanttRowMid(ws, ganttRow)
    yBottom = GetGanttSummaryBottom(ws, ganttRow)

    lineColor = GetSummaryLineColor(isCritical, isComplete)

    Set shp = ws.Shapes.AddLine(x1, yMid, x2, yMid)
    shp.Name = shapeKey & "_H"
    ApplyGanttRenderLinePlacement shp
    shp.Line.Weight = 2.25
    If hasDelta Then
        shp.Line.ForeColor.RGB = RGB(255, 192, 0)
    Else
        shp.Line.ForeColor.RGB = lineColor
    End If

    Set shp = ws.Shapes.AddLine(x1, yTop, x1, yBottom)
    shp.Name = shapeKey & "_L"
    ApplyGanttRenderLinePlacement shp
    shp.Line.Weight = 2.25
    If hasDelta Then
        shp.Line.ForeColor.RGB = RGB(255, 192, 0)
    Else
        shp.Line.ForeColor.RGB = lineColor
    End If

    Set shp = ws.Shapes.AddLine(x2, yTop, x2, yBottom)
    shp.Name = shapeKey & "_R"
    ApplyGanttRenderLinePlacement shp
    shp.Line.Weight = 2.25
    If hasDelta Then
        shp.Line.ForeColor.RGB = RGB(255, 192, 0)
    Else
        shp.Line.ForeColor.RGB = lineColor
    End If

    maxLen = 28

    displayText = Trim$(labelText)
    If Len(displayText) > maxLen Then
        displayText = Left$(displayText, maxLen) & "..."
    End If

    textLeft = x2 + 8
    textTop = GetGanttRowTop(ws, ganttRow)

    Set txtShape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, textLeft, textTop, 220, GetGanttRowHeight(ws, ganttRow))
    txtShape.Name = shapeKey & "_TXT"
    ApplyGanttRenderShapePlacement txtShape

    With txtShape
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
    End With

    With txtShape.TextFrame2.TextRange
        .Text = displayText
        .Font.Size = 10.5
        .Font.Fill.ForeColor.RGB = RGB(80, 80, 80)
    End With

End Sub


'------------------------------------------------------------------------------
' FR: Execute le helper Get Lo E Progress From Today dans le workflow de rendu GANTT.
' EN: Runs the Get Lo E Progress From Today helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetLoEProgressFromToday( _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant) As Double

    Dim durationVal As Double
    Dim rawProgress As Double

    GetLoEProgressFromToday = 0#

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    durationVal = CDbl(finishVal) - CDbl(startVal) + 1
    If durationVal <= 0 Then Exit Function

    rawProgress = (CDbl(Date) - CDbl(startVal) + 1) / durationVal

    If rawProgress < 0 Then
        GetLoEProgressFromToday = 0#
    ElseIf rawProgress > 1 Then
        GetLoEProgressFromToday = 1#
    Else
        GetLoEProgressFromToday = rawProgress
    End If

End Function

