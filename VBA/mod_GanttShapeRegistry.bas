Attribute VB_Name = "mod_GanttShapeRegistry"
Option Explicit

'===============================================================================
' MODULE : mod_GanttShapeRegistry
' DOMAINE / DOMAIN : Gantt
'
' FR
' Possede les records de Shapes, leur cache, le diff predictif et les decisions fast path/fallback.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Owns Shape records, their cache, predictive diff and fast-path/fallback decisions.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : GanttShapeRegistry_AddTaskBarRecords, GanttShapeRegistry_AddCompactTaskRecords, GanttShapeRegistry_AddMilestoneRecord, GanttShapeRegistry_AddTodayLineRecord, GanttShapeRegistry_CreateAllFromRecords, GanttShapeRegistry_CreateShapeFromRecord, GanttShapeRegistry_UpdateShapeFromRecord, GanttShapeRegistry_ShapeDiffers
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

Private gRenderExpectedNames As Object
Private gRenderExistingShapes As Object
Private gRenderSessionActive As Boolean
Private gRenderDeleteStale As Boolean
Private gRenderStyleOnly As Boolean
Private gRenderCreates As Long
Private gRenderUpdates As Long
Private gRenderDeletes As Long
Private gRenderInspections As Long
Private gRenderStyleUpdates As Long
Private gRenderGeometryUpdates As Long
Private gRenderSkips As Long
Private gRenderTypeMismatchRecreates As Long
Private gRenderLocal As Boolean
Private gRenderLocalRows As Object
Private gCanonicalRecords As Object
Private gRenderTrustCanonical As Boolean

'------------------------------------------------------------------------------
' FR: Ouvre une session differentielle pour les shapes metier Gantt.
' EN: Opens a differential session for Gantt business shapes.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_BeginRenderSession( _
    ByVal ws As Worksheet, _
    Optional ByVal deleteStale As Boolean = True, _
    Optional ByVal assumeBusinessShapesEmpty As Boolean = False, _
    Optional ByVal styleOnly As Boolean = False, _
    Optional ByVal trustCanonical As Boolean = False)

    Dim i As Long
    Dim shapeName As String

    Set gRenderExpectedNames = CreateObject("Scripting.Dictionary")
    Set gRenderExistingShapes = CreateObject("Scripting.Dictionary")
    gRenderSessionActive = True
    gRenderDeleteStale = deleteStale
    gRenderStyleOnly = styleOnly
    gRenderTrustCanonical = trustCanonical And Not gCanonicalRecords Is Nothing
    gRenderLocal = False
    Set gRenderLocalRows = Nothing
    gRenderCreates = 0
    gRenderUpdates = 0
    gRenderDeletes = 0
    gRenderInspections = 0
    gRenderStyleUpdates = 0
    gRenderGeometryUpdates = 0
    gRenderSkips = 0
    gRenderTypeMismatchRecreates = 0

    If Not assumeBusinessShapesEmpty And Not ws Is Nothing Then
        For i = 1 To ws.Shapes.Count
            shapeName = CStr(ws.Shapes(i).Name)
            If GanttShapeRegistry_IsBusinessShapeName(shapeName) Then
                gRenderExistingShapes.Add shapeName, ws.Shapes(i)
                gRenderInspections = gRenderInspections + 1
            End If
        Next i
    End If

End Sub

'------------------------------------------------------------------------------
' Opens a transaction restricted to the rows in affectedIds. No global Shapes
' scan is performed; exact stable names are resolved only when required.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_BeginLocalRenderSession( _
    ByVal ws As Worksheet, _
    ByVal affectedIds As Object, _
    ByVal rowById As Object, _
    Optional ByVal styleOnly As Boolean = False)

    Dim idVal As Variant
    Dim ganttRow As Long
    Dim dataRow As Long

    Set gRenderExpectedNames = CreateObject("Scripting.Dictionary")
    Set gRenderExistingShapes = CreateObject("Scripting.Dictionary")
    Set gRenderLocalRows = CreateObject("Scripting.Dictionary")
    gRenderSessionActive = True
    gRenderDeleteStale = Not styleOnly
    gRenderStyleOnly = styleOnly
    gRenderTrustCanonical = False
    gRenderLocal = True
    gRenderCreates = 0
    gRenderUpdates = 0
    gRenderDeletes = 0
    gRenderInspections = 0
    gRenderStyleUpdates = 0
    gRenderGeometryUpdates = 0
    gRenderSkips = 0
    gRenderTypeMismatchRecreates = 0

    If Not affectedIds Is Nothing And Not rowById Is Nothing Then
        For Each idVal In affectedIds.keys
            If rowById.Exists(CStr(idVal)) Then
                ganttRow = CLng(rowById(CStr(idVal)))
                dataRow = ganttRow - FIRST_TASK_ROW + 1
                If dataRow > 0 Then gRenderLocalRows(CStr(dataRow)) = True
            End If
        Next idVal
    End If

    Profiler_RecordOperation "GanttLocalRegistryRows", gRenderLocalRows.Count, 0#

End Sub

'------------------------------------------------------------------------------
' FR: Enregistre un nom de shape attendu par le rendu canonique courant.
' EN: Registers a shape name expected by the current canonical render.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_RegisterExpectedName(ByVal shapeName As String)

    If Not gRenderSessionActive Then Exit Sub
    If Not gRenderDeleteStale Then Exit Sub
    If Len(shapeName) = 0 Then Exit Sub
    gRenderExpectedNames(shapeName) = True

End Sub

'------------------------------------------------------------------------------
' FR: Supprime uniquement les shapes metier absentes du rendu attendu.
' EN: Deletes only business shapes absent from the expected render.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_EndRenderSession(ByVal ws As Worksheet)

    Dim shapeName As String
    Dim shapeKey As Variant

    If ws Is Nothing Then Exit Sub
    If Not gRenderSessionActive Then Exit Sub

    If gRenderDeleteStale And gRenderLocal Then
        GanttShapeRegistry_DeleteLocalStale ws
    ElseIf gRenderDeleteStale Then
        'The opening scan already owns exact references for every business
        'shape. Reuse it instead of crossing the Shapes COM collection again.
        For Each shapeKey In gRenderExistingShapes.keys
            shapeName = CStr(shapeKey)
            If Not gRenderExpectedNames.Exists(shapeName) Then
                gRenderExistingShapes(shapeName).Delete
                GanttShapeRegistry_RemoveCanonical shapeName
                gRenderDeletes = gRenderDeletes + 1
            End If
        Next shapeKey
    End If

    Profiler_RecordOperation "GanttDiffShapesInspected", gRenderInspections, 0#
    Profiler_RecordOperation "GanttDiffShapesCreated", gRenderCreates, 0#
    Profiler_RecordOperation "GanttDiffShapesUpdated", gRenderUpdates, 0#
    Profiler_RecordOperation "GanttDiffShapesDeleted", gRenderDeletes, 0#
    Profiler_RecordOperation "GanttDiffStylesUpdated", gRenderStyleUpdates, 0#
    Profiler_RecordOperation "GanttDiffGeometriesUpdated", gRenderGeometryUpdates, 0#
    Profiler_RecordOperation "GanttDiffShapesSkipped", gRenderSkips, 0#
    Profiler_RecordOperation "GanttDiffTypeMismatchRecreates", gRenderTypeMismatchRecreates, 0#

    Set gRenderExpectedNames = Nothing
    Set gRenderExistingShapes = Nothing
    gRenderSessionActive = False
    gRenderDeleteStale = False
    gRenderStyleOnly = False
    gRenderTrustCanonical = False
    gRenderLocal = False
    Set gRenderLocalRows = Nothing

End Sub

'------------------------------------------------------------------------------
' FR: Abandonne une session interrompue sans conserver d'etat differentiel stale.
' EN: Cancels an interrupted session without retaining stale differential state.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_CancelRenderSession()

    Set gRenderExpectedNames = Nothing
    Set gRenderExistingShapes = Nothing
    gRenderSessionActive = False
    gRenderDeleteStale = False
    gRenderStyleOnly = False
    gRenderTrustCanonical = False
    gRenderLocal = False
    Set gRenderLocalRows = Nothing
    gRenderCreates = 0
    gRenderUpdates = 0
    gRenderDeletes = 0
    gRenderInspections = 0
    gRenderStyleUpdates = 0
    gRenderGeometryUpdates = 0
    gRenderSkips = 0
    gRenderTypeMismatchRecreates = 0

End Sub

'------------------------------------------------------------------------------
' FR: Indique qu'une session ne peut modifier que le style des shapes existantes.
' EN: Returns whether the current session may only style existing shapes.
'------------------------------------------------------------------------------
Public Function GanttShapeRegistry_IsStyleOnlySession() As Boolean

    GanttShapeRegistry_IsStyleOnlySession = gRenderSessionActive And gRenderStyleOnly

End Function


Public Function GanttShapeRegistry_HasCanonicalRecords() As Boolean

    If gCanonicalRecords Is Nothing Then Exit Function
    GanttShapeRegistry_HasCanonicalRecords = (gCanonicalRecords.Count > 0)

End Function


Public Function GanttShapeRegistry_HasCanonicalRecord(ByVal shapeName As String) As Boolean

    If gCanonicalRecords Is Nothing Then Exit Function
    GanttShapeRegistry_HasCanonicalRecord = gCanonicalRecords.Exists(shapeName)

End Function


Public Function GanttShapeRegistry_CanApplyStyleOnlyRecord( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String) As Boolean

    Dim shp As Shape

    If ws Is Nothing Then Exit Function
    If GanttShapeRegistry_HasCanonicalRecord(shapeName) Then
        GanttShapeRegistry_CanApplyStyleOnlyRecord = True
        Exit Function
    End If

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    ' No physical target means there is no style write to perform.
    GanttShapeRegistry_CanApplyStyleOnlyRecord = (shp Is Nothing)

End Function


Public Function GanttShapeRegistry_ApplyStyleOnlyRecord( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal fillColor As Variant, _
    ByVal lineColor As Variant) As Boolean

    Dim rec As Object
    Dim newRec As Object
    Dim key As Variant

    If ws Is Nothing Then Exit Function
    If gCanonicalRecords Is Nothing Then Exit Function
    If Not gCanonicalRecords.Exists(shapeName) Then Exit Function

    Set rec = gCanonicalRecords(shapeName)
    Set newRec = CreateObject("Scripting.Dictionary")
    For Each key In rec.keys
        newRec.Add CStr(key), rec(key)
    Next key

    If CBool(newRec("IsLine")) Then
        If Not IsEmpty(lineColor) Then newRec("LineColor") = CLng(lineColor)
    Else
        If Not IsEmpty(fillColor) Then newRec("FillColor") = CLng(fillColor)
    End If

    If CBool(newRec("IsLine")) Then
        newRec("StyleKey") = CStr(newRec("Family")) & "|" & CStr(newRec("Subtype")) & "|LINE|" & CStr(newRec("LineColor")) & "|" & CStr(newRec("LineWeight"))
    Else
        newRec("StyleKey") = CStr(newRec("Family")) & "|" & CStr(newRec("Subtype")) & "|" & CStr(newRec("AutoShapeType")) & "|" & CStr(newRec("FillColor")) & "|" & CStr(newRec("LineVisible")) & "|" & CStr(newRec("LineColor")) & "|" & CStr(newRec("LineWeight"))
    End If

    GanttShapeRegistry_UpsertShapeFromRecord ws, newRec
    GanttShapeRegistry_ApplyStyleOnlyRecord = True

End Function
Public Sub GanttShapeRegistry_InvalidateCanonical(Optional ByVal reason As String = "")

    Set gCanonicalRecords = Nothing
    If Len(Trim$(reason)) > 0 Then
        Profiler_RecordOperation "GanttRegistryInvalidation_" & reason, 1, 0#
    End If

End Sub


'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_AddTaskBarRecords( _
    ByVal expected As Object, _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    ByVal progressVal As Double, _
    ByVal isCritical As Boolean, _
    ByVal shapeName As String, _
    ByVal hasDelta As Boolean, _
    ByVal isLoE As Boolean)

    Dim leftPos As Double
    Dim rightPos As Double
    Dim topPos As Double
    Dim barWidth As Double
    Dim barHeight As Double
    Dim fullBarTop As Double
    Dim fullBarHeight As Double
    Dim progressColor As Variant
    Dim progressWidth As Double
    Dim progressLeft As Double
    Dim progressTop As Double
    Dim progressHeight As Double

    If Not HasValue(startVal) Or Not HasValue(finishVal) Then Exit Sub

    leftPos = TimelineLeft(ws, projectStart, startVal)
    rightPos = TimelineRightAfterFinish(ws, projectStart, finishVal)
    If rightPos <= leftPos Then Exit Sub

    fullBarTop = GetGanttBarTop(ws, ganttRow)
    fullBarHeight = GetGanttBarHeight(ws, ganttRow)

    If isLoE Then
        barHeight = fullBarHeight / 2
        If barHeight < 2 Then barHeight = 2
        topPos = fullBarTop + ((fullBarHeight - barHeight) / 2)
    Else
        barHeight = fullBarHeight
        topPos = fullBarTop
    End If

    barWidth = rightPos - leftPos

    GanttShapeRegistry_AddShapeRecord expected, shapeName, "TASK", "BAR", msoShapeRoundedRectangle, _
        leftPos, topPos, barWidth, barHeight, True, GetTaskExpectedFillColor(startVal, finishVal, progressVal, isCritical), 0#, _
        hasDelta, RGB(255, 192, 0), 2.75

    If progressVal > 0 And progressVal < 1 Then
        progressColor = GetProgressFillColor(startVal, finishVal, progressVal)
        If Not IsEmpty(progressColor) Then
            progressWidth = barWidth * WorksheetFunction.Min(progressVal, 1)
            If progressWidth < 2 Then progressWidth = 2
            progressLeft = leftPos + 2
            progressTop = topPos + 2
            progressHeight = barHeight - 4
            If progressWidth > barWidth - 4 Then progressWidth = barWidth - 4
            If progressHeight < 2 Then progressHeight = 2

            GanttShapeRegistry_AddShapeRecord expected, shapeName & "_P", "TASK", "PROGRESS", msoShapeRoundedRectangle, _
                progressLeft, progressTop, progressWidth, progressHeight, True, CLng(progressColor), 0.25, _
                False, 0, 0#
        End If
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_AddCompactTaskRecords( _
    ByVal expected As Object, _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    ByVal progressVal As Double, _
    ByVal isCritical As Boolean, _
    ByVal shapeName As String, _
    ByVal hasDelta As Boolean)

    Dim targetCol As Long
    Dim cellWidth As Double
    Dim markerCenterX As Double
    Dim topPos As Double
    Dim sizeVal As Double
    Dim innerSize As Double
    Dim progressColor As Variant

    If Not HasValue(startVal) Then Exit Sub
    If Not HasValue(finishVal) Then finishVal = startVal

    cellWidth = GanttTimelineProjection_SlotWidth(projectStart, startVal)
    If cellWidth <= 0 Then
        targetCol = TimelineColumnFromHeaderDate_Exact(ws, projectStart, startVal)
        If targetCol < FIRST_TIMELINE_COL Then Exit Sub
        cellWidth = ws.cells(HEADER_ROW_2, targetCol).Width
    End If
    sizeVal = GetGanttCompactTaskMarkerSize(ws, ganttRow, cellWidth)
    If sizeVal < 2 Then sizeVal = 2

    markerCenterX = TimelineDateRangeMidX(ws, projectStart, startVal, finishVal)
    If markerCenterX <= 0 Then Exit Sub

    topPos = ws.cells(ganttRow, FIRST_TIMELINE_COL).Top + ((ws.rows(ganttRow).Height - sizeVal) / 2)

    GanttShapeRegistry_AddShapeRecord expected, shapeName, "TASK", "COMPACT", msoShapeOval, _
        markerCenterX - (sizeVal / 2), topPos, sizeVal, sizeVal, True, GetTaskExpectedFillColor(startVal, finishVal, progressVal, isCritical), 0#, _
        hasDelta, RGB(255, 192, 0), 2.75

    If progressVal > 0 And progressVal < 1 Then
        progressColor = GetProgressFillColor(startVal, finishVal, progressVal)
        innerSize = sizeVal * 0.5
        GanttShapeRegistry_AddShapeRecord expected, shapeName & "_P", "TASK", "COMPACT_PROGRESS", msoShapeOval, _
            markerCenterX - (innerSize / 2), topPos + ((sizeVal - innerSize) / 2), innerSize, innerSize, True, CLng(progressColor), 0.15, _
            False, 0, 0#
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_AddMilestoneRecord( _
    ByVal expected As Object, _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal progressVal As Double, _
    ByVal isCritical As Boolean, _
    ByVal shapeName As String, _
    ByVal hasDelta As Boolean)

    Dim leftPos As Double
    Dim topPos As Double
    Dim sizeVal As Double
    Dim cellMidX As Double

    If Not HasValue(startVal) Then Exit Sub

    sizeVal = GetGanttMilestoneSize(ws, ganttRow)
    cellMidX = GetTaskMidX(ws, projectStart, startVal)
    If cellMidX <= 0 Then Exit Sub

    leftPos = cellMidX - (sizeVal / 2)
    topPos = GetGanttRowTop(ws, ganttRow) + ((GetGanttRowHeight(ws, ganttRow) - sizeVal) / 2)

    GanttShapeRegistry_AddShapeRecord expected, shapeName, "MS", "MILESTONE", msoShapeDiamond, _
        leftPos, topPos, sizeVal, sizeVal, True, GetTaskExpectedFillColor(startVal, startVal, progressVal, isCritical), 0#, _
        hasDelta, RGB(255, 192, 0), 2.75

End Sub
'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_AddTodayLineRecord( _
    ByVal expected As Object, _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal rowCount As Long)

    Dim todayVal As Date
    Dim projectFinish As Date
    Dim x As Double
    Dim yTop As Double
    Dim yBottom As Double
    Dim rec As Object

    If ws Is Nothing Then Exit Sub
    If Not HasValue(projectStart) Then Exit Sub
    If totalDays < 1 Or rowCount < 1 Then Exit Sub

    todayVal = Date
    projectFinish = GetScaleProjectFinishFromSlots(projectStart, totalDays)
    If todayVal < CDate(projectStart) Or todayVal > projectFinish Then Exit Sub

    x = GetTaskMidX(ws, projectStart, todayVal)
    yTop = ws.cells(HEADER_ROW_1, FIRST_TIMELINE_COL).Top
    yBottom = GetGanttRowTop(ws, FIRST_TASK_ROW + rowCount - 1) + GetGanttRowHeight(ws, FIRST_TASK_ROW + rowCount - 1)
    If x <= 0 Or yBottom <= yTop Then Exit Sub

    Set rec = CreateObject("Scripting.Dictionary")
    rec("Name") = "TODAY_LINE"
    rec("Family") = "TODAY"
    rec("Subtype") = "LINE"
    rec("IsLine") = True
    rec("X1") = x
    rec("Y1") = yTop
    rec("X2") = x
    rec("Y2") = yBottom
    rec("Left") = x
    rec("Top") = yTop
    rec("Width") = 0#
    rec("Height") = yBottom - yTop
    rec("Visible") = True
    rec("LineColor") = RGB(255, 192, 0)
    rec("LineWeight") = 4.5
    rec("DashStyle") = msoLineDash
    rec("ZFront") = True
    expected.Add "TODAY_LINE", rec

End Sub
'------------------------------------------------------------------------------
' FR: Execute le helper Get Task Expected Fill Color dans le workflow de rendu GANTT.
' EN: Runs the Get Task Expected Fill Color helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function GetTaskExpectedFillColor( _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    ByVal progressVal As Double, _
    ByVal isCritical As Boolean) As Long

    If progressVal >= 1 Then
        GetTaskExpectedFillColor = COLOR_PROGRESS_GREEN
    Else
        GetTaskExpectedFillColor = GetTaskBaseColor(isCritical)
    End If

End Function
'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Private Sub GanttShapeRegistry_AddShapeRecord( _
    ByVal expected As Object, _
    ByVal shapeName As String, _
    ByVal familyName As String, _
    ByVal subtypeName As String, _
    ByVal autoShapeType As Long, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal widthVal As Double, _
    ByVal heightVal As Double, _
    ByVal visibleVal As Boolean, _
    ByVal fillColor As Long, _
    ByVal fillTransparency As Double, _
    ByVal lineVisible As Boolean, _
    ByVal lineColor As Long, _
    ByVal lineWeight As Double)

    Dim rec As Object

    Set rec = CreateObject("Scripting.Dictionary")
    rec("Name") = shapeName
    rec("Family") = familyName
    rec("Subtype") = subtypeName
    rec("IsLine") = False
    rec("AutoShapeType") = autoShapeType
    rec("Left") = leftPos
    rec("Top") = topPos
    rec("Width") = widthVal
    rec("Height") = heightVal
    rec("Visible") = visibleVal
    rec("FillColor") = fillColor
    rec("FillTransparency") = fillTransparency
    rec("LineVisible") = lineVisible
    rec("LineColor") = lineColor
    rec("LineWeight") = lineWeight
    rec("StyleKey") = familyName & "|" & subtypeName & "|" & CStr(autoShapeType) & "|" & CStr(fillColor) & "|" & CStr(lineVisible) & "|" & CStr(lineColor) & "|" & CStr(lineWeight)
    expected.Add shapeName, rec

End Sub

'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_CreateAllFromRecords(ByVal ws As Worksheet, ByVal records As Object)

    Dim key As Variant

    If records Is Nothing Then Exit Sub

    For Each key In records.keys
        GanttShapeRegistry_RegisterExpectedName CStr(key)
        GanttShapeRegistry_UpsertShapeFromRecord ws, records(CStr(key))
    Next key

End Sub

'------------------------------------------------------------------------------
' FR: Cree ou met a jour une shape nommee sans recreation si son type est stable.
' EN: Creates or updates a named shape without recreation when its type is stable.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_UpsertShapeFromRecord(ByVal ws As Worksheet, ByVal rec As Object)

    Dim shp As Shape
    Dim oldRec As Object
    Dim geometryDiffers As Boolean
    Dim styleDiffers As Boolean
    Dim typeMismatch As Boolean
    Dim hasCanonical As Boolean

    If ws Is Nothing Then Exit Sub
    If rec Is Nothing Then Exit Sub

    GanttShapeRegistry_RegisterExpectedName CStr(rec("Name"))

    If Not gCanonicalRecords Is Nothing Then
        If gCanonicalRecords.Exists(CStr(rec("Name"))) Then
            Set oldRec = gCanonicalRecords(CStr(rec("Name")))
            hasCanonical = True
            If GanttShapeRegistry_RecordsEqual(oldRec, rec) Then
                If gRenderTrustCanonical Then
                    If gRenderExistingShapes.Exists(CStr(rec("Name"))) Then
                        Profiler_RecordOperation "GanttRegistryTrustedEqualSkips", 1, 0#
                        Exit Sub
                    End If
                    GanttShapeRegistry_RemoveCanonical CStr(rec("Name"))
                    Set oldRec = Nothing
                    hasCanonical = False
                End If
                If gRenderLocal Then
                    On Error Resume Next
                    Set shp = ws.Shapes(CStr(rec("Name")))
                    On Error GoTo 0
                    gRenderInspections = gRenderInspections + 1

                    If Not shp Is Nothing Then
                        If Not GanttShapeRegistry_ShapeDiffers(shp, rec) Then
                            gRenderExistingShapes.Add CStr(rec("Name")), shp
                            gRenderSkips = gRenderSkips + 1
                            Profiler_RecordOperation "GanttRegistryKnownEqualSkips", 1, 0#
                            Exit Sub
                        End If
                    Else
                        GanttShapeRegistry_RemoveCanonical CStr(rec("Name"))
                    End If
                Else
                    gRenderSkips = gRenderSkips + 1
                    Profiler_RecordOperation "GanttRegistryKnownEqualSkips", 1, 0#
                    Exit Sub
                End If
            End If
        End If
    End If

    If gRenderSessionActive Then
        If gRenderExistingShapes.Exists(CStr(rec("Name"))) Then
            Set shp = gRenderExistingShapes(CStr(rec("Name")))
        End If
        If shp Is Nothing And gRenderLocal Then
            On Error Resume Next
            Set shp = ws.Shapes(CStr(rec("Name")))
            On Error GoTo 0
            If Not shp Is Nothing Then
                gRenderExistingShapes.Add CStr(rec("Name")), shp
                gRenderInspections = gRenderInspections + 1
            End If
        End If
    Else
        On Error Resume Next
        Set shp = ws.Shapes(CStr(rec("Name")))
        On Error GoTo 0
    End If

    If shp Is Nothing Then
        If gRenderStyleOnly Then
            Err.Raise vbObjectError + 441, _
                "GanttShapeRegistry_UpsertShapeFromRecord", _
                "Style-only render found a missing shape: " & CStr(rec("Name"))
        End If
        GanttShapeRegistry_CreateShapeFromRecord ws, rec
        If gRenderSessionActive Then
            Set shp = ws.Shapes(ws.Shapes.Count)
            gRenderExistingShapes.Add CStr(rec("Name")), shp
        End If
        If Not gRenderSessionActive Then
            Profiler_RecordOperation "GanttDiffShapesCreated", 1, 0#
        Else
            gRenderCreates = gRenderCreates + 1
        End If
        GanttShapeRegistry_StoreCanonical rec
        Exit Sub
    End If

    If Not gRenderSessionActive Then
        Profiler_RecordOperation "GanttDiffShapesInspected", 1, 0#
    End If

    If gRenderTrustCanonical And hasCanonical Then
        typeMismatch = GanttShapeRegistry_RecordTypeMismatch(oldRec, rec)
    Else
        typeMismatch = GanttShapeRegistry_TypeMismatch(shp, rec)
    End If

    If typeMismatch Then
        If gRenderStyleOnly Then
            Err.Raise vbObjectError + 442, _
                "GanttShapeRegistry_UpsertShapeFromRecord", _
                "Style-only render found an incompatible shape: " & CStr(rec("Name"))
        End If
        shp.Delete
        GanttShapeRegistry_CreateShapeFromRecord ws, rec
        If gRenderSessionActive Then
            Set shp = ws.Shapes(ws.Shapes.Count)
            Set gRenderExistingShapes(CStr(rec("Name"))) = shp
        End If
        If Not gRenderSessionActive Then
            Profiler_RecordOperation "GanttDiffShapesDeleted", 1, 0#
            Profiler_RecordOperation "GanttDiffShapesCreated", 1, 0#
        Else
            gRenderDeletes = gRenderDeletes + 1
            gRenderCreates = gRenderCreates + 1
            gRenderTypeMismatchRecreates = gRenderTypeMismatchRecreates + 1
        End If
    Else
        If gRenderTrustCanonical And hasCanonical Then
            If Not gRenderStyleOnly Then
                geometryDiffers = GanttShapeRegistry_RecordGeometryDiffers(oldRec, rec)
            End If
            styleDiffers = GanttShapeRegistry_RecordStyleDiffers(oldRec, rec)
        Else
            If Not gRenderStyleOnly Then
                geometryDiffers = GanttShapeRegistry_GeometryDiffers(shp, rec)
            End If
            styleDiffers = GanttShapeRegistry_StyleDiffers(shp, rec)
        End If

        If geometryDiffers Then
            If Not gRenderSessionActive Then
                Profiler_RecordOperation "GanttDiffGeometriesUpdated", 1, 0#
            Else
                gRenderGeometryUpdates = gRenderGeometryUpdates + 1
            End If
        End If
        If styleDiffers Then
            If Not gRenderSessionActive Then
                Profiler_RecordOperation "GanttDiffStylesUpdated", 1, 0#
            Else
                gRenderStyleUpdates = gRenderStyleUpdates + 1
            End If
        End If

        If geometryDiffers Then
            If gRenderTrustCanonical And hasCanonical Then
                GanttShapeRegistry_ApplyGeometryDiff shp, oldRec, rec
            Else
                GanttShapeRegistry_ApplyGeometry shp, rec
            End If
        End If
        If styleDiffers Then GanttShapeRegistry_ApplyStyle shp, rec

        If geometryDiffers Or styleDiffers Then
            If Not gRenderSessionActive Then
                Profiler_RecordOperation "GanttDiffShapesUpdated", 1, 0#
            Else
                gRenderUpdates = gRenderUpdates + 1
            End If
        ElseIf gRenderSessionActive Then
            gRenderSkips = gRenderSkips + 1
        End If
    End If

    GanttShapeRegistry_StoreCanonical rec

End Sub
'------------------------------------------------------------------------------
' FR:
' Cree une shape Excel ou une ligne a partir d'un record de registre standardise.
'
' EN:
' Creates an Excel shape or line from a standardized registry record.
'
' Entrees / Inputs:
' - Feuille GANTT et record contenant geometrie, type et style.
'
' Sorties / Outputs:
' - Shape nommee et placee, puis stylisee par GanttShapeRegistry_UpdateShapeFromRecord.
'
' Appele par / Called by:
' - GanttShapeRegistry_CreateAllFromRecords et GanttPredictive_ApplyDiff.
'
' Notes:
' - API interne du registre de rendu; a conserver coherent avec les champs de record.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_CreateShapeFromRecord(ByVal ws As Worksheet, ByVal rec As Object)

    Dim shp As Shape

    If CBool(rec("IsLine")) Then
        Set shp = ws.Shapes.AddLine(CDbl(rec("X1")), CDbl(rec("Y1")), CDbl(rec("X2")), CDbl(rec("Y2")))
        shp.Name = CStr(rec("Name"))
        ApplyGanttRenderLinePlacement shp
    Else
        Set shp = ws.Shapes.AddShape(CLng(rec("AutoShapeType")), CDbl(rec("Left")), CDbl(rec("Top")), CDbl(rec("Width")), CDbl(rec("Height")))
        shp.Name = CStr(rec("Name"))
        ApplyGanttRenderShapePlacement shp
    End If

    GanttShapeRegistry_UpdateShapeFromRecord shp, rec

End Sub
'------------------------------------------------------------------------------
' FR:
' Synchronise geometrie, visibilite et style d'une shape existante depuis son record.
'
' EN:
' Synchronizes geometry, visibility, and style of an existing shape from its record.
'
' Entrees / Inputs:
' - Shape existante et record de rendu.
'
' Sorties / Outputs:
' - Shape mise a jour sans recreation si son type reste compatible.
'
' Appele par / Called by:
' - GanttShapeRegistry_CreateShapeFromRecord et GanttPredictive_ApplyDiff.
'
' Notes:
' - Cle du refresh predictif: evite une recreation couteuse quand le type ne change pas.
'------------------------------------------------------------------------------
Public Sub GanttShapeRegistry_UpdateShapeFromRecord(ByVal shp As Shape, ByVal rec As Object)

    GanttShapeRegistry_ApplyGeometry shp, rec
    GanttShapeRegistry_ApplyStyle shp, rec

End Sub

'------------------------------------------------------------------------------
' FR: Applique uniquement la geometrie et la visibilite d'un record.
' EN: Applies only geometry and visibility from a record.
'------------------------------------------------------------------------------
Private Sub GanttShapeRegistry_ApplyGeometry(ByVal shp As Shape, ByVal rec As Object)

    If CBool(rec("IsLine")) Then
        shp.Left = CDbl(rec("Left"))
        shp.Top = CDbl(rec("Top"))
        shp.Width = CDbl(rec("Width"))
        shp.Height = CDbl(rec("Height"))
    Else
        shp.Left = CDbl(rec("Left"))
        shp.Top = CDbl(rec("Top"))
        shp.Width = CDbl(rec("Width"))
        shp.Height = CDbl(rec("Height"))
    End If

    shp.Visible = IIf(CBool(rec("Visible")), msoTrue, msoFalse)

End Sub

'------------------------------------------------------------------------------
' Applies only canonical properties that changed across a trusted render
' transaction. This avoids four redundant COM writes for a horizontal-only
' scale projection while keeping the same final record.
'------------------------------------------------------------------------------
Private Sub GanttShapeRegistry_ApplyGeometryDiff( _
    ByVal shp As Shape, _
    ByVal oldRec As Object, _
    ByVal newRec As Object)

    If gRenderTrustCanonical Or _
       Abs(CDbl(oldRec("Left")) - CDbl(newRec("Left"))) > 0.1 Then
        shp.Left = CDbl(newRec("Left"))
        Profiler_RecordOperation "GanttScaleCOMWriteLeft", 1, 0#
    End If
    If gRenderTrustCanonical Or _
       Abs(CDbl(oldRec("Top")) - CDbl(newRec("Top"))) > 0.1 Then
        shp.Top = CDbl(newRec("Top"))
        Profiler_RecordOperation "GanttScaleCOMWriteTop", 1, 0#
    End If
    If gRenderTrustCanonical Or _
       Abs(CDbl(oldRec("Width")) - CDbl(newRec("Width"))) > 0.1 Then
        shp.Width = CDbl(newRec("Width"))
        Profiler_RecordOperation "GanttScaleCOMWriteWidth", 1, 0#
    End If
    If gRenderTrustCanonical Or _
       Abs(CDbl(oldRec("Height")) - CDbl(newRec("Height"))) > 0.1 Then
        shp.Height = CDbl(newRec("Height"))
        Profiler_RecordOperation "GanttScaleCOMWriteHeight", 1, 0#
    End If
    If CBool(oldRec("Visible")) <> CBool(newRec("Visible")) Then
        shp.Visible = IIf(CBool(newRec("Visible")), msoTrue, msoFalse)
        Profiler_RecordOperation "GanttScaleCOMWriteVisible", 1, 0#
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Applique uniquement le style d'un record sans toucher a sa geometrie.
' EN: Applies only styling from a record without touching its geometry.
'------------------------------------------------------------------------------
Private Sub GanttShapeRegistry_ApplyStyle(ByVal shp As Shape, ByVal rec As Object)

    If CBool(rec("IsLine")) Then
        With shp.Line
            .ForeColor.RGB = CLng(rec("LineColor"))
            .Weight = CDbl(rec("LineWeight"))
            .DashStyle = CLng(rec("DashStyle"))
        End With
        If CBool(rec("ZFront")) Then shp.ZOrder msoBringToFront
    Else
        shp.Fill.Visible = msoTrue
        shp.Fill.ForeColor.RGB = CLng(rec("FillColor"))
        shp.Fill.Transparency = CDbl(rec("FillTransparency"))
        If CBool(rec("LineVisible")) Then
            shp.Line.Visible = msoTrue
            shp.Line.ForeColor.RGB = CLng(rec("LineColor"))
            shp.Line.Weight = CDbl(rec("LineWeight"))
        Else
            shp.Line.Visible = msoFalse
        End If
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Public Function GanttShapeRegistry_ShapeDiffers(ByVal shp As Shape, ByVal rec As Object) As Boolean

    On Error GoTo Differs

    If GanttShapeRegistry_TypeMismatch(shp, rec) Then GoTo Differs
    If GanttShapeRegistry_GeometryDiffers(shp, rec) Then GoTo Differs
    If GanttShapeRegistry_StyleDiffers(shp, rec) Then GoTo Differs

    GanttShapeRegistry_ShapeDiffers = False
    Exit Function

Differs:
    GanttShapeRegistry_ShapeDiffers = True

End Function

'------------------------------------------------------------------------------
' FR: Compare uniquement la geometrie et la visibilite d'une shape.
' EN: Compares only shape geometry and visibility.
'------------------------------------------------------------------------------
Public Function GanttShapeRegistry_GeometryDiffers(ByVal shp As Shape, ByVal rec As Object) As Boolean

    On Error GoTo Differs

    If Abs(shp.Left - CDbl(rec("Left"))) > 0.1 Then GoTo Differs
    If Abs(shp.Top - CDbl(rec("Top"))) > 0.1 Then GoTo Differs
    If Abs(shp.Width - CDbl(rec("Width"))) > 0.1 Then GoTo Differs
    If Abs(shp.Height - CDbl(rec("Height"))) > 0.1 Then GoTo Differs
    If (shp.Visible = msoTrue) <> CBool(rec("Visible")) Then GoTo Differs
    Exit Function

Differs:
    GanttShapeRegistry_GeometryDiffers = True

End Function

'------------------------------------------------------------------------------
' FR: Compare uniquement le style d'une shape.
' EN: Compares only shape styling.
'------------------------------------------------------------------------------
Public Function GanttShapeRegistry_StyleDiffers(ByVal shp As Shape, ByVal rec As Object) As Boolean

    On Error GoTo Differs

    If CBool(rec("IsLine")) Then
        If shp.Line.ForeColor.RGB <> CLng(rec("LineColor")) Then GoTo Differs
        If Abs(shp.Line.Weight - CDbl(rec("LineWeight"))) > 0.01 Then GoTo Differs
        If shp.Line.DashStyle <> CLng(rec("DashStyle")) Then GoTo Differs
    Else
        If shp.Fill.ForeColor.RGB <> CLng(rec("FillColor")) Then GoTo Differs
        If Abs(shp.Fill.Transparency - CDbl(rec("FillTransparency"))) > 0.01 Then GoTo Differs
        If (shp.Line.Visible = msoTrue) <> CBool(rec("LineVisible")) Then GoTo Differs
        If CBool(rec("LineVisible")) Then
            If shp.Line.ForeColor.RGB <> CLng(rec("LineColor")) Then GoTo Differs
            If Abs(shp.Line.Weight - CDbl(rec("LineWeight"))) > 0.01 Then GoTo Differs
        End If
    End If
    Exit Function

Differs:
    GanttShapeRegistry_StyleDiffers = True

End Function
'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Public Function GanttShapeRegistry_TypeMismatch(ByVal shp As Shape, ByVal rec As Object) As Boolean

    On Error GoTo Mismatch

    If CBool(rec("IsLine")) Then
        GanttShapeRegistry_TypeMismatch = (shp.Type <> msoLine)
    Else
        If shp.Type = msoLine Then GoTo Mismatch
        GanttShapeRegistry_TypeMismatch = (shp.autoShapeType <> CLng(rec("AutoShapeType")))
    End If
    Exit Function

Mismatch:
    GanttShapeRegistry_TypeMismatch = True

End Function

Private Sub GanttShapeRegistry_DeleteLocalStale(ByVal ws As Worksheet)

    Dim rowKey As Variant
    Dim candidates As Variant
    Dim candidate As Variant
    Dim rowText As String

    If gRenderLocalRows Is Nothing Then Exit Sub

    For Each rowKey In gRenderLocalRows.keys
        rowText = CStr(rowKey)
        candidates = Array( _
            "TASK_" & rowText, _
            "TASK_" & rowText & "_P", _
            "MS_" & rowText, _
            "SUM_" & rowText & "_H", _
            "SUM_" & rowText & "_L", _
            "SUM_" & rowText & "_R", _
            "SUM_" & rowText & "_TXT")

        For Each candidate In candidates
            If Not gRenderExpectedNames.Exists(CStr(candidate)) Then
                GanttShapeRegistry_DeleteExactShape ws, CStr(candidate)
            End If
        Next candidate
    Next rowKey

End Sub

Private Sub GanttShapeRegistry_DeleteExactShape(ByVal ws As Worksheet, ByVal shapeName As String)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If Not shp Is Nothing Then
        shp.Delete
        gRenderDeletes = gRenderDeletes + 1
    End If

    GanttShapeRegistry_RemoveCanonical shapeName

End Sub

Private Sub GanttShapeRegistry_StoreCanonical(ByVal rec As Object)

    Dim copyRec As Object
    Dim key As Variant
    Dim shapeName As String

    If rec Is Nothing Then Exit Sub
    If gCanonicalRecords Is Nothing Then Set gCanonicalRecords = CreateObject("Scripting.Dictionary")

    Set copyRec = CreateObject("Scripting.Dictionary")
    For Each key In rec.keys
        copyRec.Add CStr(key), rec(key)
    Next key

    shapeName = CStr(rec("Name"))
    If gCanonicalRecords.Exists(shapeName) Then
        Set gCanonicalRecords(shapeName) = copyRec
    Else
        gCanonicalRecords.Add shapeName, copyRec
    End If

End Sub

Private Sub GanttShapeRegistry_RemoveCanonical(ByVal shapeName As String)

    If gCanonicalRecords Is Nothing Then Exit Sub
    If gCanonicalRecords.Exists(shapeName) Then gCanonicalRecords.Remove shapeName

End Sub

Private Function GanttShapeRegistry_RecordsEqual(ByVal oldRec As Object, ByVal newRec As Object) As Boolean

    Dim key As Variant
    Dim oldValue As Variant
    Dim newValue As Variant

    If oldRec Is Nothing Or newRec Is Nothing Then Exit Function
    If oldRec.Count <> newRec.Count Then Exit Function

    For Each key In newRec.keys
        If Not oldRec.Exists(CStr(key)) Then Exit Function
        oldValue = oldRec(CStr(key))
        newValue = newRec(CStr(key))

        If IsNumeric(oldValue) And IsNumeric(newValue) Then
            If Abs(CDbl(oldValue) - CDbl(newValue)) > 0.0001 Then Exit Function
        Else
            If CStr(oldValue) <> CStr(newValue) Then Exit Function
        End If
    Next key

    GanttShapeRegistry_RecordsEqual = True

End Function

Private Function GanttShapeRegistry_RecordTypeMismatch( _
    ByVal oldRec As Object, _
    ByVal newRec As Object) As Boolean

    If CBool(oldRec("IsLine")) <> CBool(newRec("IsLine")) Then
        GanttShapeRegistry_RecordTypeMismatch = True
        Exit Function
    End If

    If Not CBool(newRec("IsLine")) Then
        GanttShapeRegistry_RecordTypeMismatch = _
            (CLng(oldRec("AutoShapeType")) <> CLng(newRec("AutoShapeType")))
    End If

End Function

Private Function GanttShapeRegistry_RecordGeometryDiffers( _
    ByVal oldRec As Object, _
    ByVal newRec As Object) As Boolean

    If Abs(CDbl(oldRec("Left")) - CDbl(newRec("Left"))) > 0.1 Then GoTo Differs
    If Abs(CDbl(oldRec("Top")) - CDbl(newRec("Top"))) > 0.1 Then GoTo Differs
    If Abs(CDbl(oldRec("Width")) - CDbl(newRec("Width"))) > 0.1 Then GoTo Differs
    If Abs(CDbl(oldRec("Height")) - CDbl(newRec("Height"))) > 0.1 Then GoTo Differs
    If CBool(oldRec("Visible")) <> CBool(newRec("Visible")) Then GoTo Differs
    Exit Function

Differs:
    GanttShapeRegistry_RecordGeometryDiffers = True

End Function

Private Function GanttShapeRegistry_RecordStyleDiffers( _
    ByVal oldRec As Object, _
    ByVal newRec As Object) As Boolean

    If CBool(oldRec("IsLine")) <> CBool(newRec("IsLine")) Then GoTo Differs

    If CBool(newRec("IsLine")) Then
        If CLng(oldRec("LineColor")) <> CLng(newRec("LineColor")) Then GoTo Differs
        If Abs(CDbl(oldRec("LineWeight")) - CDbl(newRec("LineWeight"))) > 0.01 Then GoTo Differs
        If CLng(oldRec("DashStyle")) <> CLng(newRec("DashStyle")) Then GoTo Differs
        If CBool(oldRec("ZFront")) <> CBool(newRec("ZFront")) Then GoTo Differs
    Else
        If CLng(oldRec("FillColor")) <> CLng(newRec("FillColor")) Then GoTo Differs
        If Abs(CDbl(oldRec("FillTransparency")) - CDbl(newRec("FillTransparency"))) > 0.01 Then GoTo Differs
        If CBool(oldRec("LineVisible")) <> CBool(newRec("LineVisible")) Then GoTo Differs
        If CBool(newRec("LineVisible")) Then
            If CLng(oldRec("LineColor")) <> CLng(newRec("LineColor")) Then GoTo Differs
            If Abs(CDbl(oldRec("LineWeight")) - CDbl(newRec("LineWeight"))) > 0.01 Then GoTo Differs
        End If
    End If
    Exit Function

Differs:
    GanttShapeRegistry_RecordStyleDiffers = True

End Function

'------------------------------------------------------------------------------
' FR: Indique si le Registry possede le lifecycle de cette shape metier.
' EN: Returns whether the Registry owns this business shape lifecycle.
'------------------------------------------------------------------------------
Private Function GanttShapeRegistry_IsBusinessShapeName(ByVal shapeName As String) As Boolean

    GanttShapeRegistry_IsBusinessShapeName = _
        (shapeName = "TODAY_LINE") Or _
        (Left$(shapeName, 5) = "TASK_") Or _
        (Left$(shapeName, 3) = "MS_") Or _
        (Left$(shapeName, 4) = "SUM_")

End Function
