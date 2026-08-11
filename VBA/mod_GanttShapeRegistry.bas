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

    targetCol = TimelineColumnFromHeaderDate_Exact(ws, projectStart, startVal)
    If targetCol < FIRST_TIMELINE_COL Then Exit Sub

    cellWidth = ws.Cells(HEADER_ROW_2, targetCol).Width
    sizeVal = GetGanttCompactTaskMarkerSize(ws, ganttRow, cellWidth)
    If sizeVal < 2 Then sizeVal = 2

    markerCenterX = TimelineDateRangeMidX(ws, projectStart, startVal, finishVal)
    If markerCenterX <= 0 Then Exit Sub

    topPos = ws.Cells(ganttRow, FIRST_TIMELINE_COL).Top + ((ws.Rows(ganttRow).Height - sizeVal) / 2)

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
    yTop = ws.Cells(HEADER_ROW_1, FIRST_TIMELINE_COL).Top
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

    For Each key In records.Keys
        GanttShapeRegistry_CreateShapeFromRecord ws, records(CStr(key))
    Next key

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

    If CBool(rec("IsLine")) Then
        shp.Left = CDbl(rec("Left"))
        shp.Top = CDbl(rec("Top"))
        shp.Width = CDbl(rec("Width"))
        shp.Height = CDbl(rec("Height"))
        With shp.Line
            .ForeColor.RGB = CLng(rec("LineColor"))
            .Weight = CDbl(rec("LineWeight"))
            .DashStyle = CLng(rec("DashStyle"))
        End With
        If CBool(rec("ZFront")) Then shp.ZOrder msoBringToFront
    Else
        shp.Left = CDbl(rec("Left"))
        shp.Top = CDbl(rec("Top"))
        shp.Width = CDbl(rec("Width"))
        shp.Height = CDbl(rec("Height"))
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

    shp.Visible = IIf(CBool(rec("Visible")), msoTrue, msoFalse)

End Sub
'------------------------------------------------------------------------------
' FR: Participe au registre de shapes utilise pour creer, comparer ou mettre a jour le rendu predictif GANTT.
' EN: Participates in the shape registry used to create, compare, or update predictive GANTT rendering.
'------------------------------------------------------------------------------
Public Function GanttShapeRegistry_ShapeDiffers(ByVal shp As Shape, ByVal rec As Object) As Boolean

    On Error GoTo Differs

    If GanttShapeRegistry_TypeMismatch(shp, rec) Then GoTo Differs
    If Abs(shp.Left - CDbl(rec("Left"))) > 0.1 Then GoTo Differs
    If Abs(shp.Top - CDbl(rec("Top"))) > 0.1 Then GoTo Differs
    If Abs(shp.Width - CDbl(rec("Width"))) > 0.1 Then GoTo Differs
    If Abs(shp.Height - CDbl(rec("Height"))) > 0.1 Then GoTo Differs
    If (shp.Visible = msoTrue) <> CBool(rec("Visible")) Then GoTo Differs

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

    GanttShapeRegistry_ShapeDiffers = False
    Exit Function

Differs:
    GanttShapeRegistry_ShapeDiffers = True

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
        GanttShapeRegistry_TypeMismatch = (shp.AutoShapeType <> CLng(rec("AutoShapeType")))
    End If
    Exit Function

Mismatch:
    GanttShapeRegistry_TypeMismatch = True

End Function
