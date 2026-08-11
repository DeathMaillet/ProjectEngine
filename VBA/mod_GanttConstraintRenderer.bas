Attribute VB_Name = "mod_GanttConstraintRenderer"
Option Explicit

'===============================================================================
' MODULE : mod_GanttConstraintRenderer
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
' CONTRATS / CONTRACTS : BuildGanttConstraintMapFromCalc, DrawConstraintMarkers_Gantt
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
' FR: Gere l'affichage ou la lecture des contraintes visuelles GANTT sans modifier le moteur de calcul.
' EN: Handles GANTT visual constraint display or lookup without changing the calculation engine.
'------------------------------------------------------------------------------
Public Function BuildGanttConstraintMapFromCalc(Optional ByVal includeDeadline As Boolean = False) As Object

    Dim perfScope As clsPerfScope

    Dim d As Object
    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim activeVal As String
    Dim isSummaryVal As String
    Dim drivingLogicVal As String
    Dim taskTypeVal As String
    Dim deadlineVal As Variant
    Dim hasDeadline As Boolean

    Set perfScope = Profiler_BeginScope("BuildGanttConstraintMapFromCalc", "Excel Read")

    Set d = CreateObject("Scripting.Dictionary")

    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)

    If tblCalc.DataBodyRange Is Nothing Then
        Set BuildGanttConstraintMapFromCalc = d
        Exit Function
    End If

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    RequireGanttCalcColumn mapCalc, "ID", "BuildGanttConstraintMapFromCalc"
    RequireGanttCalcColumn mapCalc, "Constraint Active", "BuildGanttConstraintMapFromCalc"
    RequireGanttCalcColumn mapCalc, "Start Constraint Type", "BuildGanttConstraintMapFromCalc"
    RequireGanttCalcColumn mapCalc, "Start Constraint Date", "BuildGanttConstraintMapFromCalc"
    RequireGanttCalcColumn mapCalc, "Finish Constraint Type", "BuildGanttConstraintMapFromCalc"
    RequireGanttCalcColumn mapCalc, "Finish Constraint Date", "BuildGanttConstraintMapFromCalc"
    If includeDeadline Then RequireGanttCalcColumn mapCalc, "Deadline", "BuildGanttConstraintMapFromCalc"

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))
        If idVal = "" Then GoTo NextRow

        isSummaryVal = ""
        drivingLogicVal = ""
        taskTypeVal = ""

        If mapCalc.Exists("IsSummary") Then isSummaryVal = UCase$(Trim$(CStr(arr(r, mapCalc("IsSummary")))))
        If mapCalc.Exists("Driving Logic") Then drivingLogicVal = UCase$(Trim$(CStr(arr(r, mapCalc("Driving Logic")))))
        If mapCalc.Exists("Task Type") Then taskTypeVal = UCase$(Trim$(CStr(arr(r, mapCalc("Task Type")))))

        If isSummaryVal = "TRUE" Or isSummaryVal = "YES" Then GoTo NextRow
        If drivingLogicVal = "SUMMARY" Or drivingLogicVal = "LOE" Then GoTo NextRow
        If InStr(1, taskTypeVal, "LEVEL OF EFFORT", vbTextCompare) > 0 Then GoTo NextRow

        activeVal = UCase$(Trim$(CStr(arr(r, mapCalc("Constraint Active")))))
        deadlineVal = Empty
        hasDeadline = False
        If includeDeadline Then
            deadlineVal = GetCellValue(arr(r, mapCalc("Deadline")))
            hasDeadline = HasValue(deadlineVal)
        End If

        If activeVal <> "YES" And Not hasDeadline Then GoTo NextRow

        d(idVal) = Array( _
            Trim$(CStr(arr(r, mapCalc("Start Constraint Type")))), _
            GetCellValue(arr(r, mapCalc("Start Constraint Date"))), _
            Trim$(CStr(arr(r, mapCalc("Finish Constraint Type")))), _
            GetCellValue(arr(r, mapCalc("Finish Constraint Date"))), _
            deadlineVal _
        )

NextRow:
    Next r

    Set BuildGanttConstraintMapFromCalc = d

End Function
'------------------------------------------------------------------------------
' FR: Verifie qu'une source Excel expose les colonnes attendues par le rendu GANTT.
' EN: Validates that an Excel source exposes the columns expected by GANTT rendering.
'------------------------------------------------------------------------------
Private Sub RequireGanttCalcColumn( _
    ByVal mapCalc As Object, _
    ByVal colName As String, _
    ByVal functionName As String)

    If Not mapCalc.Exists(colName) Then
        Err.Raise vbObjectError + 760, functionName, "Missing source column in tbl_CALC: " & colName
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Gere l'affichage ou la lecture des contraintes visuelles GANTT sans modifier le moteur de calcul.
' EN: Handles GANTT visual constraint display or lookup without changing the calculation engine.
'------------------------------------------------------------------------------
Public Sub DrawConstraintMarkers_Gantt( _
    ByVal ws As Worksheet, _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal constraintById As Object, _
    ByVal drawHardConstraints As Boolean, _
    ByVal drawDeadlineMarkers As Boolean)

    Dim perfScope As clsPerfScope

    Dim r As Long
    Dim rowCount As Long
    Dim renderableRowCount As Long
    Dim ganttRow As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim vals As Variant
    Dim startType As String
    Dim finishType As String
    Dim startDate As Variant
    Dim finishDate As Variant
    Dim deadlineDate As Variant
    Dim isLoE As Boolean
    Dim isMilestone As Boolean

    Set perfScope = Profiler_BeginScope("DrawConstraintMarkers_Gantt", "Constraint Render")

    If constraintById Is Nothing Then Exit Sub
    If constraintById.Count = 0 Then Exit Sub
    If Not HasValue(projectStart) Then Exit Sub
    If totalDays < 1 Then Exit Sub

    rowCount = UBound(dataArr, 1)

    For r = 1 To rowCount

        ganttRow = FIRST_TASK_ROW + r - 1
        If ws.rows(ganttRow).Hidden Then GoTo NextRow

        idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))
        If idVal = "" Then GoTo NextRow
        If Not constraintById.Exists(idVal) Then GoTo NextRow

        wbsVal = NormalizeWBS(CStr(dataArr(r, mapWBS("WBS"))))
        If hasChildren.Exists(wbsVal) Then GoTo NextRow

        isLoE = TaskTypeRules_IsLevelOfEffortRow(dataArr, mapWBS, r)
        If isLoE Then GoTo NextRow

        vals = constraintById(idVal)
        startType = UCase$(Trim$(CStr(vals(0))))
        startDate = vals(1)
        finishType = UCase$(Trim$(CStr(vals(2))))
        finishDate = vals(3)
        deadlineDate = Empty
        If UBound(vals) >= 4 Then deadlineDate = vals(4)

        If drawDeadlineMarkers And HasValue(deadlineDate) Then
            ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, deadlineDate, "RIGHT"
        End If

        If Not drawHardConstraints Then GoTo NextRow

        isMilestone = TaskTypeRules_IsMilestoneRow(dataArr, mapWBS, r)

        If isMilestone _
            And startType = "MUST START ON" _
            And finishType = "MUST FINISH ON" _
            And HasValue(startDate) _
            And HasValue(finishDate) _
            And CLng(CDbl(startDate)) = CLng(CDbl(finishDate)) Then

            ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, startDate, "FULL"
            GoTo NextRow
        End If

        Select Case startType
            Case "START NO EARLIER THAN"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, startDate, "LEFT"
                DrawConstraintBigCrossAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, startDate, "LEFT", CStr(r), "SNET"
            Case "START NO LATER THAN"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, startDate, "RIGHT"
                DrawConstraintSmallCrossesAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, startDate, "RIGHT", "RIGHT", CStr(r), "SNLT"
            Case "MUST START ON"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, startDate, "LEFT"
                DrawConstraintBigCrossAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, startDate, "LEFT", CStr(r), "SNET"
                DrawConstraintSmallCrossesAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, startDate, "LEFT", "RIGHT", CStr(r), "SNLT"
        End Select

        Select Case finishType
            Case "FINISH NO EARLIER THAN"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "LEFT"
                DrawConstraintSmallCrossesAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "LEFT", "LEFT", CStr(r), "FNET"
            Case "FINISH NO LATER THAN"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "RIGHT"
                DrawConstraintBigCrossAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "RIGHT", CStr(r), "FNLT"
            Case "MUST FINISH ON"
                ApplyConstraintCellBorder_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "RIGHT"
                DrawConstraintSmallCrossesAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "RIGHT", "LEFT", CStr(r), "FNET"
                DrawConstraintBigCrossAtCellEdge_Gantt ws, ganttRow, projectStart, totalDays, finishDate, "RIGHT", CStr(r), "FNLT"
        End Select

NextRow:
    Next r

End Sub
'------------------------------------------------------------------------------
' FR: Gere l'affichage ou la lecture des contraintes visuelles GANTT sans modifier le moteur de calcul.
' EN: Handles GANTT visual constraint display or lookup without changing the calculation engine.
'------------------------------------------------------------------------------
Private Function GetConstraintTargetCell_Gantt( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal constraintDate As Variant) As Range

    Dim renderDate As Variant
    Dim projectFinish As Date
    Dim targetCol As Long

    If Not HasValue(constraintDate) Then Exit Function

    renderDate = GetRenderStartForCurrentScale(constraintDate)
    If Not HasValue(renderDate) Then Exit Function

    projectFinish = GetScaleProjectFinishFromSlots(projectStart, totalDays)

    If CDate(renderDate) < CDate(projectStart) Or CDate(renderDate) > projectFinish Then Exit Function

    targetCol = TimelineColumnFromHeaderDate_Exact(ws, projectStart, renderDate)
    If targetCol < FIRST_TIMELINE_COL Then Exit Function

    Set GetConstraintTargetCell_Gantt = ws.cells(ganttRow, targetCol)

End Function
'------------------------------------------------------------------------------
' FR: Gere l'affichage ou la lecture des contraintes visuelles GANTT sans modifier le moteur de calcul.
' EN: Handles GANTT visual constraint display or lookup without changing the calculation engine.
'------------------------------------------------------------------------------
Private Sub ApplyConstraintCellBorder_Gantt( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal constraintDate As Variant, _
    ByVal borderMode As String)

    Dim targetCell As Range

    Set targetCell = GetConstraintTargetCell_Gantt(ws, ganttRow, projectStart, totalDays, constraintDate)
    If targetCell Is Nothing Then Exit Sub

    Select Case UCase$(borderMode)
        Case "LEFT"
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeLeft)
        Case "RIGHT"
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeRight)
        Case "FULL"
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeLeft)
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeRight)
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeTop)
            ApplyConstraintOneBorder_Gantt targetCell.Borders(xlEdgeBottom)
    End Select

End Sub
'------------------------------------------------------------------------------
' FR: Gere l'affichage ou la lecture des contraintes visuelles GANTT sans modifier le moteur de calcul.
' EN: Handles GANTT visual constraint display or lookup without changing the calculation engine.
'------------------------------------------------------------------------------
Private Sub ApplyConstraintOneBorder_Gantt(ByVal borderObj As Border)

    With borderObj
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(220, 0, 0)
    End With

End Sub
'------------------------------------------------------------------------------
' FR: Gere l'affichage ou la lecture des contraintes visuelles GANTT sans modifier le moteur de calcul.
' EN: Handles GANTT visual constraint display or lookup without changing the calculation engine.
'------------------------------------------------------------------------------
Private Sub DrawConstraintBigCrossAtCellEdge_Gantt( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal constraintDate As Variant, _
    ByVal edgeSide As String, _
    ByVal shapeSuffix As String, _
    ByVal token As String)

    Dim targetCell As Range
    Dim x As Double
    Dim y As Double
    Dim sizeVal As Double
    Dim offsetVal As Double

    Set targetCell = GetConstraintTargetCell_Gantt(ws, ganttRow, projectStart, totalDays, constraintDate)
    If targetCell Is Nothing Then Exit Sub

    sizeVal = 8
    offsetVal = 5

    If UCase$(edgeSide) = "RIGHT" Then
        x = targetCell.Left + targetCell.Width + offsetVal
    Else
        x = targetCell.Left - offsetVal
    End If

    y = GetGanttRowTop(ws, ganttRow) + (GetGanttRowHeight(ws, ganttRow) / 2)

    DrawConstraintCross_Gantt ws, x, y, sizeVal, "CSTR_" & shapeSuffix & "_" & token & "_", 1

End Sub
'------------------------------------------------------------------------------
' FR: Gere l'affichage ou la lecture des contraintes visuelles GANTT sans modifier le moteur de calcul.
' EN: Handles GANTT visual constraint display or lookup without changing the calculation engine.
'------------------------------------------------------------------------------
Private Sub DrawConstraintSmallCrossesAtCellEdge_Gantt( _
    ByVal ws As Worksheet, _
    ByVal ganttRow As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal constraintDate As Variant, _
    ByVal wallEdge As String, _
    ByVal crossSide As String, _
    ByVal shapeSuffix As String, _
    ByVal token As String)

    Dim targetCell As Range
    Dim wallX As Double
    Dim x As Double
    Dim yTop As Double
    Dim yBottom As Double
    Dim sizeVal As Double
    Dim offsetVal As Double
    Dim rowTop As Double
    Dim rowHeight As Double
    Dim verticalShift As Double
    Dim minY As Double
    Dim maxY As Double

    Set targetCell = GetConstraintTargetCell_Gantt(ws, ganttRow, projectStart, totalDays, constraintDate)
    If targetCell Is Nothing Then Exit Sub

    If UCase$(wallEdge) = "RIGHT" Then
        wallX = targetCell.Left + targetCell.Width
    Else
        wallX = targetCell.Left
    End If

    sizeVal = 5
    offsetVal = 4

    If UCase$(crossSide) = "RIGHT" Then
        x = wallX + offsetVal
    Else
        x = wallX - offsetVal
    End If

    rowTop = GetGanttRowTop(ws, ganttRow)
    rowHeight = GetGanttRowHeight(ws, ganttRow)
    verticalShift = rowHeight * 0.1
    minY = rowTop + (sizeVal / 2)
    maxY = rowTop + rowHeight - (sizeVal / 2)

    yTop = rowTop + (sizeVal / 2) + 0.5 - verticalShift
    If yTop < minY Then yTop = minY

    yBottom = rowTop + rowHeight - (sizeVal / 2) - 0.5 + verticalShift
    If yBottom > maxY Then yBottom = maxY

    DrawConstraintCross_Gantt ws, x, yTop, sizeVal, "CSTR_" & shapeSuffix & "_" & token & "_", 1
    DrawConstraintCross_Gantt ws, x, yBottom, sizeVal, "CSTR_" & shapeSuffix & "_" & token & "_", 3

End Sub
'------------------------------------------------------------------------------
' FR: Gere l'affichage ou la lecture des contraintes visuelles GANTT sans modifier le moteur de calcul.
' EN: Handles GANTT visual constraint display or lookup without changing the calculation engine.
'------------------------------------------------------------------------------
Private Sub DrawConstraintCross_Gantt( _
    ByVal ws As Worksheet, _
    ByVal centerX As Double, _
    ByVal centerY As Double, _
    ByVal sizeVal As Double, _
    ByVal shapeBase As String, _
    ByVal startIndex As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeMathMultiply, centerX - (sizeVal / 2), centerY - (sizeVal / 2), sizeVal, sizeVal)
    shp.Name = shapeBase & CStr(startIndex)
    ApplyGanttRenderShapePlacement shp

    With shp
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(220, 0, 0)
        .Fill.Transparency = 0
        .Line.Visible = msoFalse
    End With

    shp.ZOrder msoBringToFront

End Sub
