Attribute VB_Name = "mod_GanttSheetLayout"
Option Explicit

'===============================================================================
' MODULE : mod_GanttSheetLayout
' DOMAINE / DOMAIN : Gantt
'
' FR
' Prepare, nettoie et finalise la surface de feuille Gantt sans dessiner les objets metier.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Prepares, clears and finalizes the Gantt sheet surface without drawing business objects.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Gantt_SafeEmptyState, PrepareGanttDisplayOnlyLayout, PrepareGanttFullLayout, Gantt_HasPendingGeometryRepair, Gantt_PendingGeometryRepairCount, Gantt_RepairPendingGeometryIfNeeded, Gantt_CapturePendingGeometryRepair, Gantt_ClearPendingGeometryRepair
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"

Private Const FIRST_TIMELINE_COL As Long = 11
Private Const TITLE_ROW As Long = 1
Private Const TOGGLE_ROW_TOP As Long = 2
Private Const TOGGLE_ROW_BOTTOM As Long = 3
Private Const HEADER_ROW_1 As Long = 3
Private Const HEADER_ROW_2 As Long = 4
Private Const FIRST_TASK_ROW As Long = 5

Private Const COL_WBS As Long = 1
Private Const COL_TASK As Long = 2
Private Const COL_START As Long = 3
Private Const COL_FINISH As Long = 4
Private Const COL_TEST_START As Long = 5
Private Const COL_TEST_FINISH As Long = 6
Private Const COL_DURATION As Long = 7
Private Const COL_PROGRESS As Long = 8
Private Const COL_TEST_PROGRESS As Long = 9
Private Const COL_LOGIC As Long = 10

Private Const TEST_START_HEADER As String = "Test Start"
Private Const TEST_FINISH_HEADER As String = "Test Finish"
Private Const TEST_PROGRESS_HEADER As String = "Test %"

Private Const GANTT_ROW_HEIGHT_HEADER_1 As Double = 21
Private Const GANTT_ROW_HEIGHT_HEADER_2 As Double = 18
Private Const GANTT_ROW_HEIGHT_TASK As Double = 18

Private Const GANTT_VIEW_DETAIL As String = "DETAIL"
Private Const GANTT_VIEW_SUMMARY As String = "SUMMARY"
Private Const GANTT_SCALE_DAY As String = "DAY"
Private Const GANTT_SCALE_WEEK As String = "WEEK"
Private Const GANTT_SCALE_MONTH As String = "MONTH"

Private gPendingGanttGeometryRepair As Boolean
Private gPendingGanttGeometrySnapshot As Object

'------------------------------------------------------------------------------
' FR: Nettoie, restaure ou normalise une partie de l'etat visuel GANTT.
' EN: Cleans, restores, or normalizes part of the GANTT visual state.
'------------------------------------------------------------------------------
Private Sub RestoreGanttTestInputs(ByVal ws As Worksheet, ByVal testInputMap As Object)

    Dim perfScope As clsPerfScope

    Dim lastRow As Long
    Dim r As Long
    Dim wbsVal As String
    Dim savedVals As Variant

    Set perfScope = Profiler_BeginScope("RestoreGanttTestInputs", "Excel Cell Write")

    If testInputMap Is Nothing Then Exit Sub

    lastRow = GetLastGanttRow(ws)
    If lastRow < FIRST_TASK_ROW Then Exit Sub

    For r = FIRST_TASK_ROW To lastRow

        wbsVal = NormalizeWBS(CStr(ws.cells(r, COL_WBS).value))

        If wbsVal <> "" Then
            If testInputMap.Exists(wbsVal) Then
                savedVals = testInputMap(wbsVal)

                ws.cells(r, COL_TEST_START).value = savedVals(0)
                ws.cells(r, COL_TEST_FINISH).value = savedVals(1)
                ws.cells(r, COL_TEST_PROGRESS).value = savedVals(2)

                ws.cells(r, COL_TEST_START).NumberFormat = "dd/mm/yyyy"
                ws.cells(r, COL_TEST_FINISH).NumberFormat = "dd/mm/yyyy"
                ws.cells(r, COL_TEST_PROGRESS).NumberFormat = "0%"
            End If
        End If

    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Safe Empty State dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Safe Empty State helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub Gantt_SafeEmptyState()

    Dim wsGantt As Worksheet
    Dim wasGanttSheetCreated As Boolean
    Dim oldInternalWrite As Boolean
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean

    On Error GoTo SafeExit

    oldInternalWrite = GetGanttInternalWrite()
    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    SetGanttInternalWrite True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    EnsureGanttViewInitialized
    Set wsGantt = EnsureGanttSheet(wasGanttSheetCreated)

    ClearGanttSheet wsGantt
    SetupStaticLayout wsGantt
    If wasGanttSheetCreated Then SetupLeftPanelDefaults wsGantt
    Ensure_Gantt_Test_Buttons

    GanttDependency_ClearExpandedLinksCache
    GanttLive_SafeEmptyState

SafeExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    SetGanttInternalWrite oldInternalWrite

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Prepare Gantt Display Only Layout dans le workflow de rendu GANTT.
' EN: Runs the Prepare Gantt Display Only Layout helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub PrepareGanttDisplayOnlyLayout( _
    ByVal wsGantt As Worksheet, _
    ByVal rowCount As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal testInputMap As Object)

    If rowCount > 0 Then
        wsGantt.rows(FIRST_TASK_ROW & ":" & FIRST_TASK_ROW + rowCount - 1).Hidden = False
    End If

    ClearGanttRightPaneOnly wsGantt
    SetupTimelineLayout wsGantt, projectStart, totalDays

    If GetGanttPreserveTestInputs() Then
        RestoreGanttTestInputs wsGantt, testInputMap
    End If

    FinalizeGanttSheet_DisplayOnly wsGantt, totalDays, rowCount

End Sub

'------------------------------------------------------------------------------
' FR:
' Reconstruit le layout complet: efface la feuille, pose l'entete statique,
' genere la timeline, ecrit le panneau gauche et formate chaque ligne.
'
' EN:
' Rebuilds the full layout: clears the sheet, lays out static headers,
' builds the timeline, writes the left pane, and formats every row.
'
' Entrees / Inputs:
' - Feuille GANTT, data WBS, maps, projectStart, slots timeline et inputs TEST sauvegardes.
'
' Sorties / Outputs:
' - Panneau gauche, colonnes TEST, timeline et freeze panes prets pour le dessin shapes.
'
' Appele par / Called by:
' - RunGanttRefreshCore en mode refresh complet.
'
' Notes:
' - Ne dessine pas les barres; prepare uniquement la grille et les cellules.
'------------------------------------------------------------------------------
Public Sub PrepareGanttFullLayout( _
    ByVal wsGantt As Worksheet, _
    ByVal dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal calcDrivingMap As Object, _
    ByVal rowCount As Long, _
    ByVal projectStart As Variant, _
    ByVal totalDays As Long, _
    ByVal testInputMap As Object, _
    ByVal isNewSheet As Boolean, _
    ByVal activateGantt As Boolean)

    Dim perfScope As clsPerfScope

    Dim ganttRow As Long
    Dim r As Long
    Dim applyLeftPanelDefaults As Boolean

    Set perfScope = Profiler_BeginScope("PrepareGanttFullLayout", "Gantt Layout")

    applyLeftPanelDefaults = (isNewSheet Or IsGanttSheetLayoutEmpty(wsGantt))

    ClearGanttSheet wsGantt
    SetupStaticLayout wsGantt
    If applyLeftPanelDefaults Then SetupLeftPanelDefaults wsGantt
    SetupTimelineLayout wsGantt, projectStart, totalDays

    ganttRow = FIRST_TASK_ROW
    For r = 1 To rowCount
        WriteLeftPanelRow wsGantt, ganttRow, dataArr, r, mapWBS
        ApplyRowStyle wsGantt, ganttRow, dataArr, r, mapWBS, hasChildren, calcDrivingMap
        ganttRow = ganttRow + 1
    Next r

    If GetGanttPreserveTestInputs() Then
        RestoreGanttTestInputs wsGantt, testInputMap
    End If

    FinalizeGanttSheet wsGantt, totalDays, rowCount

    If isNewSheet And activateGantt Then
        FreezeGanttAfterFinish wsGantt, rowCount
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Has Pending Geometry Repair dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Has Pending Geometry Repair helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function Gantt_HasPendingGeometryRepair() As Boolean

    Gantt_HasPendingGeometryRepair = gPendingGanttGeometryRepair

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Pending Geometry Repair Count dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Pending Geometry Repair Count helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function Gantt_PendingGeometryRepairCount() As Long

    If gPendingGanttGeometrySnapshot Is Nothing Then Exit Function
    Gantt_PendingGeometryRepairCount = gPendingGanttGeometrySnapshot.Count

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Repair Pending Geometry If Needed dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Repair Pending Geometry If Needed helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub Gantt_RepairPendingGeometryIfNeeded()

    Dim ws As Worksheet
    Dim shapeName As Variant
    Dim geom As Variant
    Dim shp As Shape
    Dim oldScreenUpdating As Boolean
    Dim oldInternalWrite As Boolean

    If Not gPendingGanttGeometryRepair Then Exit Sub
    If gPendingGanttGeometrySnapshot Is Nothing Then
        gPendingGanttGeometryRepair = False
        Exit Sub
    End If

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)
    oldScreenUpdating = Application.ScreenUpdating
    oldInternalWrite = GetGanttInternalWrite()

    Application.ScreenUpdating = False
    SetGanttInternalWrite True

    For Each shapeName In gPendingGanttGeometrySnapshot.Keys
        Set shp = Nothing
        On Error Resume Next
        Set shp = ws.Shapes(CStr(shapeName))
        On Error GoTo SafeExit

        If Not shp Is Nothing Then
            geom = gPendingGanttGeometrySnapshot(CStr(shapeName))
            shp.Left = CDbl(geom(0))
            shp.Top = CDbl(geom(1))
            shp.Width = CDbl(geom(2))
            shp.Height = CDbl(geom(3))
            shp.Visible = CLng(geom(4))
        End If
    Next shapeName

SafeExit:
    SetGanttInternalWrite oldInternalWrite
    Application.ScreenUpdating = oldScreenUpdating
    Gantt_ClearPendingGeometryRepair

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Capture Pending Geometry Repair dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Capture Pending Geometry Repair helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub Gantt_CapturePendingGeometryRepair(ByVal ws As Worksheet)

    Dim shp As Shape
    Dim shapeName As String
    Dim snapshot As Object

    If ws Is Nothing Then Exit Sub

    Set snapshot = CreateObject("Scripting.Dictionary")

    For Each shp In ws.Shapes
        shapeName = CStr(shp.Name)
        If IsGanttGeometryRepairShape(shapeName) Then
            snapshot(shapeName) = Array(CDbl(shp.Left), CDbl(shp.Top), CDbl(shp.Width), CDbl(shp.Height), CLng(shp.Visible))
        End If
    Next shp

    If snapshot.Count > 0 Then
        Set gPendingGanttGeometrySnapshot = snapshot
        gPendingGanttGeometryRepair = True
    Else
        Gantt_ClearPendingGeometryRepair
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Clear Pending Geometry Repair dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Clear Pending Geometry Repair helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub Gantt_ClearPendingGeometryRepair()

    gPendingGanttGeometryRepair = False
    Set gPendingGanttGeometrySnapshot = Nothing

End Sub

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Public Function IsGanttGeometryRepairShape(ByVal shapeName As String) As Boolean

    IsGanttGeometryRepairShape = _
           shapeName = "TODAY_LINE" _
        Or Left$(shapeName, 5) = "TASK_" _
        Or Left$(shapeName, 3) = "MS_" _
        Or Left$(shapeName, 4) = "SUM_" _
        Or Left$(shapeName, 4) = "DEP_" _
        Or Left$(shapeName, 5) = "CSTR_"

End Function

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function IsManagedGanttShape(ByVal shapeName As String) As Boolean

    IsManagedGanttShape = _
           shapeName = "TODAY_LINE" _
        Or shapeName = "SCURVE_TODAY_LINE" _
        Or Left$(shapeName, 10) = "GANTT_BAR_" _
        Or Left$(shapeName, 13) = "GANTT_LINK_H_" _
        Or Left$(shapeName, 13) = "GANTT_LINK_V_" _
        Or Left$(shapeName, 12) = "GANTT_CP_BAR_" _
        Or Left$(shapeName, 15) = "GANTT_TEST_BOX_" _
        Or Left$(shapeName, 21) = "GANTT_PROGRESS_TEXT_" _
        Or Left$(shapeName, 4) = "SUM_" _
        Or Left$(shapeName, 3) = "MS_" _
        Or Left$(shapeName, 5) = "TASK_" _
        Or Left$(shapeName, 4) = "DEP_" _
        Or Left$(shapeName, 5) = "CSTR_"

End Function

'------------------------------------------------------------------------------
' FR: Nettoie, restaure ou normalise une partie de l'etat visuel GANTT.
' EN: Cleans, restores, or normalizes part of the GANTT visual state.
'------------------------------------------------------------------------------
Private Sub DeleteManagedGanttShapes(ByVal wsGantt As Worksheet)

    Dim perfScope As clsPerfScope

    Dim shp As Shape
    Dim i As Long

    Set perfScope = Profiler_BeginScope("DeleteManagedGanttShapes", "Shape Delete")

    For i = wsGantt.Shapes.Count To 1 Step -1
        Set shp = wsGantt.Shapes(i)
        If IsManagedGanttShape(shp.Name) Then
            shp.Delete
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Nettoie, restaure ou normalise une partie de l'etat visuel GANTT.
' EN: Cleans, restores, or normalizes part of the GANTT visual state.
'------------------------------------------------------------------------------
Private Sub ClearGanttRightPaneOnly(ByVal wsGantt As Worksheet)

    Dim perfScope As clsPerfScope

    Dim lastCol As Long

    Set perfScope = Profiler_BeginScope("ClearGanttRightPaneOnly", "Excel Clear")

    lastCol = wsGantt.cells(HEADER_ROW_2, wsGantt.Columns.Count).End(xlToLeft).Column
    If lastCol < FIRST_TIMELINE_COL Then lastCol = FIRST_TIMELINE_COL

    On Error Resume Next
    wsGantt.Range( _
        wsGantt.cells(HEADER_ROW_1, FIRST_TIMELINE_COL), _
        wsGantt.cells(HEADER_ROW_2, lastCol) _
    ).UnMerge
    On Error GoTo 0

    wsGantt.Range( _
        wsGantt.cells(HEADER_ROW_1, FIRST_TIMELINE_COL), _
        wsGantt.cells(wsGantt.rows.Count, lastCol) _
    ).ClearContents

    wsGantt.Range( _
        wsGantt.cells(HEADER_ROW_1, FIRST_TIMELINE_COL), _
        wsGantt.cells(wsGantt.rows.Count, lastCol) _
    ).Interior.Pattern = xlNone

    wsGantt.Range( _
        wsGantt.cells(HEADER_ROW_1, FIRST_TIMELINE_COL), _
        wsGantt.cells(wsGantt.rows.Count, lastCol) _
    ).Borders.LineStyle = xlNone

    DeleteManagedGanttShapes wsGantt

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Finalize Gantt Sheet dans le workflow de rendu GANTT.
' EN: Runs the Finalize Gantt Sheet helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub FinalizeGanttSheet(ByVal ws As Worksheet, ByVal totalDays As Long, ByVal rowCount As Long)

    Dim perfScope As clsPerfScope

    Dim lastCol As Long
    Dim lastRow As Long

    Set perfScope = Profiler_BeginScope("FinalizeGanttSheet", "Gantt Layout")

    lastCol = FIRST_TIMELINE_COL + totalDays - 1
    lastRow = FIRST_TASK_ROW + rowCount - 1

    ws.Range(ws.cells(FIRST_TASK_ROW, FIRST_TIMELINE_COL), ws.cells(lastRow, lastCol)).Borders(xlInsideVertical).LineStyle = xlDot
    ws.Range(ws.cells(FIRST_TASK_ROW, FIRST_TIMELINE_COL), ws.cells(lastRow, lastCol)).Borders(xlInsideHorizontal).LineStyle = xlDot

    ws.Range(ws.cells(FIRST_TASK_ROW, 1), ws.cells(lastRow, COL_LOGIC)).Borders.LineStyle = xlContinuous

End Sub

'------------------------------------------------------------------------------
' FR: Verifie et prepare une ressource GANTT requise avant le rendu ou l'interaction.
' EN: Ensures and prepares a GANTT resource required before rendering or interaction.
'------------------------------------------------------------------------------
Public Function EnsureGanttSheet(Optional ByRef isNewSheet As Boolean = False) As Worksheet

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("EnsureGanttSheet", "Gantt Layout")

    On Error Resume Next
    Set EnsureGanttSheet = ThisWorkbook.Worksheets(GANTT_SHEET)
    On Error GoTo 0

    If EnsureGanttSheet Is Nothing Then
        Set EnsureGanttSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureGanttSheet.Name = GANTT_SHEET
        isNewSheet = True
    Else
        isNewSheet = False
    End If

End Function

'------------------------------------------------------------------------------
' FR: Nettoie, restaure ou normalise une partie de l'etat visuel GANTT.
' EN: Cleans, restores, or normalizes part of the GANTT visual state.
'------------------------------------------------------------------------------
Private Sub ClearGanttSheet(ByVal ws As Worksheet)

    Dim perfScope As clsPerfScope

    Dim shp As Shape

    Set perfScope = Profiler_BeginScope("ClearGanttSheet", "Excel Clear")

    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ws.cells.Clear

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Setup Static Layout dans le workflow de rendu GANTT.
' EN: Runs the Setup Static Layout helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub SetupStaticLayout(ByVal ws As Worksheet)

    Dim perfScope As clsPerfScope


    Set perfScope = Profiler_BeginScope("SetupStaticLayout", "Gantt Layout")

    ws.cells(TITLE_ROW, COL_WBS).value = "GANTT VIEW"

    ws.Range("A" & HEADER_ROW_2).value = "WBS"
    ws.Range("B" & HEADER_ROW_2).value = "Task Name"
    ws.Range("C" & HEADER_ROW_2).value = "Start"
    ws.Range("D" & HEADER_ROW_2).value = "Finish"
    ws.Range("E" & HEADER_ROW_2).value = TEST_START_HEADER
    ws.Range("F" & HEADER_ROW_2).value = TEST_FINISH_HEADER
    ws.Range("G" & HEADER_ROW_2).value = "Duration"
    ws.Range("H" & HEADER_ROW_2).value = "%"
    ws.Range("I" & HEADER_ROW_2).value = TEST_PROGRESS_HEADER
    ws.Range("J" & HEADER_ROW_2).value = "Logic"

    ws.Range(ws.cells(TITLE_ROW, COL_WBS), ws.cells(TITLE_ROW, COL_LOGIC)).Font.Bold = True
    ws.Range("A" & HEADER_ROW_2 & ":J" & HEADER_ROW_2).Font.Bold = True

    'Pixel-stable heights.
    'Avoid 20 pt task rows: it can create cumulative visual drift between grid and shapes.
    ws.rows(TITLE_ROW).rowHeight = GANTT_ROW_HEIGHT_HEADER_1
    ws.rows(TOGGLE_ROW_TOP).rowHeight = GANTT_ROW_HEIGHT_HEADER_2
    ws.rows(TOGGLE_ROW_BOTTOM).rowHeight = GANTT_ROW_HEIGHT_HEADER_2
    ws.rows(HEADER_ROW_2).rowHeight = GANTT_ROW_HEIGHT_HEADER_2


    ws.Range("A1:FU3").Interior.Color = RGB(255, 255, 255)
    ws.Range("A1:FU3").Borders.LineStyle = xlNone

    ws.Range("A" & HEADER_ROW_2 & ":J" & HEADER_ROW_2).Interior.Color = RGB(217, 217, 217)
    ws.Range("A" & HEADER_ROW_2 & ":J" & HEADER_ROW_2).Borders.LineStyle = xlContinuous

    Gantt_ApplyLanguage

End Sub

'------------------------------------------------------------------------------
' FR: Applique les largeurs initiales du panneau gauche lors d'une creation ou migration.
' EN: Applies initial left-panel widths during creation or migration.
'------------------------------------------------------------------------------
Private Sub SetupLeftPanelDefaults(ByVal ws As Worksheet)

    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:F").ColumnWidth = 11
    ws.Columns("G").ColumnWidth = 9
    ws.Columns("H:I").ColumnWidth = 8

    'Logic column calibrated to avoid text overflow into the timeline at 55% zoom.
    'Measured target: approx. 109 px / Excel width 14.86.
    ws.Columns("J").ColumnWidth = 14.86

End Sub

'------------------------------------------------------------------------------
' FR:
' Route la construction de timeline vers Day, Week ou Month selon le toggle actif.
'
' EN:
' Routes timeline construction to Day, Week, or Month according to the active toggle.
'
' Entrees / Inputs:
' - Feuille GANTT, date de depart projet, nombre de slots.
'
' Sorties / Outputs:
' - Entetes timeline et largeurs de colonnes construits par le renderer specialise.
'
' Appele par / Called by:
' - PrepareGanttFullLayout et PrepareGanttDisplayOnlyLayout.
'
' Notes:
' - Point de bascule critique pour les modes agreges; impacts shapes, liens et registry.
'------------------------------------------------------------------------------
Private Sub SetupTimelineLayout(ByVal ws As Worksheet, ByVal projectStart As Variant, ByVal slotCount As Long)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("SetupTimelineLayout", "Timeline")

    Select Case GetGanttTimelineScaleMode()
        Case GANTT_SCALE_WEEK
            BuildTimeline_Week ws, projectStart, slotCount
        Case GANTT_SCALE_MONTH
            BuildTimeline_Month ws, projectStart, slotCount
        Case Else
            BuildTimeline_Day ws, projectStart, slotCount
    End Select

End Sub

'------------------------------------------------------------------------------
' FR:
' Construit la timeline journaliere: dates en ligne 4, mois fusionnes en ligne 3,
' largeurs jour et surlignage week-end.
'
' EN:
' Builds the daily timeline: dates on row 4, merged month headers on row 3,
' day widths, and weekend highlighting.
'
' Entrees / Inputs:
' - Date de depart projet et nombre de jours.
'
' Sorties / Outputs:
' - Entetes jour/mois, formats et colonnes de rendu Day.
'
' Appele par / Called by:
' - SetupTimelineLayout.
'
' Notes:
' - Seul mode eligible au registry predictif TEST Day.
'------------------------------------------------------------------------------
Private Sub BuildTimeline_Day(ByVal ws As Worksheet, ByVal projectStart As Variant, ByVal totalDays As Long)

    Dim perfScope As clsPerfScope

    Dim i As Long
    Dim currentDate As Date
    Dim currentMonth As Long
    Dim monthStartCol As Long
    Dim lastTimelineCol As Long
    Dim dateValues As Variant
    Dim dateRange As Range
    Dim weekendEndIndex As Long
    Dim timelineValueWrites As Long
    Dim timelineFormatWrites As Long
    Dim timelineWidthWrites As Long
    Dim weekendFormatWrites As Long
    Dim monthHeaderWrites As Long

    Set perfScope = Profiler_BeginScope("BuildTimeline_Day", "Timeline")

    If totalDays < 1 Then Exit Sub

    currentMonth = 0
    monthStartCol = FIRST_TIMELINE_COL
    lastTimelineCol = FIRST_TIMELINE_COL + totalDays - 1

    ReDim dateValues(1 To 1, 1 To totalDays)

    For i = 0 To totalDays - 1
        dateValues(1, i + 1) = CDate(CDbl(projectStart) + i)
    Next i

    Set dateRange = ws.Range(ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL), ws.cells(HEADER_ROW_2, lastTimelineCol))
    dateRange.value = dateValues
    timelineValueWrites = timelineValueWrites + 1

    With dateRange
        .NumberFormat = "dd"
        .HorizontalAlignment = xlCenter
    End With
    timelineFormatWrites = timelineFormatWrites + 2

    dateRange.EntireColumn.ColumnWidth = 4
    timelineWidthWrites = timelineWidthWrites + 1

    For i = 0 To totalDays - 1
        currentDate = CDate(CDbl(projectStart) + i)

        If Month(currentDate) <> currentMonth Then
            If i > 0 Then
                With ws.Range(ws.cells(HEADER_ROW_1, monthStartCol), ws.cells(HEADER_ROW_1, FIRST_TIMELINE_COL + i - 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .NumberFormat = "@"
                    .value = Gantt_FormatMonthYear(DateSerial(Year(CDate(CDbl(projectStart) + i - 1)), Month(CDate(CDbl(projectStart) + i - 1)), 1))
                    .Interior.Color = RGB(191, 191, 191)
                End With
                monthHeaderWrites = monthHeaderWrites + 1
            End If

            currentMonth = Month(currentDate)
            monthStartCol = FIRST_TIMELINE_COL + i
        End If

        If Weekday(currentDate, vbMonday) = 6 Then
            weekendEndIndex = i
            If i + 1 <= totalDays - 1 Then weekendEndIndex = i + 1
            ws.Range(ws.cells(FIRST_TASK_ROW, FIRST_TIMELINE_COL + i), ws.cells(FIRST_TASK_ROW + 500, FIRST_TIMELINE_COL + weekendEndIndex)).Interior.Color = RGB(242, 242, 242)
            weekendFormatWrites = weekendFormatWrites + 1
        End If
    Next i

    With ws.Range(ws.cells(HEADER_ROW_1, monthStartCol), ws.cells(HEADER_ROW_1, lastTimelineCol))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "@"
        .value = Gantt_FormatMonthYear(DateSerial(Year(CDate(CDbl(projectStart) + totalDays - 1)), Month(CDate(CDbl(projectStart) + totalDays - 1)), 1))
        .Interior.Color = RGB(191, 191, 191)
    End With
    monthHeaderWrites = monthHeaderWrites + 1

    ws.Range(ws.cells(HEADER_ROW_1, FIRST_TIMELINE_COL), ws.cells(HEADER_ROW_2, lastTimelineCol)).Borders.LineStyle = xlContinuous
    timelineFormatWrites = timelineFormatWrites + 1

    Profiler_RecordOperation "BuildTimelineDayValueBlockWrites", timelineValueWrites, 0#
    Profiler_RecordOperation "BuildTimelineDayFormatWrites", timelineFormatWrites, 0#
    Profiler_RecordOperation "BuildTimelineDayWidthWrites", timelineWidthWrites, 0#
    Profiler_RecordOperation "BuildTimelineDayWeekendFormatWrites", weekendFormatWrites, 0#
    Profiler_RecordOperation "BuildTimelineDayMonthHeaderWrites", monthHeaderWrites, 0#

End Sub

'------------------------------------------------------------------------------
' FR:
' Construit la timeline hebdomadaire ISO avec libelles semaine et entetes annee fusionnes.
'
' EN:
' Builds the ISO weekly timeline with week labels and merged year headers.
'
' Entrees / Inputs:
' - Lundi ISO de depart et nombre de semaines.
'
' Sorties / Outputs:
' - Colonnes Week compactes pour rendu agrege.
'
' Appele par / Called by:
' - SetupTimelineLayout.
'
' Notes:
' - Les dependency links sont desactives en mode agrege.
'------------------------------------------------------------------------------
Private Sub BuildTimeline_Week(ByVal ws As Worksheet, ByVal projectStart As Variant, ByVal slotCount As Long)

    Dim perfScope As clsPerfScope

    Dim i As Long
    Dim weekStart As Date
    Dim isoYear As Long
    Dim currentYear As Long
    Dim yearStartCol As Long

    Set perfScope = Profiler_BeginScope("BuildTimeline_Week", "Timeline")

    currentYear = 0
    yearStartCol = FIRST_TIMELINE_COL

    For i = 0 To slotCount - 1

        weekStart = CDate(projectStart) + (i * 7)
        isoYear = GetIsoWeekYear(weekStart)

        ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).NumberFormat = "@"
        ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).value = GetIsoWeekLabel(weekStart)
        ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).HorizontalAlignment = xlCenter
        ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).VerticalAlignment = xlCenter
        ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).ColumnWidth = 6

        If isoYear <> currentYear Then
            If i > 0 Then
                With ws.Range(ws.cells(HEADER_ROW_1, yearStartCol), ws.cells(HEADER_ROW_1, FIRST_TIMELINE_COL + i - 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .value = CStr(currentYear)
                    .Interior.Color = RGB(191, 191, 191)
                End With
            End If

            currentYear = isoYear
            yearStartCol = FIRST_TIMELINE_COL + i
        End If
    Next i

    With ws.Range(ws.cells(HEADER_ROW_1, yearStartCol), ws.cells(HEADER_ROW_1, FIRST_TIMELINE_COL + slotCount - 1))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .value = CStr(currentYear)
        .Interior.Color = RGB(191, 191, 191)
    End With

    ws.Range(ws.cells(HEADER_ROW_1, FIRST_TIMELINE_COL), ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + slotCount - 1)).Borders.LineStyle = xlContinuous

End Sub

'------------------------------------------------------------------------------
' FR:
' Construit la timeline mensuelle avec libelles mois courts et entetes annee fusionnes.
'
' EN:
' Builds the monthly timeline with short month labels and merged year headers.
'
' Entrees / Inputs:
' - Debut de mois projet et nombre de mois.
'
' Sorties / Outputs:
' - Colonnes Month compactes pour rendu agrege.
'
' Appele par / Called by:
' - SetupTimelineLayout.
'
' Notes:
' - Mode utile au pilotage macro; geometrie des taches est projetee dans les periodes.
'------------------------------------------------------------------------------
Private Sub BuildTimeline_Month(ByVal ws As Worksheet, ByVal projectStart As Variant, ByVal slotCount As Long)

    Dim perfScope As clsPerfScope

    Dim i As Long
    Dim monthStart As Date
    Dim currentYear As Long
    Dim yearStartCol As Long

    Set perfScope = Profiler_BeginScope("BuildTimeline_Month", "Timeline")

    currentYear = 0
    yearStartCol = FIRST_TIMELINE_COL

    For i = 0 To slotCount - 1

        monthStart = DateAdd("m", i, CDate(projectStart))

        ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).NumberFormat = "@"
        ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).value = Gantt_FormatMonthShort(monthStart)
        ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).HorizontalAlignment = xlCenter
        ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).VerticalAlignment = xlCenter
        ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + i).ColumnWidth = 8

        If Year(monthStart) <> currentYear Then
            If i > 0 Then
                With ws.Range(ws.cells(HEADER_ROW_1, yearStartCol), ws.cells(HEADER_ROW_1, FIRST_TIMELINE_COL + i - 1))
                    .Merge
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .value = CStr(currentYear)
                    .Interior.Color = RGB(191, 191, 191)
                End With
            End If

            currentYear = Year(monthStart)
            yearStartCol = FIRST_TIMELINE_COL + i
        End If
    Next i

    With ws.Range(ws.cells(HEADER_ROW_1, yearStartCol), ws.cells(HEADER_ROW_1, FIRST_TIMELINE_COL + slotCount - 1))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .value = CStr(currentYear)
        .Interior.Color = RGB(191, 191, 191)
    End With

    ws.Range(ws.cells(HEADER_ROW_1, FIRST_TIMELINE_COL), ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL + slotCount - 1)).Borders.LineStyle = xlContinuous

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper Write Left Panel Row dans le workflow de rendu GANTT.
' EN: Runs the Write Left Panel Row helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub WriteLeftPanelRow(ByVal ws As Worksheet, ByVal ganttRow As Long, ByRef dataArr As Variant, ByVal dataRow As Long, ByVal mapWBS As Object)

    Dim perfScope As clsPerfScope

    Dim logicVal As String
    Dim isLoE As Boolean
    Dim preserveTestInputs As Boolean
    Dim progressVal As Variant
    Dim rowValues As Variant
    Dim coreValues As Variant
    Dim durationProgressValues As Variant
    Dim formulaText As String
    Dim startRef As String
    Dim finishRef As String
    Dim batchWriteCount As Long
    Dim formatWriteCount As Long

    Set perfScope = Profiler_BeginScope("WriteLeftPanelRow", "Excel Cell Write")

    ws.rows(ganttRow).rowHeight = GANTT_ROW_HEIGHT_TASK

    logicVal = ""
    If mapWBS.Exists("Driving Logic") Then
        logicVal = CStr(dataArr(dataRow, mapWBS("Driving Logic")))
    End If

    isLoE = TaskTypeRules_IsLevelOfEffortRow(dataArr, mapWBS, dataRow)
    preserveTestInputs = GetGanttPreserveTestInputs()

    ws.cells(ganttRow, COL_WBS).NumberFormat = "@"
    formatWriteCount = formatWriteCount + 1

    If HasValue(dataArr(dataRow, mapWBS("% Progress"))) Then
        progressVal = dataArr(dataRow, mapWBS("% Progress"))
    Else
        progressVal = 0
    End If

    If isLoE Then
        progressVal = Empty
    End If

    If preserveTestInputs Then
        ReDim coreValues(1 To 1, 1 To 4)
        coreValues(1, 1) = NormalizeWBS(CStr(dataArr(dataRow, mapWBS("WBS"))))
        coreValues(1, 2) = dataArr(dataRow, mapWBS("Task Name"))
        coreValues(1, 3) = dataArr(dataRow, mapWBS("Calculated Start"))
        coreValues(1, 4) = dataArr(dataRow, mapWBS("Calculated Finish"))
        ws.Range(ws.cells(ganttRow, COL_WBS), ws.cells(ganttRow, COL_FINISH)).Value2 = coreValues
        batchWriteCount = batchWriteCount + 1

        ReDim durationProgressValues(1 To 1, 1 To 2)
        durationProgressValues(1, 1) = dataArr(dataRow, mapWBS("Calculated Duration"))
        durationProgressValues(1, 2) = progressVal
        ws.Range(ws.cells(ganttRow, COL_DURATION), ws.cells(ganttRow, COL_PROGRESS)).Value2 = durationProgressValues
        batchWriteCount = batchWriteCount + 1

        ws.cells(ganttRow, COL_LOGIC).Value2 = logicVal
        batchWriteCount = batchWriteCount + 1
    Else
        ReDim rowValues(1 To 1, 1 To COL_LOGIC)
        rowValues(1, COL_WBS) = NormalizeWBS(CStr(dataArr(dataRow, mapWBS("WBS"))))
        rowValues(1, COL_TASK) = dataArr(dataRow, mapWBS("Task Name"))
        rowValues(1, COL_START) = dataArr(dataRow, mapWBS("Calculated Start"))
        rowValues(1, COL_FINISH) = dataArr(dataRow, mapWBS("Calculated Finish"))
        rowValues(1, COL_TEST_START) = Empty
        rowValues(1, COL_TEST_FINISH) = Empty
        rowValues(1, COL_DURATION) = dataArr(dataRow, mapWBS("Calculated Duration"))
        rowValues(1, COL_PROGRESS) = progressVal
        rowValues(1, COL_TEST_PROGRESS) = Empty
        rowValues(1, COL_LOGIC) = logicVal

        ws.Range(ws.cells(ganttRow, COL_WBS), ws.cells(ganttRow, COL_LOGIC)).Value2 = rowValues
        batchWriteCount = batchWriteCount + 1
    End If

    If isLoE Then

        'Primavera-style LOE progress:
        '- Dynamic with TODAY()
        '- Based on displayed Calculated Start / Calculated Finish in columns C / D
        '- Clamped between 0% and 100%
        '- Formula is written in English/Formula syntax so Excel stores it robustly.
        startRef = "C" & CStr(ganttRow)
        finishRef = "D" & CStr(ganttRow)
        formulaText = _
            "=IF(OR(" & _
                startRef & "=""""," & _
                finishRef & "=""""),0," & _
              "MAX(0,MIN(1,(TODAY()-" & startRef & "+1)/(" & _
                               finishRef & "-" & _
                               startRef & "+1))))"

        ws.cells(ganttRow, COL_PROGRESS).Formula = formulaText
        batchWriteCount = batchWriteCount + 1

    End If

    ws.Range(ws.cells(ganttRow, COL_START), ws.cells(ganttRow, COL_TEST_FINISH)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.cells(ganttRow, COL_PROGRESS), ws.cells(ganttRow, COL_TEST_PROGRESS)).NumberFormat = "0%"
    formatWriteCount = formatWriteCount + 2

    Profiler_RecordOperation "WriteLeftPanelRowBatchWrites", batchWriteCount, 0#
    Profiler_RecordOperation "WriteLeftPanelRowFormatWrites", formatWriteCount, 0#

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Row Style sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Row Style without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyRowStyle(ByVal ws As Worksheet, ByVal ganttRow As Long, ByRef dataArr As Variant, ByVal dataRow As Long, ByVal mapWBS As Object, ByVal hasChildren As Object, ByVal calcDrivingMap As Object)

    Dim perfScope As clsPerfScope

    Dim wbs As String
    Dim levelCount As Long
    Dim idVal As String
    Dim isLeaf As Boolean
    Dim hasActual As Boolean
    Dim isLoE As Boolean
    Dim logicVal As String

    Set perfScope = Profiler_BeginScope("ApplyRowStyle", "Excel Format")

    wbs = NormalizeWBS(CStr(dataArr(dataRow, mapWBS("WBS"))))
    levelCount = WBSLevel(wbs)
    idVal = Trim$(CStr(dataArr(dataRow, mapWBS("ID"))))
    isLeaf = Not hasChildren.Exists(wbs)

    logicVal = ""
    If mapWBS.Exists("Driving Logic") Then
        logicVal = CStr(dataArr(dataRow, mapWBS("Driving Logic")))
    End If

    isLoE = TaskTypeRules_IsLevelOfEffortRow(dataArr, mapWBS, dataRow)

    hasActual = False
    If mapWBS.Exists("Actual Start") Then
        If HasValue(dataArr(dataRow, mapWBS("Actual Start"))) Then hasActual = True
    End If
    If mapWBS.Exists("Actual Finish") Then
        If HasValue(dataArr(dataRow, mapWBS("Actual Finish"))) Then hasActual = True
    End If

    ws.cells(ganttRow, COL_TASK).IndentLevel = WorksheetFunction.Min(levelCount - 1, 15)

    If hasChildren.Exists(wbs) Then
        ws.Range(ws.cells(ganttRow, 1), ws.cells(ganttRow, COL_LOGIC)).Font.Bold = True
        ws.Range(ws.cells(ganttRow, 1), ws.cells(ganttRow, COL_LOGIC)).Interior.Color = RGB(248, 248, 248)
    End If

    ApplyTestCellColoring ws, ganttRow, isLeaf, hasActual, isLoE



End Sub


'------------------------------------------------------------------------------
' FR: Execute le helper Finalize Gantt Sheet  Display Only dans le workflow de rendu GANTT.
' EN: Runs the Finalize Gantt Sheet  Display Only helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub FinalizeGanttSheet_DisplayOnly(ByVal wsGantt As Worksheet, ByVal totalDays As Long, ByVal rowCount As Long)

    Dim lastCol As Long
    Dim lastRow As Long

    lastCol = FIRST_TIMELINE_COL + totalDays - 1
    lastRow = FIRST_TASK_ROW + rowCount - 1

    If lastCol < FIRST_TIMELINE_COL Then lastCol = FIRST_TIMELINE_COL
    If lastRow < FIRST_TASK_ROW Then lastRow = FIRST_TASK_ROW

    wsGantt.Range(wsGantt.cells(HEADER_ROW_2, FIRST_TIMELINE_COL), wsGantt.cells(lastRow, lastCol)).Borders(xlInsideVertical).LineStyle = xlDot
    wsGantt.Range(wsGantt.cells(FIRST_TASK_ROW, FIRST_TIMELINE_COL), wsGantt.cells(lastRow, lastCol)).Borders(xlInsideHorizontal).LineStyle = xlDot

End Sub

'------------------------------------------------------------------------------
' FR: Execute le helper W B S Level dans le workflow de rendu GANTT.
' EN: Runs the W B S Level helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function WBSLevel(ByVal wbs As String) As Long

    If Trim$(wbs) = "" Then
        WBSLevel = 1
    Else
        WBSLevel = UBound(Split(wbs, ".")) + 1
    End If

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Freeze Gantt After Finish dans le workflow de rendu GANTT.
' EN: Runs the Freeze Gantt After Finish helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Sub FreezeGanttAfterFinish(ByVal ws As Worksheet, ByVal rowCount As Long)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("FreezeGanttAfterFinish", "Gantt Layout")

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.cells(FIRST_TASK_ROW, COL_TEST_START).Select
    ActiveWindow.FreezePanes = True

End Sub

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Public Function IsGanttSheetLayoutEmpty(ByVal ws As Worksheet) As Boolean

    If ws Is Nothing Then
        IsGanttSheetLayoutEmpty = True
    Else
        IsGanttSheetLayoutEmpty = (Trim$(CStr(ws.cells(TITLE_ROW, COL_WBS).value)) = "" And _
                                   Trim$(CStr(ws.cells(HEADER_ROW_2, COL_WBS).value)) = "" And _
                                   Trim$(CStr(ws.cells(HEADER_ROW_2, FIRST_TIMELINE_COL).value)) = "")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Verifie et prepare une ressource GANTT requise avant le rendu ou l'interaction.
' EN: Ensures and prepares a GANTT resource required before rendering or interaction.
'------------------------------------------------------------------------------
Public Sub EnsureGanttVisualLayoutReadyBeforeDrawing(ByVal ws As Worksheet, Optional ByVal allowActivation As Boolean = True, Optional ByVal allowDoEvents As Boolean = True)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("EnsureGanttVisualLayoutReadyBeforeDrawing", "Gantt Layout")

    If ws Is Nothing Then Exit Sub

    If Not allowActivation Then Exit Sub

    ws.Activate
    If allowDoEvents Then DoEvents

End Sub

'------------------------------------------------------------------------------
' FR: Nettoie, restaure ou normalise une partie de l'etat visuel GANTT.
' EN: Cleans, restores, or normalizes part of the GANTT visual state.
'------------------------------------------------------------------------------
Public Sub NormalizeGanttTestPercentCell(ByVal cell As Range)

    Dim rawVal As Variant
    Dim pctVal As Double

    rawVal = cell.value

    If Trim$(CStr(rawVal)) = "" Then Exit Sub
    If Not IsNumeric(rawVal) Then Exit Sub

    pctVal = CDbl(rawVal)

    If pctVal < 0 Then Exit Sub

    If pctVal <= 1 Then
        cell.value = pctVal
    Else
        cell.value = pctVal / 100#
    End If

    cell.NumberFormat = "0%"

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Test Cell Coloring sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Test Cell Coloring without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyTestCellColoring( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal isLeaf As Boolean, _
    ByVal hasActual As Boolean, _
    Optional ByVal isLoE As Boolean = False)

    Dim colStart As Long
    Dim colFinish As Long
    Dim colProgress As Long
    Dim rngDates As Range
    Dim rngProgress As Range

    colStart = 5
    colFinish = 6
    colProgress = 9

    Set rngDates = Union( _
        ws.cells(rowIndex, colStart), _
        ws.cells(rowIndex, colFinish))

    Set rngProgress = ws.cells(rowIndex, colProgress)

    If Not isLeaf Then
        rngDates.Interior.Pattern = xlNone
        rngProgress.Interior.Pattern = xlNone
        Exit Sub
    End If

    rngDates.Interior.Pattern = xlSolid

    If hasActual Then
        rngDates.Interior.Color = RGB(255, 255, 204)
    Else
        rngDates.Interior.Color = RGB(255, 255, 153)
    End If

    If isLoE Then
        'LOE progress is calculated automatically from dates.
        'Do not invite manual Test % input visually.
        rngProgress.Interior.Pattern = xlNone
    Else
        rngProgress.Interior.Pattern = xlSolid

        If hasActual Then
            rngProgress.Interior.Color = RGB(255, 255, 204)
        Else
            rngProgress.Interior.Color = RGB(255, 255, 153)
        End If
    End If

End Sub

'------------------------------------------------------------------------------
' FR:
' Vide differentiellement les trois colonnes de saisie TEST du Gantt.
' Retourne le nombre de cellules effectivement effacees et ne touche a aucun
' mode runtime, store de simulation ou rendu.
'
' EN:
' Differentially clears the three Gantt TEST input columns.
' Returns the number of cells actually cleared and does not mutate runtime
' modes, the simulation store, or rendering.
'------------------------------------------------------------------------------
Public Function GanttSheetLayout_ClearTestInputs() As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim oldInternalWrite As Boolean
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim stateCaptured As Boolean
    Dim testRange As Range

    On Error GoTo SafeExit

    oldInternalWrite = GetGanttInternalWrite()
    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    stateCaptured = True

    SetGanttInternalWrite True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(GANTT_SHEET)
    On Error GoTo SafeExit

    If ws Is Nothing Then GoTo SafeExit

    lastRow = GetLastGanttRow(ws)
    If lastRow < FIRST_TASK_ROW Then GoTo SafeExit

    Set testRange = ws.Range(ws.cells(FIRST_TASK_ROW, COL_TEST_START), ws.cells(lastRow, COL_TEST_START))
    GanttSheetLayout_ClearTestInputs = _
        GanttSheetLayout_ClearTestInputs + GanttSheetLayout_ClearInputRange(testRange)

    Set testRange = ws.Range(ws.cells(FIRST_TASK_ROW, COL_TEST_FINISH), ws.cells(lastRow, COL_TEST_FINISH))
    GanttSheetLayout_ClearTestInputs = _
        GanttSheetLayout_ClearTestInputs + GanttSheetLayout_ClearInputRange(testRange)

    Set testRange = ws.Range(ws.cells(FIRST_TASK_ROW, COL_TEST_PROGRESS), ws.cells(lastRow, COL_TEST_PROGRESS))
    GanttSheetLayout_ClearTestInputs = _
        GanttSheetLayout_ClearTestInputs + GanttSheetLayout_ClearInputRange(testRange)

SafeExit:
    If stateCaptured Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEvents
        SetGanttInternalWrite oldInternalWrite
    End If

End Function

'------------------------------------------------------------------------------
' FR: Wrapper historique deleguant au contrat atomique de sortie de simulation.
' EN: Historical wrapper delegating to the atomic simulation-exit contract.
'------------------------------------------------------------------------------
Public Sub Gantt_Clear_Test_State()
    GanttSimulation_ResetToNormal False
End Sub

'------------------------------------------------------------------------------
' FR: Efface une plage d'entree uniquement lorsqu'elle contient des valeurs.
' EN: Clears an input range only when it contains values.
'------------------------------------------------------------------------------
Private Function GanttSheetLayout_ClearInputRange(ByVal targetRange As Range) As Long

    If targetRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(targetRange) = 0 Then Exit Function

    GanttSheetLayout_ClearInputRange = _
        CLng(Application.WorksheetFunction.CountA(targetRange))
    targetRange.ClearContents

End Function
