Attribute VB_Name = "mod_GanttDataProvider"
Option Explicit

'===============================================================================
' MODULE : mod_GanttDataProvider
' DOMAINE / DOMAIN : Gantt
'
' FR
' Lit et projette les donnees necessaires au domaine sans porter sa politique metier.
' Ne rend aucune UI et ne lance aucun moteur.
'
' EN
' Reads and projects domain data without owning business policy.
' Does not render UI or run an engine.
'
' CONTRATS / CONTRACTS : BuildGanttTestInputMap, CountRenderableGanttRows, ResolveDisplayedProjectRange, ValidateGanttSourceColumns, GetLastGanttRow, IsEditableTestCell
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

'------------------------------------------------------------------------------
' FR: Construit une map, un index ou une structure intermediaire consommee par le rendu GANTT.
' EN: Builds a map, index, or intermediate structure consumed by GANTT rendering.
'------------------------------------------------------------------------------
Public Function BuildGanttTestInputMap(ByVal ws As Worksheet) As Object

    Dim perfScope As clsPerfScope

    Dim d As Object
    Dim lastRow As Long
    Dim r As Long
    Dim wbsVal As String

    Set perfScope = Profiler_BeginScope("BuildGanttTestInputMap", "Excel Read")

    Set d = CreateObject("Scripting.Dictionary")

    lastRow = GetLastGanttRow(ws)
    If lastRow < FIRST_TASK_ROW Then
        Set BuildGanttTestInputMap = d
        Exit Function
    End If

    For r = FIRST_TASK_ROW To lastRow

        wbsVal = NormalizeWBS(CStr(ws.cells(r, COL_WBS).value))

        If wbsVal <> "" Then
            d(wbsVal) = Array( _
                GetCellValue(ws.cells(r, COL_TEST_START).value), _
                GetCellValue(ws.cells(r, COL_TEST_FINISH).value), _
                GetCellValue(ws.cells(r, COL_TEST_PROGRESS).value))
        End If

    Next r

    Set BuildGanttTestInputMap = d

End Function



'------------------------------------------------------------------------------
' FR: Execute le helper Count Renderable Gantt Rows dans le workflow de rendu GANTT.
' EN: Runs the Count Renderable Gantt Rows helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function CountRenderableGanttRows(ByRef dataArr As Variant, ByVal mapWBS As Object) As Long

    Dim perfScope As clsPerfScope

    Dim r As Long
    Dim rowCount As Long
    Dim summaryDisplayVal As String

    Set perfScope = Profiler_BeginScope("CountRenderableGanttRows", "Gantt Scan")

    rowCount = UBound(dataArr, 1)
    EnsureGanttViewInitialized

    If GetGanttViewMode() = GANTT_VIEW_DETAIL Then
        CountRenderableGanttRows = rowCount
        Exit Function
    End If

    If mapWBS Is Nothing Then Exit Function
    If Not mapWBS.Exists("S") Then Exit Function

    For r = 1 To rowCount
        summaryDisplayVal = UCase$(Trim$(CStr(dataArr(r, mapWBS("S")))))
        If summaryDisplayVal = "Y" Then CountRenderableGanttRows = CountRenderableGanttRows + 1
    Next r

End Function



'------------------------------------------------------------------------------
' FR: Execute le helper Resolve Displayed Project Range dans le workflow de rendu GANTT.
' EN: Runs the Resolve Displayed Project Range helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Sub ResolveDisplayedProjectRange( _
    ByVal dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByRef projectStart As Variant, _
    ByRef projectFinish As Variant)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("ResolveDisplayedProjectRange", "Gantt Scan")

    GetProjectDisplayRange dataArr, mapWBS, hasChildren, baseById, testById, isTestMode, projectStart, projectFinish

End Sub
'------------------------------------------------------------------------------
' FR: Verifie qu'une source Excel expose les colonnes attendues par le rendu GANTT.
' EN: Validates that an Excel source exposes the columns expected by GANTT rendering.
'------------------------------------------------------------------------------
Public Sub ValidateGanttSourceColumns(ByVal mapWBS As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        "ID", _
        "WBS", _
        "Task Name", _
        "S", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration", _
        "% Progress", _
        "Driving Logic", _
        "Critical Path" _
    )

    For Each c In requiredCols
        If Not mapWBS.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 700, , "Missing source column in tbl_WBS: " & CStr(c)
        End If
    Next c

End Sub
'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function IsLeafById(ByVal idVal As String, ByVal calcDrivingMap As Object) As Boolean

    If idVal = "" Then
        IsLeafById = False
    ElseIf calcDrivingMap.Exists(idVal) Then
        IsLeafById = (calcDrivingMap(idVal) <> "SUMMARY")
    Else
        IsLeafById = False
    End If

End Function


'------------------------------------------------------------------------------
' FR: Execute le helper Get Last Gantt Row dans le workflow de rendu GANTT.
' EN: Runs the Get Last Gantt Row helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetLastGanttRow(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    Set lastCell = ws.Columns(COL_WBS).Find( _
        What:="*", _
        After:=ws.cells(1, COL_WBS), _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If lastCell Is Nothing Then
        GetLastGanttRow = FIRST_TASK_ROW - 1
    ElseIf lastCell.Row < FIRST_TASK_ROW Then
        GetLastGanttRow = FIRST_TASK_ROW - 1
    Else
        GetLastGanttRow = lastCell.Row
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function IsAllowedTestColumn(ByVal colNum As Long) As Boolean

    IsAllowedTestColumn = ( _
        colNum = COL_TEST_START Or _
        colNum = COL_TEST_FINISH Or _
        colNum = COL_TEST_PROGRESS)

End Function
'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Public Function IsEditableTestCell(ByVal ws As Worksheet, ByVal targetCell As Range) As Boolean

    Dim rowNum As Long
    Dim colNum As Long
    Dim lastRow As Long
    Dim wbsVal As String
    Dim wbsToId As Object
    Dim calcDrivingMap As Object
    Dim idVal As String

    IsEditableTestCell = False

    rowNum = targetCell.Row
    colNum = targetCell.Column
    lastRow = GetLastGanttRow(ws)

    If rowNum < FIRST_TASK_ROW Or rowNum > lastRow Then Exit Function
    If Not IsAllowedTestColumn(colNum) Then Exit Function

    wbsVal = NormalizeWBS(CStr(ws.cells(rowNum, COL_WBS).value))
    If wbsVal = "" Then Exit Function

    Set wbsToId = CanonicalIdentity_GetWbsToIdMap()
    If Not wbsToId.Exists(wbsVal) Then Exit Function

    idVal = CStr(wbsToId(wbsVal))
    Set calcDrivingMap = CanonicalIdentity_GetDrivingLogicByIdMap()

    If IsLeafById(idVal, calcDrivingMap) Then
        IsEditableTestCell = True
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function IsTaskContainedInSingleIsoWeek(ByVal startVal As Variant, ByVal finishVal As Variant) As Boolean

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    IsTaskContainedInSingleIsoWeek = (GetIsoWeekMonday(CDate(startVal)) = GetIsoWeekMonday(CDate(finishVal)))

End Function

'------------------------------------------------------------------------------
' FR: Construit une map, un index ou une structure intermediaire consommee par le rendu GANTT.
' EN: Builds a map, index, or intermediate structure consumed by GANTT rendering.
'------------------------------------------------------------------------------
Private Function BuildGanttBaseDisplayMap() As Object

    Set BuildGanttBaseDisplayMap = GanttLive_BuildBaseByIdMap()

End Function

'------------------------------------------------------------------------------
' FR: Construit une map, un index ou une structure intermediaire consommee par le rendu GANTT.
' EN: Builds a map, index, or intermediate structure consumed by GANTT rendering.
'------------------------------------------------------------------------------
Private Function BuildGanttTestDisplayMap() As Object

    Set BuildGanttTestDisplayMap = GanttLive_BuildTestByIdMap()

End Function

'------------------------------------------------------------------------------
' FR:
' Calcule la plage de dates visible du projet a partir des dates d'affichage base ou TEST.
'
' EN:
' Computes the visible project date range from base or TEST display dates.
'
' Entrees / Inputs:
' - data WBS, maps base/test, hasChildren, mode TEST.
'
' Sorties / Outputs:
' - projectStart et projectFinish passes par reference.
'
' Appele par / Called by:
' - ResolveDisplayedProjectRange et registry predictif.
'
' Notes:
' - Determinne le nombre de slots timeline et donc toute la geometrie de rendu.
'------------------------------------------------------------------------------
Private Sub GetProjectDisplayRange( _
    ByRef dataArr As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal baseById As Object, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByRef projectStart As Variant, _
    ByRef projectFinish As Variant)

    Dim perfScope As clsPerfScope

    Dim r As Long
    Dim idVal As String
    Dim displayStart As Variant
    Dim displayFinish As Variant

    Set perfScope = Profiler_BeginScope("GetProjectDisplayRange", "Gantt Scan")

    projectStart = Empty
    projectFinish = Empty

    For r = 1 To UBound(dataArr, 1)

        idVal = Trim$(CStr(dataArr(r, mapWBS("ID"))))

        displayStart = GanttLive_GetDisplayStart(idVal, baseById, testById, isTestMode)
        displayFinish = GanttLive_GetDisplayFinish(idVal, baseById, testById, isTestMode)

        If HasValue(displayStart) And HasValue(displayFinish) Then

            If Not HasValue(projectStart) Then
                projectStart = displayStart
            ElseIf CDbl(displayStart) < CDbl(projectStart) Then
                projectStart = displayStart
            End If

            If Not HasValue(projectFinish) Then
                projectFinish = displayFinish
            ElseIf CDbl(displayFinish) > CDbl(projectFinish) Then
                projectFinish = displayFinish
            End If

        End If
    Next r

End Sub

