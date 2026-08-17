Attribute VB_Name = "mod_GanttTimelineGeometry"
Option Explicit

'===============================================================================
' MODULE : mod_GanttTimelineGeometry
' DOMAINE / DOMAIN : Gantt
'
' FR
' Fournit les calculs purs de geometrie et de positionnement du domaine.
' Ne cree aucune shape et ne decide aucun workflow.
'
' EN
' Provides pure domain geometry and positioning calculations.
' Creates no shape and decides no workflow.
'
' CONTRATS / CONTRACTS : IsAggregatedScaleMode, GetIsoWeekMonday, GetIsoWeekLabel, Gantt_FormatMonthShort, Gantt_FormatMonthYear, GetIsoWeekYear, GetScalePeriodStart, GetScalePeriodFinish
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' Pure timeline/date/slot geometry helpers extracted from mod_Gantt.
' Keep formulas and rounding identical to the legacy renderer.
'=====================================================

Private Const FIRST_TIMELINE_COL As Long = 11
Private Const HEADER_ROW_2 As Long = 4
Private Const GANTT_SCALE_WEEK As String = "WEEK"
Private Const GANTT_SCALE_MONTH As String = "MONTH"

Private gTimelineProjectionActive As Boolean
Private gTimelineProjectionSignature As String
Private gTimelineSlotCount As Long
Private gTimelineSlotLeft() As Double
Private gTimelineSlotWidth() As Double

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function IsWeekScaleMode() As Boolean

    IsWeekScaleMode = (GetGanttTimelineScaleMode() = GANTT_SCALE_WEEK)

End Function

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function IsMonthScaleMode() As Boolean

    IsMonthScaleMode = (GetGanttTimelineScaleMode() = GANTT_SCALE_MONTH)

End Function

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Public Function IsAggregatedScaleMode() As Boolean

    IsAggregatedScaleMode = (GetGanttTimelineScaleMode() = GANTT_SCALE_WEEK Or GetGanttTimelineScaleMode() = GANTT_SCALE_MONTH)

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Iso Week Monday dans le workflow de rendu GANTT.
' EN: Runs the Get Iso Week Monday helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetIsoWeekMonday(ByVal anyDate As Date) As Date

    GetIsoWeekMonday = CDate(anyDate - Weekday(anyDate, vbMonday) + 1)

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Iso Week Label dans le workflow de rendu GANTT.
' EN: Runs the Get Iso Week Label helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetIsoWeekLabel(ByVal anyDate As Date) As String

    GetIsoWeekLabel = Gantt_FormatWeekLabel(anyDate)

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Format Week Label dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Format Week Label helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function Gantt_FormatWeekLabel(ByVal anyDate As Date) As String

    Gantt_FormatWeekLabel = IIf(Gantt_CurrentLanguage() = "FR", "S", "W") & Format$(WorksheetFunction.IsoWeekNum(anyDate), "00")

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Format Month Short dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Format Month Short helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function Gantt_FormatMonthShort(ByVal anyDate As Date) As String

    Static frMonths(1 To 12) As String
    Static enMonths(1 To 12) As String

    If frMonths(1) = "" Then
        frMonths(1) = "jan"
        frMonths(2) = "fév"
        frMonths(3) = "mar"
        frMonths(4) = "avr"
        frMonths(5) = "mai"
        frMonths(6) = "juin"
        frMonths(7) = "juil"
        frMonths(8) = "août"
        frMonths(9) = "sep"
        frMonths(10) = "oct"
        frMonths(11) = "nov"
        frMonths(12) = "déc"

        enMonths(1) = "Jan"
        enMonths(2) = "Feb"
        enMonths(3) = "Mar"
        enMonths(4) = "Apr"
        enMonths(5) = "May"
        enMonths(6) = "Jun"
        enMonths(7) = "Jul"
        enMonths(8) = "Aug"
        enMonths(9) = "Sep"
        enMonths(10) = "Oct"
        enMonths(11) = "Nov"
        enMonths(12) = "Dec"
    End If

    If Gantt_CurrentLanguage() = "FR" Then
        Gantt_FormatMonthShort = frMonths(Month(anyDate))
    Else
        Gantt_FormatMonthShort = enMonths(Month(anyDate))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Gantt  Format Month Year dans le workflow de rendu GANTT.
' EN: Runs the Gantt  Format Month Year helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function Gantt_FormatMonthYear(ByVal anyDate As Date) As String

    Gantt_FormatMonthYear = Gantt_FormatMonthShort(anyDate) & " " & CStr(Year(anyDate))

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Iso Week Year dans le workflow de rendu GANTT.
' EN: Runs the Get Iso Week Year helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetIsoWeekYear(ByVal anyDate As Date) As Long

    Dim thursdayDate As Date

    thursdayDate = GetIsoWeekMonday(anyDate) + 3
    GetIsoWeekYear = Year(thursdayDate)

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Month Start dans le workflow de rendu GANTT.
' EN: Runs the Get Month Start helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function GetMonthStart(ByVal anyDate As Date) As Date

    GetMonthStart = DateSerial(Year(anyDate), Month(anyDate), 1)

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Month Finish dans le workflow de rendu GANTT.
' EN: Runs the Get Month Finish helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function GetMonthFinish(ByVal anyDate As Date) As Date

    GetMonthFinish = DateSerial(Year(anyDate), Month(anyDate) + 1, 0)

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Scale Period Start dans le workflow de rendu GANTT.
' EN: Runs the Get Scale Period Start helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetScalePeriodStart(ByVal anyDate As Date) As Date

    If IsWeekScaleMode() Then
        GetScalePeriodStart = GetIsoWeekMonday(anyDate)
    ElseIf IsMonthScaleMode() Then
        GetScalePeriodStart = GetMonthStart(anyDate)
    Else
        GetScalePeriodStart = anyDate
    End If

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Scale Period Finish dans le workflow de rendu GANTT.
' EN: Runs the Get Scale Period Finish helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetScalePeriodFinish(ByVal anyDate As Date) As Date

    If IsWeekScaleMode() Then
        GetScalePeriodFinish = GetIsoWeekMonday(anyDate) + 6
    ElseIf IsMonthScaleMode() Then
        GetScalePeriodFinish = GetMonthFinish(anyDate)
    Else
        GetScalePeriodFinish = anyDate
    End If

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Week Scale Project Start dans le workflow de rendu GANTT.
' EN: Runs the Get Week Scale Project Start helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function GetWeekScaleProjectStart(ByVal rawProjectStart As Variant) As Variant

    If Not HasValue(rawProjectStart) Then Exit Function
    GetWeekScaleProjectStart = GetIsoWeekMonday(CDate(rawProjectStart))

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Week Scale Project Finish dans le workflow de rendu GANTT.
' EN: Runs the Get Week Scale Project Finish helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Private Function GetWeekScaleProjectFinish(ByVal rawProjectFinish As Variant) As Variant

    If Not HasValue(rawProjectFinish) Then Exit Function
    GetWeekScaleProjectFinish = GetIsoWeekMonday(CDate(rawProjectFinish)) + 6

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Scale Project Start dans le workflow de rendu GANTT.
' EN: Runs the Get Scale Project Start helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetScaleProjectStart(ByVal rawProjectStart As Variant) As Variant

    If Not HasValue(rawProjectStart) Then Exit Function
    GetScaleProjectStart = GetScalePeriodStart(CDate(rawProjectStart))

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Scale Project Finish dans le workflow de rendu GANTT.
' EN: Runs the Get Scale Project Finish helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetScaleProjectFinish(ByVal rawProjectFinish As Variant) As Variant

    If Not HasValue(rawProjectFinish) Then Exit Function
    GetScaleProjectFinish = GetScalePeriodFinish(CDate(rawProjectFinish))

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Scale Project Finish From Slots dans le workflow de rendu GANTT.
' EN: Runs the Get Scale Project Finish From Slots helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetScaleProjectFinishFromSlots(ByVal projectStart As Variant, ByVal slotCount As Long) As Date

    If IsWeekScaleMode() Then
        GetScaleProjectFinishFromSlots = CDate(projectStart) + ((slotCount - 1) * 7) + 6
    ElseIf IsMonthScaleMode() Then
        GetScaleProjectFinishFromSlots = GetMonthFinish(DateAdd("m", slotCount - 1, CDate(projectStart)))
    Else
        GetScaleProjectFinishFromSlots = CDate(CDbl(projectStart) + slotCount - 1)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne une decision de rendu ou d'etat utilisee par le workflow GANTT.
' EN: Returns a rendering or state decision used by the GANTT workflow.
'------------------------------------------------------------------------------
Private Function IsTaskContainedInSingleScalePeriod(ByVal startVal As Variant, ByVal finishVal As Variant) As Boolean

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    If IsWeekScaleMode() Then
        IsTaskContainedInSingleScalePeriod = (GetIsoWeekMonday(CDate(startVal)) = GetIsoWeekMonday(CDate(finishVal)))
    ElseIf IsMonthScaleMode() Then
        IsTaskContainedInSingleScalePeriod = (GetMonthStart(CDate(startVal)) = GetMonthStart(CDate(finishVal)))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Timeline Slot Count sans exposer de mutateur sur l'etat source.
' EN: Returns the Timeline Slot Count value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function GetTimelineSlotCount(ByVal projectStart As Variant, ByVal projectFinish As Variant) As Long

    Dim startVal As Date
    Dim finishVal As Date

    If Not HasValue(projectStart) Then Exit Function
    If Not HasValue(projectFinish) Then Exit Function

    startVal = CDate(projectStart)
    finishVal = CDate(projectFinish)

    If finishVal < startVal Then Exit Function

    If IsWeekScaleMode() Then
        GetTimelineSlotCount = CLng(DateDiff("ww", GetIsoWeekMonday(startVal), GetIsoWeekMonday(finishVal), vbMonday, vbFirstFourDays)) + 1
    ElseIf IsMonthScaleMode() Then
        GetTimelineSlotCount = CLng(DateDiff("m", GetMonthStart(startVal), GetMonthStart(finishVal))) + 1
    Else
        GetTimelineSlotCount = CLng(finishVal - startVal + 1)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Timeline Slot Index sans exposer de mutateur sur l'etat source.
' EN: Returns the Timeline Slot Index value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function GetTimelineSlotIndex(ByVal projectStart As Variant, ByVal anyDate As Variant) As Long

    Dim startVal As Date
    Dim dateVal As Date

    If Not HasValue(projectStart) Then
        GetTimelineSlotIndex = -1
        Exit Function
    End If

    If Not HasValue(anyDate) Then
        GetTimelineSlotIndex = -1
        Exit Function
    End If

    startVal = CDate(projectStart)
    dateVal = CDate(anyDate)

    If IsWeekScaleMode() Then
        GetTimelineSlotIndex = CLng(DateDiff("ww", GetIsoWeekMonday(startVal), GetIsoWeekMonday(dateVal), vbMonday, vbFirstFourDays))
    ElseIf IsMonthScaleMode() Then
        GetTimelineSlotIndex = CLng(DateDiff("m", GetMonthStart(startVal), GetMonthStart(dateVal)))
    Else
        GetTimelineSlotIndex = CLng(CDbl(dateVal) - CDbl(startVal))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Render Start For Current Scale dans le workflow de rendu GANTT.
' EN: Runs the Get Render Start For Current Scale helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetRenderStartForCurrentScale(ByVal startVal As Variant) As Variant

    If Not HasValue(startVal) Then Exit Function

    If IsAggregatedScaleMode() Then
        GetRenderStartForCurrentScale = GetScalePeriodStart(CDate(startVal))
    Else
        GetRenderStartForCurrentScale = startVal
    End If

End Function

'------------------------------------------------------------------------------
' FR: Execute le helper Get Render Finish For Current Scale dans le workflow de rendu GANTT.
' EN: Runs the Get Render Finish For Current Scale helper in the GANTT rendering workflow.
'------------------------------------------------------------------------------
Public Function GetRenderFinishForCurrentScale(ByVal finishVal As Variant) As Variant

    If Not HasValue(finishVal) Then Exit Function

    If IsAggregatedScaleMode() Then
        GetRenderFinishForCurrentScale = GetScalePeriodFinish(CDate(finishVal))
    Else
        GetRenderFinishForCurrentScale = finishVal
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Date Start X sans modifier les donnees d'entree.
' EN: Returns the Timeline Date Start X reference without mutating input data.
'------------------------------------------------------------------------------

Private Function TimelineDateStartX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    TimelineDateStartX = TimelineDateX(ws, projectStart, taskDate, False)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Date Finish X sans modifier les donnees d'entree.
' EN: Returns the Timeline Date Finish X reference without mutating input data.
'------------------------------------------------------------------------------

Private Function TimelineDateFinishX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    TimelineDateFinishX = TimelineDateX(ws, projectStart, taskDate, True)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Date Range Mid X sans modifier les donnees d'entree.
' EN: Returns the Timeline Date Range Mid X reference without mutating input data.
'------------------------------------------------------------------------------

Public Function TimelineDateRangeMidX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant) As Double

    Dim leftX As Double
    Dim rightX As Double

    If Not HasValue(startVal) Then Exit Function
    If Not HasValue(finishVal) Then Exit Function

    leftX = TimelineDateStartX(ws, projectStart, startVal)
    rightX = TimelineDateFinishX(ws, projectStart, finishVal)

    If rightX <= leftX Then Exit Function

    TimelineDateRangeMidX = leftX + ((rightX - leftX) / 2)

End Function


'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Date X sans modifier les donnees d'entree.
' EN: Returns the Timeline Date X reference without mutating input data.
'------------------------------------------------------------------------------

Private Function TimelineDateX( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant, _
    ByVal isFinishSide As Boolean) As Double

    Dim perfScope As clsPerfScope

    Dim targetCol As Long
    Dim cellLeft As Double
    Dim cellWidth As Double
    Dim dateVal As Date
    Dim periodStart As Date
    Dim periodDays As Long
    Dim offsetDays As Double

    Set perfScope = Profiler_BeginScope("TimelineDateX", "Timeline Geometry")

    If ws Is Nothing Then Exit Function
    If Not HasValue(projectStart) Then Exit Function
    If Not HasValue(taskDate) Then Exit Function

    If GanttTimelineProjection_IsActive(projectStart) Then
        TimelineDateX = GanttTimelineProjection_DateX(projectStart, taskDate, isFinishSide)
        Exit Function
    End If

    targetCol = TimelineColumnFromHeaderDate_Exact(ws, projectStart, taskDate)
    If targetCol < FIRST_TIMELINE_COL Then Exit Function

    cellLeft = ws.cells(HEADER_ROW_2, targetCol).Left
    cellWidth = ws.cells(HEADER_ROW_2, targetCol).Width

    If Not IsAggregatedScaleMode() Then
        If isFinishSide Then
            TimelineDateX = cellLeft + cellWidth
        Else
            TimelineDateX = cellLeft
        End If
        Exit Function
    End If

    dateVal = CDate(taskDate)

    If IsWeekScaleMode() Then
        periodStart = GetIsoWeekMonday(dateVal)
        periodDays = 7
    ElseIf IsMonthScaleMode() Then
        periodStart = GetMonthStart(dateVal)
        periodDays = Day(DateSerial(Year(periodStart), Month(periodStart) + 1, 0))
    Else
        periodStart = dateVal
        periodDays = 1
    End If

    offsetDays = CDbl(DateDiff("d", periodStart, dateVal))
    If isFinishSide Then offsetDays = offsetDays + 1

    If offsetDays < 0 Then offsetDays = 0
    If offsetDays > periodDays Then offsetDays = periodDays

    TimelineDateX = cellLeft + (cellWidth * (offsetDays / periodDays))

End Function

Public Sub GanttTimelineProjection_Begin( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal slotCount As Long)

    Dim perfScope As clsPerfScope
    Dim i As Long
    Dim colIndex As Long

    Set perfScope = Profiler_BeginScope("GanttTimelineProjection_Begin", "Timeline Projection")

    gTimelineProjectionActive = False
    gTimelineProjectionSignature = vbNullString
    gTimelineSlotCount = 0
    Erase gTimelineSlotLeft
    Erase gTimelineSlotWidth

    If ws Is Nothing Then Exit Sub
    If Not HasValue(projectStart) Then Exit Sub
    If slotCount < 1 Then Exit Sub

    ReDim gTimelineSlotLeft(0 To slotCount - 1)
    ReDim gTimelineSlotWidth(0 To slotCount - 1)

    For i = 0 To slotCount - 1
        colIndex = FIRST_TIMELINE_COL + i
        gTimelineSlotLeft(i) = ws.Cells(HEADER_ROW_2, colIndex).Left
        gTimelineSlotWidth(i) = ws.Cells(HEADER_ROW_2, colIndex).Width
    Next i

    gTimelineSlotCount = slotCount
    gTimelineProjectionSignature = GanttTimelineProjection_Signature(projectStart)
    gTimelineProjectionActive = True

    Profiler_RecordOperation "GanttTimelineProjectionSlotReads", slotCount, 0#

End Sub

Public Sub GanttTimelineProjection_End()

    gTimelineProjectionActive = False
    gTimelineProjectionSignature = vbNullString
    gTimelineSlotCount = 0
    Erase gTimelineSlotLeft
    Erase gTimelineSlotWidth

End Sub

Public Function GanttTimelineProjection_SlotWidth( _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Double

    Dim slotIndex As Long

    If Not GanttTimelineProjection_IsActive(projectStart) Then Exit Function
    slotIndex = GetTimelineSlotIndex(projectStart, taskDate)
    If slotIndex < 0 Or slotIndex >= gTimelineSlotCount Then Exit Function

    GanttTimelineProjection_SlotWidth = gTimelineSlotWidth(slotIndex)

End Function

Private Function GanttTimelineProjection_IsActive(ByVal projectStart As Variant) As Boolean

    If Not gTimelineProjectionActive Then Exit Function
    GanttTimelineProjection_IsActive = (gTimelineProjectionSignature = GanttTimelineProjection_Signature(projectStart))

End Function

Private Function GanttTimelineProjection_Signature(ByVal projectStart As Variant) As String

    If Not HasValue(projectStart) Then Exit Function
    GanttTimelineProjection_Signature = GetGanttTimelineScaleMode() & "|" & CStr(CDbl(projectStart)) & "|" & CStr(gTimelineSlotCount)

End Function

Private Function GanttTimelineProjection_DateX( _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant, _
    ByVal isFinishSide As Boolean) As Double

    Dim slotIndex As Long
    Dim cellLeft As Double
    Dim cellWidth As Double
    Dim dateVal As Date
    Dim periodStart As Date
    Dim periodDays As Long
    Dim offsetDays As Double

    If Not HasValue(projectStart) Then Exit Function
    If Not HasValue(taskDate) Then Exit Function

    slotIndex = GetTimelineSlotIndex(projectStart, taskDate)
    If slotIndex < 0 Or slotIndex >= gTimelineSlotCount Then Exit Function

    cellLeft = gTimelineSlotLeft(slotIndex)
    cellWidth = gTimelineSlotWidth(slotIndex)

    If Not IsAggregatedScaleMode() Then
        If isFinishSide Then
            GanttTimelineProjection_DateX = cellLeft + cellWidth
        Else
            GanttTimelineProjection_DateX = cellLeft
        End If
        Profiler_RecordOperation "GanttTimelineProjectionMemoryX", 1, 0#
        Exit Function
    End If

    dateVal = CDate(taskDate)

    If IsWeekScaleMode() Then
        periodStart = GetIsoWeekMonday(dateVal)
        periodDays = 7
    ElseIf IsMonthScaleMode() Then
        periodStart = GetMonthStart(dateVal)
        periodDays = Day(DateSerial(Year(periodStart), Month(periodStart) + 1, 0))
    Else
        periodStart = dateVal
        periodDays = 1
    End If

    offsetDays = CDbl(DateDiff("d", periodStart, dateVal))
    If isFinishSide Then offsetDays = offsetDays + 1

    If offsetDays < 0 Then offsetDays = 0
    If offsetDays > periodDays Then offsetDays = periodDays

    GanttTimelineProjection_DateX = cellLeft + (cellWidth * (offsetDays / periodDays))
    Profiler_RecordOperation "GanttTimelineProjectionMemoryX", 1, 0#

End Function
'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Right After Finish sans modifier les donnees d'entree.
' EN: Returns the Timeline Right After Finish reference without mutating input data.
'------------------------------------------------------------------------------

Public Function TimelineRightAfterFinish( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskFinish As Variant) As Double

    TimelineRightAfterFinish = TimelineDateFinishX(ws, projectStart, taskFinish)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Right sans modifier les donnees d'entree.
' EN: Returns the Timeline Right reference without mutating input data.
'------------------------------------------------------------------------------

Public Function TimelineRight( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskFinish As Variant) As Double

    TimelineRight = TimelineRightAfterFinish(ws, projectStart, taskFinish)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Timeline Target Col sans exposer de mutateur sur l'etat source.
' EN: Returns the Timeline Target Col value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function GetTimelineTargetCol( _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Long

    Dim slotOffset As Long

    If Not HasValue(projectStart) Then Exit Function
    If Not HasValue(taskDate) Then Exit Function

    slotOffset = GetTimelineSlotIndex(projectStart, taskDate)

    If slotOffset < 0 Then
        GetTimelineTargetCol = FIRST_TIMELINE_COL
    Else
        GetTimelineTargetCol = FIRST_TIMELINE_COL + slotOffset
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Left sans modifier les donnees d'entree.
' EN: Returns the Timeline Left reference without mutating input data.
'------------------------------------------------------------------------------

Public Function TimelineLeft( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskStart As Variant) As Double

    TimelineLeft = TimelineDateStartX(ws, projectStart, taskStart)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Width sans modifier les donnees d'entree.
' EN: Returns the Timeline Width reference without mutating input data.
'------------------------------------------------------------------------------

Public Function TimelineWidth( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskStart As Variant, _
    ByVal taskFinish As Variant) As Double

    Dim leftPos As Double
    Dim rightPos As Double

    If Not HasValue(projectStart) Then Exit Function
    If Not HasValue(taskStart) Then Exit Function
    If Not HasValue(taskFinish) Then Exit Function

    leftPos = TimelineLeft(ws, projectStart, taskStart)
    rightPos = TimelineRightAfterFinish(ws, projectStart, taskFinish)

    If rightPos <= leftPos Then
        TimelineWidth = 0
    Else
        TimelineWidth = rightPos - leftPos
    End If

End Function
'=====================================================
' DIRECT TIMELINE COLUMN MAPPING
'=====================================================

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Column From Date sans modifier les donnees d'entree.
' EN: Returns the Timeline Column From Date reference without mutating input data.
'------------------------------------------------------------------------------

Private Function TimelineColumnFromDate( _
    ByVal ws As Worksheet, _
    ByVal anyDate As Variant) As Long

    Dim lastCol As Long
    Dim c As Long
    Dim targetDate As Date
    Dim headerVal As Variant

    If Not HasValue(anyDate) Then Exit Function

    targetDate = CDate(anyDate)

    lastCol = ws.cells(HEADER_ROW_2, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < FIRST_TIMELINE_COL Then Exit Function

    For c = FIRST_TIMELINE_COL To lastCol

        headerVal = ws.cells(HEADER_ROW_2, c).value

        If IsAggregatedScaleMode() Then
            TimelineColumnFromDate = FIRST_TIMELINE_COL + GetTimelineSlotIndex(GetScalePeriodStart(CDate(anyDate)), CDate(anyDate))
            Exit Function
        Else
            If IsDate(headerVal) Then
                If CLng(CDate(headerVal)) = CLng(targetDate) Then
                    TimelineColumnFromDate = c
                    Exit Function
                End If
            End If
        End If

    Next c

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference Timeline Column From Header Date Exact sans modifier les donnees d'entree.
' EN: Returns the Timeline Column From Header Date Exact reference without mutating input data.
'------------------------------------------------------------------------------

Public Function TimelineColumnFromHeaderDate_Exact( _
    ByVal ws As Worksheet, _
    ByVal projectStart As Variant, _
    ByVal taskDate As Variant) As Long

    Dim perfScope As clsPerfScope
    Dim targetCol As Long

    Set perfScope = Profiler_BeginScope("TimelineColumnFromHeaderDate_Exact", "Timeline Direct Lookup")

    If ws Is Nothing Then Exit Function
    If Not HasValue(projectStart) Then Exit Function
    If Not HasValue(taskDate) Then Exit Function

    targetCol = GetTimelineTargetCol(projectStart, taskDate)
    If targetCol < FIRST_TIMELINE_COL Then Exit Function

    TimelineColumnFromHeaderDate_Exact = targetCol

End Function
