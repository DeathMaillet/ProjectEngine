Attribute VB_Name = "mod_SCurve"
Option Explicit

'===============================================================================
' MODULE : mod_SCurve
' DOMAINE / DOMAIN : S-Curve
'
' FR
' Calcule les series S-Curve, possede leurs tables et rend le graphique et sa projection Dashboard.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Calculates S-Curve series, owns their tables and renders the chart and Dashboard projection.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Run_SCurve_Engine, SCurve_SafeEmptyState, SCurve_ApplyLanguage, SCurve_SetLanguage, SCurve_CurrentLanguage, DrawSCurveTodayVerticalLine, SCurve_BuildDashboardProjection, SCurve_DrawDashboardTodayLine
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const SCURVE_SHEET As String = "SCURVE"
Private Const SCURVE_TABLE As String = "tbl_SCURVE"

Private Const CALC_SCURVE_SHEET As String = "CALC_SCURVE"
Private Const CALC_SCURVE_TABLE As String = "tbl_CALC_SCURVE"

Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"

Private Const SCURVE_CHART_SIGNATURE_NAME As String = "_SCURVE_CHART_SOURCE_SIGNATURE"
Private Const SCURVE_CHART_SIGNATURE_VERSION As String = "SCV1:"
Private Const SCURVE_CHART_SIGNATURE_KEY_LENGTH As Long = 69

Private gSCurveLanguage As String



'------------------------------------------------------------------------------
' FR: Lance le workflow SCurve Engine.
' EN: Runs the SCurve Engine workflow.
'------------------------------------------------------------------------------
Public Sub Run_SCurve_Engine()

    Dim perfScope As clsPerfScope

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim wsSCurve As Worksheet
    Dim wsCalcSCurve As Worksheet

    Dim tblWBS As ListObject
    Dim tblCalc As ListObject
    Dim tblSCurve As ListObject
    Dim tblCalcSCurve As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim mapSCurve As Object
    Dim mapCalcSCurve As Object

    Dim dataWBS As Variant
    Dim dataCalc As Variant
    Dim outCalc() As Variant
    Dim rowCount As Long
    Dim r As Long

    Dim hasChildren As Object
    Dim idToWbs As Object
    Dim calcRowById As Object

    Dim includedIds As Object
    Dim missingWeightIds As Object
    Dim blockingErrorIds As Object
    Dim excludedTaskTypeWithWeightIds As Object

    Dim baselineDailyByDate As Object
    Dim actualizedDailyByDate As Object
    Dim remainingDailyByDate As Object
    Dim allDates As Object

    Dim totalRawWeight As Double
    Dim consoleMessages As Collection
    Dim ackTokens As String

    Set perfScope = Profiler_BeginScope("Run_SCurve_Engine", "Workflow")

    On Error GoTo SafeExit


    Set consoleMessages = New Collection

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set wsSCurve = ThisWorkbook.Worksheets(SCURVE_SHEET)
    Set wsCalcSCurve = ThisWorkbook.Worksheets(CALC_SCURVE_SHEET)

    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set tblSCurve = wsSCurve.ListObjects(SCURVE_TABLE)
    Set tblCalcSCurve = wsCalcSCurve.ListObjects(CALC_SCURVE_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then
        SCurve_SafeEmptyState
        Exit Sub
    End If

    If tblCalc.DataBodyRange Is Nothing Then
        SCurve_SafeEmptyState
        Exit Sub
    End If

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set mapSCurve = CanonicalIdentity_BuildColumnMap(tblSCurve)
    Set mapCalcSCurve = CanonicalIdentity_BuildColumnMap(tblCalcSCurve)

    ValidateSCurveSourceColumns mapWBS, mapCalc
    ValidateSCurveOutputColumns mapSCurve
    ValidateCalcSCurveColumns mapCalcSCurve

    dataWBS = tblWBS.DataBodyRange.value
    dataCalc = tblCalc.DataBodyRange.value
    rowCount = UBound(dataWBS, 1)

    Set hasChildren = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set calcRowById = CanonicalIdentity_BuildCalcRowById(dataCalc, mapCalc)
    Set includedIds = CreateObject("Scripting.Dictionary")
    Set missingWeightIds = CreateObject("Scripting.Dictionary")
    Set blockingErrorIds = CreateObject("Scripting.Dictionary")
    Set excludedTaskTypeWithWeightIds = CreateObject("Scripting.Dictionary")

    Set baselineDailyByDate = CreateObject("Scripting.Dictionary")
    Set actualizedDailyByDate = CreateObject("Scripting.Dictionary")
    Set remainingDailyByDate = CreateObject("Scripting.Dictionary")
    Set allDates = CreateObject("Scripting.Dictionary")

    BuildWBSParentMaps dataCalc, mapCalc, hasChildren, idToWbs
    ResizeTableToRowCount tblCalcSCurve, rowCount
    ReDim outCalc(1 To rowCount, 1 To tblCalcSCurve.ListColumns.Count)

    totalRawWeight = 0

    '--------------------------------------------------
    ' PASS 1 : scan tasks / fill debug base / validate
    '--------------------------------------------------
    For r = 1 To rowCount

        Dim taskId As String
        Dim taskWBS As String
        Dim calcRow As Long
        Dim isLeaf As Boolean

        Dim rawWeight As Variant
        Dim baselineStart As Variant
        Dim baselineDuration As Variant
        Dim baselineFinish As Variant

        Dim calcStart As Variant
        Dim calcFinish As Variant
        Dim progressVal As Double
        Dim drivingLogic As String
        Dim taskTypeVal As String

        Dim baselineDaily As Variant
        Dim calculatedDaily As Variant
        Dim actualizedWeight As Variant

        Dim errorText As String

        taskId = Trim$(CStr(dataWBS(r, mapWBS(VTS_COL_ID))))
        baselineStart = Empty
        baselineDuration = Empty
        baselineFinish = Empty
        calcStart = Empty
        calcFinish = Empty
        drivingLogic = ""
        taskTypeVal = ""

        outCalc(r, mapCalcSCurve("ID")) = taskId

        If taskId = "" Then
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = ""
            outCalc(r, mapCalcSCurve("Error flag")) = ""
            GoTo NextPass1Row
        End If

        If Not calcRowById.Exists(taskId) Then
            blockingErrorIds(taskId) = True
            errorText = "TASK MISSING IN CALC"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        calcRow = CLng(calcRowById(taskId))
        taskWBS = Trim$(CStr(dataCalc(calcRow, mapCalc("WBS"))))
        isLeaf = Not hasChildren.Exists(NormalizeWBSLocal(taskWBS))

        rawWeight = dataWBS(r, mapWBS(VTS_COL_WEIGHT_PERCENT))
        progressVal = ProgressToDoubleLocal(dataWBS(r, mapWBS(VTS_COL_PROGRESS_PERCENT)))

        baselineStart = GetCellValue(dataCalc(calcRow, mapCalc("Baseline Start")))
        baselineDuration = GetCellValue(dataCalc(calcRow, mapCalc("Baseline Duration")))
        calcStart = GetCellValue(dataCalc(calcRow, mapCalc("Calculated Start")))
        calcFinish = GetCellValue(dataCalc(calcRow, mapCalc("Calculated Finish")))
        drivingLogic = UCase$(Trim$(CStr(dataCalc(calcRow, mapCalc("Driving Logic")))))
        taskTypeVal = Trim$(CStr(dataCalc(calcRow, mapCalc("Task Type"))))

        baselineDaily = Empty
        calculatedDaily = Empty
        actualizedWeight = Empty
        errorText = ""

        If Not isLeaf Then
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = "SUMMARY TASK EXCLUDED"
            outCalc(r, mapCalcSCurve("Error flag")) = ""
            GoTo WritePass1Row
        End If

        If IsSCurveExcludedTaskTypeLocal(taskTypeVal) Then
            If HasValue(rawWeight) Then
                excludedTaskTypeWithWeightIds(taskId) = True
            End If

            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = "TASK TYPE EXCLUDED FROM S-CURVE"
            outCalc(r, mapCalcSCurve("Error flag")) = ""
            GoTo WritePass1Row
        End If

        If Not HasValue(rawWeight) Then
            missingWeightIds(taskId) = True
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = "MISSING WEIGHT - EXCLUDED"
            outCalc(r, mapCalcSCurve("Error flag")) = ""
            GoTo WritePass1Row
        End If

        If Not IsNumeric(rawWeight) Then
            blockingErrorIds(taskId) = True
            errorText = "INVALID WEIGHT"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        If CDbl(rawWeight) <= 0 Then
            blockingErrorIds(taskId) = True
            errorText = "NON-POSITIVE WEIGHT"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        If (Not HasValue(baselineStart) Or Not HasValue(baselineDuration)) And _
           HasValue(calcStart) And HasValue(calcFinish) And drivingLogic = "BASELINE" Then

            baselineStart = calcStart
            baselineDuration = CDbl(calcFinish) - CDbl(calcStart) + 1
        End If

        If Not HasValue(baselineStart) Or Not HasValue(baselineDuration) Then
            blockingErrorIds(taskId) = True
            errorText = "MISSING BASELINE DATA"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        If CDbl(baselineDuration) <= 0 Then
            blockingErrorIds(taskId) = True
            errorText = "INVALID BASELINE DURATION"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        If Not HasValue(calcStart) Or Not HasValue(calcFinish) Then
            blockingErrorIds(taskId) = True
            errorText = "MISSING CALCULATED DATES"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        If CDbl(calcFinish) < CDbl(calcStart) Then
            blockingErrorIds(taskId) = True
            errorText = "CALCULATED FINISH BEFORE START"
            outCalc(r, mapCalcSCurve("SCurve Included")) = "NO"
            outCalc(r, mapCalcSCurve("SCurve Warning")) = errorText
            outCalc(r, mapCalcSCurve("Error flag")) = "ERROR"
            GoTo WritePass1Row
        End If

        includedIds(taskId) = True
        totalRawWeight = totalRawWeight + CDbl(rawWeight)

        baselineFinish = CDbl(baselineStart) + CDbl(baselineDuration) - 1
        baselineDaily = 1# / CDbl(baselineDuration)
        calculatedDaily = 1# / (CDbl(calcFinish) - CDbl(calcStart) + 1)

        actualizedWeight = CDbl(rawWeight) * progressVal

        outCalc(r, mapCalcSCurve("SCurve Included")) = "YES"
        outCalc(r, mapCalcSCurve("SCurve Warning")) = ""
        outCalc(r, mapCalcSCurve("Error flag")) = ""

WritePass1Row:
        outCalc(r, mapCalcSCurve("SCurve Weight Raw")) = rawWeight
        outCalc(r, mapCalcSCurve("SCurve Weight Normalized")) = ""
        outCalc(r, mapCalcSCurve("SCurve Baseline Start")) = baselineStart
        outCalc(r, mapCalcSCurve("SCurve Baseline Finish")) = baselineFinish
        outCalc(r, mapCalcSCurve("SCurve Calculated Start")) = calcStart
        outCalc(r, mapCalcSCurve("SCurve Calculated Finish")) = calcFinish
        outCalc(r, mapCalcSCurve("SCurve Progress")) = progressVal
        outCalc(r, mapCalcSCurve("SCurve Baseline Daily")) = baselineDaily
        outCalc(r, mapCalcSCurve("SCurve Calculated Daily")) = calculatedDaily
        outCalc(r, mapCalcSCurve("SCurve Actualized Weight")) = actualizedWeight

NextPass1Row:
    Next r


    tblCalcSCurve.DataBodyRange.value = outCalc
    FormatCalcSCurveColumns tblCalcSCurve, mapCalcSCurve

    If blockingErrorIds.Count > 0 Then
        SCurve_AddGroupedMessage consoleMessages, "STOP", blockingErrorIds, idToWbs, _
            "Données bloquantes pour S-curve", "corriger les champs nécessaires avant recalcul", _
            "Blocking data for S-curve", "fix required fields before recalculation"
        SCurve_SafeEmptyState
        GoTo SafeExit
    End If

    If includedIds.Count = 0 Then
        SCurve_AddConsoleMessage consoleMessages, "WARNING", _
            "Aucune tâche feuille exploitable pour la S-curve.", _
            "No valid leaf task available for S-curve.", _
            "SCURVE_NO_VALID_LEAF_TASK"
        SCurve_SafeEmptyState
        GoTo SafeExit
    End If

    If totalRawWeight <= 0 Then
        SCurve_AddConsoleMessage consoleMessages, "STOP", _
            "La somme des poids exploités pour la S-curve est nulle ou invalide.", _
            "The total usable S-curve weight is zero or invalid."
        SCurve_SafeEmptyState
        GoTo SafeExit
    End If

    If missingWeightIds.Count > 0 Then
        ackTokens = SCurve_LogGroupedWarningEvents( _
            missingWeightIds, idToWbs, _
            "SCURVE_MISSING_WEIGHT", _
            "Poids manquant sur certaines taches feuilles", _
            "Missing weight on some leaf tasks", _
            "les taches sont exclues de la S-curve ; completer Weight (%) si necessaire", _
            "tasks are excluded from the S-curve; fill Weight (%) if needed")

        SCurve_AddGroupedMessage consoleMessages, "WARNING", missingWeightIds, idToWbs, _
            "Poids manquant sur certaines tâches feuilles - non prises en compte dans la S-curve", "compléter Weight (%) si nécessaire", _
            "Missing weight on some leaf tasks - excluded from S-curve", "fill Weight (%) if needed", _
            True, _
            ackTokens
    End If

    If excludedTaskTypeWithWeightIds.Count > 0 Then
        ackTokens = SCurve_LogGroupedWarningEvents( _
            excludedTaskTypeWithWeightIds, idToWbs, _
            "SCURVE_EXCLUDED_TASK_TYPE_WITH_WEIGHT", _
            "Poids renseigne sur des Milestones ou Level of Effort", _
            "Weight entered on Milestones or Level of Effort", _
            "les taches sont exclues de la S-curve ; supprimer Weight (%) si vous voulez eviter ce warning", _
            "tasks are excluded from the S-curve; remove Weight (%) if you want to avoid this warning")

        SCurve_AddGroupedMessage consoleMessages, "WARNING", excludedTaskTypeWithWeightIds, idToWbs, _
            "Poids renseigné sur des Milestones ou Level of Effort - non pris en compte dans la S-curve", "supprimer le Weight (%) si vous voulez éviter ce warning", _
            "Weight entered on Milestones or Level of Effort - excluded from S-curve", "remove Weight (%) if you want to avoid this warning", _
            True, _
            ackTokens
    End If

    '--------------------------------------------------
    ' PASS 2 : normalize + build daily curves
    '--------------------------------------------------
    outCalc = tblCalcSCurve.DataBodyRange.value

    For r = 1 To rowCount

        Dim currentId As String
        Dim normWeight As Double
        Dim rowBaselineStart As Variant
        Dim rowBaselineFinish As Variant
        Dim rowCalcStart As Variant
        Dim rowCalcFinish As Variant
        Dim rowProgress As Double
        Dim rowRawWeight As Double
        Dim baselineDays As Long
        Dim calcDays As Long
        Dim dailyCalcUnit As Double
        Dim dailyBaselineUnit As Double
        Dim actualDays As Long
        Dim remainingDays As Long
        Dim daySerial As Long
        Dim i As Long

        currentId = Trim$(CStr(outCalc(r, mapCalcSCurve("ID"))))

        If currentId = "" Then GoTo NextPass2Row
        If Trim$(CStr(outCalc(r, mapCalcSCurve("SCurve Included")))) <> "YES" Then GoTo NextPass2Row

        rowRawWeight = CDbl(outCalc(r, mapCalcSCurve("SCurve Weight Raw")))
        normWeight = rowRawWeight / totalRawWeight
        outCalc(r, mapCalcSCurve("SCurve Weight Normalized")) = normWeight

        rowBaselineStart = outCalc(r, mapCalcSCurve("SCurve Baseline Start"))
        rowBaselineFinish = outCalc(r, mapCalcSCurve("SCurve Baseline Finish"))
        rowCalcStart = outCalc(r, mapCalcSCurve("SCurve Calculated Start"))
        rowCalcFinish = outCalc(r, mapCalcSCurve("SCurve Calculated Finish"))
        rowProgress = ProgressToDoubleLocal(outCalc(r, mapCalcSCurve("SCurve Progress")))

        baselineDays = CLng(CDbl(rowBaselineFinish) - CDbl(rowBaselineStart) + 1)
        calcDays = CLng(CDbl(rowCalcFinish) - CDbl(rowCalcStart) + 1)

        dailyBaselineUnit = normWeight / baselineDays
        dailyCalcUnit = normWeight / calcDays

        actualDays = CLng(Fix(calcDays * rowProgress + 0.0000001))
        If actualDays < 0 Then actualDays = 0
        If actualDays > calcDays Then actualDays = calcDays

        remainingDays = calcDays - actualDays

        outCalc(r, mapCalcSCurve("SCurve Baseline Daily")) = dailyBaselineUnit
        outCalc(r, mapCalcSCurve("SCurve Calculated Daily")) = dailyCalcUnit
        outCalc(r, mapCalcSCurve("SCurve Actualized Weight")) = dailyCalcUnit * actualDays

        For daySerial = CLng(CDbl(rowBaselineStart)) To CLng(CDbl(rowBaselineFinish))
            AddValueByDate baselineDailyByDate, daySerial, dailyBaselineUnit
            allDates(CStr(daySerial)) = True
        Next daySerial

        If actualDays > 0 Then
            For i = 0 To actualDays - 1
                daySerial = CLng(CDbl(rowCalcStart)) + i
                AddValueByDate actualizedDailyByDate, daySerial, dailyCalcUnit
                allDates(CStr(daySerial)) = True
            Next i
        End If

        If remainingDays > 0 Then
            For i = actualDays To calcDays - 1
                daySerial = CLng(CDbl(rowCalcStart)) + i
                AddValueByDate remainingDailyByDate, daySerial, dailyCalcUnit
                allDates(CStr(daySerial)) = True
            Next i
        End If

NextPass2Row:
    Next r


    tblCalcSCurve.DataBodyRange.value = outCalc
    FormatCalcSCurveColumns tblCalcSCurve, mapCalcSCurve

    If allDates.Count = 0 Then
        SCurve_AddConsoleMessage consoleMessages, "STOP", _
            "Aucune date exploitable pour générer la S-curve.", _
            "No usable date found to generate the S-curve."
        SCurve_SafeEmptyState
        GoTo SafeExit
    End If

    WriteSCurveDailyTable tblSCurve, mapSCurve, allDates, baselineDailyByDate, actualizedDailyByDate, remainingDailyByDate

    Ensure_SCurve_Chart

SafeExit:
    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection

        SCurve_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur VBA dans Run_SCurve_Engine" & vbCrLf & _
            "-> vérifier le dernier bloc modifié dans mod_SCurve" & vbCrLf & _
            "-> " & Err.Description, _
            "VBA error in Run_SCurve_Engine" & vbCrLf & _
            "-> check the last edited block in mod_SCurve" & vbCrLf & _
            "-> " & Err.Description
    End If


    If Not consoleMessages Is Nothing Then
        If consoleMessages.Count > 0 Then
            CalcBridge_ShowPlanningConsole consoleMessages
        End If
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Valide SCurve Source Columns et signale les incoherences detectees.
' EN: Validates SCurve Source Columns and reports detected inconsistencies.
'------------------------------------------------------------------------------
Private Sub ValidateSCurveSourceColumns(ByVal mapWBS As Object, ByVal mapCalc As Object)

    Dim requiredWbsCols As Variant
    Dim requiredCalcCols As Variant
    Dim c As Variant

    requiredWbsCols = Array( _
        VTS_COL_ID, _
        VTS_COL_WEIGHT_PERCENT, _
        VTS_COL_PROGRESS_PERCENT _
    )

    requiredCalcCols = Array( _
        "ID", _
        "WBS", _
        "Task Type", _
        "Baseline Start", _
        "Baseline Duration", _
        "Calculated Start", _
        "Calculated Finish", _
        "Driving Logic" _
    )

    For Each c In requiredWbsCols
        If Not mapWBS.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 810, , "Missing source column in tbl_WBS: " & CStr(c)
        End If
    Next c

    For Each c In requiredCalcCols
        If Not mapCalc.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 813, , "Missing source column in tbl_CALC: " & CStr(c)
        End If
    Next c

End Sub

'------------------------------------------------------------------------------
' FR: Valide SCurve Output Columns et signale les incoherences detectees.
' EN: Validates SCurve Output Columns and reports detected inconsistencies.
'------------------------------------------------------------------------------
Private Sub ValidateSCurveOutputColumns(ByVal mapSCurve As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        VTS_COL_DATE, _
        VTS_COL_DAILY_BASELINE, _
        VTS_COL_CUMULATIVE_BASELINE, _
        VTS_COL_DAILY_ACTUALIZED, _
        VTS_COL_CUMULATIVE_ACTUALIZED, _
        VTS_COL_DAILY_REMAINING_FORECAST, _
        VTS_COL_CUMULATIVE_REMAINING_FORECAST, _
        VTS_COL_CALCULATED_CURVE_SOLID, _
        VTS_COL_CALCULATED_CURVE_DASHED, _
        VTS_COL_CUMULATIVE_ACTUAL _
    )

    For Each c In requiredCols
        If Not mapSCurve.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 811, , "Missing column in tbl_SCURVE: " & CStr(c)
        End If
    Next c

End Sub

'------------------------------------------------------------------------------
' FR: Valide Calc SCurve Columns et signale les incoherences detectees.
' EN: Validates Calc SCurve Columns and reports detected inconsistencies.
'------------------------------------------------------------------------------
Private Sub ValidateCalcSCurveColumns(ByVal mapCalcSCurve As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        "ID", _
        "SCurve Included", _
        "SCurve Warning", _
        "SCurve Weight Raw", _
        "SCurve Weight Normalized", _
        "SCurve Baseline Start", _
        "SCurve Baseline Finish", _
        "SCurve Calculated Start", _
        "SCurve Calculated Finish", _
        "SCurve Progress", _
        "SCurve Baseline Daily", _
        "SCurve Calculated Daily", _
        "SCurve Actualized Weight", _
        "Error flag" _
    )

    For Each c In requiredCols
        If Not mapCalcSCurve.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 812, , "Missing column in tbl_CALC_SCURVE: " & CStr(c)
        End If
    Next c

End Sub
'------------------------------------------------------------------------------
' FR: Construit la map WBS Parent Maps a partir des donnees fournies par l'appelant.
' EN: Builds the WBS Parent Maps map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Sub BuildWBSParentMaps( _
    ByRef dataWBS As Variant, _
    ByVal mapWBS As Object, _
    ByVal hasChildren As Object, _
    ByVal idToWbs As Object)

    Dim perfScope As clsPerfScope

    Dim r As Long
    Dim rowCount As Long
    Dim currentWBS As String
    Dim parentWbs As String
    Dim taskId As String

    Set perfScope = Profiler_BeginScope("BuildWBSParentMaps", "Dictionary")

    rowCount = UBound(dataWBS, 1)

    For r = 1 To rowCount
        currentWBS = NormalizeWBSLocal(CStr(dataWBS(r, mapWBS(VTS_COL_WBS))))
        taskId = Trim$(CStr(dataWBS(r, mapWBS(VTS_COL_ID))))

        If currentWBS <> "" Then
            parentWbs = GetParentWBS(currentWBS)
            If parentWbs <> "" Then
                hasChildren(parentWbs) = True
            End If
        End If

        If taskId <> "" Then
            idToWbs(taskId) = currentWBS
        End If
    Next r

End Sub


'------------------------------------------------------------------------------
' FR: Traite la reference Resize Table To Row Count sans modifier les donnees d'entree.
' EN: Handles the Resize Table To Row Count reference without mutating input data.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ResizeTableToRowCount(ByVal tbl As ListObject, ByVal targetRows As Long)

    Dim perfScope As clsPerfScope

    Dim currentRows As Long
    Dim r As Long

    Set perfScope = Profiler_BeginScope("ResizeTableToRowCount", "Excel Table Resize")

    If tbl.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = tbl.ListRows.Count
    End If

    If targetRows = 0 Then
        Do While tbl.ListRows.Count > 0
            tbl.ListRows(tbl.ListRows.Count).Delete
        Loop
        Exit Sub
    End If

    If currentRows < targetRows Then
        For r = currentRows + 1 To targetRows
            tbl.ListRows.Add
        Next r
    ElseIf currentRows > targetRows Then
        For r = currentRows To targetRows + 1 Step -1
            tbl.ListRows(r).Delete
        Next r
    End If

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute Value By Date a la structure cible.
' EN: Adds Value By Date to the target structure.
'------------------------------------------------------------------------------
Private Sub AddValueByDate(ByVal dict As Object, ByVal dayKey As Long, ByVal valueToAdd As Double)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("AddValueByDate", "Dictionary")

    If dict.Exists(CStr(dayKey)) Then
        dict(CStr(dayKey)) = CDbl(dict(CStr(dayKey))) + valueToAdd
    Else
        dict(CStr(dayKey)) = valueToAdd
    End If

End Sub


'------------------------------------------------------------------------------
' FR: Ecrit SCurve Daily Table vers le stockage cible.
' EN: Writes SCurve Daily Table to the target storage.
'------------------------------------------------------------------------------
Private Sub WriteSCurveDailyTable( _
    ByVal tblSCurve As ListObject, _
    ByVal mapSCurve As Object, _
    ByVal allDates As Object, _
    ByVal baselineDailyByDate As Object, _
    ByVal actualizedDailyByDate As Object, _
    ByVal remainingDailyByDate As Object)

    Dim perfScope As clsPerfScope

    Dim sortedDates() As Long
    Dim idx As Long
    Dim key As Variant
    Dim outArr() As Variant

    Dim cumulativeBaseline As Double
    Dim cumulativeActualized As Double
    Dim cumulativeRemainingForecast As Double

    Dim currentDateKey As String
    Dim dailyBaseline As Double
    Dim dailyActualized As Double
    Dim dailyRemainingForecast As Double

    Dim todaySerial As Long
    Dim currentDate As Long
    Dim dateFormat As String

    Set perfScope = Profiler_BeginScope("WriteSCurveDailyTable", "Excel Table Write")

    ReDim sortedDates(1 To allDates.Count)

    idx = 0
    For Each key In allDates.Keys
        idx = idx + 1
        sortedDates(idx) = CLng(key)
    Next key

    SortLongArray sortedDates

    ResizeTableToRowCount tblSCurve, UBound(sortedDates)
    ReDim outArr(1 To UBound(sortedDates), 1 To tblSCurve.ListColumns.Count)

    cumulativeBaseline = 0
    cumulativeActualized = 0
    cumulativeRemainingForecast = 0
    todaySerial = CLng(Date)

    For idx = 1 To UBound(sortedDates)

        currentDateKey = CStr(sortedDates(idx))
        currentDate = sortedDates(idx)

        If baselineDailyByDate.Exists(currentDateKey) Then
            dailyBaseline = CDbl(baselineDailyByDate(currentDateKey))
        Else
            dailyBaseline = 0
        End If

        If actualizedDailyByDate.Exists(currentDateKey) Then
            dailyActualized = CDbl(actualizedDailyByDate(currentDateKey))
        Else
            dailyActualized = 0
        End If

        If remainingDailyByDate.Exists(currentDateKey) Then
            dailyRemainingForecast = CDbl(remainingDailyByDate(currentDateKey))
        Else
            dailyRemainingForecast = 0
        End If

        cumulativeBaseline = cumulativeBaseline + dailyBaseline
        cumulativeActualized = cumulativeActualized + dailyActualized
        cumulativeRemainingForecast = cumulativeRemainingForecast + dailyRemainingForecast

        outArr(idx, mapSCurve(VTS_COL_DATE)) = CLng(sortedDates(idx))
        outArr(idx, mapSCurve(VTS_COL_DAILY_BASELINE)) = dailyBaseline
        outArr(idx, mapSCurve(VTS_COL_CUMULATIVE_BASELINE)) = cumulativeBaseline
        outArr(idx, mapSCurve(VTS_COL_DAILY_ACTUALIZED)) = dailyActualized
        outArr(idx, mapSCurve(VTS_COL_CUMULATIVE_ACTUALIZED)) = cumulativeActualized
        outArr(idx, mapSCurve(VTS_COL_DAILY_REMAINING_FORECAST)) = dailyRemainingForecast
        outArr(idx, mapSCurve(VTS_COL_CUMULATIVE_REMAINING_FORECAST)) = cumulativeRemainingForecast
        outArr(idx, mapSCurve(VTS_COL_CALCULATED_CURVE_SOLID)) = CVErr(xlErrNA)
        outArr(idx, mapSCurve(VTS_COL_CALCULATED_CURVE_DASHED)) = CVErr(xlErrNA)
        outArr(idx, mapSCurve(VTS_COL_CUMULATIVE_ACTUAL)) = CVErr(xlErrNA)

        If currentDate <= todaySerial Then
            outArr(idx, mapSCurve(VTS_COL_CUMULATIVE_ACTUAL)) = cumulativeActualized
        End If

    Next idx

    For idx = 1 To UBound(sortedDates)

        currentDate = sortedDates(idx)

        If currentDate < todaySerial Then
            outArr(idx, mapSCurve(VTS_COL_CALCULATED_CURVE_SOLID)) = _
                CDbl(outArr(idx, mapSCurve(VTS_COL_CUMULATIVE_ACTUALIZED))) + CDbl(outArr(idx, mapSCurve(VTS_COL_CUMULATIVE_REMAINING_FORECAST)))
        Else
            outArr(idx, mapSCurve(VTS_COL_CALCULATED_CURVE_SOLID)) = CVErr(xlErrNA)
        End If

        If currentDate >= todaySerial Then
            outArr(idx, mapSCurve(VTS_COL_CALCULATED_CURVE_DASHED)) = _
                CDbl(outArr(idx, mapSCurve(VTS_COL_CUMULATIVE_ACTUALIZED))) + CDbl(outArr(idx, mapSCurve(VTS_COL_CUMULATIVE_REMAINING_FORECAST)))
        Else
            outArr(idx, mapSCurve(VTS_COL_CALCULATED_CURVE_DASHED)) = CVErr(xlErrNA)
        End If

    Next idx

    tblSCurve.DataBodyRange.value = outArr

    dateFormat = Settings_GetDateNumberFormat()
    SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange.NumberFormat = dateFormat
    SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_DAILY_BASELINE).DataBodyRange.NumberFormat = "0.00%"
    SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_CUMULATIVE_BASELINE).DataBodyRange.NumberFormat = "0.00%"
    SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_DAILY_ACTUALIZED).DataBodyRange.NumberFormat = "0.00%"
    SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_CUMULATIVE_ACTUALIZED).DataBodyRange.NumberFormat = "0.00%"
    SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_DAILY_REMAINING_FORECAST).DataBodyRange.NumberFormat = "0.00%"
    SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_CUMULATIVE_REMAINING_FORECAST).DataBodyRange.NumberFormat = "0.00%"
    SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_CALCULATED_CURVE_SOLID).DataBodyRange.NumberFormat = "0.00%"
    SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_CALCULATED_CURVE_DASHED).DataBodyRange.NumberFormat = "0.00%"
    SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_CUMULATIVE_ACTUAL).DataBodyRange.NumberFormat = "0.00%"

End Sub

'------------------------------------------------------------------------------
' FR: Traite la reference Safe Empty State sans modifier les donnees d'entree.
' EN: Handles the Safe Empty State reference without mutating input data.
'------------------------------------------------------------------------------

Public Sub SCurve_SafeEmptyState()

    Dim wsSCurve As Worksheet
    Dim wsCalcSCurve As Worksheet
    Dim tblSCurve As ListObject
    Dim tblCalcSCurve As ListObject

    On Error Resume Next
    Set wsSCurve = ThisWorkbook.Worksheets(SCURVE_SHEET)
    Set wsCalcSCurve = ThisWorkbook.Worksheets(CALC_SCURVE_SHEET)
    Set tblSCurve = wsSCurve.ListObjects(SCURVE_TABLE)
    Set tblCalcSCurve = wsCalcSCurve.ListObjects(CALC_SCURVE_TABLE)
    On Error GoTo 0

    If Not tblSCurve Is Nothing Then ResizeTableToRowCount tblSCurve, 0
    If Not tblCalcSCurve Is Nothing Then ResizeTableToRowCount tblCalcSCurve, 0

    SCurve_DeleteChartIfExists wsSCurve, "cht_SCurve"
    SCurve_DeleteShapeIfExists wsSCurve, "SCURVE_TODAY_LINE"
    SCurve_ClearChartSignature

End Sub
'------------------------------------------------------------------------------
' FR: Trie la valeur Sort Long Array en place selon l'ordre deterministic attendu par le consommateur.
' EN: Sorts the Sort Long Array value in place using the deterministic order expected by the consumer.
'------------------------------------------------------------------------------

Private Sub SortLongArray(ByRef arr() As Long)

    Dim perfScope As clsPerfScope


    Dim i As Long
    Dim j As Long
    Dim tempVal As Long

    Set perfScope = Profiler_BeginScope("SortLongArray", "Sort")

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tempVal = arr(i)
                arr(i) = arr(j)
                arr(j) = tempVal
            End If
        Next j
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Formate Calc SCurve Columns pour l'affichage ou l'ecriture.
' EN: Formats Calc SCurve Columns for display or writing.
'------------------------------------------------------------------------------
Private Sub FormatCalcSCurveColumns(ByVal tblCalcSCurve As ListObject, ByVal mapCalcSCurve As Object)

    If tblCalcSCurve.DataBodyRange Is Nothing Then Exit Sub

    tblCalcSCurve.ListColumns("SCurve Baseline Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalcSCurve.ListColumns("SCurve Baseline Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalcSCurve.ListColumns("SCurve Calculated Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalcSCurve.ListColumns("SCurve Calculated Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"

    tblCalcSCurve.ListColumns("SCurve Weight Raw").DataBodyRange.NumberFormat = "0.00%"
    tblCalcSCurve.ListColumns("SCurve Weight Normalized").DataBodyRange.NumberFormat = "0.00%"
    tblCalcSCurve.ListColumns("SCurve Progress").DataBodyRange.NumberFormat = "0.00%"
    tblCalcSCurve.ListColumns("SCurve Baseline Daily").DataBodyRange.NumberFormat = "0.0000%"
    tblCalcSCurve.ListColumns("SCurve Calculated Daily").DataBodyRange.NumberFormat = "0.0000%"
    tblCalcSCurve.ListColumns("SCurve Actualized Weight").DataBodyRange.NumberFormat = "0.00%"

End Sub

'------------------------------------------------------------------------------
' FR: Normalise WBSLocal dans un format exploitable.
' EN: Normalizes WBSLocal into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizeWBSLocal(ByVal wbs As String) As String

    NormalizeWBSLocal = Replace(Trim$(wbs), ",", ".")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Progress To Double Local sans modifier les donnees d'entree.
' EN: Returns the Progress To Double Local value without mutating input data.
'------------------------------------------------------------------------------

Private Function ProgressToDoubleLocal(ByVal v As Variant) As Double

    If IsNumeric(v) Then
        ProgressToDoubleLocal = CDbl(v)
    Else
        ProgressToDoubleLocal = 0
    End If

    If ProgressToDoubleLocal < 0 Then ProgressToDoubleLocal = 0
    If ProgressToDoubleLocal > 1 Then ProgressToDoubleLocal = 1

End Function

'------------------------------------------------------------------------------
' FR: Affiche SCurve Grouped Message pour l'utilisateur ou le diagnostic.
' EN: Shows SCurve Grouped Message for the user or diagnostics.
'------------------------------------------------------------------------------
Private Sub ShowSCurveGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String, _
    ByVal boxStyle As VbMsgBoxStyle)

    Dim consoleMessages As Collection
    Dim ackTokens As String
    Dim msgType As String

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    If (boxStyle And vbCritical) = vbCritical Then
        msgType = "STOP"
    ElseIf (boxStyle And vbExclamation) = vbExclamation Then
        msgType = "WARNING"
    Else
        msgType = "INFO"
    End If


    Set consoleMessages = New Collection

    SCurve_AddGroupedMessage consoleMessages, msgType, idsDict, idToWbs, _
        frProblem, frAction, enProblem, enAction

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree SCurve Chart si necessaire.
' EN: Ensures or creates SCurve Chart when needed.
'------------------------------------------------------------------------------
Private Sub Ensure_SCurve_Chart()

    Dim perfScope As clsPerfScope

    Dim ws As Worksheet
    Dim chartObj As ChartObject


    Set perfScope = Profiler_BeginScope("Ensure_SCurve_Chart", "Chart")

    Set ws = ThisWorkbook.Worksheets(SCURVE_SHEET)

    On Error Resume Next
    Set chartObj = ws.ChartObjects("cht_SCurve")
    On Error GoTo 0

    If chartObj Is Nothing Then
        Create_SCurve_Chart
    Else
        Update_SCurve_Chart chartObj.Chart
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Language sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Language without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub SCurve_ApplyLanguage(Optional ByVal languageCode As String = "")

    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim ch As Chart

    On Error GoTo ErrHandler

    If Trim$(languageCode) <> "" Then
        SCurve_SetLanguage languageCode
    Else
        EnsureSCurveLanguageInitialized
    End If

    Set ws = ThisWorkbook.Worksheets(SCURVE_SHEET)

    On Error Resume Next
    Set chartObj = ws.ChartObjects("cht_SCurve")
    On Error GoTo ErrHandler

    If chartObj Is Nothing Then Exit Sub
    Set ch = chartObj.Chart

    ch.HasTitle = True
    ch.ChartTitle.Text = SCurve_L("S-Curve", "S-Curve")

    If ch.SeriesCollection.Count >= 1 Then ch.SeriesCollection(1).Name = SCurve_L("Réel journalier", "Daily Actualized")
    If ch.SeriesCollection.Count >= 2 Then ch.SeriesCollection(2).Name = SCurve_L("Prévu restant journalier", "Daily Remaining Forecast")
    If ch.SeriesCollection.Count >= 3 Then ch.SeriesCollection(3).Name = SCurve_L("Référence", "Baseline")
    If ch.SeriesCollection.Count >= 4 Then ch.SeriesCollection(4).Name = SCurve_L("Calculé", "Calculated")
    If ch.SeriesCollection.Count >= 5 Then ch.SeriesCollection(5).Name = SCurve_L("Prévu", "Forecast")
    If ch.SeriesCollection.Count >= 6 Then ch.SeriesCollection(6).Name = SCurve_L("Réel", "Actual")

    On Error Resume Next
    ch.Axes(xlCategory).TickLabels.NumberFormat = SCurve_DateAxisNumberFormat()
    On Error GoTo ErrHandler

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "SCurve_ApplyLanguage", Err.Description

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Set Language dans l'etat runtime du composant.
' EN: Activates or initializes Set Language in the component runtime state.
'------------------------------------------------------------------------------

Public Sub SCurve_SetLanguage(ByVal languageCode As String)

    Select Case UCase$(Trim$(languageCode))
        Case "FR"
            gSCurveLanguage = "FR"
        Case "EN"
            gSCurveLanguage = "EN"
        Case Else
            gSCurveLanguage = "EN"
    End Select

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Language sans exposer de mutateur sur l'etat source.
' EN: Returns the Language value without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function SCurve_CurrentLanguage() As String

    EnsureSCurveLanguageInitialized
    SCurve_CurrentLanguage = gSCurveLanguage

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree SCurve Language Initialized si necessaire.
' EN: Ensures or creates SCurve Language Initialized when needed.
'------------------------------------------------------------------------------
Private Sub EnsureSCurveLanguageInitialized()

    If UCase$(Trim$(gSCurveLanguage)) <> "FR" And UCase$(Trim$(gSCurveLanguage)) <> "EN" Then
        gSCurveLanguage = "EN"
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur L sans modifier les donnees d'entree.
' EN: Returns the L value without mutating input data.
'------------------------------------------------------------------------------

Private Function SCurve_L(ByVal frText As String, ByVal enText As String) As String

    If SCurve_CurrentLanguage() = "FR" Then
        SCurve_L = frText
    Else
        SCurve_L = enText
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Date Axis Number Format sans modifier les donnees d'entree.
' EN: Returns the Date Axis Number Format value without mutating input data.
'------------------------------------------------------------------------------

Private Function SCurve_DateAxisNumberFormat() As String

    SCurve_DateAxisNumberFormat = Settings_GetDateNumberFormat()

End Function


'------------------------------------------------------------------------------
' FR: Cree SCurve Chart pour le domaine S-Curve.
' EN: Creates SCurve Chart for the S-Curve domain.
'------------------------------------------------------------------------------
Private Sub Create_SCurve_Chart()

    Dim perfScope As clsPerfScope

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim chartObj As ChartObject
    Dim ch As Chart
    Dim s As Series


    Set perfScope = Profiler_BeginScope("Create_SCurve_Chart", "Chart")
    SCurve_SetLanguage Settings_GetOwnerLanguage("SCURVE")

    Set ws = ThisWorkbook.Worksheets(SCURVE_SHEET)
    Set tbl = ws.ListObjects(SCURVE_TABLE)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("L2").Left, _
        Top:=ws.Range("L2").Top, _
        Width:=1100, _
        Height:=500)

    chartObj.Name = "cht_SCurve"
    Set ch = chartObj.Chart

    Do While ch.SeriesCollection.Count > 0
        ch.SeriesCollection(1).Delete
    Loop

    ch.HasTitle = True
    ch.ChartTitle.Text = "S-Curve"
    ch.HasLegend = True

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = SCurve_L("Réel journalier", "Daily Actualized")
        .XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
        .Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DAILY_ACTUALIZED).DataBodyRange
        .AxisGroup = xlPrimary
        .ChartType = xlColumnStacked
        .Format.Fill.ForeColor.RGB = RGB(31, 78, 121)
        .Format.Line.Visible = msoFalse
    End With

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = SCurve_L("Prévu restant journalier", "Daily Remaining Forecast")
        .XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
        .Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DAILY_REMAINING_FORECAST).DataBodyRange
        .AxisGroup = xlPrimary
        .ChartType = xlColumnStacked
        .Format.Fill.ForeColor.RGB = RGB(191, 191, 191)
        .Format.Line.Visible = msoFalse
    End With

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = SCurve_L("Référence", "Baseline")
        .XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
        .Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_CUMULATIVE_BASELINE).DataBodyRange
        .AxisGroup = xlSecondary
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(166, 166, 166)
        .Format.Line.Weight = 1.5
        .MarkerStyle = xlMarkerStyleNone
    End With

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = SCurve_L("Calculé", "Calculated")
        .XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
        .Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_CALCULATED_CURVE_SOLID).DataBodyRange
        .AxisGroup = xlSecondary
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(68, 114, 196)
        .Format.Line.Weight = 2.75
        .MarkerStyle = xlMarkerStyleNone
    End With

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = SCurve_L("Prévu", "Forecast")
        .XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
        .Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_CALCULATED_CURVE_DASHED).DataBodyRange
        .AxisGroup = xlSecondary
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(157, 195, 230)
        .Format.Line.Weight = 2.75
        .Format.Line.DashStyle = msoLineDash
        .MarkerStyle = xlMarkerStyleNone
    End With

    Set s = ch.SeriesCollection.NewSeries
    With s
        .Name = SCurve_L("Réel", "Actual")
        .XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
        .Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_CUMULATIVE_ACTUAL).DataBodyRange
        .AxisGroup = xlSecondary
        .ChartType = xlLine
        .Format.Line.ForeColor.RGB = RGB(31, 78, 121)
        .Format.Line.Weight = 2.75
        .MarkerStyle = xlMarkerStyleNone
    End With

    SCurveStoreChartSignature SCurveChartSourceKey(tbl)

    Format_SCurve_Chart ch
    DrawTodayVerticalLine ch
End Sub


'------------------------------------------------------------------------------
' FR: Actualise Update S Curve Chart sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Update S Curve Chart without changing the business rules that produce the data.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub Update_SCurve_Chart(ByVal ch As Chart)

    Dim perfScope As clsPerfScope

    Dim tbl As ListObject
    Dim rebuildNeeded As Boolean
    Dim sourceKey As String

    Set perfScope = Profiler_BeginScope("Update_SCurve_Chart", "Chart")

    Set tbl = ThisWorkbook.Worksheets(SCURVE_SHEET).ListObjects(SCURVE_TABLE)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    rebuildNeeded = False

    If ch.SeriesCollection.Count <> 6 Then
        rebuildNeeded = True
    Else
        On Error Resume Next
        rebuildNeeded = ( _
            ch.SeriesCollection(1).ChartType <> xlColumnStacked Or _
            ch.SeriesCollection(2).ChartType <> xlColumnStacked Or _
            ch.SeriesCollection(1).AxisGroup <> xlPrimary Or _
            ch.SeriesCollection(2).AxisGroup <> xlPrimary Or _
            ch.SeriesCollection(3).ChartType <> xlLine Or _
            ch.SeriesCollection(3).AxisGroup <> xlSecondary Or _
            ch.SeriesCollection(4).ChartType <> xlLine Or _
            ch.SeriesCollection(4).AxisGroup <> xlSecondary Or _
            ch.SeriesCollection(5).ChartType <> xlLine Or _
            ch.SeriesCollection(5).AxisGroup <> xlSecondary Or _
            ch.SeriesCollection(6).ChartType <> xlLine Or _
            ch.SeriesCollection(6).AxisGroup <> xlSecondary)
        On Error GoTo 0
    End If

    If rebuildNeeded Then
        ch.Parent.Delete
        Create_SCurve_Chart
        Exit Sub
    End If

    sourceKey = SCurveChartSourceKey(tbl)

    If SCurveStoredChartSignature() <> sourceKey Then
        SCurveAssignChartSeriesRanges ch, tbl
        SCurveStoreChartSignature sourceKey
    End If

    Format_SCurve_Chart ch
    DrawTodayVerticalLine ch
End Sub


'------------------------------------------------------------------------------
' FR: Traite la reference S Curve Assign Chart Series Ranges sans modifier les donnees d'entree.
' EN: Handles the S Curve Assign Chart Series Ranges reference without mutating input data.
'------------------------------------------------------------------------------

Private Sub SCurveAssignChartSeriesRanges(ByVal ch As Chart, ByVal tbl As ListObject)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("SCurveAssignChartSeriesRanges", "Chart")
    Profiler_RecordCounter "SCurveChartSourceAssignments", 1

    ch.SeriesCollection(1).XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
    ch.SeriesCollection(1).Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DAILY_ACTUALIZED).DataBodyRange

    ch.SeriesCollection(2).XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
    ch.SeriesCollection(2).Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DAILY_REMAINING_FORECAST).DataBodyRange

    ch.SeriesCollection(3).XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
    ch.SeriesCollection(3).Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_CUMULATIVE_BASELINE).DataBodyRange

    ch.SeriesCollection(4).XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
    ch.SeriesCollection(4).Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_CALCULATED_CURVE_SOLID).DataBodyRange

    ch.SeriesCollection(5).XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
    ch.SeriesCollection(5).Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_CALCULATED_CURVE_DASHED).DataBodyRange

    ch.SeriesCollection(6).XValues = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
    ch.SeriesCollection(6).Values = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_CUMULATIVE_ACTUAL).DataBodyRange

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Delete Chart If Exists dans le perimetre possede par le composant.
' EN: Resets Delete Chart If Exists within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub SCurve_DeleteChartIfExists(ByVal ws As Worksheet, ByVal chartName As String)

    Dim chartObj As ChartObject

    If ws Is Nothing Then Exit Sub
    If Len(Trim$(chartName)) = 0 Then Exit Sub

    On Error Resume Next
    Set chartObj = ws.ChartObjects(chartName)
    If Not chartObj Is Nothing Then chartObj.Delete
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Delete Shape If Exists dans le perimetre possede par le composant.
' EN: Resets Delete Shape If Exists within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub SCurve_DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)

    If ws Is Nothing Then Exit Sub
    If Len(Trim$(shapeName)) = 0 Then Exit Sub

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Clear Chart Signature dans le perimetre possede par le composant.
' EN: Resets Clear Chart Signature within the state owned by the component.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub SCurve_ClearChartSignature()

    On Error Resume Next
    ThisWorkbook.Names(SCURVE_CHART_SIGNATURE_NAME).Delete
    On Error GoTo 0

End Sub
'------------------------------------------------------------------------------
' FR: Retourne la reference S Curve Chart Source Signature sans modifier les donnees d'entree.
' EN: Returns the S Curve Chart Source Signature reference without mutating input data.
'------------------------------------------------------------------------------

Private Function SCurveChartSourceSignature(ByVal tbl As ListObject) As String

    SCurveChartSourceSignature = _
        SCurveRangeSignature(tbl, VTS_COL_DATE) & "|" & _
        SCurveRangeSignature(tbl, VTS_COL_DAILY_ACTUALIZED) & "|" & _
        SCurveRangeSignature(tbl, VTS_COL_DAILY_REMAINING_FORECAST) & "|" & _
        SCurveRangeSignature(tbl, VTS_COL_CUMULATIVE_BASELINE) & "|" & _
        SCurveRangeSignature(tbl, VTS_COL_CALCULATED_CURVE_SOLID) & "|" & _
        SCurveRangeSignature(tbl, VTS_COL_CALCULATED_CURVE_DASHED) & "|" & _
        SCurveRangeSignature(tbl, VTS_COL_CUMULATIVE_ACTUAL)

End Function

'------------------------------------------------------------------------------
' FR: Retourne la reference S Curve Range Signature sans modifier les donnees d'entree.
' EN: Returns the S Curve Range Signature reference without mutating input data.
'------------------------------------------------------------------------------

Private Function SCurveRangeSignature(ByVal tbl As ListObject, ByVal columnKey As String) As String

    Dim rng As Range

    Set rng = SchemaListColumn(tbl, VTS_TABLE_SCURVE, columnKey).DataBodyRange
    SCurveRangeSignature = columnKey & ":" & _
        rng.rows.Count & ":" & _
        rng.Address(RowAbsolute:=True, ColumnAbsolute:=True, ReferenceStyle:=xlA1, External:=True)

End Function

'------------------------------------------------------------------------------
' FR: Construit la cle fixe versionnee representant la signature source brute.
' EN: Builds the fixed versioned key representing the raw source signature.
'------------------------------------------------------------------------------
Private Function SCurveChartSourceKey(ByVal tbl As ListObject) As String

    SCurveChartSourceKey = SCURVE_CHART_SIGNATURE_VERSION & _
        DeterministicDigest_SHA256Hex(SCurveChartSourceSignature(tbl))

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur S Curve Stored Chart Signature sans modifier les donnees d'entree.
' EN: Returns the S Curve Stored Chart Signature value without mutating input data.
'------------------------------------------------------------------------------

Private Function SCurveStoredChartSignature() As String

    Dim nm As Name
    Dim rawValue As String

    On Error Resume Next
    Set nm = ThisWorkbook.Names(SCURVE_CHART_SIGNATURE_NAME)
    On Error GoTo 0

    If nm Is Nothing Then Exit Function

    rawValue = CStr(nm.RefersTo)
    If Len(rawValue) >= 3 Then
        If Left$(rawValue, 2) = "=""" And Right$(rawValue, 1) = """" Then
            rawValue = Mid$(rawValue, 3, Len(rawValue) - 3)
            rawValue = Replace(rawValue, """""", """")
        End If
    End If

    SCurveStoredChartSignature = rawValue

End Function

'------------------------------------------------------------------------------
' FR: Persiste la signature de source du graphique S-Curve utilisee pour eviter les reconstructions inutiles.
' EN: Persists the S-Curve chart source signature used to avoid unnecessary rebuilds.
'------------------------------------------------------------------------------

Private Sub SCurveStoreChartSignature(ByVal signatureKey As String)

    Dim nm As Name
    Dim refersToValue As String

    If Len(signatureKey) <> SCURVE_CHART_SIGNATURE_KEY_LENGTH Or _
       Left$(signatureKey, Len(SCURVE_CHART_SIGNATURE_VERSION)) <> _
            SCURVE_CHART_SIGNATURE_VERSION Then
        Err.Raise vbObjectError + 2713, "SCurveStoreChartSignature", _
            "Invalid fixed-length S-Curve chart signature key."
    End If

    refersToValue = "=""" & signatureKey & """"

    On Error Resume Next
    Set nm = ThisWorkbook.Names(SCURVE_CHART_SIGNATURE_NAME)
    On Error GoTo 0

    If nm Is Nothing Then
        Set nm = ThisWorkbook.Names.Add(Name:=SCURVE_CHART_SIGNATURE_NAME, RefersTo:=refersToValue)
    Else
        nm.RefersTo = refersToValue
    End If

    On Error Resume Next
    nm.Visible = False
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Formate SCurve Chart pour l'affichage ou l'ecriture.
' EN: Formats SCurve Chart for display or writing.
'------------------------------------------------------------------------------
Private Sub Format_SCurve_Chart(ByVal ch As Chart)


    With ch.Axes(xlValue, xlPrimary)
        .MinimumScale = 0
        .MaximumScale = 0.12
        .MajorUnit = 0.02
        .TickLabels.NumberFormat = "0%"
        .HasMajorGridlines = False
    End With

    With ch.Axes(xlValue, xlSecondary)
        .MinimumScale = 0
        .MaximumScale = 1
        .MajorUnit = 0.1
        .TickLabels.NumberFormat = "0%"
        .HasMajorGridlines = True
    End With

    With ch.Axes(xlCategory)
        .TickLabels.NumberFormat = SCurve_DateAxisNumberFormat()
        .TickLabels.Orientation = 45
        .CategoryType = xlCategoryScale
    End With

    ch.Legend.Position = xlLegendPositionRight

    On Error Resume Next
    ch.ChartGroups(1).GapWidth = 35
    ch.ChartGroups(1).Overlap = 100
    On Error GoTo 0

End Sub


'------------------------------------------------------------------------------
' FR: Actualise Draw Today Vertical Line sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Draw Today Vertical Line without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub DrawTodayVerticalLine(ByVal ch As Chart)

    Dim perfScope As clsPerfScope

    Dim ws As Worksheet
    Dim tbl As ListObject

    Set perfScope = Profiler_BeginScope("DrawTodayVerticalLine", "Chart")

    Set ws = ThisWorkbook.Worksheets(SCURVE_SHEET)
    Set tbl = ws.ListObjects(SCURVE_TABLE)

    DrawSCurveTodayVerticalLine ch, ws, tbl, "SCURVE_TODAY_LINE"

End Sub


'------------------------------------------------------------------------------
' FR: Actualise Draw S Curve Today Vertical Line sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Draw S Curve Today Vertical Line without changing the business rules that produce the data.
' FR - Effet de bord : cree ou met a jour des shapes Excel.
' EN - Side effect: creates or updates Excel shapes.
'------------------------------------------------------------------------------

Public Sub DrawSCurveTodayVerticalLine( _
    ByVal ch As Chart, _
    ByVal ws As Worksheet, _
    ByVal tbl As ListObject, _
    ByVal todayShapeName As String)

    Dim perfScope As clsPerfScope

    Dim shp As Shape
    Dim todaySerial As Double
    Dim minDate As Double
    Dim maxDate As Double
    Dim plotLeft As Double
    Dim plotTop As Double
    Dim plotWidth As Double
    Dim plotHeight As Double
    Dim xPos As Double
    Dim ratio As Double
    Dim chartObj As ChartObject
    Dim dateRange As Range
    Dim dateCount As Long
    Dim categoryIndex As Double
    Dim r As Long
    Dim currentDate As Double
    Dim nextDate As Double
    Dim axisBetweenCategories As Boolean

    Set perfScope = Profiler_BeginScope("DrawSCurveTodayVerticalLine", "Shape Create")

    If ch Is Nothing Then Exit Sub
    If ws Is Nothing Then Exit Sub
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    If Len(Trim$(todayShapeName)) = 0 Then Exit Sub

    Set chartObj = ch.Parent
    Set dateRange = SchemaListColumn(tbl, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange

    todaySerial = CDbl(Date)
    dateCount = dateRange.rows.Count

    If dateCount < 1 Then GoTo HideLine

    minDate = CDbl(dateRange.Cells(1, 1).value)
    maxDate = CDbl(dateRange.Cells(dateCount, 1).value)

    On Error Resume Next
    Set shp = ws.Shapes(todayShapeName)
    On Error GoTo 0

    If todaySerial < minDate Or todaySerial > maxDate Then GoTo HideLine
    If maxDate <= minDate Then GoTo HideLine
    If dateCount < 2 Then GoTo HideLine

    categoryIndex = 0#
    For r = 1 To dateCount
        currentDate = CDbl(dateRange.Cells(r, 1).value)

        If currentDate = todaySerial Then
            categoryIndex = CDbl(r)
            Exit For
        End If

        If r < dateCount Then
            nextDate = CDbl(dateRange.Cells(r + 1, 1).value)
            If todaySerial > currentDate And todaySerial < nextDate Then
                categoryIndex = CDbl(r) + ((todaySerial - currentDate) / (nextDate - currentDate))
                Exit For
            End If
        End If
    Next r

    If categoryIndex = 0# Then GoTo HideLine

    plotLeft = chartObj.Left + ch.PlotArea.InsideLeft
    plotTop = chartObj.Top + ch.PlotArea.InsideTop
    plotWidth = ch.PlotArea.InsideWidth
    plotHeight = ch.PlotArea.InsideHeight

    axisBetweenCategories = False
    On Error Resume Next
    axisBetweenCategories = ch.Axes(xlCategory).axisBetweenCategories
    On Error GoTo 0

    If axisBetweenCategories Then
        ratio = (categoryIndex - 0.5) / CDbl(dateCount)
    Else
        ratio = (categoryIndex - 1#) / CDbl(dateCount - 1)
    End If

    xPos = plotLeft + (plotWidth * ratio)

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddLine(xPos, plotTop, xPos, plotTop + plotHeight)
        shp.Name = todayShapeName
    Else
        shp.Visible = msoTrue
        shp.Left = xPos
        shp.Top = plotTop
        shp.Width = 0
        shp.Height = plotHeight
    End If

    SCurveFormatTodayLine shp
    GoTo Done

HideLine:
    If Not shp Is Nothing Then shp.Visible = msoFalse

Done:

End Sub
'------------------------------------------------------------------------------
' FR: Traite la reference S Curve Format Today Line sans modifier les donnees d'entree.
' EN: Handles the S Curve Format Today Line reference without mutating input data.
'------------------------------------------------------------------------------

Private Sub SCurveFormatTodayLine(ByVal shp As Shape)

    With shp.Line
        .ForeColor.RGB = RGB(0, 176, 80)
        .Weight = 1.5
        .DashStyle = msoLineDash
    End With

End Sub
'------------------------------------------------------------------------------
' FR: Indique si SCurve Excluded Task Type Local est vrai pour le contexte courant.
' EN: Returns whether SCurve Excluded Task Type Local is true for the current context.
'------------------------------------------------------------------------------
Private Function IsSCurveExcludedTaskTypeLocal(ByVal taskTypeValue As Variant) As Boolean

    Dim s As String

    s = UCase$(Trim$(CStr(taskTypeValue)))
    s = Replace$(s, "-", " ")
    s = Replace$(s, "_", " ")

    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace$(s, "  ", " ")
    Loop

    IsSCurveExcludedTaskTypeLocal = _
        (s = "MILESTONE") Or _
        (s = "MS") Or _
        (s = "JALON") Or _
        (s = "LEVEL OF EFFORT") Or _
        (s = "LOE")

End Function

'------------------------------------------------------------------------------
' FR: Construit la projection S-Curve consommee par Dashboard sans exposer les tables proprietaires.
' EN: Builds the S-Curve projection consumed by Dashboard without exposing owner tables.
'------------------------------------------------------------------------------
Public Function SCurve_BuildDashboardProjection() As Object

    Dim projection As Object
    Dim wsSCurve As Worksheet
    Dim wsCalcSCurve As Worksheet
    Dim tblSCurve As ListObject
    Dim tblCalcSCurve As ListObject
    Dim mapSCurve As Object
    Dim mapCalcSCurve As Object
    Dim arr As Variant
    Dim r As Long
    Dim value As Variant
    Dim dateValue As Variant
    Dim dateNumber As Double
    Dim todaySerial As Long
    Dim actualProgress As Double
    Dim plannedProgress As Double

    Set projection = CreateObject("Scripting.Dictionary")
    projection("ActualProgress") = 0#
    projection("PlannedProgress") = 0#

    Set wsSCurve = ThisWorkbook.Worksheets(SCURVE_SHEET)
    Set wsCalcSCurve = ThisWorkbook.Worksheets(CALC_SCURVE_SHEET)
    Set tblSCurve = wsSCurve.ListObjects(SCURVE_TABLE)
    Set tblCalcSCurve = wsCalcSCurve.ListObjects(CALC_SCURVE_TABLE)

    Set mapSCurve = CanonicalIdentity_BuildColumnMap(tblSCurve)
    Set mapCalcSCurve = CanonicalIdentity_BuildColumnMap(tblCalcSCurve)

    If Not tblSCurve.DataBodyRange Is Nothing Then
        If mapSCurve.Exists(VTS_COL_DATE) Then Set projection("DateRange") = SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_DATE).DataBodyRange
        If mapSCurve.Exists(VTS_COL_CUMULATIVE_BASELINE) Then Set projection("BaselineRange") = SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_CUMULATIVE_BASELINE).DataBodyRange
        If mapSCurve.Exists(VTS_COL_CUMULATIVE_ACTUAL) Then Set projection("ActualRange") = SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_CUMULATIVE_ACTUAL).DataBodyRange
        If mapSCurve.Exists(VTS_COL_CALCULATED_CURVE_DASHED) Then Set projection("ForecastRange") = SchemaListColumn(tblSCurve, VTS_TABLE_SCURVE, VTS_COL_CALCULATED_CURVE_DASHED).DataBodyRange

        If mapSCurve.Exists(VTS_COL_DATE) And mapSCurve.Exists(VTS_COL_CUMULATIVE_BASELINE) Then
            arr = tblSCurve.DataBodyRange.value
            todaySerial = CLng(Date)
            For r = 1 To UBound(arr, 1)
                dateValue = GetCellValue(arr(r, mapSCurve(VTS_COL_DATE)))
                If HasValue(dateValue) And (IsNumeric(dateValue) Or IsDate(dateValue)) Then
                    If IsDate(dateValue) Then dateNumber = CDbl(CDate(dateValue)) Else dateNumber = CDbl(dateValue)
                    If CLng(dateNumber) <= todaySerial Then
                        value = GetCellValue(arr(r, mapSCurve(VTS_COL_CUMULATIVE_BASELINE)))
                        If IsNumeric(value) Then plannedProgress = CDbl(value)
                    End If
                End If
            Next r
        End If
    End If

    If Not tblCalcSCurve.DataBodyRange Is Nothing Then
        If mapCalcSCurve.Exists("SCurve Actualized Weight") Then
            arr = tblCalcSCurve.DataBodyRange.value
            For r = 1 To UBound(arr, 1)
                value = GetCellValue(arr(r, mapCalcSCurve("SCurve Actualized Weight")))
                If IsNumeric(value) Then actualProgress = actualProgress + CDbl(value)
            Next r
        End If
    End If

    projection("ActualProgress") = actualProgress
    projection("PlannedProgress") = plannedProgress
    Set SCurve_BuildDashboardProjection = projection

End Function

'------------------------------------------------------------------------------
' FR: Dessine la ligne Today du Dashboard en resolvant la table S-Curve dans son domaine proprietaire.
' EN: Draws the Dashboard Today line while resolving the S-Curve table inside its owner domain.
'------------------------------------------------------------------------------
Public Sub SCurve_DrawDashboardTodayLine(ByVal ch As Chart, ByVal ws As Worksheet, ByVal shapeName As String)

    Dim tblSCurve As ListObject
    Set tblSCurve = ThisWorkbook.Worksheets(SCURVE_SHEET).ListObjects(SCURVE_TABLE)
    DrawSCurveTodayVerticalLine ch, ws, tblSCurve, shapeName

End Sub
