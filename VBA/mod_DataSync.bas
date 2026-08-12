Attribute VB_Name = "mod_DataSync"
Option Explicit

'===============================================================================
' MODULE : mod_DataSync
' DOMAINE / DOMAIN : WBS / CALC Synchronization
'
' FR
' Synchronise les datasets canoniques WBS, CALC et LOGIC_LINKS avant le calcul.
' Ne calcule pas les dates planning et ne rend aucune vue.
'
' EN
' Synchronizes canonical WBS, CALC and LOGIC_LINKS datasets before calculation.
' Does not calculate planning dates or render views.
'
' CONTRATS / CONTRACTS : Sync_WBS_To_CALC, Planning_WBSIsEmpty, Planning_CalcSafeEmptyState, Planning_GanttSafeEmptyState, Planning_FullSafeEmptyState, Sync_Forecast_Only, RebuildLogicLinksTable
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const LOGIC_LINKS_TABLE_NAME As String = "tbl_LOGIC_LINKS"
Private Const LOGIC_LINKS_FIRST_CELL As String = "Z1"


'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Sync WBS To CALC dans le stockage possede par le domaine.
' EN: Writes or synchronizes Sync WBS To CALC in the store owned by the domain.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: writes to an Excel table owned by the workflow.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Sub Sync_WBS_To_CALC(Optional ByVal preserveCalcOutputs As Boolean = False)

    Dim perfScope As clsPerfScope

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject

    Dim mapCalc As Object
    Dim mapWBS As Object
    Dim wbsIdRows As Object
    Dim wbsToId As Object
    Dim summaryWbsByWbs As Object

    Dim colsToCopy As Variant
    Dim colsToClear As Variant

    Dim i As Long
    Dim r As Long
    Dim maxId As Long
    Dim currentCalcRows As Long
    Dim targetRows As Long

    Dim idValue As Variant
    Dim idKey As String
    Dim wbsValue As String
    Dim wbsRowIndex As Long

    Dim predWbs As String
    Dim predRaw As String
    Dim taskTypeVal As String
    Dim calVal As String

    Dim consoleMessages As Collection

    Set perfScope = Profiler_BeginScope("Sync_WBS_To_CALC", "Excel Table Sync")

    On Error GoTo SafeExit

    Set consoleMessages = New Collection

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")

    EnsureWBSTaskTypeInputSetup tblWBS
    EnsureSummaryDisplayColumnExists tblWBS, Nothing
    EnsureWBSCalendarInputSetup tblWBS

    Ensure_Calc_Infrastructure

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    EnsureTaskTypeColumnExists tblWBS, tblCalc
    EnsureSummaryDisplayColumnExists tblWBS, tblCalc
    EnsureCalendarColumnExists tblWBS, tblCalc
    EnsureDeadlineOutputColumnsExist tblWBS, tblCalc
    EnsureLongestPathOutputColumnsExist tblWBS, tblCalc

    Set mapCalc = CreateObject("Scripting.Dictionary")
    Set mapWBS = CreateObject("Scripting.Dictionary")
    Set wbsIdRows = CreateObject("Scripting.Dictionary")
    Set wbsToId = CreateObject("Scripting.Dictionary")

    colsToCopy = Array( _
        "WBS", _
        "Task Name", _
        "Task Type", _
        "S", _
        "Cal", _
        "Predecessors WBS", _
        "Baseline Start", _
        "Baseline Duration", _
        "Baseline Finish", _
        "Actual Start", _
        "Actual Finish", _
        "Actual Duration", _
        "Forecast Start", _
        "Forecast Finish" _
    )

    colsToClear = Array( _
        "ParentID", _
        "IsSummary", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration", _
        "Driving Logic", _
        "Error flag", _
        "ErrorMsg", _
        "Critical Path", _
        "Longest Path", _
        "Total Float", _
        "Free Float", _
        "Critical Path REX", _
        "Total Float REX", _
        "Free Float REX", _
        "Deadline Float" _
    )

    For i = 1 To tblWBS.ListColumns.Count
        mapWBS(tblWBS.ListColumns(i).Name) = i
    Next i

    For i = 1 To tblCalc.ListColumns.Count
        mapCalc(tblCalc.ListColumns(i).Name) = i
    Next i

    If Not mapWBS.Exists("ID") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_WBS.", _
            "Column ID was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists("WBS") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne WBS est introuvable dans tbl_WBS.", _
            "Column WBS was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists("Task Type") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Task Type est introuvable dans tbl_WBS.", _
            "Column Task Type was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists("S") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne S est introuvable dans tbl_WBS.", _
            "Column S was not found in tbl_WBS."
        GoTo SafeExit
    End If
    If Not mapWBS.Exists("Predecessors WBS") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Predecessors WBS est introuvable dans tbl_WBS.", _
            "Column Predecessors WBS was not found in WBS."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("ID") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_CALC.", _
            "Column ID was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("WBS") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne WBS est introuvable dans tbl_CALC.", _
            "Column WBS was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("Task Type") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Task Type est introuvable dans tbl_CALC.", _
            "Column Task Type was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("S") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne S est introuvable dans tbl_CALC.", _
            "Column S was not found in tbl_CALC."
        GoTo SafeExit
    End If
    If Not mapCalc.Exists("Predecessors WBS") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Predecessors WBS est introuvable dans tbl_CALC.", _
            "Column Predecessors WBS was not found in tbl_CALC."
        GoTo SafeExit
    End If

    maxId = 0

    If Not tblWBS.DataBodyRange Is Nothing Then
        For r = 1 To tblWBS.ListRows.Count

            idValue = tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value
            wbsValue = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("WBS")).value))
            wbsValue = Replace(wbsValue, ",", ".")

            If Trim$(CStr(idValue)) <> "" Then

                If Not IsNumeric(idValue) Then
                    DataSync_AddConsoleMessage consoleMessages, "STOP", _
                        "ID non numérique détecté dans WBS : " & CStr(idValue), _
                        "Non-numeric ID detected in WBS: " & CStr(idValue)
                    GoTo SafeExit
                End If

                If CLng(idValue) < 1 Then
                    DataSync_AddConsoleMessage consoleMessages, "STOP", _
                        "ID invalide dans WBS (doit être >= 1) : " & CStr(idValue), _
                        "Invalid ID in WBS (must be >= 1): " & CStr(idValue)
                    GoTo SafeExit
                End If

                idKey = CStr(CLng(idValue))
                wbsIdRows(idKey) = r

                If CLng(idValue) > maxId Then
                    maxId = CLng(idValue)
                End If

                If wbsValue <> "" Then
                    If wbsToId.Exists(wbsValue) Then
                        DataSync_AddConsoleMessage consoleMessages, "STOP", _
                            "WBS dupliqué détecté dans WBS : " & wbsValue, _
                            "Duplicate WBS detected in WBS: " & wbsValue
                        GoTo SafeExit
                    Else
                        wbsToId(wbsValue) = CLng(idValue)
                    End If
                End If

            End If

        Next r
    End If

    Set summaryWbsByWbs = BuildWBSSummaryWbsLookup(tblWBS, mapWBS)
    NormalizeWBSSummaryDisplayValues tblWBS, mapWBS, summaryWbsByWbs
    NormalizeWBSCalendarValues tblWBS, mapWBS, summaryWbsByWbs

    If tblCalc.DataBodyRange Is Nothing Then
        currentCalcRows = 0
    Else
        currentCalcRows = tblCalc.ListRows.Count
    End If

    targetRows = maxId

    If targetRows = 0 Then
        Do While tblCalc.ListRows.Count > 0
            tblCalc.ListRows(tblCalc.ListRows.Count).Delete
        Loop
        GoTo SafeExit
    End If

    If currentCalcRows < targetRows Then
        For r = currentCalcRows + 1 To targetRows
            tblCalc.ListRows.Add
        Next r
    ElseIf currentCalcRows > targetRows Then
        For r = currentCalcRows To targetRows + 1 Step -1
            tblCalc.ListRows(r).Delete
        Next r
    End If

    For r = 1 To targetRows
        tblCalc.DataBodyRange.Cells(r, mapCalc("ID")).value = r
    Next r

    For r = 1 To targetRows

        idKey = CStr(r)

        If wbsIdRows.Exists(idKey) Then

            wbsRowIndex = wbsIdRows(idKey)

            For i = LBound(colsToCopy) To UBound(colsToCopy)
                If mapWBS.Exists(colsToCopy(i)) And mapCalc.Exists(colsToCopy(i)) Then

                    If colsToCopy(i) = "WBS" Or colsToCopy(i) = "Predecessors WBS" Then
                        With tblCalc.DataBodyRange.Cells(r, mapCalc(colsToCopy(i)))
                            .NumberFormat = "@"
                            .value = CStr(tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS(colsToCopy(i))).value)
                        End With

                    ElseIf colsToCopy(i) = "Task Type" Then
                        taskTypeVal = NormalizeTaskTypeValue(tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS("Task Type")).value)
                        With tblCalc.DataBodyRange.Cells(r, mapCalc("Task Type"))
                            .NumberFormat = "@"
                            .value = taskTypeVal
                        End With

                    ElseIf colsToCopy(i) = "S" Then
                        With tblCalc.DataBodyRange.Cells(r, mapCalc("S"))
                            .NumberFormat = "@"
                            .value = NormalizeSummaryDisplayValue(tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS("S")).value)
                        End With
                    ElseIf colsToCopy(i) = "Cal" Then
                        If IsCalendarIgnoredWBSRow(tblWBS, mapWBS, wbsRowIndex, summaryWbsByWbs) Then
                            calVal = vbNullString
                        Else
                            calVal = NormalizeCalendarType(tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS("Cal")).value)
                        End If
                        With tblCalc.DataBodyRange.Cells(r, mapCalc("Cal"))
                            .NumberFormat = "@"
                            .value = calVal
                        End With

                    Else
                        tblCalc.DataBodyRange.Cells(r, mapCalc(colsToCopy(i))).value = _
                            tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS(colsToCopy(i))).value
                    End If

                End If
            Next i
        Else

            For i = LBound(colsToCopy) To UBound(colsToCopy)
                If mapCalc.Exists(colsToCopy(i)) Then
                    tblCalc.DataBodyRange.Cells(r, mapCalc(colsToCopy(i))).ClearContents
                End If
            Next i

            For i = LBound(colsToClear) To UBound(colsToClear)
                If mapCalc.Exists(colsToClear(i)) Then
                    tblCalc.DataBodyRange.Cells(r, mapCalc(colsToClear(i))).ClearContents
                End If
            Next i

        End If

    Next r

    If Not preserveCalcOutputs Then
        For i = LBound(colsToClear) To UBound(colsToClear)
            If mapCalc.Exists(colsToClear(i)) Then
                If Not tblCalc.ListColumns(colsToClear(i)).DataBodyRange Is Nothing Then
                    tblCalc.ListColumns(colsToClear(i)).DataBodyRange.ClearContents
                End If
            End If
        Next i
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur dans Sync_WBS_To_CALC : " & Err.Description, _
            "Error in Sync_WBS_To_CALC: " & Err.Description
    End If

    If Not consoleMessages Is Nothing Then
        CalcBridge_ShowPlanningConsole consoleMessages
    End If

End Sub


'------------------------------------------------------------------------------
' FR: Indique si WBS ne contient aucune identite de tache apres restauration des formules gerees. En cas d'echec de lecture, retourne True pour declencher le Safe Empty State.
' EN: Returns whether WBS contains no task identity after managed formulas are restored. On read failure, returns True to trigger the Safe Empty State.
'------------------------------------------------------------------------------

Public Function Planning_WBSIsEmpty() As Boolean

    Dim perfScope As clsPerfScope

    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim r As Long

    Set perfScope = Profiler_BeginScope("Planning_WBSIsEmpty", "Excel Read")

    On Error GoTo FailSafe

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")

    If tblWBS.DataBodyRange Is Nothing Then
        Planning_WBSIsEmpty = True
        Exit Function
    End If

    RestoreWBSFormulaColumns tblWBS

    Planning_WBSIsEmpty = True
    For r = 1 To tblWBS.ListRows.Count
        If WBSRowHasTaskIdentity(tblWBS, r, Nothing) Then
            Planning_WBSIsEmpty = False
            Exit Function
        End If
    Next r

    Exit Function

FailSafe:
    Planning_WBSIsEmpty = True

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map WBS Row Has Task IDentity sans modifier les donnees d'entree.
' EN: Returns the WBS Row Has Task IDentity map without mutating input data.
'------------------------------------------------------------------------------

Private Function WBSRowHasTaskIdentity( _
    ByVal tblWBS As ListObject, _
    ByVal rowIndex As Long, _
    ByVal mapWBS As Object) As Boolean

    Dim perfScope As clsPerfScope

    Dim idVal As String
    Dim wbsVal As String

    Set perfScope = Profiler_BeginScope("WBSRowHasTaskIdentity", "Excel Cell Read")

    On Error GoTo SafeExit

    If tblWBS Is Nothing Then Exit Function
    If tblWBS.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > tblWBS.ListRows.Count Then Exit Function

    If Not mapWBS Is Nothing Then
        If mapWBS.Exists("ID") Then
            idVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(rowIndex, mapWBS("ID")).value))
        End If
        If mapWBS.Exists("WBS") Then
            wbsVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(rowIndex, mapWBS("WBS")).value))
        End If
    Else
        If TableHasColumn(tblWBS, "ID") Then
            idVal = Trim$(CStr(tblWBS.ListColumns("ID").DataBodyRange.Cells(rowIndex, 1).value))
        End If
        If TableHasColumn(tblWBS, "WBS") Then
            wbsVal = Trim$(CStr(tblWBS.ListColumns("WBS").DataBodyRange.Cells(rowIndex, 1).value))
        End If
    End If

    wbsVal = Replace$(wbsVal, ",", ".")
    WBSRowHasTaskIdentity = (idVal <> "" Or wbsVal <> "")

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Vide les sorties CALC possedees lorsque WBS ne contient aucun projet exploitable.
' EN: Clears owned CALC outputs when WBS contains no usable project.
'------------------------------------------------------------------------------

Public Sub Planning_CalcSafeEmptyState()

    Dim oldScreenUpdating As Boolean
    Dim oldEvents As Boolean

    On Error GoTo SafeExit

    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Ensure_Calc_Infrastructure
    Import_WBS_To_Constraints

    ClearPlanningTableRows "CALC", "tbl_CALC"
    ClearPlanningTableRows "CALC", "tbl_LOGIC_LINKS"

SafeExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating

End Sub

'------------------------------------------------------------------------------
' FR: Place le Gantt dans son etat visuel vide sans lancer le renderer normal.
' EN: Places Gantt in its safe empty visual state without running the normal renderer.
'------------------------------------------------------------------------------

Public Sub Planning_GanttSafeEmptyState()

    Dim oldScreenUpdating As Boolean
    Dim oldEvents As Boolean

    On Error GoTo SafeExit

    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    CalcState_ResetStorage
    GanttSimulation_ResetTableStorage
    Gantt_SafeEmptyState

SafeExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating

End Sub

'------------------------------------------------------------------------------
' FR: Orchestre les Safe Empty States CALC, Gantt, S-Curve et Dashboard dans l'ordre du workflow vide.
' EN: Orchestrates CALC, Gantt, S-Curve and Dashboard Safe Empty States in empty-workflow order.
'------------------------------------------------------------------------------

Public Sub Planning_FullSafeEmptyState()

    Planning_CalcSafeEmptyState
    Planning_GanttSafeEmptyState
    SCurve_SafeEmptyState

End Sub

'------------------------------------------------------------------------------
' FR: Vide ou reinitialise Planning Table Rows.
' EN: Clears or resets Planning Table Rows.
'------------------------------------------------------------------------------
Private Sub ClearPlanningTableRows(ByVal sheetName As String, ByVal tableName As String)

    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub

    Do While tbl.ListRows.Count > 0
        tbl.ListRows(tbl.ListRows.Count).Delete
    Loop

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree Task Type Column Exists si necessaire.
' EN: Ensures or creates Task Type Column Exists when needed.
'------------------------------------------------------------------------------
Private Sub EnsureTaskTypeColumnExists( _
    ByVal tblWBS As ListObject, _
    ByVal tblCalc As ListObject)

    Dim wbsTaskTypeIndex As Long
    Dim calcTaskTypeIndex As Long
    Dim newCol As ListColumn

    If tblWBS Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then Exit Sub

    '--------------------------------------------------
    ' WBS must already contain the user input column.
    '--------------------------------------------------
    wbsTaskTypeIndex = 0

    On Error Resume Next
    wbsTaskTypeIndex = tblWBS.ListColumns("Task Type").Index
    On Error GoTo 0

    If wbsTaskTypeIndex <= 0 Then
        Err.Raise vbObjectError + 2301, "EnsureTaskTypeColumnExists", _
            "Missing required WBS input column: Task Type"
    End If

    '--------------------------------------------------
    ' CALC must contain Task Type, but we must NOT insert
    ' it inside the existing table structure.
    '
    ' Reason:
    ' - The previous patch inserted Task Type after IsSummary.
    ' - That changed the internal structure/order of tbl_CALC.
    ' - Some existing analytics plumbing is clearly still sensitive
    '   to the table structure / right-side analytics area.
    '
    ' Safer rule:
    ' - If Task Type exists: do nothing.
    ' - If missing: append it at the END of tbl_CALC only.
    ' - Do not move existing columns.
    ' - Do not touch analytics columns.
    '--------------------------------------------------
    calcTaskTypeIndex = 0

    On Error Resume Next
    calcTaskTypeIndex = tblCalc.ListColumns("Task Type").Index
    On Error GoTo 0

    If calcTaskTypeIndex <= 0 Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Task Type"
    End If

    If Not tblCalc.DataBodyRange Is Nothing Then
        tblCalc.ListColumns("Task Type").DataBodyRange.NumberFormat = "@"
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree Calendar Column Exists si necessaire.
' EN: Ensures or creates Calendar Column Exists when needed.
'------------------------------------------------------------------------------
Private Sub EnsureCalendarColumnExists( _
    ByVal tblWBS As ListObject, _
    ByVal tblCalc As ListObject)

    Dim newCol As ListColumn

    If tblWBS Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then Exit Sub

    If Not TableHasColumn(tblWBS, "Cal") Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = "Cal"
    End If

    If Not TableHasColumn(tblCalc, "Cal") Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Cal"
    End If

    If Not tblWBS.DataBodyRange Is Nothing Then
        tblWBS.ListColumns("Cal").DataBodyRange.NumberFormat = "@"
    End If

    If Not tblCalc.DataBodyRange Is Nothing Then
        tblCalc.ListColumns("Cal").DataBodyRange.NumberFormat = "@"
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Verifie ou cree Deadline Output Columns Exist si necessaire.
' EN: Ensures or creates Deadline Output Columns Exist when needed.
'------------------------------------------------------------------------------
Private Sub EnsureDeadlineOutputColumnsExist( _
    ByVal tblWBS As ListObject, _
    ByVal tblCalc As ListObject)

    Dim colIndex As Long
    Dim newCol As ListColumn

    If tblWBS Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then Exit Sub

    
    colIndex = 0
    On Error Resume Next
    colIndex = tblWBS.ListColumns("Deadline").Index
    On Error GoTo 0

    If colIndex > 0 Then
        tblWBS.ListColumns("Deadline").Delete
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = tblWBS.ListColumns("Deadline Float").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = "Deadline Float"
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = tblCalc.ListColumns("Deadline").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Deadline"
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = tblCalc.ListColumns("Deadline Float").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Deadline Float"
    End If

    If Not tblWBS.DataBodyRange Is Nothing Then
        tblWBS.ListColumns("Deadline Float").DataBodyRange.NumberFormat = "0"
    End If

    If Not tblCalc.DataBodyRange Is Nothing Then
        tblCalc.ListColumns("Deadline").DataBodyRange.NumberFormat = "dd/mm/yyyy"
        tblCalc.ListColumns("Deadline Float").DataBodyRange.NumberFormat = "0"
    End If

End Sub



'------------------------------------------------------------------------------
' FR: Verifie ou cree Longest Path Output Columns Exist si necessaire.
' EN: Ensures or creates Longest Path Output Columns Exist when needed.
'------------------------------------------------------------------------------
Private Sub EnsureLongestPathOutputColumnsExist( _
    ByVal tblWBS As ListObject, _
    ByVal tblCalc As ListObject)

    Dim colIndex As Long
    Dim newCol As ListColumn

    If tblWBS Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then Exit Sub

    colIndex = 0
    On Error Resume Next
    colIndex = tblWBS.ListColumns("Longest Path").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = "Longest Path"
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = tblCalc.ListColumns("Longest Path").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Longest Path"
    End If

    If Not tblWBS.DataBodyRange Is Nothing Then
        tblWBS.ListColumns("Longest Path").DataBodyRange.NumberFormat = "@"
    End If

    If Not tblCalc.DataBodyRange Is Nothing Then
        tblCalc.ListColumns("Longest Path").DataBodyRange.NumberFormat = "@"
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Normalise Task Type Value dans un format exploitable.
' EN: Normalizes Task Type Value into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizeTaskTypeValue(ByVal rawValue As Variant) As String

    Dim s As String

    s = UCase$(Trim$(CStr(rawValue)))

    Select Case s

        Case "", "TASK", "STANDARD", "NORMAL"
            NormalizeTaskTypeValue = "Task"

        Case "MILESTONE", "MS", "JALON"
            NormalizeTaskTypeValue = "Milestone"

        Case "LEVEL OF EFFORT", "LOE", "LEVEL-OF-EFFORT", "LEVEL_OF_EFFORT"
            NormalizeTaskTypeValue = "Level of Effort"

        Case Else
            Err.Raise vbObjectError + 2302, "NormalizeTaskTypeValue", _
                "Invalid Task Type value: " & CStr(rawValue) & _
                " | Allowed values: Task, Milestone, Level of Effort"

    End Select

End Function

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Sync Forecast Only dans le stockage possede par le domaine.
' EN: Writes or synchronizes Sync Forecast Only in the store owned by the domain.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Function Sync_Forecast_Only() As Boolean

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim wbsIdRows As Object

    Dim r As Long
    Dim idValue As Variant
    Dim idKey As String
    Dim wbsRowIndex As Long

    Dim consoleMessages As Collection

    On Error GoTo SafeExit

    Sync_Forecast_Only = False
    Set consoleMessages = New Collection

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")

    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    Set mapWBS = CreateObject("Scripting.Dictionary")
    Set mapCalc = CreateObject("Scripting.Dictionary")
    Set wbsIdRows = CreateObject("Scripting.Dictionary")

    For r = 1 To tblWBS.ListColumns.Count
        mapWBS(tblWBS.ListColumns(r).Name) = r
    Next r

    For r = 1 To tblCalc.ListColumns.Count
        mapCalc(tblCalc.ListColumns(r).Name) = r
    Next r

    If Not mapWBS.Exists("ID") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_WBS.", _
            "Column ID was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("ID") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_CALC.", _
            "Column ID was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists("Forecast Start") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Forecast Start est introuvable dans tbl_WBS.", _
            "Column Forecast Start was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists("Forecast Finish") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Forecast Finish est introuvable dans tbl_WBS.", _
            "Column Forecast Finish was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("Forecast Start") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Forecast Start est introuvable dans tbl_CALC.", _
            "Column Forecast Start was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("Forecast Finish") Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Forecast Finish est introuvable dans tbl_CALC.", _
            "Column Forecast Finish was not found in tbl_CALC."
        GoTo SafeExit
    End If

    If tblCalc.DataBodyRange Is Nothing Then GoTo SafeExit
    If tblWBS.DataBodyRange Is Nothing Then GoTo SafeExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = 1 To tblWBS.ListRows.Count
        idValue = tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value
        If Trim(CStr(idValue)) <> "" Then
            idKey = CStr(idValue)
            wbsIdRows(idKey) = r
        End If
    Next r

    For r = 1 To tblCalc.ListRows.Count

        idValue = tblCalc.DataBodyRange.Cells(r, mapCalc("ID")).value

        If Trim(CStr(idValue)) <> "" Then
            idKey = CStr(idValue)

            If wbsIdRows.Exists(idKey) Then
                wbsRowIndex = wbsIdRows(idKey)

                tblCalc.DataBodyRange.Cells(r, mapCalc("Forecast Start")).value = _
                    tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS("Forecast Start")).value

                tblCalc.DataBodyRange.Cells(r, mapCalc("Forecast Finish")).value = _
                    tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS("Forecast Finish")).value
            End If
        End If

    Next r

    Sync_Forecast_Only = True

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur dans Sync_Forecast_Only : " & Err.Description, _
            "Error in Sync_Forecast_Only: " & Err.Description
        Sync_Forecast_Only = False
    End If

    If Not consoleMessages Is Nothing Then
        CalcBridge_ShowPlanningConsole consoleMessages
    End If

End Function


'------------------------------------------------------------------------------
' FR: Indique si Allowed Calculated Push Field est vrai pour le contexte courant.
' EN: Returns whether Allowed Calculated Push Field is true for the current context.
'------------------------------------------------------------------------------
Private Function IsAllowedCalculatedPushField(ByVal fieldName As String) As Boolean

    Select Case fieldName
        Case "Calculated Start", _
             "Calculated Finish", _
             "Driving Logic", _
             "Critical Path", _
             "Longest Path", _
             "Critical Path REX", _
             "Total Float", _
             "Free Float", _
             "Total Float REX", _
             "Free Float REX", _
             "Deadline Float", _
             "Error flag", _
             "ErrorMsg"

            IsAllowedCalculatedPushField = True

        Case Else
            IsAllowedCalculatedPushField = False

    End Select

End Function

'------------------------------------------------------------------------------
' FR: Actualise Apply WBS Date Formats sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply WBS Date Formats without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyWBSDateFormats(ByVal tblWBS As ListObject)

    On Error Resume Next

    tblWBS.ListColumns("Baseline Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Baseline Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Actual Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Actual Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Forecast Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Forecast Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Deadline Float").DataBodyRange.NumberFormat = "0"
    tblWBS.ListColumns("Calculated Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Calculated Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"

    On Error GoTo 0

End Sub


'------------------------------------------------------------------------------
' FR: Traite la collection Rebuild Logic Links Table sans modifier les donnees d'entree.
' EN: Handles the Rebuild Logic Links Table collection without mutating input data.
'------------------------------------------------------------------------------

Public Sub RebuildLogicLinksTable()

    Dim perfScope As clsPerfScope
    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblLinks As ListObject
    Dim mapWBS As Object
    Dim wbsToId As Object
    Dim arr As Variant
    Dim r As Long
    Dim rowCount As Long
    Dim succId As String
    Dim succWBS As String
    Dim predText As String
    Dim linksOut As Collection
    Dim allLinks As Collection
    Dim errText As String
    Dim linkRow As Object
    Dim outArr() As Variant
    Dim outCount As Long
    Dim i As Long
    Dim consoleMessages As Collection
    Dim missingPredRefs As Collection
    Dim missingPredMessageFR As String
    Dim missingPredMessageEN As String

    Set perfScope = Profiler_BeginScope("RebuildLogicLinksTable", "Excel Table Sync")

    On Error GoTo SafeExit

    Set consoleMessages = New Collection
    Set missingPredRefs = New Collection
    Set allLinks = New Collection

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblLinks = EnsureLogicLinksTable(wsCalc)

    If tblWBS.DataBodyRange Is Nothing Then
        If Not LogicLinksTableMatchesEmpty(tblLinks) Then ClearLogicLinksTableRows tblLinks
        GoTo SafeExit
    End If

    Set mapWBS = CreateObject("Scripting.Dictionary")
    For i = 1 To tblWBS.ListColumns.Count
        mapWBS(tblWBS.ListColumns(i).Name) = i
    Next i

    Set wbsToId = CanonicalIdentity_BuildWbsToIdMap(tblWBS, mapWBS)
    arr = tblWBS.DataBodyRange.value
    rowCount = UBound(arr, 1)

    For r = 1 To rowCount
        succId = Trim$(CStr(arr(r, mapWBS("ID"))))
        succWBS = NormalizeWBS(CStr(arr(r, mapWBS("WBS"))))
        predText = Trim$(CStr(arr(r, mapWBS("Predecessors WBS"))))

        If succId <> "" And succWBS <> "" Then
            If Not ParsePredecessorsText(succId, succWBS, predText, wbsToId, linksOut, errText) Then
                DataSync_AddConsoleMessage consoleMessages, "STOP", _
                    "Erreur lors de la reconstruction de tbl_LOGIC_LINKS." & vbCrLf & "-> " & errText, _
                    "Error while rebuilding tbl_LOGIC_LINKS." & vbCrLf & "-> " & errText
                GoTo SafeExit
            End If

            DataSync_CollectMissingPredecessorRefs linksOut, missingPredRefs
            For i = 1 To linksOut.Count
                allLinks.Add linksOut(i)
            Next i
        End If
    Next r

    If missingPredRefs.Count > 0 Then
        DataSync_BuildMissingPredecessorMessages missingPredRefs, missingPredMessageFR, missingPredMessageEN
        If IsMacroRunActive() Then
            RequestMacroAbort "RebuildLogicLinksTable", missingPredMessageFR, missingPredMessageEN
        Else
            DataSync_AddConsoleMessage consoleMessages, "STOP", missingPredMessageFR, missingPredMessageEN
        End If
        GoTo SafeExit
    End If

    outCount = allLinks.Count
    If outCount <= 0 Then
        If Not LogicLinksTableMatchesEmpty(tblLinks) Then ClearLogicLinksTableRows tblLinks
        GoTo SafeExit
    End If

    ReDim outArr(1 To outCount, 1 To 7)
    For i = 1 To outCount
        Set linkRow = allLinks(i)
        outArr(i, 1) = CStr(linkRow("Succ ID"))
        outArr(i, 2) = CStr(linkRow("Succ WBS"))
        outArr(i, 3) = CStr(linkRow("Pred ID"))
        outArr(i, 4) = CStr(linkRow("Pred WBS"))
        outArr(i, 5) = CStr(linkRow("Link Type"))
        outArr(i, 6) = CLng(linkRow("Lag"))
        outArr(i, 7) = CStr(linkRow("Raw Token"))
    Next i

    If Not LogicLinksTableMatchesOutput(tblLinks, outArr, outCount) Then
        RewriteLogicLinksTable tblLinks, outArr, outCount
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur VBA dans RebuildLogicLinksTable : " & Err.Description, _
            "VBA error in RebuildLogicLinksTable: " & Err.Description
    End If

    If Not consoleMessages Is Nothing Then CalcBridge_ShowPlanningConsole consoleMessages

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree Logic Links Table si necessaire.
' EN: Ensures or creates Logic Links Table when needed.
'------------------------------------------------------------------------------
Private Function EnsureLogicLinksTable(ByVal wsCalc As Worksheet) As ListObject

    Dim perfScope As clsPerfScope

    Dim tbl As ListObject
    Dim headerRange As Range
    Dim fullRange As Range

    Set perfScope = Profiler_BeginScope("EnsureLogicLinksTable", "Excel Infrastructure")

    On Error Resume Next
    Set tbl = wsCalc.ListObjects(LOGIC_LINKS_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then

        Set headerRange = wsCalc.Range(LOGIC_LINKS_FIRST_CELL).Resize(1, 7)
        WriteLogicLinksHeaders headerRange

        Set fullRange = headerRange.Resize(2, 7)
        Set tbl = wsCalc.ListObjects.Add(xlSrcRange, fullRange, , xlYes)
        tbl.Name = LOGIC_LINKS_TABLE_NAME

    Else
        WriteLogicLinksHeaders tbl.HeaderRowRange
    End If

    Set EnsureLogicLinksTable = tbl

End Function


'------------------------------------------------------------------------------
' FR: Ecrit Logic Links Headers vers le stockage cible.
' EN: Writes Logic Links Headers to the target storage.
'------------------------------------------------------------------------------
Private Sub WriteLogicLinksHeaders(ByVal headerRange As Range)

    headerRange.Cells(1, 1).value = "Succ ID"
    headerRange.Cells(1, 2).value = "Succ WBS"
    headerRange.Cells(1, 3).value = "Pred ID"
    headerRange.Cells(1, 4).value = "Pred WBS"
    headerRange.Cells(1, 5).value = "Link Type"
    headerRange.Cells(1, 6).value = "Lag"
    headerRange.Cells(1, 7).value = "Raw Token"

End Sub


'------------------------------------------------------------------------------
' FR: Vide ou reinitialise Logic Links Table Rows.
' EN: Clears or resets Logic Links Table Rows.
'------------------------------------------------------------------------------
Private Sub ClearLogicLinksTableRows(ByVal tbl As ListObject)

    Dim perfScope As clsPerfScope
    Dim targetRange As Range

    Set perfScope = Profiler_BeginScope("ClearLogicLinksTableRows", "Excel Table Resize")

    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.ClearContents

    Set targetRange = tbl.HeaderRowRange.Cells(1, 1).Resize(2, tbl.ListColumns.Count)
    If tbl.Range.Rows.Count <> 2 Then tbl.Resize targetRange

End Sub

'------------------------------------------------------------------------------
' FR: Journalise Logic Links Table Matches Empty dans l'historique planning.
' EN: Logs Logic Links Table Matches Empty into the planning history.
'------------------------------------------------------------------------------
Private Function LogicLinksTableMatchesEmpty(ByVal tbl As ListObject) As Boolean

    Dim values As Variant
    Dim c As Long

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then
        LogicLinksTableMatchesEmpty = True
        Exit Function
    End If
    If tbl.ListRows.Count <> 1 Then Exit Function

    values = tbl.DataBodyRange.value
    For c = 1 To tbl.ListColumns.Count
        If Len(CStr(values(1, c))) > 0 Then Exit Function
    Next c

    LogicLinksTableMatchesEmpty = True

End Function

'------------------------------------------------------------------------------
' FR: Journalise Logic Links Table Matches Output dans l'historique planning.
' EN: Logs Logic Links Table Matches Output into the planning history.
'------------------------------------------------------------------------------
Private Function LogicLinksTableMatchesOutput( _
    ByVal tbl As ListObject, _
    ByRef outArr() As Variant, _
    ByVal outCount As Long) As Boolean

    Dim currentArr As Variant
    Dim r As Long
    Dim c As Long

    If tbl Is Nothing Then Exit Function
    If outCount <= 0 Then
        LogicLinksTableMatchesOutput = LogicLinksTableMatchesEmpty(tbl)
        Exit Function
    End If
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If tbl.ListRows.Count <> outCount Then Exit Function

    currentArr = tbl.DataBodyRange.value
    For r = 1 To outCount
        For c = 1 To 7
            If c = 6 Then
                If Not IsNumeric(currentArr(r, c)) Then Exit Function
                If CLng(currentArr(r, c)) <> CLng(outArr(r, c)) Then Exit Function
            ElseIf StrComp(CStr(currentArr(r, c)), CStr(outArr(r, c)), vbBinaryCompare) <> 0 Then
                Exit Function
            End If
        Next c
        For c = 8 To tbl.ListColumns.Count
            If Len(CStr(currentArr(r, c))) > 0 Then Exit Function
        Next c
    Next r

    LogicLinksTableMatchesOutput = True

End Function

'------------------------------------------------------------------------------
' FR: Traite la reference Rewrite Logic Links Table sans modifier les donnees d'entree.
' EN: Handles the Rewrite Logic Links Table reference without mutating input data.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub RewriteLogicLinksTable( _
    ByVal tbl As ListObject, _
    ByRef outArr() As Variant, _
    ByVal outCount As Long)

    Dim perfScope As clsPerfScope
    Dim targetRows As Long
    Dim targetRange As Range

    Set perfScope = Profiler_BeginScope("RewriteLogicLinksTable", "Excel Table Write")

    If tbl Is Nothing Then Exit Sub
    targetRows = outCount
    If targetRows < 1 Then targetRows = 1

    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.ClearContents
    Set targetRange = tbl.HeaderRowRange.Cells(1, 1).Resize(targetRows + 1, tbl.ListColumns.Count)
    If tbl.Range.Rows.Count <> targetRows + 1 Then tbl.Resize targetRange

    If outCount > 0 Then
        Set targetRange = tbl.DataBodyRange.Resize(outCount, 7)
        targetRange.Columns(1).NumberFormat = "@"
        targetRange.Columns(2).NumberFormat = "@"
        targetRange.Columns(3).NumberFormat = "@"
        targetRange.Columns(4).NumberFormat = "@"
        targetRange.Columns(5).NumberFormat = "@"
        targetRange.Columns(7).NumberFormat = "@"
        targetRange.Columns(6).NumberFormat = "0"
        targetRange.value = outArr
        ApplyLogicLinksTableFormats tbl
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Logic Links Table Formats sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Logic Links Table Formats without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyLogicLinksTableFormats(ByVal tbl As ListObject)

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("ApplyLogicLinksTableFormats", "Excel Format")

    On Error Resume Next

    tbl.ListColumns("Succ ID").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Succ WBS").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Pred ID").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Pred WBS").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Link Type").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Raw Token").DataBodyRange.NumberFormat = "@"

    tbl.ListColumns("Lag").DataBodyRange.NumberFormat = "0"

    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la reference Table Has Column sans modifier les donnees d'entree.
' EN: Returns the Table Has Column reference without mutating input data.
'------------------------------------------------------------------------------

Private Function TableHasColumn(ByVal tbl As ListObject, ByVal columnName As String) As Boolean

    Dim perfScope As clsPerfScope

    Dim col As ListColumn

    Set perfScope = Profiler_BeginScope("TableHasColumn", "Excel Metadata")

    On Error Resume Next
    Set col = tbl.ListColumns(columnName)
    On Error GoTo 0

    TableHasColumn = Not col Is Nothing

End Function

'=====================================================
' Task Type WBS input setup
'=====================================================
' Purpose:
' - Add dropdown validation on tbl_WBS[Task Type].
' - Fill blank Task Type cells with "Task".
' - Keep existing valid user values.
'
' Allowed values:
' - Task
' - Milestone
' - Level of Effort
'
' Where to call:
' - In Sync_WBS_To_CALC, after:
'       Set tblWBS = wsWBS.ListObjects("tbl_WBS")
'   and before reading data from tbl_WBS.
'
' Call line:
'       EnsureWBSTaskTypeInputSetup tblWBS
'=====================================================


'------------------------------------------------------------------------------
' FR: Verifie ou cree WBSTask Type Input Setup si necessaire.
' EN: Ensures or creates WBSTask Type Input Setup when needed.
'------------------------------------------------------------------------------
Private Sub EnsureWBSTaskTypeInputSetup(ByVal tblWBS As ListObject)

    Dim perfScope As clsPerfScope

    Dim rng As Range
    Dim cell As Range
    Dim normalizedValue As String
    Dim rowIndex As Long

    Set perfScope = Profiler_BeginScope("EnsureWBSTaskTypeInputSetup", "Excel Validation")

    If tblWBS Is Nothing Then Exit Sub

    If Not TableHasColumn(tblWBS, "Task Type") Then
        Err.Raise vbObjectError + 2310, "EnsureWBSTaskTypeInputSetup", _
            "Missing required WBS input column: Task Type"
    End If

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    Set rng = tblWBS.ListColumns("Task Type").DataBodyRange

    'Default blank cells to Task only on rows that have an ID/WBS identity.
    For Each cell In rng.Cells

        rowIndex = cell.Row - rng.Row + 1

        If Not WBSRowHasTaskIdentity(tblWBS, rowIndex, Nothing) Then
            If Trim$(CStr(cell.value)) <> "" Then cell.ClearContents
        Else
            normalizedValue = NormalizeTaskTypeValue(cell.value)

            If Trim$(CStr(cell.value)) = "" Then
                cell.value = "Task"
            ElseIf CStr(cell.value) <> normalizedValue Then
                cell.value = normalizedValue
            End If
        End If

    Next cell

    'Apply dropdown validation.
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="Task,Milestone,Level of Effort"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Task Type"
        .InputMessage = "Choose: Task, Milestone, or Level of Effort."
        .ErrorTitle = "Invalid Task Type"
        .errorMessage = "Allowed values: Task, Milestone, Level of Effort."
        .ShowInput = True
        .ShowError = True
    End With

    rng.NumberFormat = "@"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree Summary Display Column Exists si necessaire.
' EN: Ensures or creates Summary Display Column Exists when needed.
'------------------------------------------------------------------------------
Private Sub EnsureSummaryDisplayColumnExists( _
    ByVal tblWBS As ListObject, _
    ByVal tblCalc As ListObject)

    Dim perfScope As clsPerfScope

    Dim newCol As ListColumn
    Dim targetIndex As Long

    Set perfScope = Profiler_BeginScope("EnsureSummaryDisplayColumnExists", "Excel Infrastructure")

    If Not tblWBS Is Nothing Then
        If Not TableHasColumn(tblWBS, "S") Then
            If TableHasColumn(tblWBS, "Comments") Then
                targetIndex = tblWBS.ListColumns("Comments").Index
            ElseIf TableHasColumn(tblWBS, "Task Type") Then
                targetIndex = tblWBS.ListColumns("Task Type").Index + 1
            Else
                targetIndex = tblWBS.ListColumns.Count + 1
            End If

            Set newCol = tblWBS.ListColumns.Add(Position:=targetIndex)
            newCol.Name = "S"
        End If

        If Not tblWBS.DataBodyRange Is Nothing Then
            With tblWBS.ListColumns("S").DataBodyRange
                .NumberFormat = "@"
                With .Validation
                    .Delete
                    .Add Type:=xlValidateList, _
                         AlertStyle:=xlValidAlertStop, _
                         Operator:=xlBetween, _
                         Formula1:="Y,N"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = "S"
                    .InputMessage = "Choose Y to show in Summary, N to hide."
                    .ErrorTitle = "Invalid S"
                    .errorMessage = "Allowed values: blank, Y, N."
                    .ShowInput = True
                    .ShowError = True
                End With
            End With
        End If
    End If

    If Not tblCalc Is Nothing Then
        If Not TableHasColumn(tblCalc, "S") Then
            Set newCol = tblCalc.ListColumns.Add
            newCol.Name = "S"
        End If

        If Not tblCalc.DataBodyRange Is Nothing Then
            tblCalc.ListColumns("S").DataBodyRange.NumberFormat = "@"
        End If
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Normalise Summary Display Value dans un format exploitable.
' EN: Normalizes Summary Display Value into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizeSummaryDisplayValue(ByVal rawValue As Variant) As String

    Dim txt As String

    txt = UCase$(Trim$(CStr(rawValue)))

    Select Case txt
        Case ""
            NormalizeSummaryDisplayValue = ""
        Case "Y"
            NormalizeSummaryDisplayValue = "Y"
        Case "N"
            NormalizeSummaryDisplayValue = "N"
        Case Else
            Err.Raise vbObjectError + 2311, "NormalizeSummaryDisplayValue", _
                "Invalid S value: " & CStr(rawValue) & " (allowed: blank, Y, N)"
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Default Summary Display Value sans modifier les donnees d'entree.
' EN: Returns the Default Summary Display Value map without mutating input data.
'------------------------------------------------------------------------------

Private Function DefaultSummaryDisplayValue( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal rowIndex As Long, _
    ByVal summaryWbsByWbs As Object) As String

    Dim wbsVal As String
    Dim taskTypeVal As String

    DefaultSummaryDisplayValue = "N"

    If tblWBS Is Nothing Then Exit Function
    If tblWBS.DataBodyRange Is Nothing Then Exit Function
    If mapWBS Is Nothing Then Exit Function

    If mapWBS.Exists("WBS") Then
        wbsVal = Replace$(Trim$(CStr(tblWBS.DataBodyRange.Cells(rowIndex, mapWBS("WBS")).value)), ",", ".")
        If Not summaryWbsByWbs Is Nothing Then
            If summaryWbsByWbs.Exists(wbsVal) Then
                DefaultSummaryDisplayValue = "Y"
                Exit Function
            End If
        End If
    End If

    If mapWBS.Exists("Task Type") Then
        taskTypeVal = UCase$(NormalizeTaskTypeValue(tblWBS.DataBodyRange.Cells(rowIndex, mapWBS("Task Type")).value))
        If taskTypeVal = "MILESTONE" Then DefaultSummaryDisplayValue = "Y"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise WBSSummary Display Values dans un format exploitable.
' EN: Normalizes WBSSummary Display Values into a usable format.
'------------------------------------------------------------------------------
Private Sub NormalizeWBSSummaryDisplayValues( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal summaryWbsByWbs As Object)

    Dim perfScope As clsPerfScope

    Dim r As Long
    Dim cell As Range
    Dim normalizedValue As String

    Set perfScope = Profiler_BeginScope("NormalizeWBSSummaryDisplayValues", "Excel Cell Write")

    If tblWBS Is Nothing Then Exit Sub
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If mapWBS Is Nothing Then Exit Sub
    If Not mapWBS.Exists("S") Then Exit Sub

    For r = 1 To tblWBS.ListRows.Count
        Set cell = tblWBS.DataBodyRange.Cells(r, mapWBS("S"))

        If Not WBSRowHasTaskIdentity(tblWBS, r, mapWBS) Then
            If Trim$(CStr(cell.value)) <> "" Then cell.ClearContents
        ElseIf Trim$(CStr(cell.value)) = "" Then
            cell.value = DefaultSummaryDisplayValue(tblWBS, mapWBS, r, summaryWbsByWbs)
        Else
            normalizedValue = NormalizeSummaryDisplayValue(cell.value)
            If CStr(cell.value) <> normalizedValue Then cell.value = normalizedValue
        End If
    Next r

End Sub
'------------------------------------------------------------------------------
' FR: Verifie ou cree WBSCalendar Input Setup si necessaire.
' EN: Ensures or creates WBSCalendar Input Setup when needed.
'------------------------------------------------------------------------------
Private Sub EnsureWBSCalendarInputSetup(ByVal tblWBS As ListObject)

    Dim perfScope As clsPerfScope

    Dim rng As Range
    Dim cell As Range
    Dim normalizedValue As String
    Dim newCol As ListColumn

    Set perfScope = Profiler_BeginScope("EnsureWBSCalendarInputSetup", "Excel Validation")

    If tblWBS Is Nothing Then Exit Sub

    If Not TableHasColumn(tblWBS, "Cal") Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = "Cal"
    End If

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    Set rng = tblWBS.ListColumns("Cal").DataBodyRange

    For Each cell In rng.Cells
        If Trim$(CStr(cell.value)) <> "" Then
            normalizedValue = NormalizeCalendarType(cell.value)
            If CStr(cell.value) <> normalizedValue Then cell.value = normalizedValue
        End If
    Next cell

    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=CALENDAR_7D & "," & CALENDAR_6D & "," & CALENDAR_5D
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Cal"
        .InputMessage = "Choose: 7j/7, 6j/7, or 5j/7."
        .ErrorTitle = "Invalid Cal"
        .errorMessage = "Allowed values: blank, 7j/7, 6j/7, 5j/7."
        .ShowInput = True
        .ShowError = True
    End With

    rng.NumberFormat = "@"

End Sub

'------------------------------------------------------------------------------
' FR: Construit la map WBS Summary WBS Lookup a partir des donnees fournies par l'appelant.
' EN: Builds the WBS Summary WBS Lookup map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildWBSSummaryWbsLookup( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object) As Object

    Dim perfScope As clsPerfScope

    Dim summaryWbs As Object
    Dim r As Long
    Dim wbsVal As String
    Dim parentWbs As String

    Set perfScope = Profiler_BeginScope("BuildWBSSummaryWbsLookup", "Dictionary")

    Set summaryWbs = CreateObject("Scripting.Dictionary")

    If tblWBS Is Nothing Then
        Set BuildWBSSummaryWbsLookup = summaryWbs
        Exit Function
    End If
    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildWBSSummaryWbsLookup = summaryWbs
        Exit Function
    End If
    If mapWBS Is Nothing Then
        Set BuildWBSSummaryWbsLookup = summaryWbs
        Exit Function
    End If
    If Not mapWBS.Exists("WBS") Then
        Set BuildWBSSummaryWbsLookup = summaryWbs
        Exit Function
    End If

    For r = 1 To tblWBS.ListRows.Count
        wbsVal = Replace$(Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("WBS")).value)), ",", ".")
        parentWbs = GetParentWBS(wbsVal)

        Do While parentWbs <> ""
            summaryWbs(parentWbs) = True
            parentWbs = GetParentWBS(parentWbs)
        Loop
    Next r

    Set BuildWBSSummaryWbsLookup = summaryWbs

End Function

'------------------------------------------------------------------------------
' FR: Indique si Calendar Ignored WBSRow est vrai pour le contexte courant.
' EN: Returns whether Calendar Ignored WBSRow is true for the current context.
'------------------------------------------------------------------------------
Private Function IsCalendarIgnoredWBSRow( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal rowIndex As Long, _
    ByVal summaryWbsByWbs As Object) As Boolean

    Dim wbsVal As String
    Dim taskTypeVal As String

    If tblWBS Is Nothing Then Exit Function
    If tblWBS.DataBodyRange Is Nothing Then Exit Function
    If mapWBS Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > tblWBS.ListRows.Count Then Exit Function

    If mapWBS.Exists("WBS") Then
        wbsVal = Replace$(Trim$(CStr(tblWBS.DataBodyRange.Cells(rowIndex, mapWBS("WBS")).value)), ",", ".")
        If Not summaryWbsByWbs Is Nothing Then
            If summaryWbsByWbs.Exists(wbsVal) Then
                IsCalendarIgnoredWBSRow = True
                Exit Function
            End If
        End If
    End If

    If mapWBS.Exists("Task Type") Then
        taskTypeVal = UCase$(NormalizeTaskTypeValue(tblWBS.DataBodyRange.Cells(rowIndex, mapWBS("Task Type")).value))
        IsCalendarIgnoredWBSRow = (taskTypeVal = "LEVEL OF EFFORT" Or taskTypeVal = "MILESTONE")
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise WBSCalendar Values dans un format exploitable.
' EN: Normalizes WBSCalendar Values into a usable format.
'------------------------------------------------------------------------------
Private Sub NormalizeWBSCalendarValues( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal summaryWbsByWbs As Object)

    Dim perfScope As clsPerfScope

    Dim r As Long
    Dim cell As Range
    Dim normalizedValue As String
    Dim rowIndex As Long

    Set perfScope = Profiler_BeginScope("NormalizeWBSCalendarValues", "Excel Cell Write")

    If tblWBS Is Nothing Then Exit Sub
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If mapWBS Is Nothing Then Exit Sub
    If Not mapWBS.Exists("Cal") Then Exit Sub

    For r = 1 To tblWBS.ListRows.Count
        Set cell = tblWBS.DataBodyRange.Cells(r, mapWBS("Cal"))

        If Not WBSRowHasTaskIdentity(tblWBS, r, mapWBS) Then
            If Trim$(CStr(cell.value)) <> "" Then cell.ClearContents
        Else
            normalizedValue = NormalizeCalendarType(cell.value)

            If Trim$(CStr(cell.value)) = "" Then
                If Not IsCalendarIgnoredWBSRow(tblWBS, mapWBS, r, summaryWbsByWbs) Then
                    cell.value = CALENDAR_7D
                End If
            ElseIf CStr(cell.value) <> normalizedValue Then
                cell.value = normalizedValue
            End If
        End If
    Next r

End Sub
'------------------------------------------------------------------------------
' FR: Traite la collection Data Sync Collect Missing Predecessor Refs sans modifier les donnees d'entree.
' EN: Handles the Data Sync Collect Missing Predecessor Refs collection without mutating input data.
'------------------------------------------------------------------------------

Private Sub DataSync_CollectMissingPredecessorRefs( _
    ByVal linksOut As Collection, _
    ByVal missingRefs As Collection)

    Dim linkRow As Variant
    Dim missingRow As Object
    Dim predId As String
    Dim predWbs As String
    Dim enteredToken As String

    If linksOut Is Nothing Then Exit Sub
    If missingRefs Is Nothing Then Exit Sub

    For Each linkRow In linksOut
        predId = ""
        predWbs = ""
        enteredToken = ""

        If linkRow.Exists("Pred ID") Then predId = Trim$(CStr(linkRow("Pred ID")))
        If linkRow.Exists("Pred WBS") Then predWbs = NormalizeWBS(CStr(linkRow("Pred WBS")))

        If predId = "" And predWbs <> "" Then
            If linkRow.Exists("Entered Token") Then
                enteredToken = Trim$(CStr(linkRow("Entered Token")))
            ElseIf linkRow.Exists("Raw Token") Then
                enteredToken = Trim$(CStr(linkRow("Raw Token")))
            End If

            Set missingRow = CreateObject("Scripting.Dictionary")
            missingRow("Succ ID") = CStr(linkRow("Succ ID"))
            missingRow("Succ WBS") = CStr(linkRow("Succ WBS"))
            missingRow("Entered Token") = enteredToken
            missingRow("Pred WBS") = predWbs
            missingRefs.Add missingRow
        End If
    Next linkRow

End Sub

'------------------------------------------------------------------------------
' FR: Construit la collection Data Sync Build Missing Predecessor Messages a partir des donnees fournies par l'appelant.
' EN: Builds the Data Sync Build Missing Predecessor Messages collection from data supplied by the caller.
'------------------------------------------------------------------------------

Private Sub DataSync_BuildMissingPredecessorMessages( _
    ByVal missingRefs As Collection, _
    ByRef frText As String, _
    ByRef enText As String)

    Dim idsLine As String
    Dim succWbsLine As String
    Dim tokensLine As String
    Dim predWbsLine As String

    idsLine = DataSync_BuildMissingPredInline(missingRefs, "Succ ID", 20)
    succWbsLine = DataSync_BuildMissingPredInline(missingRefs, "Succ WBS", 20)
    tokensLine = DataSync_BuildMissingPredInline(missingRefs, "Entered Token", 20)
    predWbsLine = DataSync_BuildMissingPredInline(missingRefs, "Pred WBS", 20)

    frText = _
        "Prédécesseur introuvable" & vbCrLf & vbCrLf & _
        "Les prédécesseurs suivants n'existent pas dans le planning." & vbCrLf & vbCrLf & _
        "IDs :" & vbCrLf & _
        idsLine & vbCrLf & vbCrLf & _
        "WBS :" & vbCrLf & _
        succWbsLine & vbCrLf & vbCrLf & _
        "Prédécesseurs saisis :" & vbCrLf & _
        tokensLine & vbCrLf & vbCrLf & _
        "WBS recherchées :" & vbCrLf & _
        predWbsLine & vbCrLf & vbCrLf & _
        "-> créer les tâches manquantes ou corriger les liens."

    enText = _
        "Missing predecessor" & vbCrLf & vbCrLf & _
        "The following predecessors do not exist in the planning." & vbCrLf & vbCrLf & _
        "IDs:" & vbCrLf & _
        idsLine & vbCrLf & vbCrLf & _
        "WBS:" & vbCrLf & _
        succWbsLine & vbCrLf & vbCrLf & _
        "Entered predecessors:" & vbCrLf & _
        tokensLine & vbCrLf & vbCrLf & _
        "Referenced WBS:" & vbCrLf & _
        predWbsLine & vbCrLf & vbCrLf & _
        "-> create the missing tasks or correct the links."

End Sub

'------------------------------------------------------------------------------
' FR: Construit la collection Data Sync Build Missing Pred Inline a partir des donnees fournies par l'appelant.
' EN: Builds the Data Sync Build Missing Pred Inline collection from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function DataSync_BuildMissingPredInline( _
    ByVal missingRefs As Collection, _
    ByVal fieldName As String, _
    ByVal maxItems As Long) As String

    Dim result As String
    Dim item As Variant
    Dim itemText As String
    Dim countShown As Long
    Dim totalCount As Long

    result = ""
    If missingRefs Is Nothing Then
        DataSync_BuildMissingPredInline = result
        Exit Function
    End If

    totalCount = missingRefs.Count

    For Each item In missingRefs
        countShown = countShown + 1

        If countShown <= maxItems Then
            itemText = "-"
            If item.Exists(fieldName) Then itemText = CStr(item(fieldName))

            If result <> "" Then result = result & " / "
            result = result & itemText
        Else
            Exit For
        End If
    Next item

    If totalCount > maxItems Then
        result = result & " / +" & CStr(totalCount - maxItems)
    End If

    DataSync_BuildMissingPredInline = result

End Function
'------------------------------------------------------------------------------
' FR: Ajoute la collection Data Sync Add Console Message a la structure cible fournie par l'appelant.
' EN: Adds the Data Sync Add Console Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub DataSync_AddConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

End Sub










