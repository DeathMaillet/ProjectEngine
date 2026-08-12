Attribute VB_Name = "mod_CalcState"
Option Explicit

'===============================================================================
' MODULE : mod_CalcState
' DOMAINE / DOMAIN : Calculation / Data Sync
'
' FR
' Possede l'etat runtime du domaine et ses transitions explicites.
' Ne persiste ni ne recalcule les donnees metier sauf mention contraire.
'
' EN
' Owns domain runtime state and its explicit transitions.
' Does not persist or recalculate business data unless stated otherwise.
'
' CONTRATS / CONTRACTS : Ensure_CalcState_Table, Ensure_CalcState_Table_Console, Write_CalcState_Snapshot, Write_CalcState_Snapshot_Console, CalcState_ResetStorage, CalcState_GetSignatureByIdMap, CalcState_GetRunStatus, CalcState_MarkDirty
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const CALC_STATE_SHEET_NAME As String = "CALC_STATE"
Private Const CALC_STATE_TABLE_NAME As String = "tbl_CALC_STATE"

'------------------------------------------------------------------------------
' FR: Verifie ou cree Calc State Table si necessaire.
' EN: Ensures or creates Calc State Table when needed.
'------------------------------------------------------------------------------
Public Sub Ensure_CalcState_Table()

    Dim localConsole As Collection

    Set localConsole = New Collection
    Ensure_CalcState_Table_Console localConsole

    If localConsole.Count > 0 Then
        CalcBridge_ShowPlanningConsole localConsole
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie ou cree Calc State Table Console si necessaire.
' EN: Ensures or creates Calc State Table Console when needed.
'------------------------------------------------------------------------------
Public Sub Ensure_CalcState_Table_Console(ByVal consoleMessages As Collection)

    Dim wsState As Worksheet
    Dim tblState As ListObject
    Dim wasSheetCreated As Boolean
    Dim wasTableCreated As Boolean

    On Error GoTo SafeExit

    Application.ScreenUpdating = False

    Set wsState = EnsureWorksheetExists_CalcState(CALC_STATE_SHEET_NAME, wasSheetCreated)
    Set tblState = EnsureTableWithHeaders_CalcState( _
        wsState, _
        CALC_STATE_TABLE_NAME, _
        wsState.Range("A1"), _
        BuildCalcStateHeaders(), _
        wasTableCreated)

    If wasSheetCreated Or wasTableCreated Then
        ApplyCalcStateFormats tblState
    End If

SafeExit:
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        CalcState_AddStopToConsole consoleMessages, _
            "Erreur dans Ensure_CalcState_Table : " & Err.Description, _
            "Error in Ensure_CalcState_Table: " & Err.Description
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit Calc State Snapshot vers le stockage cible.
' EN: Writes Calc State Snapshot to the target storage.
'------------------------------------------------------------------------------
Public Sub Write_CalcState_Snapshot(ByVal runStatus As String)

    Dim localConsole As Collection

    Set localConsole = New Collection
    Write_CalcState_Snapshot_Console runStatus, localConsole

    If localConsole.Count > 0 Then
        CalcBridge_ShowPlanningConsole localConsole
    End If

End Sub


'------------------------------------------------------------------------------
' FR: Ecrit Calc State Snapshot Console vers le stockage cible.
' EN: Writes Calc State Snapshot Console to the target storage.
'------------------------------------------------------------------------------
Public Sub Write_CalcState_Snapshot_Console( _
    ByVal runStatus As String, _
    ByVal consoleMessages As Collection)

    Dim perfScope As clsPerfScope

    Dim wsCalc As Worksheet
    Dim wsState As Worksheet
    Dim tblCalc As ListObject
    Dim tblState As ListObject

    Dim mapCalc As Object
    Dim mapState As Object

    Dim arrCalc As Variant
    Dim arrOut() As Variant

    Dim rowCount As Long
    Dim r As Long

    Set perfScope = Profiler_BeginScope("Write_CalcState_Snapshot_Console", "Excel Table Write")

    On Error GoTo SafeExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Ensure_CalcState_Table_Console consoleMessages

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set wsState = ThisWorkbook.Worksheets(CALC_STATE_SHEET_NAME)

    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set tblState = wsState.ListObjects(CALC_STATE_TABLE_NAME)

    Set mapCalc = BuildColumnMap_CalcState(tblCalc)
    Set mapState = BuildColumnMap_CalcState(tblState)

    ValidateCalcStateSourceColumns mapCalc
    ValidateCalcStateTargetColumns mapState

    If tblCalc.DataBodyRange Is Nothing Then
        ResizeTableToRowCount_CalcState tblState, 0
        GoTo SafeExit
    End If

    arrCalc = tblCalc.DataBodyRange.value
    rowCount = UBound(arrCalc, 1)

    ResizeTableToRowCount_CalcState tblState, rowCount
    ReDim arrOut(1 To rowCount, 1 To tblState.ListColumns.Count)

    For r = 1 To rowCount

        arrOut(r, mapState("ID")) = GetCalcStateVal(arrCalc, r, mapCalc, "ID")
        arrOut(r, mapState("ParentID")) = GetCalcStateVal(arrCalc, r, mapCalc, "ParentID")
        arrOut(r, mapState("IsSummary")) = GetCalcStateVal(arrCalc, r, mapCalc, "IsSummary")
        arrOut(r, mapState("Predecessors WBS")) = GetCalcStateVal(arrCalc, r, mapCalc, "Predecessors WBS")
        arrOut(r, mapState("Cal")) = NormalizeCalendarType(GetCalcStateVal(arrCalc, r, mapCalc, "Cal"))

        arrOut(r, mapState("Baseline Start")) = GetCalcStateVal(arrCalc, r, mapCalc, "Baseline Start")
        arrOut(r, mapState("Baseline Duration")) = GetCalcStateVal(arrCalc, r, mapCalc, "Baseline Duration")
        arrOut(r, mapState("Actual Start")) = GetCalcStateVal(arrCalc, r, mapCalc, "Actual Start")
        arrOut(r, mapState("Actual Finish")) = GetCalcStateVal(arrCalc, r, mapCalc, "Actual Finish")
        arrOut(r, mapState("Forecast Start")) = GetCalcStateVal(arrCalc, r, mapCalc, "Forecast Start")
        arrOut(r, mapState("Forecast Finish")) = GetCalcStateVal(arrCalc, r, mapCalc, "Forecast Finish")
        arrOut(r, mapState("Deadline")) = GetCalcStateVal(arrCalc, r, mapCalc, "Deadline")
        arrOut(r, mapState("Constraint Active")) = GetCalcStateVal(arrCalc, r, mapCalc, "Constraint Active")
        arrOut(r, mapState("Start Constraint Type")) = GetCalcStateVal(arrCalc, r, mapCalc, "Start Constraint Type")
        arrOut(r, mapState("Start Constraint Date")) = GetCalcStateVal(arrCalc, r, mapCalc, "Start Constraint Date")
        arrOut(r, mapState("Finish Constraint Type")) = GetCalcStateVal(arrCalc, r, mapCalc, "Finish Constraint Type")
        arrOut(r, mapState("Finish Constraint Date")) = GetCalcStateVal(arrCalc, r, mapCalc, "Finish Constraint Date")

        arrOut(r, mapState("Row Signature")) = IncrementalSignature_BuildRow(arrCalc, r, mapCalc)
        arrOut(r, mapState("Run Status")) = UCase$(Trim$(runStatus))

    Next r

    tblState.DataBodyRange.value = arrOut
    ApplyCalcStateFormats tblState

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        CalcState_AddStopToConsole consoleMessages, _
            "Erreur dans Write_CalcState_Snapshot : " & Err.Description, _
            "Error in Write_CalcState_Snapshot: " & Err.Description
    End If

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Calc State Add Stop To Console a la structure cible fournie par l'appelant.
' EN: Adds the Calc State Add Stop To Console collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub CalcState_AddStopToConsole( _
    ByVal consoleMessages As Collection, _
    ByVal frText As String, _
    ByVal enText As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg(frText, enText)

End Sub

'------------------------------------------------------------------------------
' FR: Construit la valeur Calc State Headers a partir des donnees fournies par l'appelant.
' EN: Builds the Calc State Headers value from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildCalcStateHeaders() As Variant

    Dim arr

    arr = Array( _
        "ID", _
        "ParentID", _
        "IsSummary", _
        "Predecessors WBS", _
        "Cal", _
        "Baseline Start", _
        "Baseline Duration", _
        "Actual Start", _
        "Actual Finish", _
        "Forecast Start", _
        "Forecast Finish", _
                "Deadline", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Row Signature", _
        "Run Status")

    BuildCalcStateHeaders = arr

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree Worksheet Exists Calc State si necessaire.
' EN: Ensures or creates Worksheet Exists Calc State when needed.
'------------------------------------------------------------------------------
Private Function EnsureWorksheetExists_CalcState( _
    ByVal sheetName As String, _
    Optional ByRef wasCreated As Boolean = False) As Worksheet

    Dim ws As Worksheet
    Dim wsStart As Worksheet

    wasCreated = False
    Set wsStart = ActiveSheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        wasCreated = True

        If Not wsStart Is Nothing Then
            wsStart.Activate
        End If
    End If

    Set EnsureWorksheetExists_CalcState = ws

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree Table With Headers Calc State si necessaire.
' EN: Ensures or creates Table With Headers Calc State when needed.
'------------------------------------------------------------------------------
Private Function EnsureTableWithHeaders_CalcState( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal headers As Variant, _
    Optional ByRef wasCreated As Boolean = False) As ListObject

    Dim tbl As ListObject
    Dim i As Long
    Dim headerCount As Long
    Dim rng As Range
    Dim lc As ListColumn
    Dim headerName As String

    wasCreated = False

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    headerCount = UBound(headers) - LBound(headers) + 1

    If tbl Is Nothing Then

        Set rng = ws.Range(topLeft, topLeft.Offset(1, headerCount - 1))
        rng.Clear

        For i = 0 To headerCount - 1
            rng.Cells(1, i + 1).value = headers(LBound(headers) + i)
        Next i

        Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = tableName
        wasCreated = True

        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.ClearContents
        End If

    Else

        For i = LBound(headers) To UBound(headers)
            headerName = CStr(headers(i))

            Set lc = Nothing
            On Error Resume Next
            Set lc = tbl.ListColumns(headerName)
            On Error GoTo 0

            If lc Is Nothing Then
                Set lc = tbl.ListColumns.Add
                lc.Name = headerName
            End If
        Next i

    End If

    Set EnsureTableWithHeaders_CalcState = tbl

End Function

'------------------------------------------------------------------------------
' FR: Construit la map Column Map CALC State a partir des donnees fournies par l'appelant.
' EN: Builds the Column Map CALC State map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildColumnMap_CalcState(ByVal tbl As ListObject) As Object

    Dim d As Object
    Dim i As Long

    Set d = CreateObject("Scripting.Dictionary")

    For i = 1 To tbl.ListColumns.Count
        d(tbl.ListColumns(i).Name) = i
    Next i

    Set BuildColumnMap_CalcState = d

End Function

'------------------------------------------------------------------------------
' FR: Valide Calc State Source Columns et signale les incoherences detectees.
' EN: Validates Calc State Source Columns and reports detected inconsistencies.
'------------------------------------------------------------------------------
Private Sub ValidateCalcStateSourceColumns(ByVal mapCalc As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = IncrementalSignature_RequiredColumns()

    For Each c In requiredCols
        If Not mapCalc.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 9101, _
                "ValidateCalcStateSourceColumns", _
                "Missing source column in tbl_CALC: " & CStr(c)
        End If
    Next c

End Sub

'------------------------------------------------------------------------------
' FR: Valide Calc State Target Columns et signale les incoherences detectees.
' EN: Validates Calc State Target Columns and reports detected inconsistencies.
'------------------------------------------------------------------------------
Private Sub ValidateCalcStateTargetColumns(ByVal mapState As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = Array( _
        "ID", _
        "ParentID", _
        "IsSummary", _
        "Predecessors WBS", _
        "Cal", _
        "Baseline Start", _
        "Baseline Duration", _
        "Actual Start", _
        "Actual Finish", _
        "Forecast Start", _
        "Forecast Finish", _
                "Deadline", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Row Signature", _
        "Run Status" _
    )

    For Each c In requiredCols
        If Not mapState.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 9102, _
                "ValidateCalcStateTargetColumns", _
                "Missing target column in " & CALC_STATE_TABLE_NAME & ": " & CStr(c)
        End If
    Next c

End Sub


'------------------------------------------------------------------------------
' FR: Traite la reference Resize Table To Row Count CALC State sans modifier les donnees d'entree.
' EN: Handles the Resize Table To Row Count CALC State reference without mutating input data.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ResizeTableToRowCount_CalcState(ByVal tbl As ListObject, ByVal targetRows As Long)

    Dim currentRows As Long
    Dim r As Long

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
' FR: Actualise Apply Calc State Formats sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Calc State Formats without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyCalcStateFormats(ByVal tbl As ListObject)

    On Error Resume Next

    tbl.ListColumns("ID").Range.NumberFormat = "0"
    tbl.ListColumns("ParentID").Range.NumberFormat = "0"

    tbl.ListColumns("IsSummary").Range.NumberFormat = "@"
    tbl.ListColumns("Predecessors WBS").Range.NumberFormat = "@"
    tbl.ListColumns("Cal").Range.NumberFormat = "@"
    tbl.ListColumns("Constraint Active").Range.NumberFormat = "@"
    tbl.ListColumns("Start Constraint Type").Range.NumberFormat = "@"
    tbl.ListColumns("Finish Constraint Type").Range.NumberFormat = "@"
    tbl.ListColumns("Row Signature").Range.NumberFormat = "@"
    tbl.ListColumns("Run Status").Range.NumberFormat = "@"

    tbl.ListColumns("Baseline Start").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Actual Start").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Actual Finish").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Forecast Start").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Forecast Finish").Range.NumberFormat = "dd/mm/yyyy"
        tbl.ListColumns("Deadline").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Start Constraint Date").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Finish Constraint Date").Range.NumberFormat = "dd/mm/yyyy"

    tbl.ListColumns("Baseline Duration").Range.NumberFormat = "0"

    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Retourne Calc State Val depuis le contexte calculation state.
' EN: Returns Calc State Val from the calculation state context.
'------------------------------------------------------------------------------
Private Function GetCalcStateVal( _
    ByRef arr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal colName As String) As Variant

    If Not mapCol.Exists(colName) Then
        GetCalcStateVal = Empty
        Exit Function
    End If

    GetCalcStateVal = arr(rowIdx, mapCol(colName))

End Function

'------------------------------------------------------------------------------
' FR: Construit Calc State Row Signature pour le traitement calculation state.
' EN: Builds Calc State Row Signature for calculation state processing.


'------------------------------------------------------------------------------
' FR: Normalise Incremental Signature Value dans un format exploitable.
' EN: Normalizes Incremental Signature Value into a usable format.

'------------------------------------------------------------------------------
' FR: Normalise Incremental Signature Boolean dans un format exploitable.
' EN: Normalizes Incremental Signature Boolean into a usable format.

'------------------------------------------------------------------------------
' FR: Normalise Incremental Signature Date dans un format exploitable.
' EN: Normalizes Incremental Signature Date into a usable format.

'------------------------------------------------------------------------------
' FR: Normalise Incremental Signature Number dans un format exploitable.
' EN: Normalizes Incremental Signature Number into a usable format.

'------------------------------------------------------------------------------
' FR: Normalise Incremental Signature Text dans un format exploitable.
' EN: Normalizes Incremental Signature Text into a usable format.

'------------------------------------------------------------------------------
' FR: Supprime toutes les lignes de CALC_STATE sans creer son infrastructure.
' EN: Deletes every CALC_STATE row without creating its infrastructure.
'------------------------------------------------------------------------------
Public Sub CalcState_ResetStorage()

    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_STATE_SHEET_NAME)
    If Not ws Is Nothing Then Set tbl = ws.ListObjects(CALC_STATE_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub

    ResizeTableToRowCount_CalcState tbl, 0

End Sub

'------------------------------------------------------------------------------
' FR: Lit le snapshot incremental sous forme ID -> signature sans exposer le schema CALC_STATE.
' EN: Reads the incremental snapshot as ID -> signature without exposing the CALC_STATE schema.
'------------------------------------------------------------------------------
Public Function CalcState_GetSignatureByIdMap() As Object

    Dim result As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mapState As Object
    Dim arrState As Variant
    Dim r As Long
    Dim idVal As String

    Set result = CreateObject("Scripting.Dictionary")
    On Error GoTo SafeExit
    Set ws = ThisWorkbook.Worksheets(CALC_STATE_SHEET_NAME)
    Set tbl = ws.ListObjects(CALC_STATE_TABLE_NAME)
    If tbl.DataBodyRange Is Nothing Then GoTo SafeExit

    Set mapState = BuildColumnMap_CalcState(tbl)
    If Not mapState.Exists("ID") Then Err.Raise vbObjectError + 9202, "CalcState_GetSignatureByIdMap", "Missing required column in tbl_CALC_STATE: ID"
    If Not mapState.Exists("Row Signature") Then Err.Raise vbObjectError + 9202, "CalcState_GetSignatureByIdMap", "Missing required column in tbl_CALC_STATE: Row Signature"
    arrState = tbl.DataBodyRange.value

    For r = 1 To UBound(arrState, 1)
        idVal = Trim$(CStr(arrState(r, mapState("ID"))))
        If idVal <> "" Then result(idVal) = Trim$(CStr(arrState(r, mapState("Row Signature"))))
    Next r

SafeExit:
    Set CalcState_GetSignatureByIdMap = result

End Function

'------------------------------------------------------------------------------
' FR: Retourne le statut du snapshot incremental courant, ou une chaine vide s'il est absent.
' EN: Returns the current incremental snapshot status, or an empty string when absent.
'------------------------------------------------------------------------------
Public Function CalcState_GetRunStatus() As String

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mapState As Object
    Dim arrState As Variant

    On Error GoTo SafeExit
    Set ws = ThisWorkbook.Worksheets(CALC_STATE_SHEET_NAME)
    Set tbl = ws.ListObjects(CALC_STATE_TABLE_NAME)
    If tbl.DataBodyRange Is Nothing Then Exit Function
    Set mapState = BuildColumnMap_CalcState(tbl)
    If Not mapState.Exists("Run Status") Then Exit Function
    arrState = tbl.DataBodyRange.value
    CalcState_GetRunStatus = Trim$(CStr(arrState(1, mapState("Run Status"))))

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Invalide le snapshot incremental existant en positionnant son statut a DIRTY.
' EN: Invalidates the existing incremental snapshot by setting its status to DIRTY.
'------------------------------------------------------------------------------
Public Sub CalcState_MarkDirty()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mapState As Object

    On Error GoTo SafeExit
    Set ws = ThisWorkbook.Worksheets(CALC_STATE_SHEET_NAME)
    Set tbl = ws.ListObjects(CALC_STATE_TABLE_NAME)

    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Set mapState = BuildColumnMap_CalcState(tbl)
    If Not mapState.Exists("Run Status") Then Exit Sub

    tbl.ListColumns("Run Status").DataBodyRange.value = "DIRTY"

SafeExit:
End Sub
