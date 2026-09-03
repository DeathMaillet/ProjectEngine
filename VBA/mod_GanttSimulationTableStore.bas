Attribute VB_Name = "mod_GanttSimulationTableStore"
Option Explicit

'===============================================================================
' MODULE : mod_GanttSimulationTableStore
' DOMAINE / DOMAIN : Gantt
'
' FR
' Possede le schema, la lecture, l'ecriture et le reset de son stockage Excel.
' Ne porte aucune decision de simulation ou de rendu.
'
' EN
' Owns schema, reads, writes and reset of its Excel store.
' Owns no simulation or rendering decision.
'
' CONTRATS / CONTRACTS : Ensure_CalcGanttTest_Sheet, Ensure_CalcGanttTest_Table, ValidateCalcGanttTestColumns, ResizeTableToRowCount_Generic, FormatCalcGanttTestColumns, RequireMapColumn_GanttLive, GetColumnIndex_GanttLive, GanttSimulation_ResetTableStorage
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const CALC_GANTT_TEST_SHEET As String = "CALC_GANTT_TEST"
Private Const CALC_GANTT_TEST_TABLE As String = "tbl_CALC_GANTT_TEST"


'------------------------------------------------------------------------------
' FR: Retourne la feuille CALC_GANTT_TEST ou la cree juste apres GANTT si elle manque.
' EN: Returns the CALC_GANTT_TEST sheet or creates it just after GANTT when missing.
'------------------------------------------------------------------------------
Public Function Ensure_CalcGanttTest_Sheet() As Worksheet

    Dim ws As Worksheet
    Dim wsGantt As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set wsGantt = ThisWorkbook.Worksheets(GANTT_SHEET)
        Set ws = ThisWorkbook.Worksheets.Add(After:=wsGantt)
        ws.Name = CALC_GANTT_TEST_SHEET
    End If

    Set Ensure_CalcGanttTest_Sheet = ws

End Function

'------------------------------------------------------------------------------
' FR: Garantit le schema exact de tbl_CALC_GANTT_TEST et reconstruit la table si ses colonnes divergent.
' EN: Ensures the exact tbl_CALC_GANTT_TEST schema and rebuilds the table when columns diverge.
'------------------------------------------------------------------------------
Public Function Ensure_CalcGanttTest_Table(ByVal ws As Worksheet) As ListObject

    Dim tbl As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim expectedColCount As Long
    Dim needsRebuild As Boolean

    On Error Resume Next
    Set tbl = ws.ListObjects(CALC_GANTT_TEST_TABLE)
    On Error GoTo 0

    headers = Array( _
        "ID", "WBS", "Task Type", "Cal", "Is Summary", "Parent ID", "Base Start", "Base Finish", "Base Duration", _
        "Base Progress", "Driving Logic", "Has Actual", "Test Start", "Test Finish", "Test % Raw", _
        "Test % Normalized", "Input Start", "Input Finish", "Input Duration", "Input Progress", _
        "Predecessors", "Lag", "Any Test Value", "Calc Test Start", "Calc Test Finish", _
        "Calc Test Duration", "Calc Test Progress", "Warning Flag", "Warning Text", "Error Flag")

    expectedColCount = UBound(headers) + 1

    If Not tbl Is Nothing Then
        If tbl.ListColumns.Count <> expectedColCount Then
            needsRebuild = True
        Else
            For i = LBound(headers) To UBound(headers)
                If tbl.ListColumns(i + 1).Name <> CStr(headers(i)) Then
                    needsRebuild = True
                    Exit For
                End If
            Next i
        End If
    End If

    If tbl Is Nothing Or needsRebuild Then

        ws.Cells.Clear

        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).value = headers(i)
        Next i

        Set tbl = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range(ws.Cells(1, 1), ws.Cells(2, expectedColCount)), _
            XlListObjectHasHeaders:=xlYes)

        tbl.Name = CALC_GANTT_TEST_TABLE

        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.ClearContents
        End If
    End If

    Set Ensure_CalcGanttTest_Table = tbl

End Function


'------------------------------------------------------------------------------
' FR: Verifie que tbl_CALC_GANTT_TEST expose toutes les colonnes attendues par les backends live.
' EN: Validates that tbl_CALC_GANTT_TEST exposes every column expected by live backends.
'------------------------------------------------------------------------------
Public Sub ValidateCalcGanttTestColumns(ByVal tbl As ListObject)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "ID", "WBS", "Task Type", "Cal", "Is Summary", "Parent ID", "Base Start", "Base Finish", "Base Duration", _
        "Base Progress", "Driving Logic", "Has Actual", "Test Start", "Test Finish", "Test % Raw", _
        "Test % Normalized", "Input Start", "Input Finish", "Input Duration", "Input Progress", _
        "Predecessors", "Lag", "Any Test Value", "Calc Test Start", "Calc Test Finish", _
        "Calc Test Duration", "Calc Test Progress", "Warning Flag", "Warning Text", "Error Flag")

    For i = LBound(headers) To UBound(headers)
        Call GetColumnIndex_GanttLive(tbl, CStr(headers(i)))
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Ajuste le nombre de lignes d'une table Excel a un nombre cible sans changer son schema.
' EN: Resizes an Excel table to a target row count without changing its schema.
'------------------------------------------------------------------------------
Public Sub ResizeTableToRowCount_Generic(ByVal tbl As ListObject, ByVal targetRows As Long)

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
' FR: Normalise ou formate Format Calc Gantt Test Columns selon le contrat canonique du composant.
' EN: Normalizes or formats Format Calc Gantt Test Columns according to the component contract.
'------------------------------------------------------------------------------

Public Sub FormatCalcGanttTestColumns(ByVal tbl As ListObject)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    tbl.ListColumns("ID").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("WBS").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Task Type").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Cal").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Is Summary").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Parent ID").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Driving Logic").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Has Actual").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Predecessors").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Any Test Value").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Warning Flag").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Warning Text").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Error Flag").DataBodyRange.NumberFormat = "@"
    tbl.ListColumns("Base Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Base Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Test Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Test Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Input Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Input Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Calc Test Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Calc Test Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"

    tbl.ListColumns("Base Progress").DataBodyRange.NumberFormat = "0%"
    tbl.ListColumns("Test % Raw").DataBodyRange.NumberFormat = "0%"
    tbl.ListColumns("Test % Normalized").DataBodyRange.NumberFormat = "0%"
    tbl.ListColumns("Input Progress").DataBodyRange.NumberFormat = "0%"
    tbl.ListColumns("Calc Test Progress").DataBodyRange.NumberFormat = "0%"

End Sub

'------------------------------------------------------------------------------
' FR: Leve une erreur bilingue standardisee quand une colonne live obligatoire manque.
' EN: Raises a standardized bilingual error when a required live column is missing.
'------------------------------------------------------------------------------
Private Sub RaiseMissingGanttLiveColumn( _
    ByVal sourceName As String, _
    ByVal colName As String, _
    ByVal functionName As String)

    Err.Raise vbObjectError + 913, functionName, _
        "FR:" & vbCrLf & _
        "Colonne requise introuvable pour GANTT live/scenario" & vbCrLf & vbCrLf & _
        "Table/source : " & sourceName & vbCrLf & _
        "Colonne : " & colName & vbCrLf & _
        "Fonction : " & functionName & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        "Required column missing for GANTT live/scenario" & vbCrLf & vbCrLf & _
        "Source/table: " & sourceName & vbCrLf & _
        "Column: " & colName & vbCrLf & _
        "Function: " & functionName

End Sub

'------------------------------------------------------------------------------
' FR: Retourne l'index d'une colonne mappee ou declenche l'erreur de schema GanttLive.
' EN: Returns a mapped column index or raises the GanttLive schema error.
'------------------------------------------------------------------------------
Public Function RequireMapColumn_GanttLive( _
    ByVal mapObj As Object, _
    ByVal sourceName As String, _
    ByVal colName As String, _
    ByVal functionName As String) As Long

    If mapObj Is Nothing Then
        RaiseMissingGanttLiveColumn sourceName, colName, functionName
    End If

    If Not mapObj.Exists(colName) Then
        RaiseMissingGanttLiveColumn sourceName, colName, functionName
    End If

    RequireMapColumn_GanttLive = CLng(mapObj(colName))

End Function

'------------------------------------------------------------------------------
' FR: Retrouve l'index d'une colonne directement dans une ListObject live.
' EN: Finds a column index directly from a live ListObject.
'------------------------------------------------------------------------------
Public Function GetColumnIndex_GanttLive(ByVal tbl As ListObject, ByVal colName As String) As Long

    Dim i As Long

    For i = 1 To tbl.ListColumns.Count
        If tbl.ListColumns(i).Name = colName Then
            GetColumnIndex_GanttLive = i
            Exit Function
        End If
    Next i

    RaiseMissingGanttLiveColumn tbl.Name, colName, "GetColumnIndex_GanttLive"

End Function
'------------------------------------------------------------------------------
' FR: Teste la presence d'une colonne dans une table avec comparaison insensible a la casse.
' EN: Checks whether a table has a column using case-insensitive comparison.
'------------------------------------------------------------------------------
Private Function TableHasColumn_GanttLive(ByVal tbl As ListObject, ByVal colName As String) As Boolean

    Dim i As Long

    For i = 1 To tbl.ListColumns.Count
        If StrComp(CStr(tbl.ListColumns(i).Name), colName, vbTextCompare) = 0 Then
            TableHasColumn_GanttLive = True
            Exit Function
        End If
    Next i

End Function

'------------------------------------------------------------------------------
' FR: Supprime toutes les lignes du store de simulation sans creer son infrastructure.
' EN: Deletes every simulation-store row without creating its infrastructure.
'------------------------------------------------------------------------------
Public Sub GanttSimulation_ResetTableStorage()

    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    If Not ws Is Nothing Then Set tbl = ws.ListObjects(CALC_GANTT_TEST_TABLE)
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub

    ResizeTableToRowCount_Generic tbl, 0

End Sub

'------------------------------------------------------------------------------
' FR: Vide les valeurs de simulation en conservant les lignes et le schema de la table.
' EN: Clears simulation values while preserving the table rows and schema.
'------------------------------------------------------------------------------
Public Sub GanttSimulation_ClearResults()

    Call GanttSimulation_ClearResultsIfAny

End Sub

'------------------------------------------------------------------------------
' FR:
' Vide le store uniquement lorsqu'il contient encore au moins une valeur.
' Retourne True lorsqu'une mutation a réellement été effectuée.
'
' EN:
' Clears the store only when it still contains at least one value.
' Returns True when a mutation was actually performed.
'------------------------------------------------------------------------------
Public Function GanttSimulation_ClearResultsIfAny() As Boolean

    Dim ws As Worksheet
    Dim tbl As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    If Not ws Is Nothing Then Set tbl = ws.ListObjects(CALC_GANTT_TEST_TABLE)
    On Error GoTo 0

    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(tbl.DataBodyRange) = 0 Then Exit Function

    tbl.DataBodyRange.ClearContents
    GanttSimulation_ClearResultsIfAny = True

End Function

'------------------------------------------------------------------------------
' FR: Construit les resultats de simulation par ID pour le renderer et les transactions live.
' EN: Builds simulation results by ID for the renderer and live transactions.
'------------------------------------------------------------------------------
Public Function GanttSimulation_BuildResultByIdMap() As Object

    Dim d As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mapTest As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")
    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    Set tbl = ws.ListObjects(CALC_GANTT_TEST_TABLE)
    If tbl.DataBodyRange Is Nothing Then GoTo SafeExit

    Set mapTest = CanonicalIdentity_BuildColumnMap(tbl)
    arr = tbl.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapTest("ID"))))
        If idVal <> "" Then
            d(idVal) = Array( _
                GetCellValue(arr(r, mapTest("Calc Test Start"))), _
                GetCellValue(arr(r, mapTest("Calc Test Finish"))), _
                GetCellValue(arr(r, mapTest("Calc Test Duration"))), _
                GetCellValue(arr(r, mapTest("Calc Test Progress"))), _
                UCase$(Trim$(CStr(arr(r, mapTest("Is Summary"))))), _
                Trim$(CStr(arr(r, mapTest("Error Flag")))), _
                Trim$(CStr(arr(r, mapTest("Any Test Value")))) _
            )
        End If
    Next r

SafeExit:
    Set GanttSimulation_BuildResultByIdMap = d

End Function

'------------------------------------------------------------------------------
' FR: Construit les valeurs de baseline scenario par ID sans exposer le schema du store.
' EN: Builds scenario baseline values by ID without exposing the store schema.
'------------------------------------------------------------------------------
Public Function GanttSimulation_BuildScenarioBaselineById() As Object

    Dim d As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mapTest As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim summaryValue As String
    Dim taskTypeValue As String
    Dim isSummary As Boolean
    Dim isLOE As Boolean
    Dim requiredColumn As Variant

    Set d = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    Set tbl = ws.ListObjects(CALC_GANTT_TEST_TABLE)
    If tbl.DataBodyRange Is Nothing Then
        Set GanttSimulation_BuildScenarioBaselineById = d
        Exit Function
    End If

    Set mapTest = CanonicalIdentity_BuildColumnMap(tbl)
    For Each requiredColumn In Array("ID", "Task Type", "Is Summary", "Calc Test Start", "Calc Test Duration", "Input Progress")
        If Not mapTest.Exists(CStr(requiredColumn)) Then Err.Raise vbObjectError + 1291, , "Missing scenario column: " & CStr(requiredColumn)
    Next requiredColumn

    arr = tbl.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapTest("ID"))))
        If idVal <> "" And Not d.Exists(idVal) Then
            summaryValue = UCase$(Trim$(CStr(arr(r, mapTest("Is Summary")))))
            isSummary = (summaryValue = "YES" Or summaryValue = "TRUE" Or summaryValue = "1" Or summaryValue = "SUMMARY")
            taskTypeValue = UCase$(Trim$(CStr(arr(r, mapTest("Task Type")))))
            isLOE = (taskTypeValue = "LEVEL OF EFFORT" Or taskTypeValue = "LOE" Or taskTypeValue = "LEVEL-OF-EFFORT" Or taskTypeValue = "LEVEL_OF_EFFORT")
            d(idVal) = Array( _
                GetCellValue(arr(r, mapTest("Calc Test Start"))), _
                GetCellValue(arr(r, mapTest("Calc Test Duration"))), _
                GetCellValue(arr(r, mapTest("Input Progress"))), _
                isSummary, _
                isLOE)
        End If
    Next r

    Set GanttSimulation_BuildScenarioBaselineById = d

End Function

'------------------------------------------------------------------------------
' FR: Indique si le store de simulation contient au moins une erreur de calcul.
' EN: Returns whether the simulation store contains at least one calculation error.
'------------------------------------------------------------------------------
Public Function GanttSimulation_HasErrors() As Boolean

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mapTest As Object
    Dim arr As Variant
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(CALC_GANTT_TEST_SHEET)
    Set tbl = ws.ListObjects(CALC_GANTT_TEST_TABLE)

    On Error GoTo SafeExit
    If tbl.DataBodyRange Is Nothing Then Exit Function
    Set mapTest = CanonicalIdentity_BuildColumnMap(tbl)
    If Not mapTest.Exists("Error Flag") Then Exit Function
    arr = tbl.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, mapTest("Error Flag"))))) = "ERROR" Then
            GanttSimulation_HasErrors = True
            Exit Function
        End If
    Next r
    Exit Function

SafeExit:
    GanttSimulation_HasErrors = True

End Function
