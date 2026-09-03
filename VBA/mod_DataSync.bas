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
    Dim summaryWbsByWbs As Object

    Dim colsToCopy As Variant
    Dim colsToClear As Variant
    Dim wbsData As Variant
    Dim inputColumns As Object
    Dim missingIdRows As Object

    Dim i As Long
    Dim r As Long
    Dim maxId As Long
    Dim targetRows As Long

    Dim idValue As Variant
    Dim idKey As String
    Dim wbsValue As String
    Dim wbsRowIndex As Long

    Dim consoleMessages As Collection
    Dim previousCalculation As XlCalculation
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousStatusBar As Variant
    Dim applicationStateCaptured As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim resizeOperations As Long
    Dim valueBlockWrites As Long
    Dim skippedValueBlocks As Long
    Dim formatRangeWrites As Long
    Dim outputClearWrites As Long
    Dim targetedCalculations As Long

    Set perfScope = Profiler_BeginScope("Sync_WBS_To_CALC", "Excel Table Sync")

    On Error GoTo SafeExit

    Set consoleMessages = New Collection

    previousCalculation = Application.Calculation
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousStatusBar = Application.StatusBar
    applicationStateCaptured = True

    Application.Calculation = xlCalculationManual
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
    Set mapWBS = SchemaBuildColumnKeyMap(tblWBS, VTS_TABLE_WBS)
    Set wbsIdRows = CreateObject("Scripting.Dictionary")
    Set inputColumns = CreateObject("Scripting.Dictionary")
    Set missingIdRows = CreateObject("Scripting.Dictionary")

    colsToCopy = Array( _
        VTS_COL_WBS, _
        VTS_COL_TASK_NAME, _
        VTS_COL_TASK_TYPE, _
        VTS_COL_S, _
        VTS_COL_CAL, _
        VTS_COL_PREDECESSORS_WBS, _
        VTS_COL_BASELINE_START, _
        VTS_COL_BASELINE_DURATION, _
        VTS_COL_BASELINE_FINISH, _
        VTS_COL_ACTUAL_START, _
        VTS_COL_ACTUAL_FINISH, _
        VTS_COL_ACTUAL_DURATION, _
        VTS_COL_FORECAST_START, _
        VTS_COL_FORECAST_FINISH _
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
        "Longest Path REX", _
        "Total Float REX", _
        "Free Float REX", _
        "Deadline Float" _
    )

    For i = 1 To tblCalc.ListColumns.Count
        mapCalc(tblCalc.ListColumns(i).Name) = i
    Next i

    If Not mapWBS.Exists(VTS_COL_ID) Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_WBS.", _
            "Column ID was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists(VTS_COL_WBS) Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne WBS est introuvable dans tbl_WBS.", _
            "Column WBS was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists(VTS_COL_TASK_TYPE) Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Task Type est introuvable dans tbl_WBS.", _
            "Column Task Type was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists(VTS_COL_S) Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne S est introuvable dans tbl_WBS.", _
            "Column S was not found in tbl_WBS."
        GoTo SafeExit
    End If
    If Not mapWBS.Exists(VTS_COL_PREDECESSORS_WBS) Then
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

    DataSync_CalculateManagedWBSFormulaRanges tblWBS, targetedCalculations
    If Not tblWBS.DataBodyRange Is Nothing Then wbsData = tblWBS.DataBodyRange.value

    maxId = DataSync_IndexWBSRows( _
        wbsData, mapWBS, wbsIdRows, consoleMessages)
    If maxId < 0 Then GoTo SafeExit

    Set summaryWbsByWbs = DataSync_BuildSummaryWbsLookupFromArray(wbsData, mapWBS)
    NormalizeWBSSummaryDisplayValues tblWBS, mapWBS, summaryWbsByWbs
    NormalizeWBSCalendarValues tblWBS, mapWBS, summaryWbsByWbs

    'The two normalization contracts may update WBS. Refresh the source buffer
    'once before constructing CALC so the observable legacy result is preserved.
    If Not tblWBS.DataBodyRange Is Nothing Then wbsData = tblWBS.DataBodyRange.value

    targetRows = maxId

    If targetRows = 0 Then
        DataSync_ResizeCalcTable tblCalc, 0, resizeOperations
        GoTo SafeExit
    End If

    DataSync_ResizeCalcTable tblCalc, targetRows, resizeOperations

    DataSync_BuildCalcInputColumns _
        wbsData, mapWBS, mapCalc, wbsIdRows, summaryWbsByWbs, _
        targetRows, colsToCopy, inputColumns, missingIdRows

    DataSync_WriteCalcInputBlocks _
        tblCalc, mapCalc, inputColumns, _
        valueBlockWrites, skippedValueBlocks

    DataSync_ApplyCalcInputFormats tblCalc, formatRangeWrites

    If Not preserveCalcOutputs Then
        DataSync_ClearCalcOutputColumns _
            tblCalc, mapCalc, colsToClear, outputClearWrites
    ElseIf missingIdRows.Count > 0 Then
        DataSync_ClearMissingIdOutputs _
            tblCalc, mapCalc, colsToClear, missingIdRows, outputClearWrites
    End If

SafeExit:
    errorNumber = Err.Number
    errorDescription = Err.Description

    Profiler_RecordOperation "DataSyncResizeOperations", resizeOperations, 0#
    Profiler_RecordOperation "DataSyncValueBlockWrites", valueBlockWrites, 0#
    Profiler_RecordOperation "DataSyncSkippedValueBlocks", skippedValueBlocks, 0#
    Profiler_RecordOperation "DataSyncFormatRangeWrites", formatRangeWrites, 0#
    Profiler_RecordOperation "DataSyncOutputClearWrites", outputClearWrites, 0#
    Profiler_RecordOperation "DataSyncTargetedCalculations", targetedCalculations, 0#

    If applicationStateCaptured Then
        On Error Resume Next
        Application.Calculation = previousCalculation
        Application.EnableEvents = previousEnableEvents
        Application.ScreenUpdating = previousScreenUpdating
        Application.StatusBar = previousStatusBar
        On Error GoTo 0
    End If

    If errorNumber <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur dans Sync_WBS_To_CALC : " & errorDescription, _
            "Error in Sync_WBS_To_CALC: " & errorDescription
    End If

    If Not consoleMessages Is Nothing Then
        CalcBridge_ShowPlanningConsole consoleMessages
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Calcule une seule fois les colonnes de formules WBS lues par le Sync.
' EN: Calculates once the managed WBS formula columns read by the sync.
'------------------------------------------------------------------------------
Private Sub DataSync_CalculateManagedWBSFormulaRanges( _
    ByVal tblWBS As ListObject, _
    ByRef calculationCount As Long)

    Dim formulaRange As Range
    Dim columnKey As Variant
    Dim targetRange As Range

    If tblWBS Is Nothing Then Exit Sub
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    For Each columnKey In Array( _
        VTS_COL_BASELINE_FINISH, VTS_COL_ACTUAL_DURATION, VTS_COL_CALCULATED_DURATION)
        Set targetRange = SchemaListColumn(tblWBS, VTS_TABLE_WBS, CStr(columnKey)).DataBodyRange
        If Not targetRange Is Nothing Then
            If formulaRange Is Nothing Then
                Set formulaRange = targetRange
            Else
                Set formulaRange = Union(formulaRange, targetRange)
            End If
        End If
    Next columnKey

    If Not formulaRange Is Nothing Then
        formulaRange.Calculate
        calculationCount = calculationCount + 1
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Indexe les lignes WBS par ID depuis le buffer memoire et retourne l'ID maximal.
' EN: Indexes WBS rows by ID from the memory buffer and returns the maximum ID.
'------------------------------------------------------------------------------
Private Function DataSync_IndexWBSRows( _
    ByVal wbsData As Variant, _
    ByVal mapWBS As Object, _
    ByVal wbsIdRows As Object, _
    ByVal consoleMessages As Collection) As Long

    Dim r As Long
    Dim idValue As Variant
    Dim idNumber As Long
    Dim idKey As String
    Dim wbsValue As String
    Dim wbsToId As Object

    DataSync_IndexWBSRows = 0
    If Not IsArray(wbsData) Then Exit Function
    Set wbsToId = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(wbsData, 1)
        idValue = wbsData(r, mapWBS(VTS_COL_ID))
        wbsValue = Replace$(Trim$(CStr(wbsData(r, mapWBS(VTS_COL_WBS)))), ",", ".")

        If Trim$(CStr(idValue)) <> "" Then
            If Not IsNumeric(idValue) Then
                DataSync_AddConsoleMessage consoleMessages, "STOP", _
                    "ID non numerique detecte dans WBS : " & CStr(idValue), _
                    "Non-numeric ID detected in WBS: " & CStr(idValue)
                DataSync_IndexWBSRows = -1
                Exit Function
            End If

            idNumber = CLng(idValue)
            If idNumber < 1 Then
                DataSync_AddConsoleMessage consoleMessages, "STOP", _
                    "ID invalide dans WBS (doit etre >= 1) : " & CStr(idValue), _
                    "Invalid ID in WBS (must be >= 1): " & CStr(idValue)
                DataSync_IndexWBSRows = -1
                Exit Function
            End If

            idKey = CStr(idNumber)
            wbsIdRows(idKey) = r
            If idNumber > DataSync_IndexWBSRows Then DataSync_IndexWBSRows = idNumber

            If wbsValue <> "" Then
                If wbsToId.Exists(wbsValue) Then
                    DataSync_AddConsoleMessage consoleMessages, "STOP", _
                        "WBS duplique detecte dans WBS : " & wbsValue, _
                        "Duplicate WBS detected in WBS: " & wbsValue
                    DataSync_IndexWBSRows = -1
                    Exit Function
                End If
                wbsToId(wbsValue) = idNumber
            End If
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Redimensionne tbl_CALC en une operation lorsque la taille cible change.
' EN: Resizes tbl_CALC in one operation when the target size changes.
'------------------------------------------------------------------------------
Private Sub DataSync_ResizeCalcTable( _
    ByVal tblCalc As ListObject, _
    ByVal targetRows As Long, _
    ByRef resizeCount As Long)

    Dim targetRange As Range

    If tblCalc Is Nothing Then Exit Sub
    If targetRows < 0 Then targetRows = 0
    If tblCalc.ListRows.Count = targetRows Then Exit Sub

    If targetRows = 0 Then
        Do While tblCalc.ListRows.Count > 0
            tblCalc.ListRows(tblCalc.ListRows.Count).Delete
        Loop
        resizeCount = resizeCount + 1
        Exit Sub
    End If

    Set targetRange = tblCalc.HeaderRowRange.Cells(1, 1).Resize( _
        targetRows + 1, tblCalc.ListColumns.Count)
    tblCalc.Resize targetRange
    resizeCount = resizeCount + 1

End Sub

'------------------------------------------------------------------------------
' FR: Construit en memoire l'etat final des colonnes d'entree CALC.
' EN: Builds the final CALC input-column state in memory.
'------------------------------------------------------------------------------
Private Sub DataSync_BuildCalcInputColumns( _
    ByVal wbsData As Variant, _
    ByVal mapWBS As Object, _
    ByVal mapCalc As Object, _
    ByVal wbsIdRows As Object, _
    ByVal summaryWbsByWbs As Object, _
    ByVal targetRows As Long, _
    ByVal copiedColumns As Variant, _
    ByVal inputColumns As Object, _
    ByVal missingIdRows As Object)

    Dim columnKey As Variant
    Dim calcColumnName As String
    Dim columnValues() As Variant
    Dim calcRow As Long
    Dim wbsRow As Long
    Dim taskTypeValue As String
    Dim wbsValue As String
    Dim sourceValue As Variant

    ReDim columnValues(1 To targetRows, 1 To 1)
    For calcRow = 1 To targetRows
        columnValues(calcRow, 1) = calcRow
    Next calcRow
    inputColumns("ID") = columnValues

    For Each columnKey In copiedColumns
        calcColumnName = SchemaColumnTitle( _
            VTS_TABLE_WBS, CStr(columnKey), VTS_LANG_EN)
        If Not mapWBS.Exists(CStr(columnKey)) _
            Or Not mapCalc.Exists(calcColumnName) Then
            GoTo NextColumn
        End If

        ReDim columnValues(1 To targetRows, 1 To 1)

        For calcRow = 1 To targetRows
            If wbsIdRows.Exists(CStr(calcRow)) Then
                wbsRow = CLng(wbsIdRows(CStr(calcRow)))
                sourceValue = wbsData(wbsRow, mapWBS(CStr(columnKey)))

                Select Case CStr(columnKey)
                    Case VTS_COL_WBS, VTS_COL_PREDECESSORS_WBS
                        columnValues(calcRow, 1) = CStr(sourceValue)

                    Case VTS_COL_TASK_TYPE
                        columnValues(calcRow, 1) = NormalizeTaskTypeValue(sourceValue)

                    Case VTS_COL_S
                        columnValues(calcRow, 1) = NormalizeSummaryDisplayValue(sourceValue)

                    Case VTS_COL_CAL
                        wbsValue = Replace$(Trim$(CStr(wbsData(wbsRow, mapWBS(VTS_COL_WBS)))), ",", ".")
                        taskTypeValue = UCase$(NormalizeTaskTypeValue( _
                            wbsData(wbsRow, mapWBS(VTS_COL_TASK_TYPE))))

                        If summaryWbsByWbs.Exists(wbsValue) _
                            Or taskTypeValue = "LEVEL OF EFFORT" _
                            Or taskTypeValue = "MILESTONE" Then
                            columnValues(calcRow, 1) = vbNullString
                        Else
                            columnValues(calcRow, 1) = NormalizeCalendarType(sourceValue)
                        End If

                    Case Else
                        columnValues(calcRow, 1) = sourceValue
                End Select
            Else
                columnValues(calcRow, 1) = Empty
                missingIdRows(CStr(calcRow)) = True
            End If
        Next calcRow

        inputColumns(calcColumnName) = columnValues
NextColumn:
    Next columnKey

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit les colonnes d'entree CALC par groupes contigus seulement si elles divergent.
' EN: Writes CALC input columns in contiguous groups only when they differ.
'------------------------------------------------------------------------------
Private Sub DataSync_WriteCalcInputBlocks( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal inputColumns As Object, _
    ByRef writeCount As Long, _
    ByRef skippedCount As Long)

    Dim sortedNames As Variant
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim i As Long
    Dim currentIndex As Long
    Dim previousIndex As Long

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    sortedNames = DataSync_SortedInputColumnNames(mapCalc, inputColumns)
    groupStart = LBound(sortedNames)
    previousIndex = CLng(mapCalc(CStr(sortedNames(groupStart))))

    For i = groupStart + 1 To UBound(sortedNames)
        currentIndex = CLng(mapCalc(CStr(sortedNames(i))))
        If currentIndex <> previousIndex + 1 Then
            groupEnd = i - 1
            DataSync_WriteOneInputBlock tblCalc, mapCalc, inputColumns, _
                sortedNames, groupStart, groupEnd, writeCount, skippedCount
            groupStart = i
        End If
        previousIndex = currentIndex
    Next i

    DataSync_WriteOneInputBlock tblCalc, mapCalc, inputColumns, _
        sortedNames, groupStart, UBound(sortedNames), writeCount, skippedCount

End Sub

'------------------------------------------------------------------------------
' FR: Trie les noms des colonnes d'entree selon leur position reelle dans CALC.
' EN: Sorts input column names by their actual CALC position.
'------------------------------------------------------------------------------
Private Function DataSync_SortedInputColumnNames( _
    ByVal mapCalc As Object, _
    ByVal inputColumns As Object) As Variant

    Dim names() As Variant
    Dim key As Variant
    Dim i As Long
    Dim j As Long
    Dim swapValue As Variant

    ReDim names(0 To inputColumns.Count - 1)
    For Each key In inputColumns.Keys
        names(i) = CStr(key)
        i = i + 1
    Next key

    For i = LBound(names) To UBound(names) - 1
        For j = i + 1 To UBound(names)
            If CLng(mapCalc(CStr(names(j)))) < CLng(mapCalc(CStr(names(i)))) Then
                swapValue = names(i)
                names(i) = names(j)
                names(j) = swapValue
            End If
        Next j
    Next i

    DataSync_SortedInputColumnNames = names

End Function

'------------------------------------------------------------------------------
' FR: Compare puis ecrit un groupe contigu de colonnes CALC en une affectation.
' EN: Compares and writes one contiguous CALC column group in one assignment.
'------------------------------------------------------------------------------
Private Sub DataSync_WriteOneInputBlock( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal inputColumns As Object, _
    ByVal sortedNames As Variant, _
    ByVal firstNameIndex As Long, _
    ByVal lastNameIndex As Long, _
    ByRef writeCount As Long, _
    ByRef skippedCount As Long)

    Dim targetRange As Range
    Dim expectedValues() As Variant
    Dim currentValues As Variant
    Dim columnValues As Variant
    Dim rowCount As Long
    Dim columnCount As Long
    Dim r As Long
    Dim c As Long
    Dim firstColumn As Long

    rowCount = tblCalc.ListRows.Count
    columnCount = lastNameIndex - firstNameIndex + 1
    firstColumn = CLng(mapCalc(CStr(sortedNames(firstNameIndex))))

    ReDim expectedValues(1 To rowCount, 1 To columnCount)
    For c = 1 To columnCount
        columnValues = inputColumns(CStr(sortedNames(firstNameIndex + c - 1)))
        For r = 1 To rowCount
            expectedValues(r, c) = columnValues(r, 1)
        Next r
    Next c

    Set targetRange = tblCalc.DataBodyRange.Cells(1, firstColumn).Resize( _
        rowCount, columnCount)
    currentValues = targetRange.value

    If DataSync_RangeValuesEqual(currentValues, expectedValues, rowCount, columnCount) Then
        skippedCount = skippedCount + 1
    Else
        targetRange.value = expectedValues
        writeCount = writeCount + 1
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Compare deux buffers de plage sans conversion destructive des types VBA.
' EN: Compares two range buffers without destructive VBA type conversion.
'------------------------------------------------------------------------------
Private Function DataSync_RangeValuesEqual( _
    ByVal currentValues As Variant, _
    ByVal expectedValues As Variant, _
    ByVal rowCount As Long, _
    ByVal columnCount As Long) As Boolean

    Dim r As Long
    Dim c As Long

    If rowCount = 1 And columnCount = 1 And Not IsArray(currentValues) Then
        DataSync_RangeValuesEqual = DataSync_ValuesEqual( _
            currentValues, expectedValues(1, 1))
        Exit Function
    End If

    For r = 1 To rowCount
        For c = 1 To columnCount
            If Not DataSync_ValuesEqual(currentValues(r, c), expectedValues(r, c)) Then
                Exit Function
            End If
        Next c
    Next r

    DataSync_RangeValuesEqual = True

End Function

'------------------------------------------------------------------------------
' FR: Compare deux valeurs Excel en preservant les distinctions utiles au moteur.
' EN: Compares two Excel values while preserving distinctions relevant to the engine.
'------------------------------------------------------------------------------
Private Function DataSync_ValuesEqual( _
    ByVal leftValue As Variant, _
    ByVal rightValue As Variant) As Boolean

    Dim leftType As VbVarType
    Dim rightType As VbVarType

    If IsError(leftValue) Or IsError(rightValue) Then
        If IsError(leftValue) And IsError(rightValue) Then
            DataSync_ValuesEqual = ( _
                DataSync_ErrorCode(leftValue) = DataSync_ErrorCode(rightValue))
        End If
        Exit Function
    End If

    If IsNull(leftValue) Or IsNull(rightValue) Then
        DataSync_ValuesEqual = (IsNull(leftValue) And IsNull(rightValue))
        Exit Function
    End If

    If DataSync_IsBlankValue(leftValue) Or DataSync_IsBlankValue(rightValue) Then
        DataSync_ValuesEqual = ( _
            DataSync_IsBlankValue(leftValue) And DataSync_IsBlankValue(rightValue))
        Exit Function
    End If

    leftType = VarType(leftValue)
    rightType = VarType(rightValue)

    If leftType <> vbString And rightType <> vbString Then
        If IsNumeric(leftValue) And IsNumeric(rightValue) Then
            DataSync_ValuesEqual = (CDbl(leftValue) = CDbl(rightValue))
            Exit Function
        End If
    End If

    If leftType <> rightType Then Exit Function

    Select Case leftType
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbDate
            DataSync_ValuesEqual = (CDbl(leftValue) = CDbl(rightValue))
        Case vbBoolean
            DataSync_ValuesEqual = (CBool(leftValue) = CBool(rightValue))
        Case Else
            DataSync_ValuesEqual = (CStr(leftValue) = CStr(rightValue))
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Reconnait les deux representations Excel equivalentes d'une cellule vide.
' EN: Recognizes Excel's two equivalent representations of an empty cell.
'------------------------------------------------------------------------------
Private Function DataSync_IsBlankValue(ByVal value As Variant) As Boolean

    If IsEmpty(value) Then
        DataSync_IsBlankValue = True
    ElseIf VarType(value) = vbString Then
        DataSync_IsBlankValue = (Len(CStr(value)) = 0)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne le code d'une erreur Excel pour une comparaison exacte en memoire.
' EN: Returns an Excel error code for exact in-memory comparison.
'------------------------------------------------------------------------------
Private Function DataSync_ErrorCode(ByVal errorValue As Variant) As Long

    On Error GoTo UnknownError
    DataSync_ErrorCode = CLng(Application.WorksheetFunction.Error_Type(errorValue))
    Exit Function

UnknownError:
    DataSync_ErrorCode = -1

End Function

'------------------------------------------------------------------------------
' FR: Applique les formats des colonnes d'entree par plages completes de table.
' EN: Applies input-column formats to complete table ranges.
'------------------------------------------------------------------------------
Private Sub DataSync_ApplyCalcInputFormats( _
    ByVal tblCalc As ListObject, _
    ByRef formatWriteCount As Long)

    Dim columnName As Variant

    For Each columnName In Array( _
        "WBS", "Task Name", "Task Type", "S", "Cal", "Predecessors WBS")
        DataSync_EnsureColumnFormat tblCalc, CStr(columnName), "@", formatWriteCount
    Next columnName

    For Each columnName In Array( _
        "Baseline Start", "Baseline Finish", "Actual Start", "Actual Finish", _
        "Forecast Start", "Forecast Finish")
        DataSync_EnsureColumnFormat tblCalc, CStr(columnName), _
            "dd/mm/yyyy", formatWriteCount
    Next columnName

    For Each columnName In Array("ID", "Baseline Duration", "Actual Duration")
        DataSync_EnsureColumnFormat tblCalc, CStr(columnName), "0", formatWriteCount
    Next columnName

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit un format de colonne uniquement lorsque la plage diverge.
' EN: Writes a column format only when the range differs.
'------------------------------------------------------------------------------
Private Sub DataSync_EnsureColumnFormat( _
    ByVal tblCalc As ListObject, _
    ByVal columnName As String, _
    ByVal expectedFormat As String, _
    ByRef formatWriteCount As Long)

    Dim targetRange As Range
    Dim currentFormat As Variant

    If Not TableHasColumn(tblCalc, columnName) Then Exit Sub
    Set targetRange = tblCalc.ListColumns(columnName).DataBodyRange
    If targetRange Is Nothing Then Exit Sub

    currentFormat = targetRange.NumberFormat
    If IsNull(currentFormat) Or CStr(currentFormat) <> expectedFormat Then
        targetRange.NumberFormat = expectedFormat
        formatWriteCount = formatWriteCount + 1
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Invalide les colonnes de sortie CALC par plages de colonnes.
' EN: Invalidates CALC output columns by column ranges.
'------------------------------------------------------------------------------
Private Sub DataSync_ClearCalcOutputColumns( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal outputColumns As Variant, _
    ByRef clearCount As Long)

    Dim columnName As Variant
    Dim targetRange As Range
    Dim clearRange As Range

    For Each columnName In outputColumns
        If mapCalc.Exists(CStr(columnName)) Then
            Set targetRange = tblCalc.ListColumns(CStr(columnName)).DataBodyRange
            If Not targetRange Is Nothing Then
                If clearRange Is Nothing Then
                    Set clearRange = targetRange
                Else
                    Set clearRange = Union(clearRange, targetRange)
                End If
            End If
        End If
    Next columnName

    If Not clearRange Is Nothing Then
        clearRange.ClearContents
        clearCount = clearCount + 1
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Efface les sorties uniquement pour les IDs absents lorsque leur conservation est demandee.
' EN: Clears outputs only for missing IDs when output preservation is requested.
'------------------------------------------------------------------------------
Private Sub DataSync_ClearMissingIdOutputs( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal outputColumns As Variant, _
    ByVal missingIdRows As Object, _
    ByRef writeCount As Long)

    Dim columnName As Variant
    Dim targetRange As Range
    Dim values As Variant
    Dim rowKey As Variant
    Dim changed As Boolean
    Dim rowCount As Long

    rowCount = tblCalc.ListRows.Count
    For Each columnName In outputColumns
        If mapCalc.Exists(CStr(columnName)) Then
            Set targetRange = tblCalc.ListColumns(CStr(columnName)).DataBodyRange
            If Not targetRange Is Nothing Then
                values = targetRange.value
                changed = False

                For Each rowKey In missingIdRows.Keys
                    If rowCount = 1 And Not IsArray(values) Then
                        If Not DataSync_ValuesEqual(values, Empty) Then
                            values = Empty
                            changed = True
                        End If
                    ElseIf Not DataSync_ValuesEqual(values(CLng(rowKey), 1), Empty) Then
                        values(CLng(rowKey), 1) = Empty
                        changed = True
                    End If
                Next rowKey

                If changed Then
                    targetRange.value = values
                    writeCount = writeCount + 1
                End If
            End If
        End If
    Next columnName

End Sub

'------------------------------------------------------------------------------
' FR: Construit l'ensemble des WBS parents depuis le buffer source.
' EN: Builds the parent-WBS set from the source buffer.
'------------------------------------------------------------------------------
Private Function DataSync_BuildSummaryWbsLookupFromArray( _
    ByVal wbsData As Variant, _
    ByVal mapWBS As Object) As Object

    Dim result As Object
    Dim r As Long
    Dim wbsValue As String
    Dim parentWbs As String

    Set result = CreateObject("Scripting.Dictionary")
    If Not IsArray(wbsData) Then
        Set DataSync_BuildSummaryWbsLookupFromArray = result
        Exit Function
    End If

    For r = 1 To UBound(wbsData, 1)
        wbsValue = Replace$(Trim$(CStr(wbsData(r, mapWBS(VTS_COL_WBS)))), ",", ".")
        parentWbs = GetParentWBS(wbsValue)

        Do While parentWbs <> ""
            result(parentWbs) = True
            parentWbs = GetParentWBS(parentWbs)
        Loop
    Next r

    Set DataSync_BuildSummaryWbsLookupFromArray = result

End Function


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
        If mapWBS.Exists(VTS_COL_ID) Then
            idVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(rowIndex, mapWBS(VTS_COL_ID)).value))
        End If
        If mapWBS.Exists(VTS_COL_WBS) Then
            wbsVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(rowIndex, mapWBS(VTS_COL_WBS)).value))
        End If
    Else
        If TableHasColumn(tblWBS, SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_ID)) Then
            idVal = Trim$(CStr(SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_ID).DataBodyRange.Cells(rowIndex, 1).value))
        End If
        If TableHasColumn(tblWBS, SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_WBS)) Then
            wbsVal = Trim$(CStr(SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_WBS).DataBodyRange.Cells(rowIndex, 1).value))
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
    wbsTaskTypeIndex = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_TASK_TYPE).Index
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

    If Not TableHasColumn(tblWBS, SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CAL)) Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CAL)
    End If

    If Not TableHasColumn(tblCalc, "Cal") Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Cal"
    End If

    If Not tblWBS.DataBodyRange Is Nothing Then
        SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_CAL).DataBodyRange.NumberFormat = "@"
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
    colIndex = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_DEADLINE_FLOAT).Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_DEADLINE_FLOAT)
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
        SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_DEADLINE_FLOAT).DataBodyRange.NumberFormat = "0"
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
    colIndex = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_LONGEST_PATH).Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_LONGEST_PATH)
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = tblCalc.ListColumns("Longest Path").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblCalc.ListColumns.Add
        newCol.Name = "Longest Path"
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_LONGEST_PATH_REX).Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblWBS.ListColumns.Add(Position:=SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_LONGEST_PATH).Index + 1)
        newCol.Name = SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_LONGEST_PATH_REX)
    End If

    colIndex = 0
    On Error Resume Next
    colIndex = tblCalc.ListColumns("Longest Path REX").Index
    On Error GoTo 0

    If colIndex <= 0 Then
        Set newCol = tblCalc.ListColumns.Add(Position:=tblCalc.ListColumns("Longest Path").Index + 1)
        newCol.Name = "Longest Path REX"
    End If

    If Not tblWBS.DataBodyRange Is Nothing Then
        SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_LONGEST_PATH).DataBodyRange.NumberFormat = "@"
        SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_LONGEST_PATH_REX).DataBodyRange.NumberFormat = "@"
    End If

    If Not tblCalc.DataBodyRange Is Nothing Then
        tblCalc.ListColumns("Longest Path").DataBodyRange.NumberFormat = "@"
        tblCalc.ListColumns("Longest Path REX").DataBodyRange.NumberFormat = "@"
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

    Set mapWBS = SchemaBuildColumnKeyMap(tblWBS, VTS_TABLE_WBS)
    Set mapCalc = CreateObject("Scripting.Dictionary")
    Set wbsIdRows = CreateObject("Scripting.Dictionary")

    For r = 1 To tblCalc.ListColumns.Count
        mapCalc(tblCalc.ListColumns(r).Name) = r
    Next r

    If Not mapWBS.Exists(VTS_COL_ID) Then
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

    If Not mapWBS.Exists(VTS_COL_FORECAST_START) Then
        DataSync_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne Forecast Start est introuvable dans tbl_WBS.", _
            "Column Forecast Start was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapWBS.Exists(VTS_COL_FORECAST_FINISH) Then
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
        idValue = tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_ID)).value
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
                    tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS(VTS_COL_FORECAST_START)).value

                tblCalc.DataBodyRange.Cells(r, mapCalc("Forecast Finish")).value = _
                    tblWBS.DataBodyRange.Cells(wbsRowIndex, mapWBS(VTS_COL_FORECAST_FINISH)).value
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
             "Longest Path REX", _
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

    Dim dateFormat As String

    On Error Resume Next

    dateFormat = Settings_GetDateNumberFormat()

    SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_BASELINE_START).DataBodyRange.NumberFormat = dateFormat
    SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_BASELINE_FINISH).DataBodyRange.NumberFormat = dateFormat
    SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_ACTUAL_START).DataBodyRange.NumberFormat = dateFormat
    SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_ACTUAL_FINISH).DataBodyRange.NumberFormat = dateFormat
    SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_FORECAST_START).DataBodyRange.NumberFormat = dateFormat
    SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_FORECAST_FINISH).DataBodyRange.NumberFormat = dateFormat
    SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_DEADLINE_FLOAT).DataBodyRange.NumberFormat = "0"
    SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_CALCULATED_START).DataBodyRange.NumberFormat = dateFormat
    SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_CALCULATED_FINISH).DataBodyRange.NumberFormat = dateFormat

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

    Set mapWBS = SchemaBuildColumnKeyMap(tblWBS, VTS_TABLE_WBS)

    Set wbsToId = CanonicalIdentity_BuildWbsToIdMap(tblWBS, mapWBS)
    arr = tblWBS.DataBodyRange.value
    rowCount = UBound(arr, 1)

    For r = 1 To rowCount
        succId = Trim$(CStr(arr(r, mapWBS(VTS_COL_ID))))
        succWBS = NormalizeWBS(CStr(arr(r, mapWBS(VTS_COL_WBS))))
        predText = Trim$(CStr(arr(r, mapWBS(VTS_COL_PREDECESSORS_WBS))))

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
    Dim tableData As Variant
    Dim outputValues() As Variant
    Dim normalizedValue As String
    Dim rowIndex As Long
    Dim rowCount As Long
    Dim idIndex As Long
    Dim wbsIndex As Long
    Dim taskTypeIndex As Long
    Dim hasIdentity As Boolean
    Dim changed As Boolean

    Set perfScope = Profiler_BeginScope("EnsureWBSTaskTypeInputSetup", "Excel Validation")

    If tblWBS Is Nothing Then Exit Sub

    If Not TableHasColumn(tblWBS, SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_TASK_TYPE)) Then
        Err.Raise vbObjectError + 2310, "EnsureWBSTaskTypeInputSetup", _
            "Missing required WBS input column: Task Type"
    End If

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    Set rng = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_TASK_TYPE).DataBodyRange
    tableData = tblWBS.DataBodyRange.value
    rowCount = tblWBS.ListRows.Count
    idIndex = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_ID).Index
    wbsIndex = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_WBS).Index
    taskTypeIndex = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_TASK_TYPE).Index
    ReDim outputValues(1 To rowCount, 1 To 1)

    'Default blank cells to Task only on rows that have an ID/WBS identity.
    For rowIndex = 1 To rowCount
        hasIdentity = ( _
            Trim$(CStr(tableData(rowIndex, idIndex))) <> "" _
            Or Replace$(Trim$(CStr(tableData(rowIndex, wbsIndex))), ",", ".") <> "")

        If Not hasIdentity Then
            outputValues(rowIndex, 1) = Empty
            If Trim$(CStr(tableData(rowIndex, taskTypeIndex))) <> "" Then changed = True
        Else
            normalizedValue = NormalizeTaskTypeValue(tableData(rowIndex, taskTypeIndex))
            outputValues(rowIndex, 1) = normalizedValue

            If CStr(tableData(rowIndex, taskTypeIndex)) <> normalizedValue Then changed = True
        End If
    Next rowIndex

    If changed Then
        rng.value = outputValues
        Profiler_RecordOperation "DataSyncWbsTaskTypeBlockWrites", 1, 0#
    End If

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
        If Not TableHasColumn(tblWBS, SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_S)) Then
            If TableHasColumn(tblWBS, SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_COMMENTS)) Then
                targetIndex = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_COMMENTS).Index
            ElseIf TableHasColumn(tblWBS, SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_TASK_TYPE)) Then
                targetIndex = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_TASK_TYPE).Index + 1
            Else
                targetIndex = tblWBS.ListColumns.Count + 1
            End If

            Set newCol = tblWBS.ListColumns.Add(Position:=targetIndex)
            newCol.Name = SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_S)
        End If

        If Not tblWBS.DataBodyRange Is Nothing Then
            With SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_S).DataBodyRange
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
' FR: Normalise WBSSummary Display Values dans un format exploitable.
' EN: Normalizes WBSSummary Display Values into a usable format.
'------------------------------------------------------------------------------
Private Sub NormalizeWBSSummaryDisplayValues( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal summaryWbsByWbs As Object)

    Dim perfScope As clsPerfScope

    Dim r As Long
    Dim tableData As Variant
    Dim outputValues() As Variant
    Dim normalizedValue As String
    Dim defaultValue As String
    Dim rowCount As Long
    Dim idIndex As Long
    Dim wbsIndex As Long
    Dim taskTypeIndex As Long
    Dim summaryIndex As Long
    Dim hasIdentity As Boolean
    Dim wbsValue As String
    Dim taskTypeValue As String
    Dim changed As Boolean

    Set perfScope = Profiler_BeginScope("NormalizeWBSSummaryDisplayValues", "Excel Cell Write")

    If tblWBS Is Nothing Then Exit Sub
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If mapWBS Is Nothing Then Exit Sub
    If Not mapWBS.Exists(VTS_COL_S) Then Exit Sub

    tableData = tblWBS.DataBodyRange.value
    rowCount = tblWBS.ListRows.Count
    idIndex = CLng(mapWBS(VTS_COL_ID))
    wbsIndex = CLng(mapWBS(VTS_COL_WBS))
    taskTypeIndex = CLng(mapWBS(VTS_COL_TASK_TYPE))
    summaryIndex = CLng(mapWBS(VTS_COL_S))
    ReDim outputValues(1 To rowCount, 1 To 1)

    For r = 1 To rowCount
        wbsValue = Replace$(Trim$(CStr(tableData(r, wbsIndex))), ",", ".")
        hasIdentity = (Trim$(CStr(tableData(r, idIndex))) <> "" Or wbsValue <> "")

        If Not hasIdentity Then
            outputValues(r, 1) = Empty
            If Trim$(CStr(tableData(r, summaryIndex))) <> "" Then changed = True
        ElseIf Trim$(CStr(tableData(r, summaryIndex))) = "" Then
            defaultValue = "N"
            If summaryWbsByWbs.Exists(wbsValue) Then
                defaultValue = "Y"
            Else
                taskTypeValue = UCase$(NormalizeTaskTypeValue(tableData(r, taskTypeIndex)))
                If taskTypeValue = "MILESTONE" Then defaultValue = "Y"
            End If

            outputValues(r, 1) = defaultValue
            changed = True
        Else
            normalizedValue = NormalizeSummaryDisplayValue(tableData(r, summaryIndex))
            outputValues(r, 1) = normalizedValue
            If CStr(tableData(r, summaryIndex)) <> normalizedValue Then changed = True
        End If
    Next r

    If changed Then
        SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_S).DataBodyRange.value = outputValues
        Profiler_RecordOperation "DataSyncWbsSummaryBlockWrites", 1, 0#
    End If

End Sub
'------------------------------------------------------------------------------
' FR: Verifie ou cree WBSCalendar Input Setup si necessaire.
' EN: Ensures or creates WBSCalendar Input Setup when needed.
'------------------------------------------------------------------------------
Private Sub EnsureWBSCalendarInputSetup(ByVal tblWBS As ListObject)

    Dim perfScope As clsPerfScope

    Dim rng As Range
    Dim newCol As ListColumn

    Set perfScope = Profiler_BeginScope("EnsureWBSCalendarInputSetup", "Excel Validation")

    If tblWBS Is Nothing Then Exit Sub

    If Not TableHasColumn(tblWBS, SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CAL)) Then
        Set newCol = tblWBS.ListColumns.Add
        newCol.Name = SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CAL)
    End If

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    Set rng = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_CAL).DataBodyRange

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
' FR: Normalise WBSCalendar Values dans un format exploitable.
' EN: Normalizes WBSCalendar Values into a usable format.
'------------------------------------------------------------------------------
Private Sub NormalizeWBSCalendarValues( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal summaryWbsByWbs As Object)

    Dim perfScope As clsPerfScope

    Dim r As Long
    Dim tableData As Variant
    Dim outputValues() As Variant
    Dim normalizedValue As String
    Dim rowCount As Long
    Dim idIndex As Long
    Dim wbsIndex As Long
    Dim taskTypeIndex As Long
    Dim calendarIndex As Long
    Dim hasIdentity As Boolean
    Dim ignored As Boolean
    Dim wbsValue As String
    Dim taskTypeValue As String
    Dim changed As Boolean

    Set perfScope = Profiler_BeginScope("NormalizeWBSCalendarValues", "Excel Cell Write")

    If tblWBS Is Nothing Then Exit Sub
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If mapWBS Is Nothing Then Exit Sub
    If Not mapWBS.Exists(VTS_COL_CAL) Then Exit Sub

    tableData = tblWBS.DataBodyRange.value
    rowCount = tblWBS.ListRows.Count
    idIndex = CLng(mapWBS(VTS_COL_ID))
    wbsIndex = CLng(mapWBS(VTS_COL_WBS))
    taskTypeIndex = CLng(mapWBS(VTS_COL_TASK_TYPE))
    calendarIndex = CLng(mapWBS(VTS_COL_CAL))
    ReDim outputValues(1 To rowCount, 1 To 1)

    For r = 1 To rowCount
        wbsValue = Replace$(Trim$(CStr(tableData(r, wbsIndex))), ",", ".")
        hasIdentity = (Trim$(CStr(tableData(r, idIndex))) <> "" Or wbsValue <> "")

        If Not hasIdentity Then
            outputValues(r, 1) = Empty
            If Trim$(CStr(tableData(r, calendarIndex))) <> "" Then changed = True
        Else
            taskTypeValue = UCase$(NormalizeTaskTypeValue(tableData(r, taskTypeIndex)))
            ignored = summaryWbsByWbs.Exists(wbsValue) _
                Or taskTypeValue = "LEVEL OF EFFORT" _
                Or taskTypeValue = "MILESTONE"

            normalizedValue = NormalizeCalendarType(tableData(r, calendarIndex))
            If Trim$(CStr(tableData(r, calendarIndex))) = "" Then
                If ignored Then
                    outputValues(r, 1) = Empty
                Else
                    outputValues(r, 1) = CALENDAR_7D
                    changed = True
                End If
            Else
                outputValues(r, 1) = normalizedValue
                If CStr(tableData(r, calendarIndex)) <> normalizedValue Then changed = True
            End If
        End If
    Next r

    If changed Then
        SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_CAL).DataBodyRange.value = outputValues
        Profiler_RecordOperation "DataSyncWbsCalendarBlockWrites", 1, 0#
    End If

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










