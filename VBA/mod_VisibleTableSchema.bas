Attribute VB_Name = "mod_VisibleTableSchema"
Option Explicit

' Central schema for user-visible Excel tables.
' Physical headers are resolved from the persisted language owner.

Public Const VTS_LANG_EN As String = "EN"
Public Const VTS_LANG_FR As String = "FR"

Public Const VTS_TABLE_WBS As String = "tbl_WBS"
Public Const VTS_TABLE_SCURVE As String = "tbl_SCURVE"
Public Const VTS_TABLE_CONSTRAINTS As String = "tbl_CONSTRAINTS"
Public Const VTS_TABLE_EVENT_HISTORY As String = "tbl_EVENT_HISTORY"
Public Const VTS_TABLE_EVENT_ACK As String = "tbl_EVENT_ACK"

Public Const VTS_OWNER_WBS As String = "WBS"
Public Const VTS_OWNER_SCURVE As String = "SCURVE"
Public Const VTS_OWNER_CONSTRAINTS As String = "CONSTRAINTS"
Public Const VTS_OWNER_EVENT_HISTORY As String = "EVENT_HISTORY"

Public Const VTS_COL_ID As String = "ID"
Public Const VTS_COL_WBS As String = "WBS"
Public Const VTS_COL_TASK_NAME As String = "TASK_NAME"
Public Const VTS_COL_TASK_DESCRIPTION As String = "TASK_DESCRIPTION"
Public Const VTS_COL_SUPPLIER As String = "SUPPLIER"
Public Const VTS_COL_DISCIPLINE As String = "DISCIPLINE"
Public Const VTS_COL_PROJECT As String = "PROJECT"
Public Const VTS_COL_TASK_TYPE As String = "TASK_TYPE"
Public Const VTS_COL_COMMENTS As String = "COMMENTS"
Public Const VTS_COL_PREDECESSORS_WBS As String = "PREDECESSORS_WBS"
Public Const VTS_COL_WEIGHT_PERCENT As String = "WEIGHT_PERCENT"
Public Const VTS_COL_PROGRESS_PERCENT As String = "PROGRESS_PERCENT"
Public Const VTS_COL_BASELINE_START As String = "BASELINE_START"
Public Const VTS_COL_BASELINE_DURATION As String = "BASELINE_DURATION"
Public Const VTS_COL_BASELINE_FINISH As String = "BASELINE_FINISH"
Public Const VTS_COL_ACTUAL_START As String = "ACTUAL_START"
Public Const VTS_COL_ACTUAL_FINISH As String = "ACTUAL_FINISH"
Public Const VTS_COL_ACTUAL_DURATION As String = "ACTUAL_DURATION"
Public Const VTS_COL_FORECAST_START As String = "FORECAST_START"
Public Const VTS_COL_FORECAST_FINISH As String = "FORECAST_FINISH"
Public Const VTS_COL_CALCULATED_START As String = "CALCULATED_START"
Public Const VTS_COL_CALCULATED_FINISH As String = "CALCULATED_FINISH"
Public Const VTS_COL_CALCULATED_DURATION As String = "CALCULATED_DURATION"
Public Const VTS_COL_START_VARIANCE As String = "START_VARIANCE"
Public Const VTS_COL_FINISH_VARIANCE As String = "FINISH_VARIANCE"
Public Const VTS_COL_DURATION_VARIANCE As String = "DURATION_VARIANCE"
Public Const VTS_COL_DRIVING_LOGIC As String = "DRIVING_LOGIC"
Public Const VTS_COL_CRITICAL_PATH As String = "CRITICAL_PATH"
Public Const VTS_COL_CRITICAL_PATH_REX As String = "CRITICAL_PATH_REX"
Public Const VTS_COL_LONGEST_PATH As String = "LONGEST_PATH"
Public Const VTS_COL_LONGEST_PATH_REX As String = "LONGEST_PATH_REX"
Public Const VTS_COL_TOTAL_FLOAT As String = "TOTAL_FLOAT"
Public Const VTS_COL_FREE_FLOAT As String = "FREE_FLOAT"
Public Const VTS_COL_TOTAL_FLOAT_REX As String = "TOTAL_FLOAT_REX"
Public Const VTS_COL_FREE_FLOAT_REX As String = "FREE_FLOAT_REX"
Public Const VTS_COL_DEADLINE_FLOAT As String = "DEADLINE_FLOAT"
Public Const VTS_COL_CAL As String = "CAL"
Public Const VTS_COL_S As String = "S"

Public Const VTS_COL_DATE As String = "DATE"
Public Const VTS_COL_DAILY_BASELINE As String = "DAILY_BASELINE"
Public Const VTS_COL_CUMULATIVE_BASELINE As String = "CUMULATIVE_BASELINE"
Public Const VTS_COL_DAILY_ACTUALIZED As String = "DAILY_ACTUALIZED"
Public Const VTS_COL_CUMULATIVE_ACTUALIZED As String = "CUMULATIVE_ACTUALIZED"
Public Const VTS_COL_DAILY_REMAINING_FORECAST As String = "DAILY_REMAINING_FORECAST"
Public Const VTS_COL_CUMULATIVE_REMAINING_FORECAST As String = "CUMULATIVE_REMAINING_FORECAST"
Public Const VTS_COL_CALCULATED_CURVE_SOLID As String = "CALCULATED_CURVE_SOLID"
Public Const VTS_COL_CALCULATED_CURVE_DASHED As String = "CALCULATED_CURVE_DASHED"
Public Const VTS_COL_CUMULATIVE_ACTUAL As String = "CUMULATIVE_ACTUAL"

Public Const VTS_COL_IS_SUMMARY As String = "IS_SUMMARY"
Public Const VTS_COL_START_CONSTRAINT_TYPE As String = "START_CONSTRAINT_TYPE"
Public Const VTS_COL_START_CONSTRAINT_DATE As String = "START_CONSTRAINT_DATE"
Public Const VTS_COL_FINISH_CONSTRAINT_TYPE As String = "FINISH_CONSTRAINT_TYPE"
Public Const VTS_COL_FINISH_CONSTRAINT_DATE As String = "FINISH_CONSTRAINT_DATE"
Public Const VTS_COL_ACTIVE As String = "ACTIVE"
Public Const VTS_COL_COMMENT As String = "COMMENT"
Public Const VTS_COL_DEADLINE As String = "DEADLINE"

Public Const VTS_COL_HOUR As String = "HOUR"
Public Const VTS_COL_SEVERITY As String = "SEVERITY"
Public Const VTS_COL_MESSAGE As String = "MESSAGE"
Public Const VTS_COL_ACKNOWLEDGED As String = "ACKNOWLEDGED"
Public Const VTS_COL_HASH As String = "HASH"
Public Const VTS_COL_EVENT_TYPE As String = "EVENT_TYPE"
Public Const VTS_COL_ACKNOWLEDGED_AT As String = "ACKNOWLEDGED_AT"
Public Const VTS_COL_ACKNOWLEDGED_BY As String = "ACKNOWLEDGED_BY"
Public Const VTS_COL_FR_MESSAGE As String = "FR_MESSAGE"
Public Const VTS_COL_EN_MESSAGE As String = "EN_MESSAGE"

Private Const VTS_ERROR_BASE As Long = vbObjectError + 5290
Private gPhysicalLanguageOverrides As Object

Public Function SchemaOwnerForTable(ByVal tableKey As String) As String
    Select Case tableKey
        Case VTS_TABLE_WBS
            SchemaOwnerForTable = VTS_OWNER_WBS
        Case VTS_TABLE_SCURVE
            SchemaOwnerForTable = VTS_OWNER_SCURVE
        Case VTS_TABLE_CONSTRAINTS
            SchemaOwnerForTable = VTS_OWNER_CONSTRAINTS
        Case VTS_TABLE_EVENT_HISTORY, VTS_TABLE_EVENT_ACK
            SchemaOwnerForTable = VTS_OWNER_EVENT_HISTORY
        Case Else
            SchemaRaiseUnknownTable "SchemaOwnerForTable", tableKey
    End Select
End Function

Public Function SchemaPhysicalHeaderLanguageForTable(ByVal tableKey As String) As String
    Dim ownerKey As String
    Dim overrideLanguage As String

    If Not gPhysicalLanguageOverrides Is Nothing Then
        If gPhysicalLanguageOverrides.Exists(tableKey) Then
            overrideLanguage = CStr(gPhysicalLanguageOverrides(tableKey))
            If overrideLanguage = VTS_LANG_FR Or overrideLanguage = VTS_LANG_EN Then
                SchemaPhysicalHeaderLanguageForTable = overrideLanguage
                Exit Function
            End If
        End If
    End If

    ownerKey = SchemaOwnerForTable(tableKey)
    Select Case ownerKey
        Case VTS_OWNER_WBS
            SchemaPhysicalHeaderLanguageForTable = Settings_GetOwnerLanguage("WBS")
        Case VTS_OWNER_SCURVE
            SchemaPhysicalHeaderLanguageForTable = Settings_GetOwnerLanguage("SCURVE")
        Case VTS_OWNER_CONSTRAINTS
            SchemaPhysicalHeaderLanguageForTable = Settings_GetOwnerLanguage("CONSTRAINTS")
        Case VTS_OWNER_EVENT_HISTORY
            SchemaPhysicalHeaderLanguageForTable = Settings_GetOwnerLanguage("EVENT")
        Case Else
            SchemaRaiseUnknownTable "SchemaPhysicalHeaderLanguageForTable", tableKey
    End Select
End Function

Public Sub Schema_SetPhysicalHeaderLanguageOverride( _
    ByVal tableKey As String, _
    ByVal languageCode As String)

    Dim normalizedLanguage As String

    Select Case UCase$(Trim$(languageCode))
        Case VTS_LANG_FR
            normalizedLanguage = VTS_LANG_FR
        Case VTS_LANG_EN
            normalizedLanguage = VTS_LANG_EN
        Case Else
            Err.Raise VTS_ERROR_BASE + 20, "Schema_SetPhysicalHeaderLanguageOverride", _
                "Unsupported transient schema language '" & languageCode & _
                "' for table '" & tableKey & "'. Expected EN or FR."
    End Select

    Call SchemaOwnerForTable(tableKey)
    If gPhysicalLanguageOverrides Is Nothing Then
        Set gPhysicalLanguageOverrides = CreateObject("Scripting.Dictionary")
        gPhysicalLanguageOverrides.CompareMode = vbBinaryCompare
    End If
    gPhysicalLanguageOverrides(tableKey) = normalizedLanguage
End Sub

Public Sub Schema_ClearPhysicalHeaderLanguageOverrides()
    Set gPhysicalLanguageOverrides = Nothing
End Sub

Public Function SchemaCurrentColumnTitle(ByVal tableKey As String, ByVal columnKey As String) As String
    SchemaCurrentColumnTitle = SchemaColumnTitle( _
        tableKey, columnKey, SchemaPhysicalHeaderLanguageForTable(tableKey))
End Function

Public Function SchemaCanonicalEnglishColumnTitle( _
    ByVal tableKey As String, _
    ByVal columnKey As String) As String

    ' Bridge from a visible-table business key to a technical table whose
    ' physical schema remains permanently English.
    SchemaCanonicalEnglishColumnTitle = _
        SchemaColumnTitle(tableKey, columnKey, VTS_LANG_EN)
End Function

Public Function SchemaColumnTitle( _
    ByVal tableKey As String, _
    ByVal columnKey As String, _
    ByVal languageCode As String) As String

    Dim useFrench As Boolean
    useFrench = SchemaUseFrench(languageCode, tableKey, columnKey)

    Select Case tableKey
        Case VTS_TABLE_WBS
            SchemaColumnTitle = SchemaWBSColumnTitle(columnKey, useFrench)
        Case VTS_TABLE_SCURVE
            SchemaColumnTitle = SchemaSCurveColumnTitle(columnKey, useFrench)
        Case VTS_TABLE_CONSTRAINTS
            SchemaColumnTitle = SchemaConstraintsColumnTitle(columnKey, useFrench)
        Case VTS_TABLE_EVENT_HISTORY
            SchemaColumnTitle = SchemaEventHistoryColumnTitle(columnKey, useFrench)
        Case VTS_TABLE_EVENT_ACK
            SchemaColumnTitle = SchemaEventAckColumnTitle(columnKey, useFrench)
        Case Else
            Err.Raise VTS_ERROR_BASE, "SchemaColumnTitle", _
                "Unknown visible-table key '" & tableKey & "' for column key '" & _
                columnKey & "' and language '" & languageCode & "'."
    End Select

    If Len(SchemaColumnTitle) = 0 Then
        Err.Raise VTS_ERROR_BASE + 2, "SchemaColumnTitle", _
            "Unknown visible-table column key '" & columnKey & _
            "' for table '" & tableKey & "' and language '" & languageCode & "'."
    End If
End Function

Public Function SchemaListColumn( _
    ByVal tableObject As ListObject, _
    ByVal tableKey As String, _
    ByVal columnKey As String) As ListColumn

    Dim expectedTitle As String

    If tableObject Is Nothing Then
        Err.Raise VTS_ERROR_BASE + 3, "SchemaListColumn", _
            "The ListObject is Nothing for table key '" & tableKey & "'."
    End If

    If StrComp(tableObject.Name, tableKey, vbBinaryCompare) <> 0 Then
        Err.Raise VTS_ERROR_BASE + 4, "SchemaListColumn", _
            "ListObject mismatch. Expected '" & tableKey & "' but received '" & _
            tableObject.Name & "'."
    End If

    expectedTitle = SchemaCurrentColumnTitle(tableKey, columnKey)

    On Error Resume Next
    Set SchemaListColumn = tableObject.ListColumns(expectedTitle)
    On Error GoTo 0

    If SchemaListColumn Is Nothing Then
        Err.Raise VTS_ERROR_BASE + 5, "SchemaListColumn", _
            "The column key '" & columnKey & "' could not be resolved in table '" & _
            tableKey & "'. Expected physical header: '" & expectedTitle & "'."
    End If
End Function

Public Function SchemaColumnIndex( _
    ByVal tableObject As ListObject, _
    ByVal tableKey As String, _
    ByVal columnKey As String) As Long

    SchemaColumnIndex = SchemaListColumn(tableObject, tableKey, columnKey).Index
End Function

Public Function SchemaIsVisibleTable(ByVal tableKey As String) As Boolean
    Select Case tableKey
        Case VTS_TABLE_WBS, _
             VTS_TABLE_SCURVE, _
             VTS_TABLE_CONSTRAINTS, _
             VTS_TABLE_EVENT_HISTORY, _
             VTS_TABLE_EVENT_ACK
            SchemaIsVisibleTable = True
    End Select
End Function

Public Function SchemaBuildColumnKeyMap( _
    ByVal tableObject As ListObject, _
    ByVal tableKey As String) As Object

    Dim result As Object
    Dim columnKeys As Variant
    Dim columnKey As Variant

    If tableObject Is Nothing Then
        Err.Raise VTS_ERROR_BASE + 6, "SchemaBuildColumnKeyMap", _
            "The ListObject is Nothing for table key '" & tableKey & "'."
    End If
    If StrComp(tableObject.Name, tableKey, vbBinaryCompare) <> 0 Then
        Err.Raise VTS_ERROR_BASE + 7, "SchemaBuildColumnKeyMap", _
            "ListObject mismatch. Expected '" & tableKey & "' but received '" & _
            tableObject.Name & "'."
    End If

    columnKeys = SchemaColumnKeys(tableKey)
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbBinaryCompare

    For Each columnKey In columnKeys
        result.Add CStr(columnKey), _
            SchemaColumnIndex(tableObject, tableKey, CStr(columnKey))
    Next columnKey

    Set SchemaBuildColumnKeyMap = result
End Function

Public Function SchemaColumnKeys(ByVal tableKey As String) As Variant
    Select Case tableKey
        Case VTS_TABLE_WBS
            SchemaColumnKeys = Array( _
                VTS_COL_ID, VTS_COL_WBS, VTS_COL_TASK_NAME, VTS_COL_TASK_DESCRIPTION, _
                VTS_COL_SUPPLIER, VTS_COL_DISCIPLINE, VTS_COL_PROJECT, VTS_COL_TASK_TYPE, _
                VTS_COL_COMMENTS, VTS_COL_PREDECESSORS_WBS, VTS_COL_WEIGHT_PERCENT, _
                VTS_COL_PROGRESS_PERCENT, VTS_COL_BASELINE_START, VTS_COL_BASELINE_DURATION, _
                VTS_COL_BASELINE_FINISH, VTS_COL_ACTUAL_START, VTS_COL_ACTUAL_FINISH, _
                VTS_COL_ACTUAL_DURATION, VTS_COL_FORECAST_START, VTS_COL_FORECAST_FINISH, _
                VTS_COL_CALCULATED_START, VTS_COL_CALCULATED_FINISH, _
                VTS_COL_CALCULATED_DURATION, VTS_COL_START_VARIANCE, VTS_COL_FINISH_VARIANCE, _
                VTS_COL_DURATION_VARIANCE, VTS_COL_DRIVING_LOGIC, VTS_COL_CRITICAL_PATH, _
                VTS_COL_CRITICAL_PATH_REX, VTS_COL_LONGEST_PATH, VTS_COL_LONGEST_PATH_REX, _
                VTS_COL_TOTAL_FLOAT, VTS_COL_FREE_FLOAT, VTS_COL_TOTAL_FLOAT_REX, _
                VTS_COL_FREE_FLOAT_REX, VTS_COL_DEADLINE_FLOAT, VTS_COL_CAL, VTS_COL_S)
        Case VTS_TABLE_SCURVE
            SchemaColumnKeys = Array( _
                VTS_COL_DATE, VTS_COL_DAILY_BASELINE, VTS_COL_CUMULATIVE_BASELINE, _
                VTS_COL_DAILY_ACTUALIZED, VTS_COL_CUMULATIVE_ACTUALIZED, _
                VTS_COL_DAILY_REMAINING_FORECAST, VTS_COL_CUMULATIVE_REMAINING_FORECAST, _
                VTS_COL_CALCULATED_CURVE_SOLID, VTS_COL_CALCULATED_CURVE_DASHED, _
                VTS_COL_CUMULATIVE_ACTUAL)
        Case VTS_TABLE_CONSTRAINTS
            SchemaColumnKeys = Array( _
                VTS_COL_ID, VTS_COL_WBS, VTS_COL_TASK_NAME, VTS_COL_TASK_DESCRIPTION, _
                VTS_COL_TASK_TYPE, VTS_COL_IS_SUMMARY, VTS_COL_CALCULATED_START, _
                VTS_COL_CALCULATED_FINISH, VTS_COL_CALCULATED_DURATION, VTS_COL_DRIVING_LOGIC, _
                VTS_COL_START_CONSTRAINT_TYPE, VTS_COL_START_CONSTRAINT_DATE, _
                VTS_COL_FINISH_CONSTRAINT_TYPE, VTS_COL_FINISH_CONSTRAINT_DATE, _
                VTS_COL_DEADLINE, VTS_COL_ACTIVE, VTS_COL_COMMENT)
        Case VTS_TABLE_EVENT_HISTORY
            SchemaColumnKeys = Array( _
                VTS_COL_DATE, VTS_COL_HOUR, VTS_COL_SEVERITY, VTS_COL_MESSAGE, _
                VTS_COL_ACKNOWLEDGED)
        Case VTS_TABLE_EVENT_ACK
            SchemaColumnKeys = Array( _
                VTS_COL_HASH, VTS_COL_WBS, VTS_COL_EVENT_TYPE, VTS_COL_SEVERITY, _
                VTS_COL_FR_MESSAGE, VTS_COL_EN_MESSAGE, VTS_COL_ACKNOWLEDGED, _
                VTS_COL_ACKNOWLEDGED_AT, VTS_COL_ACKNOWLEDGED_BY, VTS_COL_COMMENT, _
                VTS_COL_TASK_NAME)
        Case Else
            SchemaRaiseUnknownTable "SchemaColumnKeys", tableKey
    End Select
End Function

Public Function SchemaStructuredRowReference( _
    ByVal tableKey As String, _
    ByVal columnKey As String) As String

    SchemaStructuredRowReference = "[@[" & _
        SchemaCurrentColumnTitle(tableKey, columnKey) & "]]"
End Function

Private Function SchemaUseFrench( _
    ByVal languageCode As String, _
    ByVal tableKey As String, _
    ByVal columnKey As String) As Boolean
    Select Case UCase$(Trim$(languageCode))
        Case VTS_LANG_EN
            SchemaUseFrench = False
        Case VTS_LANG_FR
            SchemaUseFrench = True
        Case Else
            Err.Raise VTS_ERROR_BASE + 1, "SchemaColumnTitle", _
                "Unsupported schema language '" & languageCode & _
                "' for table '" & tableKey & "' and column key '" & columnKey & _
                "'. Expected EN or FR."
    End Select
End Function

Private Sub SchemaRaiseUnknownTable(ByVal sourceName As String, ByVal tableKey As String)
    Err.Raise VTS_ERROR_BASE, sourceName, _
        "Unknown visible-table key '" & tableKey & "'."
End Sub

Private Function SchemaWBSColumnTitle(ByVal columnKey As String, ByVal useFrench As Boolean) As String
    Select Case columnKey
        Case VTS_COL_ID: SchemaWBSColumnTitle = "ID"
        Case VTS_COL_WBS: SchemaWBSColumnTitle = "WBS"
        Case VTS_COL_TASK_NAME: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Task Name", "Nom de tâche")
        Case VTS_COL_TASK_DESCRIPTION: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Task Description", "Description")
        Case VTS_COL_SUPPLIER: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Supplier", "Fournisseur")
        Case VTS_COL_DISCIPLINE: SchemaWBSColumnTitle = "Discipline"
        Case VTS_COL_PROJECT: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Project", "Projet")
        Case VTS_COL_TASK_TYPE: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Task Type", "Type de tâche")
        Case VTS_COL_COMMENTS: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Comments", "Commentaires")
        Case VTS_COL_PREDECESSORS_WBS: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Predecessors WBS", "Antécédent WBS")
        Case VTS_COL_WEIGHT_PERCENT: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Weight (%)", "Poids [%]")
        Case VTS_COL_PROGRESS_PERCENT: SchemaWBSColumnTitle = SchemaChoose(useFrench, "% Progress", "Progrès [%]")
        Case VTS_COL_BASELINE_START: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Baseline Start", "Début référence")
        Case VTS_COL_BASELINE_DURATION: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Baseline Duration", "Durée référence")
        Case VTS_COL_BASELINE_FINISH: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Baseline Finish", "Fin référence")
        Case VTS_COL_ACTUAL_START: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Actual Start", "Début réel")
        Case VTS_COL_ACTUAL_FINISH: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Actual Finish", "Fin réelle")
        Case VTS_COL_ACTUAL_DURATION: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Actual Duration", "Durée réelle")
        Case VTS_COL_FORECAST_START: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Forecast Start", "Début prévu")
        Case VTS_COL_FORECAST_FINISH: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Forecast Finish", "Fin prévue")
        Case VTS_COL_CALCULATED_START: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Calculated Start", "Début calculé")
        Case VTS_COL_CALCULATED_FINISH: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Calculated Finish", "Fin calculée")
        Case VTS_COL_CALCULATED_DURATION: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Calculated Duration", "Durée calculée")
        Case VTS_COL_START_VARIANCE: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Start Variance", "Écart début")
        Case VTS_COL_FINISH_VARIANCE: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Finish Variance", "Écart fin")
        Case VTS_COL_DURATION_VARIANCE: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Duration Variance", "Écart durée")
        Case VTS_COL_DRIVING_LOGIC: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Driving Logic", "Logique pilotante")
        Case VTS_COL_CRITICAL_PATH: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Critical Path", "Chemin critique")
        Case VTS_COL_CRITICAL_PATH_REX: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Critical Path REX", "Chemin critique REX")
        Case VTS_COL_LONGEST_PATH: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Longest Path", "Chemin le plus long")
        Case VTS_COL_LONGEST_PATH_REX: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Longest Path REX", "Chemin le plus long REX")
        Case VTS_COL_TOTAL_FLOAT: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Total Float", "Marge totale")
        Case VTS_COL_FREE_FLOAT: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Free Float", "Marge libre")
        Case VTS_COL_TOTAL_FLOAT_REX: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Total Float REX", "Marge totale REX")
        Case VTS_COL_FREE_FLOAT_REX: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Free Float REX", "Marge libre REX")
        Case VTS_COL_DEADLINE_FLOAT: SchemaWBSColumnTitle = SchemaChoose(useFrench, "Deadline Float", "Marge deadline")
        Case VTS_COL_CAL: SchemaWBSColumnTitle = "Cal"
        Case VTS_COL_S: SchemaWBSColumnTitle = "S"
    End Select
End Function

Private Function SchemaSCurveColumnTitle(ByVal columnKey As String, ByVal useFrench As Boolean) As String
    Select Case columnKey
        Case VTS_COL_DATE: SchemaSCurveColumnTitle = "Date"
        Case VTS_COL_DAILY_BASELINE: SchemaSCurveColumnTitle = SchemaChoose(useFrench, "Daily Baseline", "Référence journalière")
        Case VTS_COL_CUMULATIVE_BASELINE: SchemaSCurveColumnTitle = SchemaChoose(useFrench, "Cumulative Baseline", "Référence cumulée")
        Case VTS_COL_DAILY_ACTUALIZED: SchemaSCurveColumnTitle = SchemaChoose(useFrench, "Daily Actualized", "Réel journalier")
        Case VTS_COL_CUMULATIVE_ACTUALIZED: SchemaSCurveColumnTitle = SchemaChoose(useFrench, "Cumulative Actualized", "Réel actualisé cumulé")
        Case VTS_COL_DAILY_REMAINING_FORECAST: SchemaSCurveColumnTitle = SchemaChoose(useFrench, "Daily Remaining Forecast", "Prévu restant journalier")
        Case VTS_COL_CUMULATIVE_REMAINING_FORECAST: SchemaSCurveColumnTitle = SchemaChoose(useFrench, "Cumulative Remaining Forecast", "Prévu restant cumulé")
        Case VTS_COL_CALCULATED_CURVE_SOLID: SchemaSCurveColumnTitle = SchemaChoose(useFrench, "Calculated Curve Solid", "Courbe calculée pleine")
        Case VTS_COL_CALCULATED_CURVE_DASHED: SchemaSCurveColumnTitle = SchemaChoose(useFrench, "Calculated Curve Dashed", "Courbe calculée pointillée")
        Case VTS_COL_CUMULATIVE_ACTUAL: SchemaSCurveColumnTitle = SchemaChoose(useFrench, "Cumulative Actual", "Réel cumulé")
    End Select
End Function

Private Function SchemaConstraintsColumnTitle(ByVal columnKey As String, ByVal useFrench As Boolean) As String
    Select Case columnKey
        Case VTS_COL_ID: SchemaConstraintsColumnTitle = "ID"
        Case VTS_COL_WBS: SchemaConstraintsColumnTitle = "WBS"
        Case VTS_COL_TASK_NAME: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Task Name", "Nom de tâche")
        Case VTS_COL_TASK_DESCRIPTION: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Task Description", "Description")
        Case VTS_COL_TASK_TYPE: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Task Type", "Type de tâche")
        Case VTS_COL_IS_SUMMARY: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Is Summary", "Est summary")
        Case VTS_COL_START_CONSTRAINT_TYPE: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Start Constraint Type", "Type contrainte début")
        Case VTS_COL_START_CONSTRAINT_DATE: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Start Constraint Date", "Date contrainte début")
        Case VTS_COL_FINISH_CONSTRAINT_TYPE: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Finish Constraint Type", "Type contrainte fin")
        Case VTS_COL_FINISH_CONSTRAINT_DATE: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Finish Constraint Date", "Date contrainte fin")
        Case VTS_COL_ACTIVE: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Active", "Actif")
        Case VTS_COL_COMMENT: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Comment", "Commentaire")
        Case VTS_COL_CALCULATED_START: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Calculated Start", "Début calculé")
        Case VTS_COL_CALCULATED_FINISH: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Calculated Finish", "Fin calculée")
        Case VTS_COL_CALCULATED_DURATION: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Calculated Duration", "Durée calculée")
        Case VTS_COL_DRIVING_LOGIC: SchemaConstraintsColumnTitle = SchemaChoose(useFrench, "Driving Logic", "Logique pilotante")
        Case VTS_COL_DEADLINE: SchemaConstraintsColumnTitle = "Deadline"
    End Select
End Function

Private Function SchemaEventHistoryColumnTitle(ByVal columnKey As String, ByVal useFrench As Boolean) As String
    Select Case columnKey
        Case VTS_COL_DATE: SchemaEventHistoryColumnTitle = "Date"
        Case VTS_COL_HOUR: SchemaEventHistoryColumnTitle = SchemaChoose(useFrench, "Hour", "Heure")
        Case VTS_COL_SEVERITY: SchemaEventHistoryColumnTitle = SchemaChoose(useFrench, "Severity", "Sévérité")
        Case VTS_COL_MESSAGE: SchemaEventHistoryColumnTitle = "Message"
        Case VTS_COL_ACKNOWLEDGED: SchemaEventHistoryColumnTitle = SchemaChoose(useFrench, "Acknowledged", "Acquitté")
    End Select
End Function

Private Function SchemaEventAckColumnTitle(ByVal columnKey As String, ByVal useFrench As Boolean) As String
    Select Case columnKey
        Case VTS_COL_HASH: SchemaEventAckColumnTitle = SchemaChoose(useFrench, "Hash", "Identifiant unique")
        Case VTS_COL_WBS: SchemaEventAckColumnTitle = "WBS"
        Case VTS_COL_EVENT_TYPE: SchemaEventAckColumnTitle = SchemaChoose(useFrench, "Event Type", "Type")
        Case VTS_COL_SEVERITY: SchemaEventAckColumnTitle = SchemaChoose(useFrench, "Severity", "Sévérité")
        Case VTS_COL_FR_MESSAGE: SchemaEventAckColumnTitle = "FR Message"
        Case VTS_COL_EN_MESSAGE: SchemaEventAckColumnTitle = "EN Message"
        Case VTS_COL_ACKNOWLEDGED: SchemaEventAckColumnTitle = SchemaChoose(useFrench, "Acknowledged", "Acquitté")
        Case VTS_COL_ACKNOWLEDGED_AT: SchemaEventAckColumnTitle = SchemaChoose(useFrench, "Acknowledged At", "Date et heure")
        Case VTS_COL_ACKNOWLEDGED_BY: SchemaEventAckColumnTitle = SchemaChoose(useFrench, "Acknowledged By", "Acquitté par")
        Case VTS_COL_COMMENT: SchemaEventAckColumnTitle = SchemaChoose(useFrench, "Comment", "Commentaire")
        Case VTS_COL_TASK_NAME: SchemaEventAckColumnTitle = SchemaChoose(useFrench, "Task Name", "Nom de tâche")
    End Select
End Function

Private Function SchemaChoose( _
    ByVal useFrench As Boolean, _
    ByVal englishTitle As String, _
    ByVal frenchTitle As String) As String

    If useFrench Then
        SchemaChoose = frenchTitle
    Else
        SchemaChoose = englishTitle
    End If
End Function
