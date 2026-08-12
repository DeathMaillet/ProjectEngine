Attribute VB_Name = "mod_Constraints"
Option Explicit

'===============================================================================
' MODULE : mod_Constraints
' DOMAINE / DOMAIN : Constraints
'
' FR
' Importe les contraintes WBS, valide leurs regles et synchronise leurs colonnes vers CALC.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Imports WBS constraints, validates their rules and synchronizes their columns to CALC.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Handle_Constraints_Change, Import_WBS_To_Constraints, Sync_Constraints_To_CALC, Sync_Constraints_To_CALC_ForWorkflow, Constraints_DeactivateAll
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Handle_Constraints_Change
'===============================================================================




Private Const CONSTRAINTS_SHEET_NAME As String = "CONSTRAINTS"
Private Const CONSTRAINTS_TABLE_NAME As String = "tbl_CONSTRAINTS"
Private Const CONSTRAINTS_TOP_LEFT As String = "A2"

Private Const WBS_SHEET_NAME As String = "WBS"
Private Const WBS_TABLE_NAME As String = "tbl_WBS"

'------------------------------------------------------------------------------
' FR: Traite un changement ou evenement pour Constraints Change.
' EN: Handles a change or event for Constraints Change.
'------------------------------------------------------------------------------
Public Sub Handle_Constraints_Change(ByVal ws As Worksheet, ByVal Target As Range)

    Dim tbl As ListObject
    Dim rngDeadline As Range
    Dim rngToCheck As Range
    Dim cell As Range
    Dim cellValue As String

    On Error GoTo SafeExit

    Set tbl = ws.ListObjects(CONSTRAINTS_TABLE_NAME)
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Set rngDeadline = tbl.ListColumns("Deadline").DataBodyRange
    Set rngToCheck = Intersect(Target, rngDeadline)
    If rngToCheck Is Nothing Then Exit Sub

    For Each cell In rngToCheck
        cellValue = Trim$(CStr(cell.value))
        If cellValue <> "" Then
            If Not IsDate(cell.value) Then
                Application.EnableEvents = False
                Application.Undo

                CalcBridge_ShowSingleConsoleMessage "STOP", _
                    "Deadline invalide." & vbCrLf & _
                    "-> saisir une date valide ou laisser vide.", _
                    "Invalid deadline." & vbCrLf & _
                    "-> enter a valid date or leave blank."
                GoTo SafeExit
            End If
        End If
    Next cell

SafeExit:
    Application.EnableEvents = True

End Sub


'------------------------------------------------------------------------------
' FR: Traite la collection Import WBS To Constraints sans modifier les donnees d'entree.
' EN: Handles the Import WBS To Constraints collection without mutating input data.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Public Sub Import_WBS_To_Constraints()

    Dim perfScope As clsPerfScope

    Dim wsWBS As Worksheet
    Dim wsConstraints As Worksheet
    Dim tblWBS As ListObject
    Dim tblConstraints As ListObject

    Dim mapWBS As Object
    Dim mapConstraints As Object
    Dim existingById As Object
    Dim summaryByWbs As Object

    Dim arrWBS As Variant
    Dim outArr() As Variant
    Dim rowCountWBS As Long
    Dim targetRows As Long
    Dim outRow As Long
    Dim r As Long
    Dim consoleMessages As Collection

    Set perfScope = Profiler_BeginScope("Import_WBS_To_Constraints", "Excel Table Sync")

    On Error GoTo ErrHandler

    Set consoleMessages = New Collection

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET_NAME)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE_NAME)

    Set wsConstraints = EnsureConstraintsWorksheet()
    Set tblConstraints = EnsureConstraintsTable(wsConstraints)

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    Set mapConstraints = CanonicalIdentity_BuildColumnMap(tblConstraints)

    RequireColumns_Constraints mapWBS, Array( _
        "ID", "WBS", "Task Name", "Task Description", "Task Type", _
        "Calculated Start", "Calculated Finish", "Calculated Duration", "Driving Logic"), "tbl_WBS"

    RequireColumns_Constraints mapConstraints, ConstraintsHeaders(), CONSTRAINTS_TABLE_NAME

    Set existingById = BuildExistingConstraintRows(tblConstraints, mapConstraints)

    If tblWBS.DataBodyRange Is Nothing Then
        rowCountWBS = 0
        Set summaryByWbs = CreateObject("Scripting.Dictionary")
    Else
        arrWBS = tblWBS.DataBodyRange.value
        rowCountWBS = UBound(arrWBS, 1)
        Set summaryByWbs = BuildSummaryWbsMap(arrWBS, mapWBS)
    End If

    If rowCountWBS > 0 Then
        targetRows = CountConstraintSourceRows(arrWBS, mapWBS)
    Else
        targetRows = 0
    End If
    ResizeTableToRowCount_Constraints tblConstraints, targetRows

    If targetRows = 0 Then
        ApplyConstraintsFormats tblConstraints
        GoTo SafeExit
    End If

    ReDim outArr(1 To targetRows, 1 To tblConstraints.ListColumns.Count)

    outRow = 0

    If rowCountWBS > 0 Then
        For r = 1 To rowCountWBS
            If ConstraintSourceRowHasIdentity(arrWBS, r, mapWBS) Then
                outRow = outRow + 1
                FillConstraintOutputRowFromWBS outArr, outRow, arrWBS, r, mapWBS, mapConstraints, existingById, summaryByWbs
            End If
        Next r
    End If

    tblConstraints.DataBodyRange.value = outArr
    ApplyConstraintsFormats tblConstraints

SafeExit:
    If Not consoleMessages Is Nothing Then
        If consoleMessages.Count > 0 Then CalcBridge_ShowPlanningConsole consoleMessages
    End If
    Exit Sub

ErrHandler:
    If consoleMessages Is Nothing Then Set consoleMessages = New Collection
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg( _
            "Erreur dans Import_WBS_To_Constraints" & vbCrLf & _
            "-> " & Err.Description, _
            "Error in Import_WBS_To_Constraints" & vbCrLf & _
            "-> " & Err.Description)
    Resume SafeExit

End Sub

'------------------------------------------------------------------------------
' FR: Synchronise Constraints To CALC entre les tables du planning.
' EN: Synchronizes Constraints To CALC between planning tables.
'------------------------------------------------------------------------------
Public Sub Sync_Constraints_To_CALC()

    BeginPlanningEventRun "Sync_Constraints_To_CALC"
    Sync_Constraints_To_CALC_Impl True

End Sub

'------------------------------------------------------------------------------
' FR: Synchronise Constraints To CALC For Workflow entre les tables du planning.
' EN: Synchronizes Constraints To CALC For Workflow between planning tables.
'------------------------------------------------------------------------------
Public Function Sync_Constraints_To_CALC_ForWorkflow( _
    ByVal consoleMessages As Collection) As Boolean

    Sync_Constraints_To_CALC_ForWorkflow = _
        Sync_Constraints_To_CALC_Impl(False, consoleMessages)

End Function


'------------------------------------------------------------------------------
' FR: Synchronise Constraints To CALC Impl entre les tables du planning.
' EN: Synchronizes Constraints To CALC Impl between planning tables.
'------------------------------------------------------------------------------
Private Function Sync_Constraints_To_CALC_Impl( _
    ByVal showConsole As Boolean, _
    Optional ByVal externalConsoleMessages As Collection) As Boolean

    Dim perfScope As clsPerfScope

    Dim wsCalc As Worksheet
    Dim wsConstraints As Worksheet
    Dim tblCalc As ListObject
    Dim tblConstraints As ListObject

    Dim mapCalc As Object
    Dim mapConstraints As Object
    Dim calcRowById As Object

    Dim arrCalc As Variant
    Dim arrConstraints As Variant

    Dim r As Long
    Dim calcRow As Long
    Dim idVal As String
    Dim activeVal As String
    Dim consoleMessages As Collection

    Set perfScope = Profiler_BeginScope("Sync_Constraints_To_CALC_Impl", "Excel Table Sync")

    On Error GoTo ErrHandler

    If externalConsoleMessages Is Nothing Then
        Set consoleMessages = New Collection
    Else
        Set consoleMessages = externalConsoleMessages
    End If

    Ensure_Calc_Infrastructure consoleMessages

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    On Error Resume Next
    Set wsConstraints = ThisWorkbook.Worksheets(CONSTRAINTS_SHEET_NAME)
    If Not wsConstraints Is Nothing Then
        Set tblConstraints = wsConstraints.ListObjects(CONSTRAINTS_TABLE_NAME)
    End If
    On Error GoTo ErrHandler

    If tblConstraints Is Nothing Then
        Err.Raise vbObjectError + 8610, "Sync_Constraints_To_CALC", _
            "Missing table " & CONSTRAINTS_TABLE_NAME & ". Run Import_WBS_To_Constraints first."
    End If

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set mapConstraints = CanonicalIdentity_BuildColumnMap(tblConstraints)

    RequireColumns_Constraints mapCalc, CalcConstraintHeaders(), "tbl_CALC"
    RequireColumns_Constraints mapConstraints, Array( _
        "ID", _
        "WBS", _
        "Task Name", _
        "Task Type", _
        "Is Summary", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Deadline", _
        "Active"), CONSTRAINTS_TABLE_NAME

    If tblCalc.DataBodyRange Is Nothing Then
        If Not tblConstraints.DataBodyRange Is Nothing Then
            If HasActiveConstraints(tblConstraints, mapConstraints) Then
                Err.Raise vbObjectError + 8611, "Sync_Constraints_To_CALC", _
                    "Active constraints exist but tbl_CALC is empty."
            End If
        End If
        GoTo SafeExit
    End If

    arrCalc = tblCalc.DataBodyRange.value
    Set calcRowById = CanonicalIdentity_BuildCalcRowById(arrCalc, mapCalc)

    ClearCalcConstraintColumns arrCalc, mapCalc

    If Not tblConstraints.DataBodyRange Is Nothing Then
        arrConstraints = tblConstraints.DataBodyRange.value

        ValidateActiveConstraints arrConstraints, mapConstraints, calcRowById, consoleMessages

        For r = 1 To UBound(arrConstraints, 1)
            idVal = Trim$(CStr(arrConstraints(r, mapConstraints("ID"))))
            activeVal = NormalizeActiveValue(arrConstraints(r, mapConstraints("Active")))

            If HasValue(arrConstraints(r, mapConstraints("Deadline"))) And _
                Not HasConstraintDate(arrConstraints(r, mapConstraints("Deadline"))) Then
                Err.Raise vbObjectError + 8614, "Sync_Constraints_To_CALC", _
                    "Invalid deadline in tbl_CONSTRAINTS for ID: " & idVal
            End If

            If idVal <> "" Then
                If calcRowById.Exists(idVal) Then
                    If HasConstraintDate(arrConstraints(r, mapConstraints("Deadline"))) Then
                        calcRow = CLng(calcRowById(idVal))
                        arrCalc(calcRow, mapCalc("Deadline")) = arrConstraints(r, mapConstraints("Deadline"))
                    End If
                End If
            End If

            If activeVal = "YES" Then
                If idVal = "" Then
                    Err.Raise vbObjectError + 8612, "Sync_Constraints_To_CALC", _
                        "Active constraint row has an empty ID."
                End If

                If IsConstraintSummaryRow(arrConstraints(r, mapConstraints("Is Summary"))) Then GoTo NextConstraintRow
                If IsConstraintLevelOfEffort(arrConstraints(r, mapConstraints("Task Type"))) Then GoTo NextConstraintRow

                If Not calcRowById.Exists(idVal) Then
                    Err.Raise vbObjectError + 8613, "Sync_Constraints_To_CALC", _
                        "Active constraint references ID not found in tbl_CALC: " & idVal
                End If

                If IsActiveConstraintEmpty(arrConstraints, r, mapConstraints) Then GoTo NextConstraintRow

                calcRow = CLng(calcRowById(idVal))

                arrCalc(calcRow, mapCalc("Constraint Active")) = "Yes"
                arrCalc(calcRow, mapCalc("Start Constraint Type")) = _
                    arrConstraints(r, mapConstraints("Start Constraint Type"))
                arrCalc(calcRow, mapCalc("Start Constraint Date")) = _
                    arrConstraints(r, mapConstraints("Start Constraint Date"))
                arrCalc(calcRow, mapCalc("Finish Constraint Type")) = _
                    arrConstraints(r, mapConstraints("Finish Constraint Type"))
                arrCalc(calcRow, mapCalc("Finish Constraint Date")) = _
                    arrConstraints(r, mapConstraints("Finish Constraint Date"))
            End If

NextConstraintRow:
        Next r
    End If

    WriteCalcConstraintColumns tblCalc, mapCalc, arrCalc
    ApplyCalcConstraintFormats tblCalc

SafeExit:
    Sync_Constraints_To_CALC_Impl = (Err.Number = 0)

    If showConsole And Not consoleMessages Is Nothing Then
        If consoleMessages.Count > 0 Then CalcBridge_ShowPlanningConsole consoleMessages
    End If
    Exit Function

ErrHandler:
    If consoleMessages Is Nothing Then Set consoleMessages = New Collection

    If Err.Source = "ValidateActiveConstraints" Then
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", Err.Description
    Else
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
            BiMsg( _
                "Erreur dans Sync_Constraints_To_CALC" & vbCrLf & _
                "-> " & Err.Description, _
                "Error in Sync_Constraints_To_CALC" & vbCrLf & _
                "-> " & Err.Description)
    End If

    Resume SafeExit

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree Constraints Worksheet si necessaire.
' EN: Ensures or creates Constraints Worksheet when needed.
'------------------------------------------------------------------------------
Private Function EnsureConstraintsWorksheet() As Worksheet

    Dim ws As Worksheet
    Dim wsStart As Worksheet

    Set wsStart = ActiveSheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONSTRAINTS_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONSTRAINTS_SHEET_NAME
        If Not wsStart Is Nothing Then wsStart.Activate
    End If

    Set EnsureConstraintsWorksheet = ws

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree Constraints Table si necessaire.
' EN: Ensures or creates Constraints Table when needed.
'------------------------------------------------------------------------------
Private Function EnsureConstraintsTable(ByVal ws As Worksheet) As ListObject

    Dim tbl As ListObject
    Dim headers As Variant
    Dim headerCount As Long
    Dim rng As Range
    Dim i As Long
    Dim lc As ListColumn
    Dim headerName As String

    headers = ConstraintsHeaders()
    headerCount = UBound(headers) - LBound(headers) + 1

    On Error Resume Next
    Set tbl = ws.ListObjects(CONSTRAINTS_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then
        Set rng = ws.Range(CONSTRAINTS_TOP_LEFT).Resize(2, headerCount)
        rng.Clear

        For i = 0 To headerCount - 1
            rng.Cells(1, i + 1).value = headers(LBound(headers) + i)
        Next i

        Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = CONSTRAINTS_TABLE_NAME

        If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.ClearContents
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

    ApplyConstraintsFormats tbl
    Set EnsureConstraintsTable = tbl

End Function
'------------------------------------------------------------------------------
' FR: Retourne la valeur Constraints Headers sans modifier les donnees d'entree.
' EN: Returns the Constraints Headers value without mutating input data.
'------------------------------------------------------------------------------

Private Function ConstraintsHeaders() As Variant

    ConstraintsHeaders = Array( _
        "ID", _
        "WBS", _
        "Task Name", _
        "Task Description", _
        "Task Type", _
        "Is Summary", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration", _
        "Driving Logic", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Deadline", _
        "Active", _
        "Comment")

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Calc Constraint Headers sans modifier les donnees d'entree.
' EN: Returns the Calc Constraint Headers value without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcConstraintHeaders() As Variant

    CalcConstraintHeaders = Array( _
        "ID", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Deadline")

End Function
'------------------------------------------------------------------------------
' FR: Valide Require Columns Constraints et applique la politique d'erreur definie par le composant.
' EN: Validates Require Columns Constraints and applies the component's defined failure policy.
'------------------------------------------------------------------------------

Private Sub RequireColumns_Constraints( _
    ByVal mapCol As Object, _
    ByVal requiredCols As Variant, _
    ByVal tableName As String)

    Dim c As Variant

    For Each c In requiredCols
        If Not mapCol.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 8601, "RequireColumns_Constraints", _
                "Missing required column in " & tableName & ": " & CStr(c)
        End If
    Next c

End Sub
'------------------------------------------------------------------------------
' FR: Vide ou reinitialise Calc Constraint Columns.
' EN: Clears or resets Calc Constraint Columns.
'------------------------------------------------------------------------------
Private Sub ClearCalcConstraintColumns( _
    ByRef arrCalc As Variant, _
    ByVal mapCalc As Object)

    Dim r As Long

    For r = 1 To UBound(arrCalc, 1)
        arrCalc(r, mapCalc("Constraint Active")) = "No"
        arrCalc(r, mapCalc("Start Constraint Type")) = Empty
        arrCalc(r, mapCalc("Start Constraint Date")) = Empty
        arrCalc(r, mapCalc("Finish Constraint Type")) = Empty
        arrCalc(r, mapCalc("Finish Constraint Date")) = Empty
        arrCalc(r, mapCalc("Deadline")) = Empty
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Valide Active Constraints et signale les incoherences detectees.
' EN: Validates Active Constraints and reports detected inconsistencies.
'------------------------------------------------------------------------------
Private Sub ValidateActiveConstraints( _
    ByRef arrConstraints As Variant, _
    ByVal mapConstraints As Object, _
    ByVal calcRowById As Object, _
    ByVal consoleMessages As Collection)

    Dim r As Long
    Dim idVal As String
    Dim startType As String
    Dim finishType As String
    Dim eventHashVal As String

    For r = 1 To UBound(arrConstraints, 1)
        If NormalizeActiveValue(arrConstraints(r, mapConstraints("Active"))) <> "YES" Then GoTo NextRow

        idVal = Trim$(CStr(arrConstraints(r, mapConstraints("ID"))))
        startType = Trim$(CStr(arrConstraints(r, mapConstraints("Start Constraint Type"))))
        finishType = Trim$(CStr(arrConstraints(r, mapConstraints("Finish Constraint Type"))))

        If idVal = "" Then
            Err.Raise vbObjectError + 8620, "ValidateActiveConstraints", _
                "Active constraint row has an empty ID."
        End If

        If IsConstraintSummaryRow(arrConstraints(r, mapConstraints("Is Summary"))) Then
            eventHashVal = BuildPlanningEventHash( _
                "WARNING", _
                "CONSTRAINT_PARENT_IGNORED", _
                "Contrainte active ignoree sur une tache parent", _
                "Active constraint ignored on a summary task", _
                "les contraintes sur taches parent ne sont pas exportees vers CALC", _
                "constraints on summary tasks are not exported to CALC", _
                "ValidateActiveConstraints", _
                "CONSTRAINTS", _
                "tbl_CONSTRAINTS", _
                Trim$(CStr(arrConstraints(r, mapConstraints("ID")))), _
                Trim$(CStr(arrConstraints(r, mapConstraints("WBS")))), _
                Trim$(CStr(arrConstraints(r, mapConstraints("Task Name")))))

            LogPlanningEvent _
                "WARNING", _
                "CONSTRAINT_PARENT_IGNORED", _
                eventHashVal, _
                "Contrainte active ignoree sur une tache parent", _
                "Active constraint ignored on a summary task", _
                "les contraintes sur taches parent ne sont pas exportees vers CALC", _
                "constraints on summary tasks are not exported to CALC", _
                "ValidateActiveConstraints", _
                "CONSTRAINTS", _
                "tbl_CONSTRAINTS", _
                Trim$(CStr(arrConstraints(r, mapConstraints("ID")))), _
                Trim$(CStr(arrConstraints(r, mapConstraints("WBS")))), _
                Trim$(CStr(arrConstraints(r, mapConstraints("Task Name"))))

            AddConstraintWarning consoleMessages, _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Contrainte active ignoree sur une tache parent", _
                    "Active constraint ignored on a summary task", _
                    "les contraintes sur taches parent ne sont pas exportees vers CALC", _
                    "constraints on summary tasks are not exported to CALC"), _
                True, _
                "CONSTRAINT_PARENT_IGNORED", _
                eventHashVal
            GoTo NextRow
        End If

        If IsConstraintLevelOfEffort(arrConstraints(r, mapConstraints("Task Type"))) Then
            eventHashVal = BuildPlanningEventHash( _
                "WARNING", _
                "CONSTRAINT_LOE_IGNORED", _
                "Contrainte active ignoree sur une tache Level of Effort", _
                "Active constraint ignored on a Level of Effort task", _
                "les contraintes sur LOE ne sont pas exportees vers CALC", _
                "constraints on LOE tasks are not exported to CALC", _
                "ValidateActiveConstraints", _
                "CONSTRAINTS", _
                "tbl_CONSTRAINTS", _
                Trim$(CStr(arrConstraints(r, mapConstraints("ID")))), _
                Trim$(CStr(arrConstraints(r, mapConstraints("WBS")))), _
                Trim$(CStr(arrConstraints(r, mapConstraints("Task Name")))))

            LogPlanningEvent _
                "WARNING", _
                "CONSTRAINT_LOE_IGNORED", _
                eventHashVal, _
                "Contrainte active ignoree sur une tache Level of Effort", _
                "Active constraint ignored on a Level of Effort task", _
                "les contraintes sur LOE ne sont pas exportees vers CALC", _
                "constraints on LOE tasks are not exported to CALC", _
                "ValidateActiveConstraints", _
                "CONSTRAINTS", _
                "tbl_CONSTRAINTS", _
                Trim$(CStr(arrConstraints(r, mapConstraints("ID")))), _
                Trim$(CStr(arrConstraints(r, mapConstraints("WBS")))), _
                Trim$(CStr(arrConstraints(r, mapConstraints("Task Name"))))

            AddConstraintWarning consoleMessages, _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Contrainte active ignoree sur une tache Level of Effort", _
                    "Active constraint ignored on a Level of Effort task", _
                    "les contraintes sur LOE ne sont pas exportees vers CALC", _
                    "constraints on LOE tasks are not exported to CALC"), _
                True, _
                "CONSTRAINT_LOE_IGNORED", _
                eventHashVal
            GoTo NextRow
        End If

        If HasValue(arrConstraints(r, mapConstraints("Deadline"))) And _
            Not HasConstraintDate(arrConstraints(r, mapConstraints("Deadline"))) Then
            Err.Raise vbObjectError + 8630, "ValidateActiveConstraints", _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Deadline renseignee invalide", _
                    "Invalid deadline")
        End If
        If Not calcRowById.Exists(idVal) Then
            Err.Raise vbObjectError + 8621, "ValidateActiveConstraints", _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Contrainte active sur un ID absent de CALC", _
                    "Active constraint references an ID not found in CALC")
        End If

        If startType <> "" And _
            startType <> "Start No Earlier Than" And _
            startType <> "Start No Later Than" And _
            startType <> "Must Start On" Then
            Err.Raise vbObjectError + 8624, "ValidateActiveConstraints", _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Type de contrainte debut non reconnu", _
                    "Unknown start constraint type")
        End If

        If finishType <> "" And _
            finishType <> "Finish No Earlier Than" And _
            finishType <> "Finish No Later Than" And _
            finishType <> "Must Finish On" Then
            Err.Raise vbObjectError + 8625, "ValidateActiveConstraints", _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Type de contrainte fin non reconnu", _
                    "Unknown finish constraint type")
        End If

        If startType <> "" And Not HasConstraintDate(arrConstraints(r, mapConstraints("Start Constraint Date"))) Then
            Err.Raise vbObjectError + 8626, "ValidateActiveConstraints", _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Type de contrainte debut renseigne sans date", _
                    "Start constraint type defined without constraint date")
        End If

        If HasConstraintDate(arrConstraints(r, mapConstraints("Start Constraint Date"))) And startType = "" Then
            Err.Raise vbObjectError + 8627, "ValidateActiveConstraints", _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Date de contrainte debut renseignee sans type", _
                    "Start constraint date defined without constraint type")
        End If

        If finishType <> "" And Not HasConstraintDate(arrConstraints(r, mapConstraints("Finish Constraint Date"))) Then
            Err.Raise vbObjectError + 8628, "ValidateActiveConstraints", _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Type de contrainte fin renseigne sans date", _
                    "Finish constraint type defined without constraint date")
        End If

        If HasConstraintDate(arrConstraints(r, mapConstraints("Finish Constraint Date"))) And finishType = "" Then
            Err.Raise vbObjectError + 8629, "ValidateActiveConstraints", _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Date de contrainte fin renseignee sans type", _
                    "Finish constraint date defined without constraint type")
        End If

        If IsActiveConstraintEmpty(arrConstraints, r, mapConstraints) Then
            AddConstraintWarning consoleMessages, _
                BuildConstraintValidationMessage( _
                    arrConstraints, r, mapConstraints, _
                    "Contrainte active vide ignoree", _
                    "Active empty constraint ignored", _
                    "aucune contrainte debut/fin n'est definie ; la ligne n'est pas exportee vers CALC", _
                    "no start/finish constraint is defined ; the row is not exported to CALC")
        End If

NextRow:
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Indique si Active Constraint Empty est vrai pour le contexte courant.
' EN: Returns whether Active Constraint Empty is true for the current context.
'------------------------------------------------------------------------------
Private Function IsActiveConstraintEmpty( _
    ByRef arrConstraints As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapConstraints As Object) As Boolean

    IsActiveConstraintEmpty = _
        Trim$(CStr(arrConstraints(rowIdx, mapConstraints("Start Constraint Type")))) = "" And _
        Not HasConstraintDate(arrConstraints(rowIdx, mapConstraints("Start Constraint Date"))) And _
        Trim$(CStr(arrConstraints(rowIdx, mapConstraints("Finish Constraint Type")))) = "" And _
        Not HasConstraintDate(arrConstraints(rowIdx, mapConstraints("Finish Constraint Date")))

End Function

'------------------------------------------------------------------------------
' FR: Indique si Constraint Date est vrai pour le contexte courant.
' EN: Returns whether Constraint Date is true for the current context.
'------------------------------------------------------------------------------
Private Function HasConstraintDate(ByVal value As Variant) As Boolean

    If IsEmpty(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If Trim$(CStr(value)) = "" Then Exit Function

    If IsDate(value) Or IsNumeric(value) Then
        HasConstraintDate = True
    End If

End Function

'------------------------------------------------------------------------------
' FR: Indique si Constraint Summary Row est vrai pour le contexte courant.
' EN: Returns whether Constraint Summary Row is true for the current context.
'------------------------------------------------------------------------------
Private Function IsConstraintSummaryRow(ByVal value As Variant) As Boolean

    Dim txt As String

    If IsEmpty(value) Or IsNull(value) Then Exit Function

    txt = LCase$(Trim$(CStr(value)))
    IsConstraintSummaryRow = (txt = "true" Or txt = "yes" Or txt = "1" Or txt = "summary")

End Function

'------------------------------------------------------------------------------
' FR: Indique si Constraint Level Of Effort est vrai pour le contexte courant.
' EN: Returns whether Constraint Level Of Effort is true for the current context.
'------------------------------------------------------------------------------
Private Function IsConstraintLevelOfEffort(ByVal value As Variant) As Boolean

    IsConstraintLevelOfEffort = (LCase$(Trim$(CStr(value))) = "level of effort")

End Function

'------------------------------------------------------------------------------
' FR: Indique si Active Constraints est vrai pour le contexte courant.
' EN: Returns whether Active Constraints is true for the current context.
'------------------------------------------------------------------------------
Private Function HasActiveConstraints( _
    ByVal tblConstraints As ListObject, _
    ByVal mapConstraints As Object) As Boolean

    Dim arrConstraints As Variant
    Dim r As Long

    If tblConstraints.DataBodyRange Is Nothing Then Exit Function

    arrConstraints = tblConstraints.DataBodyRange.value

    For r = 1 To UBound(arrConstraints, 1)
        If NormalizeActiveValue(arrConstraints(r, mapConstraints("Active"))) = "YES" Then
            If IsActiveConstraintEmpty(arrConstraints, r, mapConstraints) Then GoTo NextRow

            HasActiveConstraints = True
            Exit Function
        End If

NextRow:
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Normalise Active Value dans un format exploitable.
' EN: Normalizes Active Value into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizeActiveValue(ByVal rawValue As Variant) As String

    Dim s As String

    s = UCase$(Trim$(CStr(rawValue)))

    Select Case s
        Case "YES", "Y", "TRUE", "1", "OUI"
            NormalizeActiveValue = "YES"
        Case Else
            NormalizeActiveValue = "NO"
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Actualise Apply Calc Constraint Formats sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Calc Constraint Formats without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyCalcConstraintFormats(ByVal tblCalc As ListObject)

    On Error Resume Next

    tblCalc.ListColumns("Constraint Active").DataBodyRange.NumberFormat = "@"
    tblCalc.ListColumns("Start Constraint Type").DataBodyRange.NumberFormat = "@"
    tblCalc.ListColumns("Start Constraint Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Finish Constraint Type").DataBodyRange.NumberFormat = "@"
    tblCalc.ListColumns("Finish Constraint Date").DataBodyRange.NumberFormat = "dd/mm/yyyy"

        tblCalc.ListColumns("Deadline").DataBodyRange.NumberFormat = "dd/mm/yyyy"

    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit Calc Constraint Columns vers le stockage cible.
' EN: Writes Calc Constraint Columns to the target storage.
'------------------------------------------------------------------------------
Private Sub WriteCalcConstraintColumns( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByRef arrCalc As Variant)

    tblCalc.ListColumns("Constraint Active").DataBodyRange.value = _
        GetConstraintColumnArray(arrCalc, mapCalc("Constraint Active"))
    tblCalc.ListColumns("Start Constraint Type").DataBodyRange.value = _
        GetConstraintColumnArray(arrCalc, mapCalc("Start Constraint Type"))
    tblCalc.ListColumns("Start Constraint Date").DataBodyRange.value = _
        GetConstraintColumnArray(arrCalc, mapCalc("Start Constraint Date"))
    tblCalc.ListColumns("Finish Constraint Type").DataBodyRange.value = _
        GetConstraintColumnArray(arrCalc, mapCalc("Finish Constraint Type"))
    tblCalc.ListColumns("Finish Constraint Date").DataBodyRange.value = _
        GetConstraintColumnArray(arrCalc, mapCalc("Finish Constraint Date"))

    tblCalc.ListColumns("Deadline").DataBodyRange.value = _
        GetConstraintColumnArray(arrCalc, mapCalc("Deadline"))

End Sub

'------------------------------------------------------------------------------
' FR: Retourne Constraint Column Array depuis le contexte constraints.
' EN: Returns Constraint Column Array from the constraints context.
'------------------------------------------------------------------------------
Private Function GetConstraintColumnArray( _
    ByRef sourceArr As Variant, _
    ByVal colIndex As Long) As Variant

    Dim result() As Variant
    Dim r As Long

    ReDim result(1 To UBound(sourceArr, 1), 1 To 1)

    For r = 1 To UBound(sourceArr, 1)
        result(r, 1) = sourceArr(r, colIndex)
    Next r

    GetConstraintColumnArray = result

End Function


'------------------------------------------------------------------------------
' FR: Construit la map Existing Constraint Rows a partir des donnees fournies par l'appelant.
' EN: Builds the Existing Constraint Rows map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildExistingConstraintRows( _
    ByVal tbl As ListObject, _
    ByVal mapConstraints As Object) As Object

    Dim perfScope As clsPerfScope

    Dim d As Object
    Dim rowData As Object
    Dim arr As Variant
    Dim r As Long
    Dim c As Variant
    Dim idVal As String
    Dim headers As Variant

    Set perfScope = Profiler_BeginScope("BuildExistingConstraintRows", "Dictionary")

    Set d = CreateObject("Scripting.Dictionary")
    headers = ConstraintsHeaders()

    If tbl.DataBodyRange Is Nothing Then
        Set BuildExistingConstraintRows = d
        Exit Function
    End If

    arr = tbl.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapConstraints("ID"))))

        If idVal <> "" Then
            If Not d.Exists(idVal) Then
                Set rowData = CreateObject("Scripting.Dictionary")

                For Each c In headers
                    If mapConstraints.Exists(CStr(c)) Then
                        rowData(CStr(c)) = arr(r, mapConstraints(CStr(c)))
                    Else
                        rowData(CStr(c)) = Empty
                    End If
                Next c

                d.Add idVal, rowData
            End If
        End If
    Next r

    Set BuildExistingConstraintRows = d

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Count Constraint Source Rows sans exposer de mutateur sur l'etat source.
' EN: Returns the Count Constraint Source Rows map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function CountConstraintSourceRows( _
    ByRef arrWBS As Variant, _
    ByVal mapWBS As Object) As Long

    Dim r As Long

    For r = 1 To UBound(arrWBS, 1)
        If ConstraintSourceRowHasIdentity(arrWBS, r, mapWBS) Then
            CountConstraintSourceRows = CountConstraintSourceRows + 1
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR: Retourne la map Constraint Source Row Has IDentity sans modifier les donnees d'entree.
' EN: Returns the Constraint Source Row Has IDentity map without mutating input data.
'------------------------------------------------------------------------------

Private Function ConstraintSourceRowHasIdentity( _
    ByRef arrWBS As Variant, _
    ByVal rowIndex As Long, _
    ByVal mapWBS As Object) As Boolean

    Dim idVal As String
    Dim wbsVal As String

    If mapWBS Is Nothing Then Exit Function

    If mapWBS.Exists("ID") Then idVal = Trim$(CStr(arrWBS(rowIndex, mapWBS("ID"))))
    If mapWBS.Exists("WBS") Then wbsVal = NormalizeWBS(CStr(arrWBS(rowIndex, mapWBS("WBS"))))

    ConstraintSourceRowHasIdentity = (idVal <> "" Or wbsVal <> "")

End Function
'------------------------------------------------------------------------------
' FR: Construit la map Summary WBS Map a partir des donnees fournies par l'appelant.
' EN: Builds the Summary WBS Map map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function BuildSummaryWbsMap( _
    ByRef arrWBS As Variant, _
    ByVal mapWBS As Object) As Object

    Dim d As Object
    Dim r As Long
    Dim wbsVal As String
    Dim parentWbs As String

    Set d = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arrWBS, 1)
        wbsVal = NormalizeWBS(CStr(arrWBS(r, mapWBS("WBS"))))
        parentWbs = GetParentWBS(wbsVal)

        If parentWbs <> "" Then d(parentWbs) = True
    Next r

    Set BuildSummaryWbsMap = d

End Function

'------------------------------------------------------------------------------
' FR: Alimente la map Constraint Output Row From WBS dans la structure cible fournie par l'appelant.
' EN: Populates the Constraint Output Row From WBS map in the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub FillConstraintOutputRowFromWBS( _
    ByRef outArr() As Variant, _
    ByVal outRow As Long, _
    ByRef arrWBS As Variant, _
    ByVal wbsRow As Long, _
    ByVal mapWBS As Object, _
    ByVal mapConstraints As Object, _
    ByVal existingById As Object, _
    ByVal summaryByWbs As Object)

    Dim idVal As String
    Dim wbsVal As String
    Dim existingRow As Object

    idVal = Trim$(CStr(arrWBS(wbsRow, mapWBS("ID"))))
    wbsVal = NormalizeWBS(CStr(arrWBS(wbsRow, mapWBS("WBS"))))

    If existingById.Exists(idVal) Then
        Set existingRow = existingById(idVal)
    Else
        Set existingRow = Nothing
    End If

    outArr(outRow, mapConstraints("ID")) = idVal
    outArr(outRow, mapConstraints("WBS")) = wbsVal
    outArr(outRow, mapConstraints("Task Name")) = arrWBS(wbsRow, mapWBS("Task Name"))
    outArr(outRow, mapConstraints("Task Description")) = arrWBS(wbsRow, mapWBS("Task Description"))
    outArr(outRow, mapConstraints("Task Type")) = arrWBS(wbsRow, mapWBS("Task Type"))
    outArr(outRow, mapConstraints("Is Summary")) = IIf(summaryByWbs.Exists(wbsVal), "TRUE", "FALSE")
    outArr(outRow, mapConstraints("Calculated Start")) = arrWBS(wbsRow, mapWBS("Calculated Start"))
    outArr(outRow, mapConstraints("Calculated Finish")) = arrWBS(wbsRow, mapWBS("Calculated Finish"))
    outArr(outRow, mapConstraints("Calculated Duration")) = arrWBS(wbsRow, mapWBS("Calculated Duration"))
    outArr(outRow, mapConstraints("Driving Logic")) = arrWBS(wbsRow, mapWBS("Driving Logic"))

    CopyPreservedConstraintFields outArr, outRow, mapConstraints, existingRow

End Sub

'------------------------------------------------------------------------------
' FR: Alimente la map Constraint Output Row From Existing dans la structure cible fournie par l'appelant.
' EN: Populates the Constraint Output Row From Existing map in the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub FillConstraintOutputRowFromExisting( _
    ByRef outArr() As Variant, _
    ByVal outRow As Long, _
    ByVal mapConstraints As Object, _
    ByVal existingRow As Object)

    Dim c As Variant

    For Each c In ConstraintsHeaders()
        If mapConstraints.Exists(CStr(c)) Then
            If Not existingRow Is Nothing Then
                If existingRow.Exists(CStr(c)) Then
                    If Not IsPlanningMirrorField(CStr(c)) Then
                        outArr(outRow, mapConstraints(CStr(c))) = existingRow(CStr(c))
                    End If
                End If
            End If
        End If
    Next c

End Sub

'------------------------------------------------------------------------------
' FR: Indique si Planning Mirror Field est vrai pour le contexte courant.
' EN: Returns whether Planning Mirror Field is true for the current context.
'------------------------------------------------------------------------------
Private Function IsPlanningMirrorField(ByVal fieldName As String) As Boolean

    Select Case fieldName
        Case "Calculated Start", _
             "Calculated Finish", _
             "Calculated Duration", _
             "Driving Logic"
            IsPlanningMirrorField = True
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Traite la map Copy Preserved Constraint Fields sans modifier les donnees d'entree.
' EN: Handles the Copy Preserved Constraint Fields map without mutating input data.
'------------------------------------------------------------------------------

Private Sub CopyPreservedConstraintFields( _
    ByRef outArr() As Variant, _
    ByVal outRow As Long, _
    ByVal mapConstraints As Object, _
    ByVal existingRow As Object)

    Dim fields As Variant
    Dim f As Variant

    fields = Array( _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date", _
        "Deadline", _
        "Active", _
        "Comment")

    For Each f In fields
        If mapConstraints.Exists(CStr(f)) Then
            If existingRow Is Nothing Then
                If CStr(f) = "Active" Then
                    outArr(outRow, mapConstraints(CStr(f))) = "No"
                Else
                    outArr(outRow, mapConstraints(CStr(f))) = Empty
                End If
            ElseIf existingRow.Exists(CStr(f)) Then
                outArr(outRow, mapConstraints(CStr(f))) = existingRow(CStr(f))
            End If
        End If
    Next f

End Sub


'------------------------------------------------------------------------------
' FR: Traite la reference Resize Table To Row Count Constraints sans modifier les donnees d'entree.
' EN: Handles the Resize Table To Row Count Constraints reference without mutating input data.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ResizeTableToRowCount_Constraints( _
    ByVal tbl As ListObject, _
    ByVal targetRows As Long)

    Dim perfScope As clsPerfScope

    Dim currentRows As Long
    Dim r As Long

    Set perfScope = Profiler_BeginScope("ResizeTableToRowCount_Constraints", "Excel Table Resize")

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
' FR: Actualise Apply Constraints Formats sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Constraints Formats without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub ApplyConstraintsFormats(ByVal tbl As ListObject)

    On Error Resume Next

    tbl.ListColumns("ID").Range.NumberFormat = "0"
    tbl.ListColumns("WBS").Range.NumberFormat = "@"
    tbl.ListColumns("Task Name").Range.NumberFormat = "@"
    tbl.ListColumns("Task Description").Range.NumberFormat = "@"
    tbl.ListColumns("Task Type").Range.NumberFormat = "@"
    tbl.ListColumns("Is Summary").Range.NumberFormat = "@"
    tbl.ListColumns("Calculated Start").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Calculated Finish").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Calculated Duration").Range.NumberFormat = "0"
    tbl.ListColumns("Driving Logic").Range.NumberFormat = "@"

    tbl.ListColumns("Start Constraint Type").Range.NumberFormat = "@"
    tbl.ListColumns("Start Constraint Date").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Finish Constraint Type").Range.NumberFormat = "@"
    tbl.ListColumns("Finish Constraint Date").Range.NumberFormat = "dd/mm/yyyy"
    tbl.ListColumns("Deadline").Range.NumberFormat = "dd/mm/yyyy"
tbl.ListColumns("Active").Range.NumberFormat = "@"
    tbl.ListColumns("Comment").Range.NumberFormat = "@"

    ApplyConstraintsValidation tbl
    tbl.Range.Columns.AutoFit

    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply Constraints Validation sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Constraints Validation without changing the business rules that produce the data.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub ApplyConstraintsValidation(ByVal tbl As ListObject)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    With tbl.ListColumns("Start Constraint Type").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Operator:=xlBetween, _
             Formula1:="Start No Earlier Than,Start No Later Than,Must Start On"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Start Constraint Type"
        .InputMessage = "Choose blank, Start No Earlier Than, Start No Later Than, or Must Start On."
        .ErrorTitle = "Unknown Start Constraint Type"
        .errorMessage = "Use Start No Earlier Than, Start No Later Than, or Must Start On."
        .ShowInput = True
        .ShowError = True
    End With

    With tbl.ListColumns("Finish Constraint Type").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Operator:=xlBetween, _
             Formula1:="Finish No Earlier Than,Finish No Later Than,Must Finish On"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Finish Constraint Type"
        .InputMessage = "Choose blank, Finish No Earlier Than, Finish No Later Than, or Must Finish On."
        .ErrorTitle = "Unknown Finish Constraint Type"
        .errorMessage = "Use Finish No Earlier Than, Finish No Later Than, or Must Finish On."
        .ShowInput = True
        .ShowError = True
    End With

    With tbl.ListColumns("Active").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Operator:=xlBetween, _
             Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Active"
        .InputMessage = "Choose Yes or No."
        .ErrorTitle = "Unknown Active value"
        .errorMessage = "Recommended values: Yes or No."
        .ShowInput = True
        .ShowError = True
    End With

End Sub

'------------------------------------------------------------------------------
' FR: Desactive toutes les contraintes existantes sans creer la feuille ou la table.
' EN: Deactivates all existing constraints without creating the sheet or table.
'------------------------------------------------------------------------------
Public Sub Constraints_DeactivateAll()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim activeCol As ListColumn
    Dim col As ListColumn
    Dim normalizedName As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONSTRAINTS_SHEET_NAME)
    If Not ws Is Nothing Then Set tbl = ws.ListObjects(CONSTRAINTS_TABLE_NAME)
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For Each col In tbl.ListColumns
        normalizedName = UCase$(Trim$(CStr(col.Name)))
        Select Case normalizedName
            Case "CONSTRAINT ACTIVE", "ACTIVE", "IS ACTIVE"
                Set activeCol = col
                Exit For
        End Select
    Next col

    If activeCol Is Nothing Then Exit Sub
    If activeCol.DataBodyRange Is Nothing Then Exit Sub

    activeCol.DataBodyRange.value = "No"

End Sub
