Attribute VB_Name = "mod_CalcEngineCoreBridge"
Option Explicit

'===============================================================================
' MODULE : mod_CalcEngineCoreBridge
' DOMAINE / DOMAIN : Core Bridge Orchestration
'
' FR
' Orchestre la synchronisation, le Core, les analytics et les writers du workflow planning.
' Delegue le calcul, les diagnostics et les ecritures a leurs services proprietaires.
'
' EN
' Orchestrates synchronization, Core, analytics and writers for the planning workflow.
' Delegates calculation, diagnostics and writes to their owning services.
'
' CONTRATS / CONTRACTS : Run_Calc_Engine_CoreBridge
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

'------------------------------------------------------------------------------
' FR: Lance le workflow Calc Engine Core Bridge.
' EN: Runs the Calc Engine Core Bridge workflow.
'------------------------------------------------------------------------------
Public Sub Run_Calc_Engine_CoreBridge(Optional ByVal forceFullRecalcOverride As Boolean = False)

    Dim perfScope As clsPerfScope


    Const PARTIAL_RECALC_MAX_IMPACT_RATIO As Double = 0.3

    Dim wsCalc As Worksheet
    Dim wsWBS As Worksheet
    Dim tblCalc As ListObject
    Dim tblWBS As ListObject

    Dim mapCalc As Object
    Dim mapWBS As Object
    Dim dataArr As Variant
    Dim linksBySuccId As Object
    Dim executionNetwork As clsCompiledExecutionNetwork

    Dim changedIds As Object
    Dim forceFullRecalc As Boolean

    Dim succByPred As Object
    Dim parentById As Object
    Dim impactedIds As Object
    Dim usePartialWrite As Boolean

    Dim totalRows As Long
    Dim impactedCount As Long
    Dim impactRatio As Double

    Dim runAnalytics As Boolean
    Dim consoleMessages As Collection

    Set perfScope = Profiler_BeginScope("Run_Calc_Engine_CoreBridge", "Planning")

    On Error GoTo ErrHandler

    AppEvents_EnsureInitialized
    BeginMacroRun "Run_Calc_Engine_CoreBridge"

    Set consoleMessages = New Collection

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")

    If Planning_WBSIsEmpty() Then
        Planning_CalcSafeEmptyState
        Application.StatusBar = "No project data - calculation outputs cleared."
        GoTo SafeExit
    End If
    Ensure_Calc_Infrastructure consoleMessages
    Ensure_Analytics_Toggle
    AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterEnsureCalc"

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    'Preserve existing calculated outputs during sync.
    'Full mode still clears/rebuilds inside Run_Calc_Core.
    'Partial mode needs previous calculated dates available as predecessor state.
    Sync_WBS_To_CALC True
    AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterSync"

    Import_WBS_To_Constraints
    AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterConstraintImport"

    If Not Sync_Constraints_To_CALC_ForWorkflow(consoleMessages) Then
        Write_CalcState_Snapshot_Console "ERROR", consoleMessages
        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo SafeExit
    End If
    AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterConstraintSync"

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    If Not ValidateCalcAfterSync(tblCalc) Then
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
            BiMsg( _
                "Le sync WBS -> CALC a échoué ou a laissé tbl_CALC dans un état invalide." & vbCrLf & _
                "Le calcul est arrêté pour éviter un faux succès.", _
                "WBS -> CALC sync failed or left tbl_CALC in an invalid state." & vbCrLf & _
                "Calculation stopped to avoid a false success.")
        Write_CalcState_Snapshot_Console "ERROR", consoleMessages
        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo SafeExit
    End If

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)

    FillCalcParentAndSummaryFromWBS tblCalc, tblWBS
    CalcBridge_ShowParentDateWarnings tblCalc, mapCalc, consoleMessages

    Set changedIds = Get_Changed_TaskIds(forceFullRecalc)

    If forceFullRecalcOverride Then
        forceFullRecalc = True
        Set changedIds = CreateObject("Scripting.Dictionary")
        Debug.Print "FULL RECALC FORCED BY USER"
    ElseIf forceFullRecalc Then
        Debug.Print "FULL RECALC FORCED"
    Else
        Debug.Print "Changed tasks count: " & changedIds.Count
    End If

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    dataArr = tblCalc.DataBodyRange.value

    If Not forceFullRecalc Then
        If Not changedIds Is Nothing Then
            If changedIds.Count = 0 Then

                Debug.Print "NO CHANGE DETECTED - CORE SKIPPED"

                CalcBridge_AppendTaskTypeWarnings consoleMessages, tblWBS, mapWBS, tblCalc, mapCalc
                Write_CalcState_Snapshot_Console "OK", consoleMessages

                CalcBridge_AddInfoMessage consoleMessages, _
                    "Aucune modification détectée : calcul moteur non relancé.", _
                    "No change detected: core calculation was not rerun."

                If Not Profiler_ShouldSuppressUserInterface() Then
                    CalcBridge_ShowPlanningConsole consoleMessages
                End If
                GoTo SafeExit

            End If
        End If
    End If

    ApplyCalcDateFormats tblCalc

    RebuildLogicLinksTable
    If IsMacroAbortRequested() Then GoTo SafeExit

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    dataArr = tblCalc.DataBodyRange.value

    Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded(tblCalc)
    Set executionNetwork = CompileExecutionNetwork(dataArr, mapCalc, linksBySuccId)

    If CalcBridge_PreCore_CheckLOEAsPredecessor(tblCalc, mapCalc, consoleMessages, linksBySuccId) Then
        Write_CalcState_Snapshot_Console "ERROR", consoleMessages
        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo SafeExit
    End If

    Set succByPred = Build_Successor_Map(linksBySuccId)
    Set parentById = BuildParentByIdMap_FromCalc(dataArr, mapCalc, Core_BuildRowById(dataArr, mapCalc))
    Set impactedIds = Build_Impacted_TaskIds(changedIds, succByPred, parentById)

    Debug.Print "Impacted tasks count: " & impactedIds.Count

    usePartialWrite = False

    If Not forceFullRecalc Then
        If Not changedIds Is Nothing Then
            If changedIds.Count > 0 Then
                If Not impactedIds Is Nothing Then
                    If impactedIds.Count > 0 Then

                        totalRows = 0
                        impactedCount = impactedIds.Count
                        impactRatio = 1#

                        If Not tblCalc.DataBodyRange Is Nothing Then
                            totalRows = tblCalc.ListRows.Count
                        End If

                        If totalRows > 0 Then
                            impactRatio = CDbl(impactedCount) / CDbl(totalRows)
                        End If

                        Debug.Print "Partial impact ratio: " & Format$(impactRatio, "0.00%")

                        If impactRatio <= PARTIAL_RECALC_MAX_IMPACT_RATIO Then
                            usePartialWrite = True
                        Else
                            usePartialWrite = False
                            Debug.Print "PARTIAL DISABLED - IMPACT RATIO ABOVE " & Format$(PARTIAL_RECALC_MAX_IMPACT_RATIO, "0%")
                        End If

                    End If
                End If
            End If
        End If
    End If

    If usePartialWrite Then

        Debug.Print "PARTIAL CORE MODE ENABLED"
        Debug.Print "Partial impacted tasks count: " & impactedIds.Count

        Run_Calc_Core dataArr, mapCalc, linksBySuccId, impactedIds, , , , executionNetwork

        WriteCoreOutputsToCalc_Partial tblCalc, mapCalc, dataArr, impactedIds
        WriteCoreDrivingLogicToCalc_Partial tblCalc, mapCalc, dataArr, impactedIds

    Else

        Debug.Print "FULL CORE MODE"

        Run_Calc_Core dataArr, mapCalc, linksBySuccId, , , , , executionNetwork

        WriteCoreOutputsToCalc tblCalc, mapCalc, dataArr
        WriteCoreDrivingLogicToCalc tblCalc, mapCalc, dataArr

    End If

    ApplyCalcDateFormats tblCalc
    AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterCoreWrite"

    If CalcBridge_HasCoreErrors(tblCalc) Then
        CalcBridge_AppendCoreErrorMessages consoleMessages, tblCalc
        Write_CalcState_Snapshot_Console "ERROR", consoleMessages
        CalcBridge_ShowPlanningConsole consoleMessages
        GoTo SafeExit
    End If

    'Non-blocking warnings only after a clean core calculation.
    'Warnings are collected into the planning console and shown once at the end.
    runAnalytics = IsAnalyticsEnabled()

    CalcBridge_AppendTaskTypeWarnings consoleMessages, tblWBS, mapWBS, tblCalc, mapCalc

    'Analytics are controlled by the WBS user toggle and stored in VBA memory.
    'When enabled: analytics are recomputed globally using existing analytics engine.
    'When disabled: analytics outputs are cleared to avoid stale CP / float values.
    'No partial analytics calculation is implemented at this stage.

    If runAnalytics Then
        Run_Analytics_Only consoleMessages, executionNetwork, linksBySuccId
        AbortIfRequested "Run_Calc_Engine_CoreBridge.AfterAnalytics"
    Else
        Clear_Analytics_Outputs
    End If

    If usePartialWrite Then
        'Partial WBS push for core outputs only (dates / driving logic).
        'Analytics are handled separately by toggle logic above.
        Push_Calculated_Back_To_WBS_Partial impactedIds
    Else
        Push_Calculated_Back_To_WBS
    End If

    'Variances are recomputed globally.
    'This is deliberate for consistency because they depend on baseline references
    'and summary roll-ups.
    Compute_And_Push_Variances consoleMessages

    Write_CalcState_Snapshot_Console "OK", consoleMessages

    CalcBridge_AddInfoMessage consoleMessages, _
        "Calcul terminé avec succès.", _
        "Calculation completed successfully."

    CalcBridge_ShowPlanningConsole consoleMessages

SafeExit:
    If IsMacroAbortRequested() Then
        ShowAbortMessageOnce
        EndMacroRun
        Exit Sub
    End If

    EndMacroRun
    Exit Sub

ErrHandler:
    If consoleMessages Is Nothing Then Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg( _
            "Erreur dans Run_Calc_Engine_CoreBridge" & vbCrLf & _
            "-> " & Err.Description, _
            "Error in Run_Calc_Engine_CoreBridge" & vbCrLf & _
            "-> " & Err.Description)

    On Error Resume Next
    Write_CalcState_Snapshot_Console "ERROR", consoleMessages
    CalcBridge_ShowPlanningConsole consoleMessages
    On Error GoTo 0

    Resume SafeExit

End Sub

'------------------------------------------------------------------------------
' FR: Construit une synthese bornee des erreurs Core presentes dans CALC. En cas d'echec de lecture, retourne un texte fallback explicite au lieu de masquer l'erreur.
' EN: Builds a bounded summary of Core errors stored in CALC. If the read fails, returns an explicit fallback message instead of hiding the error.
'------------------------------------------------------------------------------

Private Function CalcBridge_GetCoreErrorSummary( _
    ByVal tblCalc As ListObject, _
    Optional ByVal maxRows As Long = 12) As String

    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long
    Dim countShown As Long
    Dim msg As String

    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String
    Dim errMsg As String

    On Error GoTo FailSafe

    CalcBridge_GetCoreErrorSummary = ""

    If tblCalc Is Nothing Then Exit Function
    If tblCalc.DataBodyRange Is Nothing Then Exit Function

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)

    If Not mapCalc.Exists("Error flag") Then Exit Function
    If Not mapCalc.Exists("ErrorMsg") Then Exit Function
    If Not mapCalc.Exists("ID") Then Exit Function

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        If UCase$(Trim$(CStr(arr(r, mapCalc("Error flag"))))) = "ERROR" Then

            idVal = Trim$(CStr(arr(r, mapCalc("ID"))))
            errMsg = Trim$(CStr(arr(r, mapCalc("ErrorMsg"))))

            If mapCalc.Exists("WBS") Then
                wbsVal = Trim$(CStr(arr(r, mapCalc("WBS"))))
            Else
                wbsVal = ""
            End If

            countShown = countShown + 1

            If msg <> "" Then msg = msg & vbCrLf

            msg = msg & "- ID " & idVal
            If wbsVal <> "" Then msg = msg & " / WBS " & wbsVal
            If errMsg <> "" Then msg = msg & " : " & errMsg

            If countShown >= maxRows Then Exit For

        End If

    Next r

    If msg = "" Then
        msg = "Error flag detected, but no detailed message was available."
    End If

    CalcBridge_GetCoreErrorSummary = msg
    Exit Function

FailSafe:
    CalcBridge_GetCoreErrorSummary = "Unable to build core error summary."

End Function

'------------------------------------------------------------------------------
' FR: Detecte les lignes CALC marquees ERROR avant tout push WBS. La verification est fail-closed : toute erreur de lecture bloque la suite du workflow.
' EN: Detects CALC rows marked ERROR before any WBS push. The check is fail-closed: any read failure blocks the remaining workflow.
'------------------------------------------------------------------------------

Private Function CalcBridge_HasCoreErrors(ByVal tblCalc As ListObject) As Boolean

    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long

    On Error GoTo FailSafe

    CalcBridge_HasCoreErrors = False

    If tblCalc Is Nothing Then Exit Function
    If tblCalc.DataBodyRange Is Nothing Then Exit Function

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)

    If Not mapCalc.Exists("Error flag") Then Exit Function

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, mapCalc("Error flag"))))) = "ERROR" Then
            CalcBridge_HasCoreErrors = True
            Exit Function
        End If
    Next r

    Exit Function

FailSafe:
    CalcBridge_HasCoreErrors = True

End Function
