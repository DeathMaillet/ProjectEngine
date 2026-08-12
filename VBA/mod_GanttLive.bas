Attribute VB_Name = "mod_GanttLive"
Option Explicit

'===============================================================================
' MODULE : mod_GanttLive
' DOMAINE / DOMAIN : Gantt
'
' FR
' Conserve la facade historique des workflows TEST, SCENARIO et LOCK et leurs APIs renderer/drag.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Retains the historical TEST, SCENARIO and LOCK facade and its renderer/drag APIs.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : GanttLive_SafeEmptyState, Run_Gantt_Test_Engine, GanttLive_RunTestTransaction, GanttLive_BuildTestByIdMap, GanttLive_BuildBaseByIdMap, GanttLive_HasRenderableTestDelta, GanttLive_GetDisplayStart, GanttLive_GetDisplayFinish
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"


Private Const GANTT_FIRST_TASK_ROW As Long = 4
Private Const GANTT_COL_WBS As Long = 1

Private Const COL_TEST_START As Long = 5
Private Const COL_TEST_FINISH As Long = 6
Private Const COL_TEST_PROGRESS As Long = 9
'------------------------------------------------------------------------------
' FR: Ramene le live Gantt dans un etat neutre en vidant les modes et resultats de simulation.
' EN: Returns live Gantt to a neutral state by clearing modes and simulation results.
'------------------------------------------------------------------------------
Public Sub GanttLive_SafeEmptyState()

    GanttLive_ClearTestRenderRequest
    GanttLive_ClearActiveSimulationMode
    ClearCalcGanttTestResults

End Sub
'------------------------------------------------------------------------------
' FR:
' Facade publique stable du bouton/workflow TEST; delegue l'execution au service
' Gantt TEST sans changer le contrat appele par Excel, Drag ou Lock.
'
' EN:
' Stable public facade for the TEST button/workflow; delegates execution to the
' Gantt TEST service without changing the contract used by Excel, Drag, or Lock.
'
' Entrees / Inputs:
' - Colonnes TEST du GANTT et options transactionnelles existantes.
'
' Sorties / Outputs:
' - Memes sorties transactionnelles que le moteur TEST historique.
'
' Appele par / Called by:
' - Boutons/workflows GANTT TEST.
' - GanttLive_RunTestTransaction.
' - Run_Gantt_Lock_Changes avant de verrouiller les modifications.
'
' Notes:
' - Wrapper public permanent: conserver le nom pour OnAction et appels existants.
'------------------------------------------------------------------------------
Public Sub Run_Gantt_Test_Engine( _
    Optional ByVal silentMode As Boolean = False, _
    Optional ByRef transactionSucceeded As Variant, _
    Optional ByRef transactionMessages As Variant, _
    Optional ByRef transactionGanttRebuilt As Variant, _
    Optional ByVal recordSilentMessages As Boolean = True)

    GanttTestService_RunTestEngine _
        silentMode, _
        transactionSucceeded, _
        transactionMessages, _
        transactionGanttRebuilt, _
        recordSilentMessages

    GanttUiControls_EnsureCanonical ThisWorkbook.Worksheets(GANTT_SHEET)

End Sub
'------------------------------------------------------------------------------
' FR:
' Execute le moteur TEST en mode silencieux pour les workflows qui ont besoin d'un
' resultat transactionnel plutot que d'un affichage console direct.
'
' EN:
' Runs the TEST engine silently for workflows that need a transactional result
' instead of direct console display.
'
' Entrees / Inputs:
' - References de sortie consoleMessages et ganttRebuilt.
'
' Sorties / Outputs:
' - Booleen succes/echec, collection de messages et indicateur de reconstruction GANTT.
'
' Appele par / Called by:
' - Drag/Lock/orchestrateurs pouvant enchainer simulation et decision metier.
'
' Notes:
' - Wrapper public stable autour de Run_Gantt_Test_Engine; a conserver si le moteur bouge.
'------------------------------------------------------------------------------
Public Function GanttLive_RunTestTransaction( _
    ByRef consoleMessages As Collection, _
    ByRef ganttRebuilt As Boolean) As Boolean

    Dim transactionSucceeded As Variant
    Dim transactionMessages As Variant
    Dim transactionGanttRebuilt As Variant

    Set consoleMessages = New Collection
    ganttRebuilt = False

    Run_Gantt_Test_Engine _
        True, _
        transactionSucceeded, _
        transactionMessages, _
        transactionGanttRebuilt, _
        False

    If IsObject(transactionMessages) Then
        Set consoleMessages = transactionMessages
    End If

    If Not IsEmpty(transactionGanttRebuilt) Then
        ganttRebuilt = CBool(transactionGanttRebuilt)
    End If

    If Not IsEmpty(transactionSucceeded) Then
        GanttLive_RunTestTransaction = CBool(transactionSucceeded)
    End If

End Function
'=====================================================
' PUBLIC HELPERS FOR mod_Gantt RENDERER
'=====================================================

'------------------------------------------------------------------------------
' FR:
' Expose au renderer GANTT les resultats TEST/SCENARIO calcules par ID depuis
' tbl_CALC_GANTT_TEST.
'
' EN:
' Exposes TEST/SCENARIO calculated results by ID from tbl_CALC_GANTT_TEST
' to the GANTT renderer.
'
' Entrees / Inputs:
' - tbl_CALC_GANTT_TEST avec les colonnes Calc Test et flags.
'
' Sorties / Outputs:
' - Dictionnaire ID -> start, finish, duration, progress, summary flag, error flag, any-test flag.
'
' Appele par / Called by:
' - mod_Gantt renderer et helpers de comparaison live.
'
' Notes:
' - Lecture Excel uniquement; retourne une map vide si la table n'existe pas ou est vide.
'------------------------------------------------------------------------------
Public Function GanttLive_BuildTestByIdMap() As Object

    Dim renderMode As String

    renderMode = GanttLive_GetPendingRenderMode()
    If renderMode = "" Then renderMode = GanttLive_GetActiveSimulationMode()

    If renderMode = "TEST" And GanttTestAnalyticsSnapshot_IsCurrent() Then
        Set GanttLive_BuildTestByIdMap = GanttTestAnalyticsSnapshot_GetProjectionById()
    Else
        Set GanttLive_BuildTestByIdMap = GanttSimulation_BuildResultByIdMap()
    End If

End Function

' FR: Retourne la classification analytique du meme snapshot que les dates TEST.
' EN: Returns the analytics classification from the same snapshot as TEST dates.
' FR: Indique si le rendu courant doit lire les analytics du snapshot TEST.
' EN: Returns whether the current render must read analytics from the TEST snapshot.
Public Function GanttLive_IsTestAnalyticsProjectionActive() As Boolean

    Dim renderMode As String

    renderMode = GanttLive_GetPendingRenderMode()
    If renderMode = "" Then renderMode = GanttLive_GetActiveSimulationMode()
    GanttLive_IsTestAnalyticsProjectionActive = _
        (renderMode = "TEST" And GanttTestAnalyticsSnapshot_IsCurrent())

End Function

Public Function GanttLive_GetDisplayAnalyticsPath( _
    ByVal idVal As String, _
    ByVal testById As Object, _
    ByVal isTestMode As Boolean, _
    ByVal pathColumnName As String) As String

    Dim renderMode As String

    If Not isTestMode Then Exit Function
    renderMode = GanttLive_GetPendingRenderMode()
    If renderMode = "" Then renderMode = GanttLive_GetActiveSimulationMode()
    If renderMode <> "TEST" Then Exit Function
    If testById Is Nothing Then Exit Function
    If Not testById.Exists(idVal) Then Exit Function
    If UBound(testById(idVal)) < 10 Then Exit Function

    Select Case pathColumnName
        Case "Critical Path"
            GanttLive_GetDisplayAnalyticsPath = CStr(testById(idVal)(9))
        Case "Longest Path"
            GanttLive_GetDisplayAnalyticsPath = CStr(testById(idVal)(10))
    End Select

End Function

'------------------------------------------------------------------------------
' FR:
' Expose au renderer GANTT l'etat de reference WBS par ID pour comparer le rendu
' standard et le rendu simule.
'
' EN:
' Exposes the WBS baseline state by ID to the GANTT renderer so standard and
' simulated rendering can be compared.
'
' Entrees / Inputs:
' - tbl_WBS avec dates calculees, duree, progress et calendrier.
'
' Sorties / Outputs:
' - Dictionnaire ID -> start, finish, duration, progress, calendar.
'
' Appele par / Called by:
' - mod_Gantt renderer, delta TEST, lock simulation.
'
' Notes:
' - Lecture Excel uniquement; ne modifie aucune table.
'------------------------------------------------------------------------------
Public Function GanttLive_BuildBaseByIdMap() As Object

    Dim perfScope As clsPerfScope

    Dim d As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mapWBS As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set perfScope = Profiler_BeginScope("GanttLive_BuildBaseByIdMap", "Excel Read")

    Set d = CreateObject("Scripting.Dictionary")

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tbl = ws.ListObjects(WBS_TABLE)

    If tbl.DataBodyRange Is Nothing Then
        Set GanttLive_BuildBaseByIdMap = d
        Exit Function
    End If

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tbl)
    arr = tbl.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapWBS("ID"))))

        If idVal <> "" Then
            d(idVal) = Array( _
                GetCellValue(arr(r, mapWBS("Calculated Start"))), _
                GetCellValue(arr(r, mapWBS("Calculated Finish"))), _
                GetCellValue(arr(r, mapWBS("Calculated Duration"))), _
                GetCellValue(arr(r, mapWBS("% Progress"))), _
                NormalizeCalendarType(arr(r, mapWBS("Cal"))) _
            )
        End If
    Next r

SafeExit:
    Set GanttLive_BuildBaseByIdMap = d

End Function

'------------------------------------------------------------------------------
' FR: Compare base et simulation pour savoir si une tache merite un contour/overlay GANTT.
' EN: Compares base and simulation to decide whether a task deserves a GANTT outline/overlay.
'------------------------------------------------------------------------------
Public Function GanttLive_HasRenderableTestDelta(ByVal idVal As String, ByVal baseById As Object, ByVal testById As Object) As Boolean

    Dim baseData As Variant
    Dim testData As Variant

    If idVal = "" Then Exit Function
    If Not baseById.Exists(idVal) Then Exit Function
    If Not testById.Exists(idVal) Then Exit Function

    baseData = baseById(idVal)
    testData = testById(idVal)

    ' Pas de contour si la ligne test est en erreur
    If Trim$(CStr(testData(5))) <> "" Then Exit Function

    ' Contour si le rendu simulé affiché diffère du rendu de base
    If ValuesDiffer(baseData(0), testData(0)) Then
        GanttLive_HasRenderableTestDelta = True
        Exit Function
    End If

    If ValuesDiffer(baseData(1), testData(1)) Then
        GanttLive_HasRenderableTestDelta = True
        Exit Function
    End If

    If ValuesDiffer(baseData(3), testData(3)) Then
        GanttLive_HasRenderableTestDelta = True
        Exit Function
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la date de debut a afficher, priorite au TEST valide puis fallback base.
' EN: Returns the start date to display, preferring valid TEST then falling back to base.
'------------------------------------------------------------------------------
Public Function GanttLive_GetDisplayStart(ByVal idVal As String, ByVal baseById As Object, ByVal testById As Object, ByVal isTestMode As Boolean) As Variant

    If isTestMode Then
        If testById.Exists(idVal) Then
            If Trim$(CStr(testById(idVal)(5))) = "" Then
                If HasValue(testById(idVal)(0)) Then
                    GanttLive_GetDisplayStart = testById(idVal)(0)
                    Exit Function
                End If
            End If
        End If
    End If

    If baseById.Exists(idVal) Then
        GanttLive_GetDisplayStart = baseById(idVal)(0)
    Else
        GanttLive_GetDisplayStart = Empty
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la date de fin a afficher, priorite au TEST valide puis fallback base.
' EN: Returns the finish date to display, preferring valid TEST then falling back to base.
'------------------------------------------------------------------------------
Public Function GanttLive_GetDisplayFinish(ByVal idVal As String, ByVal baseById As Object, ByVal testById As Object, ByVal isTestMode As Boolean) As Variant

    If isTestMode Then
        If testById.Exists(idVal) Then
            If Trim$(CStr(testById(idVal)(5))) = "" Then
                If HasValue(testById(idVal)(1)) Then
                    GanttLive_GetDisplayFinish = testById(idVal)(1)
                    Exit Function
                End If
            End If
        End If
    End If

    If baseById.Exists(idVal) Then
        GanttLive_GetDisplayFinish = baseById(idVal)(1)
    Else
        GanttLive_GetDisplayFinish = Empty
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la duree a afficher, priorite au TEST valide puis fallback base.
' EN: Returns the duration to display, preferring valid TEST then falling back to base.
'------------------------------------------------------------------------------
Public Function GanttLive_GetDisplayDuration(ByVal idVal As String, ByVal baseById As Object, ByVal testById As Object, ByVal isTestMode As Boolean) As Variant

    If isTestMode Then
        If testById.Exists(idVal) Then
            If Trim$(CStr(testById(idVal)(5))) = "" Then
                If HasValue(testById(idVal)(2)) Then
                    GanttLive_GetDisplayDuration = testById(idVal)(2)
                    Exit Function
                End If
            End If
        End If
    End If

    If baseById.Exists(idVal) Then
        GanttLive_GetDisplayDuration = baseById(idVal)(2)
    Else
        GanttLive_GetDisplayDuration = Empty
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne le pourcentage a afficher, priorite au TEST valide puis fallback base.
' EN: Returns the progress value to display, preferring valid TEST then falling back to base.
'------------------------------------------------------------------------------
Public Function GanttLive_GetDisplayProgress(ByVal idVal As String, ByVal baseById As Object, ByVal testById As Object, ByVal isTestMode As Boolean) As Variant

    If isTestMode Then
        If testById.Exists(idVal) Then
            If Trim$(CStr(testById(idVal)(5))) = "" Then
                If HasValue(testById(idVal)(3)) Then
                    GanttLive_GetDisplayProgress = testById(idVal)(3)
                    Exit Function
                End If
            End If
        End If
    End If

    If baseById.Exists(idVal) Then
        GanttLive_GetDisplayProgress = baseById(idVal)(3)
    Else
        GanttLive_GetDisplayProgress = Empty
    End If

End Function



'------------------------------------------------------------------------------
' FR:
' Facade publique stable du bouton Lock Changes; conserve le contrat Excel
' historique et delegue le workflow LOCK au service dedie.
'
' EN:
' Stable public facade for the Lock Changes button; keeps the historical Excel
' contract and delegates the LOCK workflow to the dedicated service.
'
' Entrees / Inputs:
' - Colonnes TEST du GANTT et etat courant WBS/CALC existants.
'
' Sorties / Outputs:
' - Memes sorties et messages que le workflow LOCK historique.
'
' Appele par / Called by:
' - Bouton Lock Changes / workflow utilisateur GANTT.
'
' Notes:
' - Wrapper public permanent: conserver le nom pour OnAction.
'------------------------------------------------------------------------------
Public Sub Run_Gantt_Lock_Changes()

    GanttLockService_RunLockChanges
    GanttUiControls_EnsureCanonical ThisWorkbook.Worksheets(GANTT_SHEET)

End Sub



'------------------------------------------------------------------------------
' FR: Compare les valeurs TEST saisies aux valeurs de base pour filtrer les faux changements.
' EN: Compares entered TEST values with base values to filter out non-changes.
'------------------------------------------------------------------------------
Public Function HasMeaningfulLockDelta( _
    ByVal baseStart As Variant, _
    ByVal baseFinish As Variant, _
    ByVal baseProgress As Variant, _
    ByVal testStart As Variant, _
    ByVal testFinish As Variant, _
    ByVal testProgress As Variant) As Boolean

    If HasValue(testStart) Then
        If ValuesDiffer(baseStart, testStart) Then
            HasMeaningfulLockDelta = True
            Exit Function
        End If
    End If

    If HasValue(testFinish) Then
        If ValuesDiffer(baseFinish, testFinish) Then
            HasMeaningfulLockDelta = True
            Exit Function
        End If
    End If

    If HasValue(testProgress) Then
        If ValuesDiffer(baseProgress, testProgress) Then
            HasMeaningfulLockDelta = True
            Exit Function
        End If
    End If

End Function



'------------------------------------------------------------------------------
' FR: Vide les resultats de tbl_CALC_GANTT_TEST sans supprimer la table.
' EN: Clears tbl_CALC_GANTT_TEST results without deleting the table.
'------------------------------------------------------------------------------
Public Sub ClearCalcGanttTestResults()

    GanttSimulation_ClearResults

End Sub


'------------------------------------------------------------------------------
' FR:
' Facade publique stable du workflow SCENARIO; delegue l'execution au service
' Gantt SCENARIO sans changer le contrat appele par Excel ou les orchestrateurs.
'
' EN:
' Stable public facade for the SCENARIO workflow; delegates execution to the
' Gantt SCENARIO service without changing the contract used by Excel or orchestrators.
'
' Entrees / Inputs:
' - Colonnes TEST/SCENARIO du GANTT et options transactionnelles existantes.
'
' Sorties / Outputs:
' - Memes sorties transactionnelles que le moteur SCENARIO historique.
'
' Appele par / Called by:
' - Boutons/workflows SCENARIO.
' - GanttLive_RunScenarioTransaction.
'
' Notes:
' - Wrapper public permanent: conserver le nom pour OnAction et appels existants.
'------------------------------------------------------------------------------
Public Sub Run_Gantt_Scenario_Engine( _
    Optional ByVal silentMode As Boolean = False, _
    Optional ByRef transactionSucceeded As Variant, _
    Optional ByRef transactionMessages As Variant, _
    Optional ByRef transactionGanttRebuilt As Variant, _
    Optional ByVal recordSilentMessages As Boolean = True)

    GanttScenarioService_RunScenarioEngine _
        silentMode, _
        transactionSucceeded, _
        transactionMessages, _
        transactionGanttRebuilt, _
        recordSilentMessages

    GanttUiControls_EnsureCanonical ThisWorkbook.Worksheets(GANTT_SHEET)

End Sub
'------------------------------------------------------------------------------
' FR:
' Execute le moteur SCENARIO en mode silencieux pour fournir succes, messages
' et statut de rebuild a un orchestrateur.
'
' EN:
' Runs the SCENARIO engine silently to provide success, messages, and rebuild
' status to an orchestrator.
'
' Entrees / Inputs:
' - References de sortie consoleMessages et ganttRebuilt.
'
' Sorties / Outputs:
' - Booleen succes/echec, messages et indicateur de reconstruction GANTT.
'
' Appele par / Called by:
' - Workflows SCENARIO transactionnels.
'
' Notes:
' - Wrapper public stable autour de Run_Gantt_Scenario_Engine.
'------------------------------------------------------------------------------
Public Function GanttLive_RunScenarioTransaction( _
    ByRef consoleMessages As Collection, _
    ByRef ganttRebuilt As Boolean) As Boolean

    Dim transactionSucceeded As Variant
    Dim transactionMessages As Variant
    Dim transactionGanttRebuilt As Variant

    Set consoleMessages = New Collection
    ganttRebuilt = False

    Run_Gantt_Scenario_Engine _
        True, _
        transactionSucceeded, _
        transactionMessages, _
        transactionGanttRebuilt, _
        False

    If IsObject(transactionMessages) Then
        Set consoleMessages = transactionMessages
    End If

    If Not IsEmpty(transactionGanttRebuilt) Then
        ganttRebuilt = CBool(transactionGanttRebuilt)
    End If

    If Not IsEmpty(transactionSucceeded) Then
        GanttLive_RunScenarioTransaction = CBool(transactionSucceeded)
    End If

End Function
