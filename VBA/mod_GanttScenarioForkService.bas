Attribute VB_Name = "mod_GanttScenarioForkService"
Option Explicit

'===============================================================================
' MODULE : mod_GanttScenarioForkService
' DOMAINE / DOMAIN : Gantt
'
' FR
' Possede le workflow specialise indique par son nom et expose ses contrats stables.
' Ne possede pas les domaines appeles en dependance.
'
' EN
' Owns the named specialized workflow and exposes its stable contracts.
' Does not own the domains it calls as dependencies.
'
' CONTRATS / CONTRACTS : CreateScenarioPlanningFromCurrentScenario, InitializeScenarioPlanningCopyFromCurrentWorkbook
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : InitializeScenarioPlanningCopyFromCurrentWorkbook
'===============================================================================

Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"



'------------------------------------------------------------------------------
' FR:
' Lorsque l'utilisateur tente de lock en mode SCENARIO, propose de creer une copie
' de classeur dont le scenario courant devient la nouvelle baseline.
'
' EN:
' When the user tries to lock in SCENARIO mode, offers to create a workbook copy
' where the current scenario becomes the new baseline.
'
' Entrees / Inputs:
' - Mode simulation SCENARIO actif et classeur source enregistre.
'
' Sorties / Outputs:
' - Nouveau classeur .xlsm cree, ouvert puis initialise par Application.Run.
'
' Appele par / Called by:
' - Run_Gantt_Lock_Changes.
'
' Notes:
' - Point Application.Run inter-workbook; le nom public initialise doit rester stable.
'------------------------------------------------------------------------------
Public Sub CreateScenarioPlanningFromCurrentScenario()

    Dim answer As VbMsgBoxResult
    Dim newPath As String
    Dim newWb As Workbook
    Dim oldScreenUpdating As Boolean
    Dim oldAlerts As Boolean
    Dim oldEvents As Boolean
    Dim macroName As String

    answer = MsgBox( _
        "Vous êtes actuellement en mode scénario." & vbCrLf & vbCrLf & _
        "Le lock direct n'est pas autorisé en mode scénario." & vbCrLf & vbCrLf & _
        "Voulez-vous créer un nouveau planning scénario basé sur l'état calculé actuel ?" & vbCrLf & vbCrLf & _
        "Le nouveau fichier :" & vbCrLf & _
        "* utilisera le scénario actuel comme nouvelle baseline ;" & vbCrLf & _
        "* conservera les % Progress du scénario ;" & vbCrLf & _
        "* videra Actual et Forecast ;" & vbCrLf & _
        "* désactivera les contraintes ;" & vbCrLf & _
        "* videra l'historique et les ACK ;" & vbCrLf & _
        "* sortira du mode scénario.", _
        vbQuestion + vbYesNo, _
        "Créer un planning scénario")

    If answer <> vbYes Then Exit Sub

    If Trim$(ThisWorkbook.Path) = "" Then
        MsgBox "Le fichier source doit être enregistré avant de créer un planning scénario.", vbExclamation, "Créer un planning scénario"
        Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents

    On Error GoTo Fail

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    newPath = BuildScenarioPlanningCopyPath()
    ThisWorkbook.SaveCopyAs newPath

    Set newWb = Application.Workbooks.Open(newPath, UpdateLinks:=0, ReadOnly:=False)
    Application.EnableEvents = True

    macroName = "'" & Replace(newWb.Name, "'", "''") & "'!InitializeScenarioPlanningCopyFromCurrentWorkbook"
    Application.Run macroName

    newWb.Activate
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    MsgBox "Erreur pendant la création du planning scénario :" & vbCrLf & Err.Description, vbCritical, "Créer un planning scénario"

End Sub

'------------------------------------------------------------------------------
' FR: Construit un chemin unique horodate pour la copie de planning scenario.
' EN: Builds a unique timestamped path for the scenario planning copy.
'------------------------------------------------------------------------------
Private Function BuildScenarioPlanningCopyPath() As String

    Dim folderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim extName As String
    Dim dotPos As Long
    Dim candidate As String
    Dim suffix As Long
    Dim stamp As String

    folderPath = ThisWorkbook.Path
    fileName = ThisWorkbook.Name
    dotPos = InStrRev(fileName, ".")
    stamp = Format$(Now, "yyyymmdd_hhnn")

    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
        extName = Mid$(fileName, dotPos)
    Else
        baseName = fileName
        extName = ".xlsm"
    End If

    candidate = folderPath & Application.PathSeparator & baseName & "_SCENARIO_" & stamp & extName
    suffix = 1

    Do While Len(Dir$(candidate)) > 0
        suffix = suffix + 1
        candidate = folderPath & Application.PathSeparator & baseName & "_SCENARIO_" & stamp & "_" & Format$(suffix, "00") & extName
    Loop

    BuildScenarioPlanningCopyPath = candidate

End Function

'------------------------------------------------------------------------------
' FR:
' Initialise la copie scenario: pousse les resultats scenario dans la baseline WBS,
' vide les etats runtime, desactive les contraintes et lance un recalcul force.
'
' EN:
' Initializes the scenario copy: pushes scenario results into WBS baseline,
' clears runtime state, deactivates constraints, and runs a forced recalculation.
'
' Entrees / Inputs:
' - Nouveau classeur ouvert par CreateScenarioPlanningFromCurrentScenario.
'
' Sorties / Outputs:
' - WBS converti en nouvelle baseline, runtime nettoye, classeur sauvegarde.
'
' Appele par / Called by:
' - Application.Run depuis la copie scenario nouvellement ouverte.
'
' Notes:
' - API publique callback inter-workbook; conserver un wrapper permanent si refactor.
'------------------------------------------------------------------------------
Public Sub InitializeScenarioPlanningCopyFromCurrentWorkbook()

    Dim oldScreenUpdating As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim inputGuardToken As Long
    Dim forcedUpdateGuardToken As Long

    On Error GoTo Fail

    oldScreenUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    inputGuardToken = OpenAuthorizedWBSWriteScope( _
        "ScenarioFork", Array("Baseline Start", "Baseline Duration", "Forecast Start", "Forecast Finish", "Actual Start", "Actual Finish", "% Progress"))
    Application.EnableEvents = True
    ApplyScenarioBaselineToWBS_CurrentWorkbook
    Application.EnableEvents = False
    CloseAuthorizedWBSWriteScope inputGuardToken
    inputGuardToken = 0

    Constraints_DeactivateAll
    EventHistory_ResetRuntimeStorage
    GanttSimulation_ResetToNormal False

    Application.EnableEvents = oldEvents
    forcedUpdateGuardToken = OpenAuthorizedWBSWriteScope( _
        "ScenarioForkForcedUpdate", ScenarioForkCalculatedWBSColumns())
    Run_Forced_Planning_Update

    ThisWorkbook.Save
    CloseAuthorizedWBSWriteScope forcedUpdateGuardToken
    forcedUpdateGuardToken = 0

CleanExit:
    On Error Resume Next
    CloseAuthorizedWBSWriteScope forcedUpdateGuardToken
    CloseAuthorizedWBSWriteScope inputGuardToken
    On Error GoTo 0
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    Exit Sub

Fail:
    MsgBox "Erreur pendant l'initialisation du nouveau planning scénario :" & vbCrLf & Err.Description, vbCritical, "Planning scénario"
    Resume CleanExit

End Sub

'------------------------------------------------------------------------------
' FR: Liste les colonnes WBS calculees autorisees pendant l'update force de la copie scenario.
' EN: Lists WBS calculated columns authorized during the forced update of the scenario copy.
'------------------------------------------------------------------------------
Private Function ScenarioForkCalculatedWBSColumns() As Variant

    ScenarioForkCalculatedWBSColumns = Array( _
        "Baseline Finish", _
        "Actual Duration", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration", _
        "Start Variance", _
        "Finish Variance", _
        "Duration Variance", _
        "Driving Logic", _
        "Critical Path", _
        "Longest Path", _
        "Critical Path REX", _
        "Total Float", _
        "Free Float", _
        "Total Float REX", _
        "Free Float REX", _
        "Deadline Float")

End Function

' FR:
' Convertit les resultats calcules du scenario en baseline WBS de la copie,
' tout en vidant Actual/Forecast et en evitant summaries/LOE pour les inputs.
'
' EN:
' Converts calculated scenario results into the copy's WBS baseline, while
' clearing Actual/Forecast and avoiding summaries/LOE for input writes.
'
' Entrees / Inputs:
' - tbl_WBS et tbl_CALC_GANTT_TEST de la copie scenario.
'
' Sorties / Outputs:
' - Baseline Start/Duration et progress WBS mis a jour; Actual/Forecast effaces.
'
' Appele par / Called by:
' - InitializeScenarioPlanningCopyFromCurrentWorkbook.
'
' Notes:
' - Ecrit dans WBS sous garde BeginAuthorizedWBSWrite du workflow scenario.
'------------------------------------------------------------------------------
Private Sub ApplyScenarioBaselineToWBS_CurrentWorkbook()

    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim scenarioById As Object
    Dim arrWBS As Variant
    Dim scenarioValues As Variant
    Dim r As Long
    Dim idVal As String
    Dim scenarioStart As Variant
    Dim scenarioDuration As Variant
    Dim scenarioProgress As Variant
    Dim isScenarioSummary As Boolean
    Dim isScenarioLOE As Boolean
    Dim writeScenarioInputs As Boolean
    Dim hasValidScenarioBaseline As Boolean

    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    RequireScenarioForkColumns mapWBS
    Set scenarioById = GanttSimulation_BuildScenarioBaselineById()
    If scenarioById.Count = 0 Then Exit Sub
    arrWBS = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arrWBS, 1)
        idVal = Trim$(CStr(arrWBS(r, mapWBS("ID"))))
        If idVal <> "" And scenarioById.Exists(idVal) Then
            scenarioValues = scenarioById(idVal)
            scenarioStart = scenarioValues(0)
            scenarioDuration = scenarioValues(1)
            scenarioProgress = scenarioValues(2)
            isScenarioSummary = CBool(scenarioValues(3))
            isScenarioLOE = CBool(scenarioValues(4))
            writeScenarioInputs = (Not isScenarioSummary) And (Not isScenarioLOE)
            hasValidScenarioBaseline = HasValue(scenarioStart) And IsScenarioForkPositiveDuration(scenarioDuration)

            If writeScenarioInputs Then
                If hasValidScenarioBaseline Then
                    tblWBS.DataBodyRange.cells(r, mapWBS("Baseline Start")).value = scenarioStart
                    tblWBS.DataBodyRange.cells(r, mapWBS("Baseline Duration")).value = scenarioDuration
                End If
                If HasValue(scenarioProgress) Then tblWBS.DataBodyRange.cells(r, mapWBS("% Progress")).value = scenarioProgress
            End If

            tblWBS.DataBodyRange.cells(r, mapWBS("Forecast Start")).ClearContents
            tblWBS.DataBodyRange.cells(r, mapWBS("Forecast Finish")).ClearContents
            tblWBS.DataBodyRange.cells(r, mapWBS("Actual Start")).ClearContents
            tblWBS.DataBodyRange.cells(r, mapWBS("Actual Finish")).ClearContents
        End If
    Next r

    tblWBS.ListColumns("Baseline Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Baseline Duration").DataBodyRange.NumberFormat = "0"
    tblWBS.ListColumns("Forecast Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Forecast Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Actual Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("Actual Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblWBS.ListColumns("% Progress").DataBodyRange.NumberFormat = "0%"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie les colonnes necessaires pour transformer un scenario en nouvelle baseline WBS.
' EN: Validates the columns required to turn a scenario into a new WBS baseline.
'------------------------------------------------------------------------------
Private Sub RequireScenarioForkColumns(ByVal mapWBS As Object)

    Dim c As Variant
    For Each c In Array("ID", "Baseline Start", "Baseline Duration", "Forecast Start", "Forecast Finish", "Actual Start", "Actual Finish", "% Progress")
        If Not mapWBS.Exists(CStr(c)) Then Err.Raise vbObjectError + 1290, , "Missing WBS column: " & CStr(c)
    Next c

End Sub

'------------------------------------------------------------------------------
' FR: Valide qu'une duree scenario est numerique et strictement positive.
' EN: Validates that a scenario duration is numeric and strictly positive.
'------------------------------------------------------------------------------
Private Function IsScenarioForkPositiveDuration(ByVal value As Variant) As Boolean

    If Not HasValue(value) Then Exit Function
    If Not IsNumeric(value) Then Exit Function

    IsScenarioForkPositiveDuration = (CDbl(value) > 0)

End Function
