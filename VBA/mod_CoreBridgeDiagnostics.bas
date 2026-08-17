Attribute VB_Name = "mod_CoreBridgeDiagnostics"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgeDiagnostics
' DOMAINE / DOMAIN : Core Bridge
'
' FR
' Projette les erreurs Core et analytics en messages structures, puis les route vers console et historique.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Projects Core and analytics errors into structured messages and routes them to console and history.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : CalcBridge_AppendCoreErrorMessages, CalcBridge_AppendCoreErrorMessagesFromData, CalcBridge_ShowGroupedErrorMessage, CalcBridge_AddInfoMessage, CalcBridge_AddGroupedWarningToCollection, CalcBridge_AddConsoleMessage, CalcBridge_ShowPlanningConsole, CalcBridge_RecordPlanningMessages
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================



'------------------------------------------------------------------------------
' FR: Ajoute la collection Constraint Root Messages a la structure cible fournie par l'appelant.
' EN: Adds the Constraint Root Messages collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub CalcBridge_AddConstraintRootMessages( _
    ByVal consoleMessages As Collection, _
    ByVal constraintMessagesById As Object)

    Dim idVal As Variant

    If consoleMessages Is Nothing Then Exit Sub
    If constraintMessagesById Is Nothing Then Exit Sub

    For Each idVal In constraintMessagesById.Keys
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
            CalcBridge_ToPMConstraintMessage(CStr(constraintMessagesById(CStr(idVal))))
    Next idVal

End Sub


'------------------------------------------------------------------------------
' FR: Lit les erreurs bloquantes de CALC et ajoute leurs diagnostics structures a la collection console. En cas de source illisible, ajoute un STOP fallback fail-closed.
' EN: Reads blocking CALC errors and appends their structured diagnostics to the console collection. If the source is unreadable, appends a fail-closed fallback STOP.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AppendCoreErrorMessages( _
    ByVal consoleMessages As Collection, _
    ByVal tblCalc As ListObject)

    Dim mapCalc As Object
    Dim arr As Variant

    On Error GoTo FailSafe

    If consoleMessages Is Nothing Then Exit Sub
    If tblCalc Is Nothing Then GoTo FailSafe
    If tblCalc.DataBodyRange Is Nothing Then GoTo FailSafe

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    arr = tblCalc.DataBodyRange.value

    CalcBridge_AppendCoreErrorMessagesFromData consoleMessages, arr, mapCalc, Nothing, "PROD"
    Exit Sub

FailSafe:
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg( _
            "Calcul arrete : le moteur a detecte des erreurs bloquantes." & vbCrLf & _
            "Impossible de reconstruire le message detaille." & vbCrLf & _
            "-> verifier les colonnes Error flag et ErrorMsg dans tbl_CALC." & vbCrLf & _
            "-> aucune donnee calculee n'a ete repoussee vers WBS.", _
            "Calculation stopped: the engine detected blocking errors." & vbCrLf & _
            "Unable to rebuild the detailed message." & vbCrLf & _
            "-> check Error flag and ErrorMsg columns in tbl_CALC." & vbCrLf & _
            "-> no calculated data was pushed back to WBS.")

End Sub


'------------------------------------------------------------------------------
' FR: Projette les erreurs d'un dataset Core en messages racine, dependance, contrainte et cascade sans recalculer le planning. Les donnees invalides produisent un STOP fallback.
' EN: Projects Core dataset errors into root, dependency, constraint and cascade messages without recalculating planning. Invalid input produces a fallback STOP.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AppendCoreErrorMessagesFromData( _
    ByVal consoleMessages As Collection, _
    ByRef dataArr As Variant, _
    ByVal mapCore As Object, _
    Optional ByVal rootErrorIds As Object = Nothing, _
    Optional ByVal contextMode As String = "PROD", _
    Optional ByVal dependencyDiagnostics As Object = Nothing, _
    Optional ByVal constraintDiagnostics As Object = Nothing, _
    Optional ByVal cascadeDiagnostics As Object = Nothing)

    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long

    Dim idToWbs As Object
    Dim idToTaskName As Object

    Dim errMissingPred As Object
    Dim errCycle As Object
    Dim errUnsupportedLinkType As Object
    Dim errActualStartConflict As Object
    Dim errActualFinishConflict As Object
    Dim errForecastConflict As Object
    Dim errForecastFinishConflict As Object
    Dim errMissingDuration As Object
    Dim errStartNotComputable As Object
    Dim errFinishBeforeStart As Object
    Dim errLOEAsPredecessor As Object
    Dim errLOEMissingSS As Object
    Dim errLOEMissingFF As Object
    Dim errLOEInvalidLink As Object
    Dim errOtherRoot As Object
    Dim errConstraintRootMessages As Object
    Dim cycleDetailMessage As String

    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String
    Dim errMsg As String
    Dim hasSpecificRootError As Boolean
    Dim contextKey As String

    On Error GoTo FailSafe

    If consoleMessages Is Nothing Then Exit Sub
    If Not IsArray(dataArr) Then GoTo FailSafe
    If mapCore Is Nothing Then GoTo FailSafe

    Set mapCalc = mapCore
    arr = dataArr
    contextKey = UCase$(Trim$(contextMode))
    Set errMissingPred = CreateObject("Scripting.Dictionary")
    Set errCycle = CreateObject("Scripting.Dictionary")
    Set errUnsupportedLinkType = CreateObject("Scripting.Dictionary")
    Set errActualStartConflict = CreateObject("Scripting.Dictionary")
    Set errActualFinishConflict = CreateObject("Scripting.Dictionary")
    Set errForecastConflict = CreateObject("Scripting.Dictionary")
    Set errForecastFinishConflict = CreateObject("Scripting.Dictionary")
    Set errMissingDuration = CreateObject("Scripting.Dictionary")
    Set errStartNotComputable = CreateObject("Scripting.Dictionary")
    Set errFinishBeforeStart = CreateObject("Scripting.Dictionary")
    Set errLOEAsPredecessor = CreateObject("Scripting.Dictionary")
    Set errLOEMissingSS = CreateObject("Scripting.Dictionary")
    Set errLOEMissingFF = CreateObject("Scripting.Dictionary")
    Set errLOEInvalidLink = CreateObject("Scripting.Dictionary")
    Set errOtherRoot = CreateObject("Scripting.Dictionary")
    Set errConstraintRootMessages = CreateObject("Scripting.Dictionary")
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set idToTaskName = CreateObject("Scripting.Dictionary")

    If Not mapCalc.Exists("ID") Then GoTo FailSafe
    If Not mapCalc.Exists("Error flag") Then GoTo FailSafe
    If Not mapCalc.Exists("ErrorMsg") Then GoTo FailSafe

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))

        If idVal <> "" Then

            If mapCalc.Exists("WBS") Then
                wbsVal = Trim$(CStr(arr(r, mapCalc("WBS"))))
            Else
                wbsVal = ""
            End If

            idToWbs(idVal) = wbsVal
            If mapCalc.Exists("Task Name") Then
                taskNameVal = Trim$(CStr(arr(r, mapCalc("Task Name"))))
            Else
                taskNameVal = ""
            End If
            idToTaskName(idVal) = taskNameVal

            If UCase$(Trim$(CStr(arr(r, mapCalc("Error flag"))))) = "ERROR" Then

                errMsg = Trim$(CStr(arr(r, mapCalc("ErrorMsg"))))

                'Inherited errors remain visible in the source table,
                'but are not highlighted in the main popup.
                If CalcBridge_IsInheritedCoreError(errMsg) Then
                    GoTo NextRow
                End If

                If Not rootErrorIds Is Nothing Then
                    If Not rootErrorIds.Exists(idVal) Then GoTo NextRow
                End If

                Select Case True

                    Case InStr(1, errMsg, "LOE cannot be used as predecessor", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "Blocked by invalid LOE predecessor", vbTextCompare) > 0
                        errLOEAsPredecessor(idVal) = True

                    Case InStr(1, errMsg, "LOE must have at least one SS predecessor", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE SS predecessor missing", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE SS predecessor not found", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE SS predecessor start not available", vbTextCompare) > 0
                        errLOEMissingSS(idVal) = True

                    Case InStr(1, errMsg, "LOE must have at least one FF predecessor", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE FF predecessor missing", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE FF predecessor not found", vbTextCompare) > 0 Or _
                         InStr(1, errMsg, "LOE FF predecessor finish not available", vbTextCompare) > 0
                        errLOEMissingFF(idVal) = True

                    Case InStr(1, errMsg, "LOE only supports SS and FF predecessors", vbTextCompare) > 0
                        errLOEInvalidLink(idVal) = True

                    Case InStr(1, errMsg, "Missing predecessor", vbTextCompare) > 0
                        errMissingPred(idVal) = True

                    Case InStr(1, errMsg, "Cycle detected", vbTextCompare) > 0
                        errCycle(idVal) = True
                        If cycleDetailMessage = "" Then cycleDetailMessage = errMsg

                    Case InStr(1, errMsg, "Unsupported link type", vbTextCompare) > 0
                        errUnsupportedLinkType(idVal) = True

                    Case InStr(1, errMsg, "Actual Start violates dependencies", vbTextCompare) > 0
                        errActualStartConflict(idVal) = True

                    Case InStr(1, errMsg, "Actual Finish violates finish constraints", vbTextCompare) > 0
                        errActualFinishConflict(idVal) = True

                    Case InStr(1, errMsg, "Forecast Start violates dependencies", vbTextCompare) > 0
                        errForecastConflict(idVal) = True

                    Case InStr(1, errMsg, "Forecast Finish violates finish constraints", vbTextCompare) > 0
                        errForecastFinishConflict(idVal) = True

                    Case InStr(1, errMsg, "Baseline Duration missing", vbTextCompare) > 0
                        errMissingDuration(idVal) = True

                    Case InStr(1, errMsg, "Start date not computable", vbTextCompare) > 0
                        errStartNotComputable(idVal) = True

                    Case InStr(1, errMsg, "Finish before start", vbTextCompare) > 0
                        errFinishBeforeStart(idVal) = True

                    Case CalcBridge_IsConstraintCoreError(errMsg)
                        errConstraintRootMessages(idVal) = errMsg

                    Case Else
                        errOtherRoot(idVal) = True

                End Select

            End If
        End If

NextRow:
    Next r

    hasSpecificRootError = _
        (errMissingPred.Count > 0) Or _
        (errUnsupportedLinkType.Count > 0) Or _
        (errCycle.Count > 0) Or _
        (errActualStartConflict.Count > 0) Or _
        (errActualFinishConflict.Count > 0) Or _
        (errForecastConflict.Count > 0) Or _
        (errForecastFinishConflict.Count > 0) Or _
        (errMissingDuration.Count > 0) Or _
        (errStartNotComputable.Count > 0) Or _
        (errFinishBeforeStart.Count > 0) Or _
        (errConstraintRootMessages.Count > 0) Or _
        (errLOEAsPredecessor.Count > 0) Or _
        (errLOEMissingSS.Count > 0) Or _
        (errLOEMissingFF.Count > 0) Or _
        (errLOEInvalidLink.Count > 0)

    If errLOEAsPredecessor.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEAsPredecessor, idToWbs, _
            "LOE utilisée comme prédécesseur", _
            "supprimer la LOE de la logique amont ; une LOE est pilotée par le réseau mais ne doit pas piloter d'autres tâches", _
            "LOE used as predecessor", _
            "remove the LOE from upstream logic; a LOE is driven by the network but must not drive other tasks"
    End If

    If errLOEMissingSS.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEMissingSS, idToWbs, _
            "LOE sans lien SS exploitable", _
            "ajouter au moins un prédécesseur SS valide pour définir le début de la LOE", _
            "LOE without usable SS link", _
            "add at least one valid SS predecessor to define the LOE start"
    End If

    If errLOEMissingFF.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEMissingFF, idToWbs, _
            "LOE sans lien FF exploitable", _
            "ajouter au moins un prédécesseur FF valide pour définir la fin de la LOE", _
            "LOE without usable FF link", _
            "add at least one valid FF predecessor to define the LOE finish"
    End If

    If errLOEInvalidLink.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errLOEInvalidLink, idToWbs, _
            "Lien invalide sur LOE", _
            "une LOE accepte uniquement des liens SS et FF", _
            "Invalid link on LOE", _
            "a LOE only supports SS and FF links"
    End If

    If errMissingPred.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errMissingPred, idToWbs, _
            "Prédécesseur introuvable", _
            "vérifier la colonne Predecessors WBS", _
            "Missing predecessor", _
            "check the Predecessors WBS column"
    End If

    If errUnsupportedLinkType.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errUnsupportedLinkType, idToWbs, _
            "Type de lien non supporté par le moteur", _
            "corriger le type de lien dans Predecessors WBS ou tbl_LOGIC_LINKS", _
            "Link type not supported by the engine", _
            "fix the link type in Predecessors WBS or tbl_LOGIC_LINKS"
    End If

    If errCycle.Count > 0 Then
        If CalcBridge_CycleDetailMessageFromCoreError(cycleDetailMessage) <> "" Then
            CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
                CalcBridge_CycleDetailMessageFromCoreError(cycleDetailMessage)
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errCycle, idToWbs, _
                "Boucle de dépendance détectée", _
                "corriger la colonne Predecessors WBS", _
                "Dependency cycle detected", _
                "fix the Predecessors WBS column"
        End If
    End If

    If errActualStartConflict.Count > 0 Then
        CalcBridge_AddUpstreamStopToCollection consoleMessages, errActualStartConflict, idToWbs, _
            "Actual Start incompatible avec les d" & ChrW$(233) & "pendances amont", _
            "corriger Actual Start, la logique amont ou le lag", _
            "Actual Start is incompatible with upstream dependencies", _
            "fix Actual Start, upstream logic, or lag"
    End If

    If errActualFinishConflict.Count > 0 Then
        If Not CalcBridge_TryAddConstraintDiagnosticStops(consoleMessages, errActualFinishConflict, idToWbs, idToTaskName, constraintDiagnostics, cascadeDiagnostics, contextKey) Then
            CalcBridge_AddUpstreamStopToCollection consoleMessages, errActualFinishConflict, idToWbs, _
                "Actual Finish incompatible avec les contraintes de fin amont", _
                "corriger Actual Finish, la logique amont ou le lag", _
                "Actual Finish is incompatible with upstream finish constraints", _
                "fix Actual Finish, upstream logic, or lag"
        End If
    End If

    If errForecastConflict.Count > 0 Then
        If contextKey = "TEST" Or contextKey = "SCENARIO" Then
            If Not CalcBridge_TryAddForecastStartDependencyDiagnosticStops( _
                consoleMessages, errForecastConflict, idToWbs, idToTaskName, dependencyDiagnostics, contextKey) Then

                CalcBridge_AddGroupedStopToCollection consoleMessages, errForecastConflict, idToWbs, _
                    "Test Start incompatible avec les d" & ChrW$(233) & "pendances amont", _
                    "corriger Test Start ou la logique amont", _
                    "Test Start is incompatible with upstream dependencies", _
                    "fix Test Start or upstream logic"
            End If
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errForecastConflict, idToWbs, _
                "Forecast Start incompatible avec les d" & ChrW$(233) & "pendances amont", _
                "corriger Forecast Start ou la logique amont", _
                "Forecast Start is incompatible with upstream dependencies", _
                "fix Forecast Start or upstream logic"
        End If
    End If
    If errForecastFinishConflict.Count > 0 Then
        If CalcBridge_TryAddConstraintDiagnosticStops(consoleMessages, errForecastFinishConflict, idToWbs, idToTaskName, constraintDiagnostics, cascadeDiagnostics, contextKey) Then
            'Structured constraint diagnostic already rendered.
        ElseIf contextKey = "TEST" Or contextKey = "SCENARIO" Then
            CalcBridge_AddGroupedStopToCollection consoleMessages, errForecastFinishConflict, idToWbs, _
                "Test Finish incompatible avec les contraintes de fin amont", _
                "corriger Test Finish ou la logique amont", _
                "Test Finish is incompatible with upstream finish constraints", _
                "fix Test Finish or upstream logic"
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errForecastFinishConflict, idToWbs, _
                "Forecast Finish incompatible avec les contraintes de fin amont", _
                "corriger Forecast Finish ou la logique amont", _
                "Forecast Finish is incompatible with upstream finish constraints", _
                "fix Forecast Finish or upstream logic"
        End If
    End If
    If errConstraintRootMessages.Count > 0 Then
        If Not CalcBridge_TryAddConstraintDiagnosticStops(consoleMessages, errConstraintRootMessages, idToWbs, idToTaskName, constraintDiagnostics, cascadeDiagnostics, contextKey) Then
            CalcBridge_AddConstraintRootMessages consoleMessages, errConstraintRootMessages
        End If
    End If

    If errMissingDuration.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errMissingDuration, idToWbs, _
            "Baseline Duration manquante", _
            "compléter la durée baseline", _
            "Missing Baseline Duration", _
            "please fill in Baseline Duration"
    End If

    If errStartNotComputable.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errStartNotComputable, idToWbs, _
            "Date de début non déterminable", _
            "vérifier les dépendances ou la baseline", _
            "Start date not computable", _
            "check dependencies or baseline"
    End If

    If errFinishBeforeStart.Count > 0 Then
        CalcBridge_AddGroupedStopToCollection consoleMessages, errFinishBeforeStart, idToWbs, _
            "Fin incompatible avec le d" & ChrW$(233) & "but", _
            "corriger les dates ou la dur" & ChrW$(233) & "e", _
            "Finish is incompatible with start", _
            "fix dates or duration"
    End If

    If errOtherRoot.Count > 0 And Not hasSpecificRootError Then
        If contextKey = "TEST" Then
            CalcBridge_AddGroupedStopToCollection consoleMessages, errOtherRoot, idToWbs, _
                "Erreur de calcul dans le moteur live", _
                "corriger les valeurs test ou la logique amont", _
                "Calculation error in live engine", _
                "fix test values or upstream logic"
        ElseIf contextKey = "SCENARIO" Then
            CalcBridge_AddGroupedStopToCollection consoleMessages, errOtherRoot, idToWbs, _
                "Erreur de calcul dans le sc" & ChrW$(233) & "nario", _
                "corriger les valeurs de test ou la logique amont", _
                "Calculation error in scenario", _
                "fix test values or upstream logic"
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errOtherRoot, idToWbs, _
                "Erreur bloquante d" & ChrW$(233) & "tect" & ChrW$(233) & "e par le moteur", _
                "v" & ChrW$(233) & "rifier ErrorMsg dans tbl_CALC pour le d" & ChrW$(233) & "tail technique", _
                "Blocking error detected by the engine", _
                "check ErrorMsg in tbl_CALC for technical details"
        End If
    End If
    Exit Sub

FailSafe:
    CalcBridge_AddConsoleMessage consoleMessages, "STOP", _
        BiMsg( _
            "Calcul arrêté : le moteur a détecté des erreurs bloquantes." & vbCrLf & _
            "Impossible de reconstruire le message détaillé." & vbCrLf & _
            "-> vérifier les colonnes Error flag et ErrorMsg dans tbl_CALC." & vbCrLf & _
            "-> aucune donnée calculée n'a été repoussée vers WBS.", _
            "Calculation stopped: the engine detected blocking errors." & vbCrLf & _
            "Unable to rebuild the detailed message." & vbCrLf & _
            "-> check Error flag and ErrorMsg columns in tbl_CALC." & vbCrLf & _
            "-> no calculated data was pushed back to WBS.")

End Sub


'------------------------------------------------------------------------------
' FR: Projette la collection Grouped Error Message vers l'interface autorisee par la politique runtime.
' EN: Projects the Grouped Error Message collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Public Sub CalcBridge_ShowGroupedErrorMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String)

    Dim consoleMessages As Collection

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    Set consoleMessages = New Collection

    CalcBridge_AddGroupedStopToCollection consoleMessages, idsDict, idToWbs, _
        frProblem, frAction, enProblem, enAction

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Info Message a la structure cible fournie par l'appelant.
' EN: Adds the Info Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddInfoMessage( _
    ByVal messages As Collection, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim msg As String

    If messages Is Nothing Then Exit Sub

    msg = _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

    CalcBridge_AddConsoleMessage messages, "INFO", msg

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Grouped Warning To Collection a la structure cible fournie par l'appelant.
' EN: Adds the Grouped Warning To Collection collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddGroupedWarningToCollection( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal ackTokens As String = "")

    Dim idsLine As String
    Dim wbsLine As String
    Dim msg As String

    If messages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    idsLine = CalcBridge_BuildInlineList(idsDict, 20)
    wbsLine = CalcBridge_BuildInlineWBSList(idsDict, idToWbs, 20)

    msg = _
        "FR:" & vbCrLf & _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "IDs : " & idsLine & vbCrLf & _
        "WBS : " & wbsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "IDs: " & idsLine & vbCrLf & _
        "WBS: " & wbsLine

    CalcBridge_AddConsoleMessage messages, "WARNING", msg, historyHandled, ackTokens:=ackTokens

End Sub


'------------------------------------------------------------------------------
' FR: Construit la map Console Item a partir des donnees fournies par l'appelant.
' EN: Builds the Console Item map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CalcBridge_CreateConsoleItem( _
    ByVal msgType As String, _
    ByVal msgText As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal eventType As String = "", _
    Optional ByVal eventHash As String = "", _
    Optional ByVal ackTokens As String = "") As Object

    Dim item As Object

    Set item = CreateObject("Scripting.Dictionary")
    item("Type") = UCase$(Trim$(msgType))
    item("Message") = CStr(msgText)
    item("HistoryHandled") = historyHandled
    If Trim$(eventType) <> "" Then item("EventType") = Trim$(eventType)
    If Trim$(eventHash) <> "" Then item("Hash") = Trim$(eventHash)
    If Trim$(ackTokens) <> "" Then item("AckTokens") = Trim$(ackTokens)

    Set CalcBridge_CreateConsoleItem = item

End Function


'------------------------------------------------------------------------------
' FR: Ajoute la collection Console Message a la structure cible fournie par l'appelant.
' EN: Adds the Console Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddConsoleMessage( _
    ByVal targetMessages As Collection, _
    ByVal msgType As String, _
    ByVal msgText As String, _
    Optional ByVal historyHandled As Boolean = False, _
    Optional ByVal eventType As String = "", _
    Optional ByVal eventHash As String = "", _
    Optional ByVal ackTokens As String = "")

    If targetMessages Is Nothing Then Exit Sub
    If Trim$(msgText) = "" Then Exit Sub

    targetMessages.Add CalcBridge_CreateConsoleItem(msgType, msgText, historyHandled, eventType, eventHash, ackTokens)

End Sub


'------------------------------------------------------------------------------
' FR: Projette la collection Planning Console vers l'interface autorisee par la politique runtime.
' EN: Projects the Planning Console collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Public Sub CalcBridge_ShowPlanningConsole(ByVal messages As Collection)

    Dim historyMessages As Collection
    Dim displayMessages As Collection
    Dim historyErrorMessage As String
    Dim prepScope As clsPerfScope
    Dim showScope As clsPerfScope

    If messages Is Nothing Then Exit Sub
    If messages.Count = 0 Then Exit Sub

    Profiler_RecordOperation "PlanningConsole_ENGINE_READY", 1, 0#
    Profiler_RecordOperation "PlanningConsole_PREP_START", 1, 0#
    Set prepScope = Profiler_BeginScope("PlanningConsole_PrepareBeforeShow", "Planning Console")

    Set historyMessages = MessageEngine_PrepareConsoleMessages(messages)

    If Not PlanningEvents_LogConsoleMessagesSafe( _
        historyMessages, _
        "CalcBridge_ShowPlanningConsole", _
        historyErrorMessage) Then

        CalcBridge_AddConsoleMessage historyMessages, "STOP", historyErrorMessage, True
    End If

    If Profiler_ShouldSuppressUserInterface() Then Exit Sub

    If ShouldDeferCurrentWorkflowDisplayToRoot("CalcBridge_ShowPlanningConsole") Then
        DeferPlanningWorkflowDisplayMessages historyMessages
        Exit Sub
    End If

    If IsPlanningWorkflowStopOnlyDisplay() Then
        Set displayMessages = CalcBridge_FilterConsoleMessagesBySeverity(historyMessages, "STOP")
    Else
        Set displayMessages = MessageEngine_PrepareDisplayMessages(historyMessages)
    End If
    If Not MessageEngine_ShouldShowConsole(displayMessages) Then Exit Sub
    If Not CanCurrentWorkflowDisplay("CalcBridge_ShowPlanningConsole") Then Exit Sub

    If PlanningConsolePolicy_IsNonInteractive() Then
        Set prepScope = Nothing
        RunButtonsTrace_Checkpoint "ConsolePolicy", "Planning console noninteractive capture start"
        PlanningConsolePolicy_CaptureDisplayMessages displayMessages, "Planning console"
        RunButtonsTrace_Checkpoint "ConsolePolicy", "Planning console noninteractive capture returned"
        Exit Sub
    End If

    Load frmPlanningMessages
    RunButtonsTrace_Checkpoint "Console", "Planning console modal load start"
    frmPlanningMessages.LoadMessages displayMessages, "Planning console"
    frmPlanningMessages.PrepareForImmediateShow
    RunButtonsTrace_Checkpoint "Console", "Planning console ready before show"
    Profiler_RecordOperation "PlanningConsole_PREP_END", 1, 0#
    Set prepScope = Nothing

    Profiler_RecordOperation "PlanningConsole_SHOW_START", 1, 0#
    RunButtonsTrace_Checkpoint "Console", "Planning console modal show start"
    Set showScope = Profiler_BeginScope("PlanningConsole_ShowModalLifetime", "Planning Console")
    frmPlanningMessages.Show vbModal
    Set showScope = Nothing
    RunButtonsTrace_Checkpoint "Console", "Planning console modal show returned"
    Profiler_RecordOperation "PlanningConsole_SHOW_RETURN", 1, 0#

End Sub


'------------------------------------------------------------------------------
' FR: Retourne la collection Filter Console Messages By Severity sans modifier les donnees d'entree.
' EN: Returns the Filter Console Messages By Severity collection without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcBridge_FilterConsoleMessagesBySeverity( _
    ByVal messages As Collection, _
    ByVal severity As String) As Collection

    Dim result As Collection
    Dim item As Variant

    Set result = New Collection

    If messages Is Nothing Then
        Set CalcBridge_FilterConsoleMessagesBySeverity = result
        Exit Function
    End If

    For Each item In messages
        If UCase$(Trim$(CStr(item("Type")))) = UCase$(Trim$(severity)) Then
            result.Add item
        End If
    Next item

    Set CalcBridge_FilterConsoleMessagesBySeverity = result

End Function


'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Record Planning Messages dans le stockage possede par le domaine.
' EN: Writes or synchronizes Record Planning Messages in the store owned by the domain.
'------------------------------------------------------------------------------

Public Function CalcBridge_RecordPlanningMessages( _
    ByVal messages As Collection, _
    Optional ByVal sourceProcedure As String = "CalcBridge_RecordPlanningMessages") As Boolean

    Dim historyMessages As Collection
    Dim historyErrorMessage As String

    If messages Is Nothing Then
        CalcBridge_RecordPlanningMessages = True
        Exit Function
    End If

    If messages.Count = 0 Then
        CalcBridge_RecordPlanningMessages = True
        Exit Function
    End If

    Set historyMessages = MessageEngine_PrepareConsoleMessages(messages)
    If historyMessages.Count = 0 Then
        CalcBridge_RecordPlanningMessages = True
        Exit Function
    End If

    CalcBridge_RecordPlanningMessages = PlanningEvents_LogConsoleMessagesSafe( _
        historyMessages, _
        sourceProcedure, _
        historyErrorMessage)

    If Not CalcBridge_RecordPlanningMessages Then
        Debug.Print "Runtime history logging failed in " & sourceProcedure & ": " & historyErrorMessage
    End If

End Function


'------------------------------------------------------------------------------
' FR: Retourne la collection Try Add Constraint Diagnostic Stops sans modifier les donnees d'entree.
' EN: Returns the Try Add Constraint Diagnostic Stops collection without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcBridge_TryAddConstraintDiagnosticStops( _
    ByVal messages As Collection, _
    ByVal constraintMessagesById As Object, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object, _
    ByVal constraintDiagnostics As Object, _
    ByVal cascadeDiagnostics As Object, _
    ByVal contextKey As String) As Boolean

    Dim key As Variant
    Dim diag As Object
    Dim detailMsg As String
    Dim addedAny As Boolean

    If messages Is Nothing Then Exit Function
    If constraintMessagesById Is Nothing Then Exit Function
    If constraintMessagesById.Count = 0 Then Exit Function

    For Each key In constraintMessagesById.Keys
        detailMsg = ""

        If Not constraintDiagnostics Is Nothing Then
            If constraintDiagnostics.Exists(CStr(key)) Then
                If IsObject(constraintDiagnostics(CStr(key))) Then
                    Set diag = constraintDiagnostics(CStr(key))
                    detailMsg = CalcBridge_BuildConstraintDiagnosticMessage( _
                        diag, idToWbs, idToTaskName, cascadeDiagnostics, contextKey)
                End If
            End If
        End If

        If Trim$(detailMsg) = "" Then
            detailMsg = CalcBridge_ToPMConstraintMessage(CStr(constraintMessagesById(CStr(key))))
        End If

        If Trim$(detailMsg) <> "" Then
            CalcBridge_AddConsoleMessage messages, "STOP", detailMsg
            addedAny = True
        End If
    Next key

    CalcBridge_TryAddConstraintDiagnosticStops = addedAny

End Function


'------------------------------------------------------------------------------
' FR: Retourne la collection Try Add Forecast Start Dependency Diagnostic Stops sans modifier les donnees d'entree.
' EN: Returns the Try Add Forecast Start Dependency Diagnostic Stops collection without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcBridge_TryAddForecastStartDependencyDiagnosticStops( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object, _
    ByVal dependencyDiagnostics As Object, _
    ByVal contextKey As String) As Boolean

    Dim key As Variant
    Dim diag As Object
    Dim detailMsg As String
    Dim addedAny As Boolean

    If messages Is Nothing Then Exit Function
    If idsDict Is Nothing Then Exit Function
    If idsDict.Count = 0 Then Exit Function
    If dependencyDiagnostics Is Nothing Then Exit Function

    For Each key In idsDict.Keys
        If dependencyDiagnostics.Exists(CStr(key)) Then
            If IsObject(dependencyDiagnostics(CStr(key))) Then
                Set diag = dependencyDiagnostics(CStr(key))
                detailMsg = CalcBridge_BuildForecastStartDependencyDiagnosticMessage( _
                    diag, idToWbs, idToTaskName, contextKey)

                If Trim$(detailMsg) <> "" Then
                    CalcBridge_AddConsoleMessage messages, "STOP", detailMsg
                    addedAny = True
                End If
            End If
        End If
    Next key

    CalcBridge_TryAddForecastStartDependencyDiagnosticStops = addedAny

End Function

'------------------------------------------------------------------------------
' FR: Ajoute la collection Grouped Stop To Collection a la structure cible fournie par l'appelant.
' EN: Adds the Grouped Stop To Collection collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddGroupedStopToCollection( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String)

    If messages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    CalcBridge_AddConsoleMessage messages, "STOP", _
        CalcBridge_BuildGroupedMessage(idsDict, idToWbs, frProblem, frAction, enProblem, enAction)

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Upstream Stop To Collection a la structure cible fournie par l'appelant.
' EN: Adds the Upstream Stop To Collection collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub CalcBridge_AddUpstreamStopToCollection( _
    ByVal messages As Collection, _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String)

    Dim itemsLine As String
    Dim msg As String

    If messages Is Nothing Then Exit Sub
    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    itemsLine = CalcBridge_BuildUpstreamViolationItems(idsDict, idToWbs, 20)

    msg = _
        "FR:" & vbCrLf & _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "Tâches : " & itemsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "Tasks: " & itemsLine

    CalcBridge_AddConsoleMessage messages, "STOP", msg

End Sub


'------------------------------------------------------------------------------
' FR: Projette la collection Single Console Message vers l'interface autorisee par la politique runtime.
' EN: Projects the Single Console Message collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Public Sub CalcBridge_ShowSingleConsoleMessage( _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim consoleMessages As Collection

    Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Or Show Raw Console Message a la structure cible fournie par l'appelant.
' EN: Adds the Or Show Raw Console Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddOrShowRawConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal msgText As String)

    Dim localMessages As Collection

    If Trim$(CStr(msgText)) = "" Then Exit Sub

    If consoleMessages Is Nothing Then
        Set localMessages = New Collection
        CalcBridge_AddConsoleMessage localMessages, msgType, msgText
        CalcBridge_ShowPlanningConsole localMessages
    Else
        CalcBridge_AddConsoleMessage consoleMessages, msgType, msgText
    End If

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute la collection Or Show Console Message a la structure cible fournie par l'appelant.
' EN: Adds the Or Show Console Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub CalcBridge_AddOrShowConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    CalcBridge_AddOrShowRawConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

End Sub



'------------------------------------------------------------------------------
' FR: Indique si la valeur Inherited Core Error satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Inherited Core Error value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function CalcBridge_IsInheritedCoreError(ByVal errMsg As String) As Boolean

    Dim txt As String

    txt = Trim$(CStr(errMsg))

    CalcBridge_IsInheritedCoreError = _
        (InStr(1, txt, "Blocked by predecessor error", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Blocked by predecessor chain", vbTextCompare) > 0)

End Function

'------------------------------------------------------------------------------
' FR: Indique si la valeur Constraint Core Error satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Constraint Core Error value satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function CalcBridge_IsConstraintCoreError(ByVal errMsg As String) As Boolean

    Dim txt As String

    txt = Trim$(CStr(errMsg))

    CalcBridge_IsConstraintCoreError = _
        (InStr(1, txt, "avant contrainte debut", vbTextCompare) > 0) Or _
        (InStr(1, txt, "before start constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "apres contrainte debut max", vbTextCompare) > 0) Or _
        (InStr(1, txt, "after latest start constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "different de contrainte Must Start On", vbTextCompare) > 0) Or _
        (InStr(1, txt, "differs from Must Start On constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "avant contrainte fin", vbTextCompare) > 0) Or _
        (InStr(1, txt, "before finish constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "apres contrainte fin max", vbTextCompare) > 0) Or _
        (InStr(1, txt, "after latest finish constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "contrainte Must Finish On", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Must Finish On constraint", vbTextCompare) > 0) Or _
        (InStr(1, txt, "contraintes Must Start On / Must Finish On", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Must Start On / Must Finish On constraints", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Type de contrainte debut non reconnu", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Unknown start constraint type", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Type de contrainte fin non reconnu", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Unknown finish constraint type", vbTextCompare) > 0)

End Function

