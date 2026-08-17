Attribute VB_Name = "mod_CoreBridgePreCoreValidation"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgePreCoreValidation
' DOMAINE / DOMAIN : Core Bridge
'
' FR
' Detecte avant Core les predecessors manquants, cycles et usages LOE bloquants, puis marque CALC.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Detects missing predecessors, cycles and blocking LOE uses before Core, then marks CALC.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : ValidateCalcAfterSync, CalcBridge_PreCore_CheckMissingPredecessors, CalcBridge_PreCore_CheckCycles, CalcBridge_PreCore_CheckLOEAsPredecessor
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Valide Calc After Sync et signale les incoherences detectees.
' EN: Validates Calc After Sync and reports detected inconsistencies.
'------------------------------------------------------------------------------
Public Function ValidateCalcAfterSync(ByVal tblCalc As ListObject) As Boolean

    Dim mapCalc As Object
    Dim arr As Variant
    Dim hasAnyId As Boolean
    Dim r As Long

    On Error GoTo Fail

    If tblCalc Is Nothing Then GoTo Fail
    If tblCalc.DataBodyRange Is Nothing Then GoTo Fail

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)

    If Not mapCalc.Exists("ID") Then GoTo Fail
    If Not mapCalc.Exists("Calculated Start") Then GoTo Fail
    If Not mapCalc.Exists("Calculated Finish") Then GoTo Fail

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If Trim$(CStr(arr(r, mapCalc("ID")))) <> "" Then
            hasAnyId = True
            Exit For
        End If
    Next r

    If Not hasAnyId Then GoTo Fail

    ValidateCalcAfterSync = True
    Exit Function

Fail:
    ValidateCalcAfterSync = False

End Function


'------------------------------------------------------------------------------
' FR: Valide Pre Core Check Missing Predecessors et applique la politique d'erreur definie par le composant.
' EN: Validates Pre Core Check Missing Predecessors and applies the component's defined failure policy.
'------------------------------------------------------------------------------

Public Function CalcBridge_PreCore_CheckMissingPredecessors( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object, _
    Optional ByVal consoleMessages As Collection) As Boolean

    Dim network As clsParsedPlanningNetwork
    Dim link As clsParsedPlanningLink

    Dim idToWbs As Object
    Dim errMissingPred As Object

    Dim r As Long
    Dim succId As String
    Dim predId As String

    On Error GoTo FailSafe

    CalcBridge_PreCore_CheckMissingPredecessors = False

    Set errMissingPred = CreateObject("Scripting.Dictionary")
    Set idToWbs = CalcBridge_PreCore_BuildIdToWbsFromCalc(tblCalc, mapCalc)

    Set network = ParsedPlanningNetwork_LoadCanonical()

    If Not network.HasColumn("Succ ID") Then Exit Function
    If Not network.HasColumn("Pred ID") Then Exit Function

    For r = 1 To network.Count

        Set link = network.Item(r)
        succId = link.SuccId
        predId = link.PredId

        If succId <> "" Then
            If predId = "" Then
                errMissingPred(succId) = True
            ElseIf Not rowById.Exists(predId) Then
                errMissingPred(succId) = True
            End If
        End If

    Next r

    If errMissingPred.Count > 0 Then

        CalcBridge_PreCore_MarkErrorIdsInCalc tblCalc, mapCalc, errMissingPred, "Missing predecessor"

        If consoleMessages Is Nothing Then
            CalcBridge_ShowGroupedErrorMessage errMissingPred, idToWbs, _
                "Prédécesseur introuvable dans tbl_LOGIC_LINKS", _
                "vérifier la table des liens logiques", _
                "Missing predecessor in tbl_LOGIC_LINKS", _
                "check the logical links table"
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, errMissingPred, idToWbs, _
                "Prédécesseur introuvable dans tbl_LOGIC_LINKS", _
                "vérifier la table des liens logiques", _
                "Missing predecessor in tbl_LOGIC_LINKS", _
                "check the logical links table"
        End If

        CalcBridge_PreCore_CheckMissingPredecessors = True
    End If

    Exit Function

FailSafe:
    CalcBridge_AddOrShowConsoleMessage consoleMessages, "STOP", _
        "Erreur pendant le contrôle des prédécesseurs." & vbCrLf & _
        "-> calcul arrêté avant écriture WBS.", _
        "Error while checking predecessors." & vbCrLf & _
        "-> calculation stopped before WBS write."

    CalcBridge_PreCore_CheckMissingPredecessors = True

End Function

'------------------------------------------------------------------------------
' FR: Valide Pre Core Check Cycles et applique la politique d'erreur definie par le composant.
' EN: Validates Pre Core Check Cycles and applies the component's defined failure policy.
'------------------------------------------------------------------------------

Public Function CalcBridge_PreCore_CheckCycles( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal validIds As Object, _
    ByVal linksBySuccId As Object, _
    Optional ByVal consoleMessages As Collection) As Boolean

    Dim cycleIds As Object
    Dim idToWbs As Object

    On Error GoTo FailSafe

    CalcBridge_PreCore_CheckCycles = False

    Set cycleIds = CalcBridge_PreCore_DetectCycleIds(validIds, linksBySuccId)

    If cycleIds.Count > 0 Then

        Set idToWbs = CalcBridge_PreCore_BuildIdToWbsFromCalc(tblCalc, mapCalc)

        CalcBridge_PreCore_MarkErrorIdsInCalc tblCalc, mapCalc, cycleIds, "Cycle detected"

        If consoleMessages Is Nothing Then
            CalcBridge_ShowGroupedErrorMessage cycleIds, idToWbs, _
                "Boucle de dépendance détectée", _
                "corriger la colonne Predecessors WBS", _
                "Dependency cycle detected", _
                "fix the Predecessors WBS column"
        Else
            CalcBridge_AddGroupedStopToCollection consoleMessages, cycleIds, idToWbs, _
                "Boucle de dépendance détectée", _
                "corriger la colonne Predecessors WBS", _
                "Dependency cycle detected", _
                "fix the Predecessors WBS column"
        End If

        CalcBridge_PreCore_CheckCycles = True
    End If

    Exit Function

FailSafe:
    CalcBridge_AddOrShowConsoleMessage consoleMessages, "STOP", _
        "Erreur pendant le contrôle des cycles." & vbCrLf & _
        "-> calcul arrêté avant écriture WBS.", _
        "Error while checking dependency cycles." & vbCrLf & _
        "-> calculation stopped before WBS write."

    CalcBridge_PreCore_CheckCycles = True

End Function

'------------------------------------------------------------------------------
' FR: Detecte la collection Pre Core Detect Cycle IDs sans modifier le dataset analyse.
' EN: Detects the Pre Core Detect Cycle IDs collection without mutating the analyzed dataset.
'------------------------------------------------------------------------------

Private Function CalcBridge_PreCore_DetectCycleIds( _
    ByVal validIds As Object, _
    ByVal linksBySuccId As Object) As Object

    Dim childrenByPred As Object
    Dim state As Object
    Dim stackIndex As Object
    Dim stack As Collection
    Dim cycleIds As Object

    Dim idKey As Variant
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String
    Dim colChildren As Collection

    Set childrenByPred = CreateObject("Scripting.Dictionary")
    Set state = CreateObject("Scripting.Dictionary")
    Set stackIndex = CreateObject("Scripting.Dictionary")
    Set stack = New Collection
    Set cycleIds = CreateObject("Scripting.Dictionary")

    For Each idKey In validIds.Keys

        Set colChildren = New Collection
        childrenByPred.Add CStr(idKey), colChildren

        state(CStr(idKey)) = 0

    Next idKey

    If Not linksBySuccId Is Nothing Then

        For Each succId In linksBySuccId.Keys

            If validIds.Exists(CStr(succId)) Then

                For Each oneLink In linksBySuccId(CStr(succId))

                    predId = Core_GetLinkPredId(oneLink)

                    If predId <> "" Then
                        If validIds.Exists(predId) Then

                            If Not childrenByPred.Exists(predId) Then
                                Set colChildren = New Collection
                                childrenByPred.Add predId, colChildren
                            End If

                            childrenByPred(predId).Add CStr(succId)

                        End If
                    End If

                Next oneLink

            End If

        Next succId

    End If

    For Each idKey In validIds.Keys

        If CLng(state(CStr(idKey))) = 0 Then

            CalcBridge_PreCore_DFS_Cycle _
                CStr(idKey), childrenByPred, state, stackIndex, stack, cycleIds

            If cycleIds.Count > 0 Then Exit For

        End If

    Next idKey

    Set CalcBridge_PreCore_DetectCycleIds = cycleIds

End Function

'------------------------------------------------------------------------------
' FR: Traite la collection Pre Core DFS Cycle sans modifier les donnees d'entree.
' EN: Handles the Pre Core DFS Cycle collection without mutating input data.
'------------------------------------------------------------------------------

Private Sub CalcBridge_PreCore_DFS_Cycle( _
    ByVal currentId As String, _
    ByVal childrenByPred As Object, _
    ByVal state As Object, _
    ByVal stackIndex As Object, _
    ByVal stack As Collection, _
    ByVal cycleIds As Object)

    Dim childId As Variant
    Dim i As Long
    Dim startIdx As Long

    If cycleIds.Count > 0 Then Exit Sub

    state(currentId) = 1

    stack.Add currentId
    stackIndex(currentId) = stack.Count

    If childrenByPred.Exists(currentId) Then

        For Each childId In childrenByPred(currentId)

            If cycleIds.Count > 0 Then Exit For

            If Not state.Exists(CStr(childId)) Then
                state(CStr(childId)) = 0
            End If

            If CLng(state(CStr(childId))) = 0 Then

                CalcBridge_PreCore_DFS_Cycle _
                    CStr(childId), childrenByPred, state, stackIndex, stack, cycleIds

            ElseIf CLng(state(CStr(childId))) = 1 Then

                If stackIndex.Exists(CStr(childId)) Then

                    startIdx = CLng(stackIndex(CStr(childId)))

                    For i = startIdx To stack.Count
                        cycleIds(CStr(stack(i))) = True
                    Next i

                    Exit For

                End If

            End If

        Next childId

    End If

    state(currentId) = 2

    If stack.Count > 0 Then
        If CStr(stack(stack.Count)) = currentId Then
            stack.Remove stack.Count
        End If
    End If

    If stackIndex.Exists(currentId) Then
        stackIndex.Remove currentId
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Construit la map Pre Core Build ID To WBS From CALC a partir des donnees fournies par l'appelant.
' EN: Builds the Pre Core Build ID To WBS From CALC map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CalcBridge_PreCore_BuildIdToWbsFromCalc( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object) As Object

    Dim idToWbs As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim taskNameVal As String

    Set idToWbs = CreateObject("Scripting.Dictionary")

    If tblCalc Is Nothing Then
        Set CalcBridge_PreCore_BuildIdToWbsFromCalc = idToWbs
        Exit Function
    End If

    If tblCalc.DataBodyRange Is Nothing Then
        Set CalcBridge_PreCore_BuildIdToWbsFromCalc = idToWbs
        Exit Function
    End If

    If Not mapCalc.Exists("ID") Then
        Set CalcBridge_PreCore_BuildIdToWbsFromCalc = idToWbs
        Exit Function
    End If

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))

        If idVal <> "" Then

            If mapCalc.Exists("WBS") Then
                wbsVal = Trim$(CStr(arr(r, mapCalc("WBS"))))
            Else
                wbsVal = "-"
            End If

            idToWbs(idVal) = wbsVal

        End If

    Next r

    Set CalcBridge_PreCore_BuildIdToWbsFromCalc = idToWbs

End Function

'------------------------------------------------------------------------------
' FR: Marque la collection Pre Core Mark Error IDs In CALC dans la structure cible sans recalculer la condition source.
' EN: Marks the Pre Core Mark Error IDs In CALC collection in the target structure without recalculating the source condition.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub CalcBridge_PreCore_MarkErrorIdsInCalc( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal errorIds As Object, _
    ByVal errMsg As String)

    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Error flag") Then Exit Sub
    If Not mapCalc.Exists("ErrorMsg") Then Exit Sub

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)

        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))

        If idVal <> "" Then
            If errorIds.Exists(idVal) Then
                arr(r, mapCalc("Error flag")) = "ERROR"
                arr(r, mapCalc("ErrorMsg")) = errMsg
            End If
        End If

    Next r

    tblCalc.DataBodyRange.value = arr

End Sub



'------------------------------------------------------------------------------
' FR: Valide Pre Core Check LOE As Predecessor et applique la politique d'erreur definie par le composant.
' EN: Validates Pre Core Check LOE As Predecessor and applies the component's defined failure policy.
'------------------------------------------------------------------------------

Public Function CalcBridge_PreCore_CheckLOEAsPredecessor( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    Optional ByVal consoleMessages As Collection, _
    Optional ByVal sharedLinksBySuccId As Object = Nothing) As Boolean

    Dim arrCalc As Variant
    Dim rowById As Object
    Dim idToWbs As Object
    Dim loeIds As Object
    Dim errIds As Object
    Dim linksBySuccId As Object
    Dim explicitDetails As Collection
    Dim parentDetails As Collection

    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String
    Dim sourceParentId As String
    Dim predRow As Long
    Dim detail As Object
    Dim localMessages As Collection

    On Error GoTo FailSafe

    CalcBridge_PreCore_CheckLOEAsPredecessor = False

    If tblCalc Is Nothing Then Exit Function
    If tblCalc.DataBodyRange Is Nothing Then Exit Function
    If mapCalc Is Nothing Then Exit Function

    If Not mapCalc.Exists("ID") Then Exit Function
    If Not mapCalc.Exists("Task Type") Then Exit Function

    arrCalc = tblCalc.DataBodyRange.value

    Set rowById = Core_BuildRowById(arrCalc, mapCalc)
    Set idToWbs = CalcBridge_PreCore_BuildIdToWbsFromCalc(tblCalc, mapCalc)
    Set loeIds = CreateObject("Scripting.Dictionary")
    Set errIds = CreateObject("Scripting.Dictionary")
    Set explicitDetails = New Collection
    Set parentDetails = New Collection

    For Each succId In rowById.Keys
        predRow = CLng(rowById(CStr(succId)))
        If TaskTypeRules_IsLevelOfEffortRow(arrCalc, mapCalc, predRow) Then
            loeIds(CStr(succId)) = True
        End If
    Next succId

    If loeIds.Count = 0 Then Exit Function

    If sharedLinksBySuccId Is Nothing Then
        Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded(tblCalc)
    Else
        Set linksBySuccId = sharedLinksBySuccId
    End If
    If linksBySuccId Is Nothing Then Exit Function

    For Each succId In linksBySuccId.Keys
        For Each oneLink In linksBySuccId(CStr(succId))
            predId = Core_GetLinkPredId(oneLink)

            If predId <> "" Then
                If loeIds.Exists(predId) Then
                    errIds(predId) = True
                    errIds(CStr(succId)) = True

                    sourceParentId = Core_GetLinkSummarySourceId(oneLink)
                    Set detail = CalcBridge_BuildLOEPredecessorDetail( _
                        CStr(succId), predId, sourceParentId, oneLink, idToWbs)

                    If sourceParentId <> "" Then
                        parentDetails.Add detail
                    Else
                        explicitDetails.Add detail
                    End If
                End If
            End If
        Next oneLink
    Next succId

    If errIds.Count > 0 Then

        CalcBridge_PreCore_MarkErrorIdsInCalc tblCalc, mapCalc, errIds, "LOE cannot be used as predecessor"

        If consoleMessages Is Nothing Then
            Set localMessages = New Collection
        Else
            Set localMessages = consoleMessages
        End If

        If explicitDetails.Count > 0 Then
            CalcBridge_AddConsoleMessage localMessages, "STOP", _
                CalcBridge_BuildLOEExplicitPredecessorMessage(explicitDetails)
        End If

        If parentDetails.Count > 0 Then
            CalcBridge_AddConsoleMessage localMessages, "STOP", _
                CalcBridge_BuildLOEParentPredecessorMessage(parentDetails)
        End If

        If consoleMessages Is Nothing Then CalcBridge_ShowPlanningConsole localMessages

        CalcBridge_PreCore_CheckLOEAsPredecessor = True

    End If

    Exit Function

FailSafe:
    CalcBridge_AddOrShowConsoleMessage consoleMessages, "STOP", _
        "Erreur pendant le controle des LOE utilisees comme predecesseur." & vbCrLf & _
        "-> calcul arrete avant ecriture WBS.", _
        "Error while checking LOE used as predecessor." & vbCrLf & _
        "-> calculation stopped before WBS write."

    CalcBridge_PreCore_CheckLOEAsPredecessor = True

End Function

'------------------------------------------------------------------------------
' FR: Construit la map LOE Predecessor Detail a partir des donnees fournies par l'appelant.
' EN: Builds the LOE Predecessor Detail map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CalcBridge_BuildLOEPredecessorDetail( _
    ByVal succId As String, _
    ByVal loeId As String, _
    ByVal sourceParentId As String, _
    ByRef oneLink As Variant, _
    ByVal idToWbs As Object) As Object

    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d("Succ ID") = succId
    d("Succ WBS") = CalcBridge_GetWbsForId(idToWbs, succId)
    d("LOE ID") = loeId
    d("LOE WBS") = CalcBridge_GetWbsForId(idToWbs, loeId)
    d("Parent ID") = sourceParentId
    d("Parent WBS") = CalcBridge_GetWbsForId(idToWbs, sourceParentId)
    d("Link Type") = Core_GetLinkType(oneLink)
    d("Lag") = Core_GetLinkLag(oneLink)

    Set CalcBridge_BuildLOEPredecessorDetail = d

End Function






