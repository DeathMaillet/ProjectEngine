Attribute VB_Name = "mod_GanttSimulationMessages"
Option Explicit

'===============================================================================
' MODULE : mod_GanttSimulationMessages
' DOMAINE / DOMAIN : Gantt
'
' FR
' Transforme les resultats TEST/SCENARIO/LOCK en messages bilingues sans recalcul metier.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Transforms TEST/SCENARIO/LOCK results into bilingual messages without business recalculation.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : ShowGanttLiveGroupedMessage, GanttLive_IsInheritedCoreError, GanttLive_RemoveDerivedLOERootErrors, GanttLive_CalcGanttTestHasErrors, GanttLive_HasConsoleCollection, GanttLive_AddVbaOrStructuredError, GanttLive_AddBiConsoleMessage
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================



'------------------------------------------------------------------------------
' FR: Publie un message groupe live dans la console existante ou dans une console locale.
' EN: Publishes a grouped live message to the existing console or to a local console.
'------------------------------------------------------------------------------
Public Sub ShowGanttLiveGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String, _
    ByVal boxStyle As VbMsgBoxStyle, _
    Optional ByVal consoleMessages As Variant)

    Dim msgType As String
    Dim localMessages As Collection

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    msgType = GanttLive_MessageTypeFromMsgBoxStyle(boxStyle)

    If GanttLive_HasConsoleCollection(consoleMessages) Then
        CalcBridge_AddConsoleMessage consoleMessages, msgType, _
            BuildGanttLiveGroupedMessage(idsDict, idToWbs, frProblem, frAction, enProblem, enAction)
    Else
        Set localMessages = New Collection
        CalcBridge_AddConsoleMessage localMessages, msgType, _
            BuildGanttLiveGroupedMessage(idsDict, idToWbs, frProblem, frAction, enProblem, enAction)
        CalcBridge_ShowPlanningConsole localMessages
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Construit le texte bilingue d'un message groupe avec IDs et WBS limites.
' EN: Builds bilingual grouped-message text with capped ID and WBS lists.
'------------------------------------------------------------------------------
Private Function BuildGanttLiveGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String) As String

    Dim idsLine As String
    Dim wbsLine As String

    idsLine = BuildInlineList_GanttLive(idsDict, 20)
    wbsLine = BuildInlineWBSList_GanttLive(idsDict, idToWbs, 20)

    BuildGanttLiveGroupedMessage = _
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

End Function

'------------------------------------------------------------------------------
' FR: Formate une liste compacte d'IDs avec limite d'affichage.
' EN: Formats a compact capped ID list.
'------------------------------------------------------------------------------
Private Function BuildInlineList_GanttLive(ByVal idsDict As Object, ByVal maxItems As Long) As String

    Dim result As String
    Dim key As Variant
    Dim countShown As Long
    Dim totalCount As Long

    result = ""
    countShown = 0
    totalCount = idsDict.Count

    For Each key In idsDict.Keys
        countShown = countShown + 1
        If countShown <= maxItems Then
            If result <> "" Then result = result & " / "
            result = result & CStr(key)
        Else
            Exit For
        End If
    Next key

    If totalCount > maxItems Then
        result = result & " / +" & CStr(totalCount - maxItems)
    End If

    BuildInlineList_GanttLive = result

End Function

'------------------------------------------------------------------------------
' FR: Formate une liste compacte de WBS correspondant a une liste d'IDs.
' EN: Formats a compact WBS list corresponding to an ID list.
'------------------------------------------------------------------------------
Private Function BuildInlineWBSList_GanttLive(ByVal idsDict As Object, ByVal idToWbs As Object, ByVal maxItems As Long) As String

    Dim result As String
    Dim key As Variant
    Dim countShown As Long
    Dim totalCount As Long
    Dim itemText As String

    result = ""
    countShown = 0
    totalCount = idsDict.Count

    For Each key In idsDict.Keys
        countShown = countShown + 1
        If countShown <= maxItems Then
            If idToWbs.Exists(CStr(key)) Then
                itemText = NormalizeWBS(CStr(idToWbs(CStr(key))))
            Else
                itemText = "-"
            End If

            If result <> "" Then result = result & " / "
            result = result & itemText
        Else
            Exit For
        End If
    Next key

    If totalCount > maxItems Then
        result = result & " / +" & CStr(totalCount - maxItems)
    End If

    BuildInlineWBSList_GanttLive = result

End Function


'------------------------------------------------------------------------------
' FR: Identifie les erreurs Core de cascade heritees afin de garder seulement les causes racines.
' EN: Identifies inherited Core cascade errors so only root causes are kept.
'------------------------------------------------------------------------------
Public Function GanttLive_IsInheritedCoreError(ByVal errMsg As String) As Boolean

    Dim txt As String

    txt = Trim$(CStr(errMsg))

    GanttLive_IsInheritedCoreError = _
        (InStr(1, txt, "Blocked by predecessor error", vbTextCompare) > 0) Or _
        (InStr(1, txt, "Blocked by predecessor chain", vbTextCompare) > 0)

End Function

'------------------------------------------------------------------------------
' FR: Retire des causes racines les erreurs LOE qui ne font que refleter un predecesseur deja en erreur.
' EN: Removes LOE errors from root causes when they only reflect an already failing predecessor.
'------------------------------------------------------------------------------
Public Sub GanttLive_RemoveDerivedLOERootErrors( _
    ByRef dataCore As Variant, _
    ByVal mapCore As Object, _
    ByVal errorIds As Object, _
    ByVal rootErrorIds As Object)

    Dim r As Long
    Dim idVal As String
    Dim errMsg As String
    Dim removeIds As Object
    Dim oneId As Variant

    If errorIds Is Nothing Then Exit Sub
    If rootErrorIds Is Nothing Then Exit Sub
    If rootErrorIds.Count = 0 Then Exit Sub

    Set removeIds = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(dataCore, 1)
        idVal = Trim$(CStr(dataCore(r, mapCore("ID"))))
        If idVal <> "" Then
            If rootErrorIds.Exists(idVal) Then
                errMsg = Trim$(CStr(dataCore(r, mapCore("ErrorMsg"))))
                If GanttLive_IsDerivedLOEPredecessorError(errMsg, errorIds) Then removeIds(idVal) = True
            End If
        End If
    Next r

    For Each oneId In removeIds.Keys
        If rootErrorIds.Exists(CStr(oneId)) Then rootErrorIds.Remove CStr(oneId)
    Next oneId

End Sub

'------------------------------------------------------------------------------
' FR: Verifie si une erreur LOE pointe vers un predecesseur deja present dans les erreurs Core.
' EN: Checks whether an LOE error points to a predecessor already present in Core errors.
'------------------------------------------------------------------------------
Private Function GanttLive_IsDerivedLOEPredecessorError(ByVal errMsg As String, ByVal errorIds As Object) As Boolean

    Dim predId As String

    predId = GanttLive_ExtractLOEBlockedPredecessorId(errMsg)
    If predId = "" Then Exit Function

    GanttLive_IsDerivedLOEPredecessorError = errorIds.Exists(predId)

End Function

'------------------------------------------------------------------------------
' FR: Extrait l'ID predecesseur depuis un message Core de blocage LOE.
' EN: Extracts the predecessor ID from a Core LOE blocking message.
'------------------------------------------------------------------------------
Private Function GanttLive_ExtractLOEBlockedPredecessorId(ByVal errMsg As String) As String

    Dim txt As String
    Dim marker As String
    Dim pos As Long
    Dim tail As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    txt = Trim$(CStr(errMsg))
    If InStr(1, txt, "LOE blocked by SS predecessor error: ID", vbTextCompare) = 0 And _
       InStr(1, txt, "LOE blocked by FF predecessor error: ID", vbTextCompare) = 0 Then Exit Function

    marker = "error: ID"
    pos = InStr(1, txt, marker, vbTextCompare)
    If pos = 0 Then Exit Function

    tail = Trim$(Mid$(txt, pos + Len(marker)))
    For i = 1 To Len(tail)
        ch = Mid$(tail, i, 1)
        If ch >= "0" And ch <= "9" Then
            result = result & ch
        ElseIf result <> "" Then
            Exit For
        End If
    Next i

    GanttLive_ExtractLOEBlockedPredecessorId = result

End Function

'------------------------------------------------------------------------------
' FR: Publie un message bilingue listant les taches amont/aval en violation.
' EN: Publishes a bilingual message listing upstream/downstream violation tasks.
'------------------------------------------------------------------------------
Private Sub ShowGanttLiveUpstreamViolationMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String, _
    ByVal boxStyle As VbMsgBoxStyle, _
    Optional ByVal consoleMessages As Variant)

    Dim itemsLine As String
    Dim msg As String
    Dim msgType As String
    Dim localMessages As Collection

    If idsDict Is Nothing Then Exit Sub
    If idsDict.Count = 0 Then Exit Sub

    itemsLine = BuildGanttLiveUpstreamViolationItems(idsDict, idToWbs, 20)
    msgType = GanttLive_MessageTypeFromMsgBoxStyle(boxStyle)

    msg = _
        "FR:" & vbCrLf & _
        frProblem & vbCrLf & _
        "-> " & frAction & vbCrLf & vbCrLf & _
        "Tâches : " & itemsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enProblem & vbCrLf & _
        "-> " & enAction & vbCrLf & vbCrLf & _
        "Tasks: " & itemsLine

    If GanttLive_HasConsoleCollection(consoleMessages) Then
        CalcBridge_AddConsoleMessage consoleMessages, msgType, msg
    Else
        Set localMessages = New Collection
        CalcBridge_AddConsoleMessage localMessages, msgType, msg
        CalcBridge_ShowPlanningConsole localMessages
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Formate une liste compacte ID/WBS pour les messages de violation live.
' EN: Formats a compact ID/WBS list for live violation messages.
'------------------------------------------------------------------------------
Private Function BuildGanttLiveUpstreamViolationItems( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal maxItems As Long) As String

    Dim result As String
    Dim key As Variant
    Dim countShown As Long
    Dim totalCount As Long
    Dim wbsVal As String

    result = ""
    countShown = 0
    totalCount = idsDict.Count

    For Each key In idsDict.Keys

        countShown = countShown + 1

        If countShown <= maxItems Then

            If Not idToWbs Is Nothing Then
                If idToWbs.Exists(CStr(key)) Then
                    wbsVal = NormalizeWBS(CStr(idToWbs(CStr(key))))
                Else
                    wbsVal = "-"
                End If
            Else
                wbsVal = "-"
            End If

            If result <> "" Then result = result & " / "
            result = result & CStr(key) & " (" & wbsVal & ")"

        Else
            Exit For
        End If

    Next key

    If totalCount > maxItems Then
        result = result & " / +" & CStr(totalCount - maxItems)
    End If

    BuildGanttLiveUpstreamViolationItems = result

End Function

'------------------------------------------------------------------------------
' FR: Detecte les erreurs presentes dans tbl_CALC_GANTT_TEST avant un lock.
' EN: Detects errors present in tbl_CALC_GANTT_TEST before lock.
'------------------------------------------------------------------------------
Public Function GanttLive_CalcGanttTestHasErrors(ByVal tblTest As ListObject) As Boolean

    Dim mapTest As Object
    Dim arr As Variant
    Dim r As Long

    On Error GoTo SafeExit

    GanttLive_CalcGanttTestHasErrors = False

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then Exit Function

    Set mapTest = CanonicalIdentity_BuildColumnMap(tblTest)

    If Not mapTest.Exists("Error Flag") Then Exit Function

    arr = tblTest.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If UCase$(Trim$(CStr(arr(r, mapTest("Error Flag"))))) = "ERROR" Then
            GanttLive_CalcGanttTestHasErrors = True
            Exit Function
        End If
    Next r

    Exit Function

SafeExit:
    GanttLive_CalcGanttTestHasErrors = True

End Function

'------------------------------------------------------------------------------
' FR: Verifie si un parametre optionnel contient une collection console utilisable.
' EN: Checks whether an optional argument contains a usable console collection.
'------------------------------------------------------------------------------
Public Function GanttLive_HasConsoleCollection(Optional ByVal consoleMessages As Variant) As Boolean

    On Error GoTo SafeExit

    If IsMissing(consoleMessages) Then Exit Function
    If IsObject(consoleMessages) Then
        If Not consoleMessages Is Nothing Then
            GanttLive_HasConsoleCollection = True
        End If
    End If

SafeExit:
End Function

'------------------------------------------------------------------------------
' FR: Convertit un style MsgBox en type de message console STOP/WARNING/INFO.
' EN: Converts a MsgBox style into a STOP/WARNING/INFO console message type.
'------------------------------------------------------------------------------
Private Function GanttLive_MessageTypeFromMsgBoxStyle(ByVal boxStyle As VbMsgBoxStyle) As String

    If (boxStyle And vbCritical) = vbCritical Then
        GanttLive_MessageTypeFromMsgBoxStyle = "STOP"
    ElseIf (boxStyle And vbExclamation) = vbExclamation Then
        GanttLive_MessageTypeFromMsgBoxStyle = "WARNING"
    Else
        GanttLive_MessageTypeFromMsgBoxStyle = "INFO"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Detecte un message deja structure en blocs FR/EN.
' EN: Detects a message already structured as FR/EN blocks.
'------------------------------------------------------------------------------
Private Function GanttLive_IsStructuredBiMessage(ByVal msgText As String) As Boolean

    Dim txt As String

    txt = LTrim$(CStr(msgText))

    GanttLive_IsStructuredBiMessage = _
        (Left$(txt, 3) = "FR:" And InStr(1, txt, "EN:", vbTextCompare) > 0)

End Function

'------------------------------------------------------------------------------
' FR: Ajoute a la console une erreur VBA, en preservant les messages FR/EN deja structures.
' EN: Adds a VBA error to the console while preserving already structured FR/EN messages.
'------------------------------------------------------------------------------
Public Sub GanttLive_AddVbaOrStructuredError( _
    ByVal consoleMessages As Collection, _
    ByVal functionName As String, _
    ByVal errDescription As String)

    If consoleMessages Is Nothing Then Exit Sub

    If GanttLive_IsStructuredBiMessage(errDescription) Then
        CalcBridge_AddConsoleMessage consoleMessages, "STOP", Trim$(CStr(errDescription))
    Else
        GanttLive_AddBiConsoleMessage consoleMessages, "STOP", _
            "Erreur VBA dans " & functionName & vbCrLf & _
            "-> " & errDescription, _
            "VBA error in " & functionName & vbCrLf & _
            "-> " & errDescription
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute un message console bilingue FR/EN avec le type fourni.
' EN: Adds a bilingual FR/EN console message with the supplied type.
'------------------------------------------------------------------------------
Public Sub GanttLive_AddBiConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim msg As String

    If consoleMessages Is Nothing Then Exit Sub

    msg = _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

    CalcBridge_AddConsoleMessage consoleMessages, msgType, msg

End Sub

