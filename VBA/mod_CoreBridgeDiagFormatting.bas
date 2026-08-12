Attribute VB_Name = "mod_CoreBridgeDiagFormatting"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgeDiagFormatting
' DOMAINE / DOMAIN : Core Bridge
'
' FR
' Construit les textes bilingues des diagnostics CoreBridge sans lire Excel ni recalculer le planning.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Builds bilingual CoreBridge diagnostic text without reading Excel or recalculating planning.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : CalcBridge_BuildGroupedMessage, CalcBridge_BuildInlineList, CalcBridge_BuildInlineWBSList, CalcBridge_BuildUpstreamViolationItems, CalcBridge_CycleDetailMessageFromCoreError, CalcBridge_ToPMConstraintMessage, CalcBridge_BuildMissingBaselineRexMessage, CalcBridge_BuildLOEExplicitPredecessorMessage
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Construit un message bilingue groupe pour des IDs/WBS.
' EN: Builds a bilingual grouped message for IDs/WBS values.
'------------------------------------------------------------------------------
Public Function CalcBridge_BuildGroupedMessage( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal frProblem As String, _
    ByVal frAction As String, _
    ByVal enProblem As String, _
    ByVal enAction As String) As String

    Dim idsLine As String
    Dim wbsLine As String

    idsLine = CalcBridge_BuildInlineList(idsDict, 20)
    wbsLine = CalcBridge_BuildInlineWBSList(idsDict, idToWbs, 20)

    CalcBridge_BuildGroupedMessage = _
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
' FR: Formate une liste compacte d'IDs pour un diagnostic groupe.
' EN: Formats a compact ID list for a grouped diagnostic.
'------------------------------------------------------------------------------
Public Function CalcBridge_BuildInlineList( _
    ByVal idsDict As Object, _
    ByVal maxItems As Long) As String

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

    CalcBridge_BuildInlineList = result

End Function

'------------------------------------------------------------------------------
' FR: Formate une liste compacte de WBS correspondant a des IDs.
' EN: Formats a compact WBS list matching IDs.
'------------------------------------------------------------------------------
Public Function CalcBridge_BuildInlineWBSList( _
    ByVal idsDict As Object, _
    ByVal idToWbs As Object, _
    ByVal maxItems As Long) As String

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

            If Not idToWbs Is Nothing Then
                If idToWbs.Exists(CStr(key)) Then
                    itemText = CStr(idToWbs(CStr(key)))
                Else
                    itemText = "-"
                End If
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

    CalcBridge_BuildInlineWBSList = result

End Function

'------------------------------------------------------------------------------
' FR: Formate les couples ID/WBS des violations amont.
' EN: Formats ID/WBS pairs for upstream violations.
'------------------------------------------------------------------------------
Public Function CalcBridge_BuildUpstreamViolationItems( _
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
                    wbsVal = CStr(idToWbs(CStr(key)))
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

    CalcBridge_BuildUpstreamViolationItems = result

End Function

'------------------------------------------------------------------------------
' FR: Lit une valeur texte dans un dictionnaire diagnostic.
' EN: Reads a text value from a diagnostic dictionary.
'------------------------------------------------------------------------------
Private Function CalcBridge_DiagString(ByVal diag As Object, ByVal key As String) As String

    If diag Is Nothing Then Exit Function
    If diag.Exists(key) Then CalcBridge_DiagString = Trim$(CStr(diag(key)))

End Function

'------------------------------------------------------------------------------
' FR: Lit une valeur brute dans un dictionnaire diagnostic.
' EN: Reads a raw value from a diagnostic dictionary.
'------------------------------------------------------------------------------
Private Function CalcBridge_DiagValue(ByVal diag As Object, ByVal key As String) As Variant

    If diag Is Nothing Then Exit Function
    If diag.Exists(key) Then CalcBridge_DiagValue = diag(key)

End Function

'------------------------------------------------------------------------------
' FR: Formate une valeur diagnostic quelconque en texte affichable.
' EN: Formats any diagnostic value as display text.
'------------------------------------------------------------------------------
Private Function CalcBridge_DiagnosticAnyValueText(ByVal value As Variant) As String

    If Not HasValue(value) Then
        CalcBridge_DiagnosticAnyValueText = "-"
    ElseIf IsDate(value) Then
        CalcBridge_DiagnosticAnyValueText = Format$(CDate(value), "dd/mm/yyyy")
    Else
        CalcBridge_DiagnosticAnyValueText = CStr(value)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Formate le libelle lisible d'une tache dans un diagnostic.
' EN: Formats the readable task label used in diagnostics.
'------------------------------------------------------------------------------
Private Function CalcBridge_DiagnosticTaskLabel( _
    ByVal taskId As String, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object) As String

    Dim wbsVal As String
    Dim nameVal As String

    If Not idToWbs Is Nothing Then
        If idToWbs.Exists(taskId) Then wbsVal = Trim$(CStr(idToWbs(taskId)))
    End If

    If Not idToTaskName Is Nothing Then
        If idToTaskName.Exists(taskId) Then nameVal = Trim$(CStr(idToTaskName(taskId)))
    End If

    If wbsVal <> "" And nameVal <> "" Then
        CalcBridge_DiagnosticTaskLabel = wbsVal & " " & nameVal
    ElseIf wbsVal <> "" Then
        CalcBridge_DiagnosticTaskLabel = wbsVal
    ElseIf nameVal <> "" Then
        CalcBridge_DiagnosticTaskLabel = nameVal
    Else
        CalcBridge_DiagnosticTaskLabel = "ID " & taskId
    End If

End Function

'------------------------------------------------------------------------------
' FR: Formate une date de diagnostic ou un tiret si absente.
' EN: Formats a diagnostic date or a dash when absent.
'------------------------------------------------------------------------------
Private Function CalcBridge_DiagnosticDateText(ByVal dateValue As Variant) As String

    If HasValue(dateValue) Then
        CalcBridge_DiagnosticDateText = Format$(CDate(dateValue), "dd/mm/yyyy")
    Else
        CalcBridge_DiagnosticDateText = "-"
    End If

End Function

'------------------------------------------------------------------------------
' FR: Formate un lag diagnostic avec signe explicite.
' EN: Formats a diagnostic lag with an explicit sign.
'------------------------------------------------------------------------------
Private Function CalcBridge_FormatDiagnosticLag(ByVal lagValue As Double) As String

    Dim lagText As String

    lagText = Replace$(Format$(Abs(lagValue), "0.##"), ",", ".")

    If lagValue >= 0# Then
        CalcBridge_FormatDiagnosticLag = "+" & lagText
    Else
        CalcBridge_FormatDiagnosticLag = "-" & lagText
    End If

End Function


'------------------------------------------------------------------------------
' FR: Retourne la valeur Cycle Detail Message From Core Error sans modifier les donnees d'entree.
' EN: Returns the Cycle Detail Message From Core Error value without mutating input data.
'------------------------------------------------------------------------------

Public Function CalcBridge_CycleDetailMessageFromCoreError(ByVal errMsg As String) As String

    Dim markerPos As Long
    Dim msg As String

    msg = CStr(errMsg)
    markerPos = InStr(1, msg, "FR:", vbTextCompare)

    If markerPos > 0 Then
        CalcBridge_CycleDetailMessageFromCoreError = Trim$(Mid$(msg, markerPos))
    Else
        CalcBridge_CycleDetailMessageFromCoreError = ""
    End If

End Function


'------------------------------------------------------------------------------
' FR: Retourne la valeur To PM Constraint Message sans modifier les donnees d'entree.
' EN: Returns the To PM Constraint Message value without mutating input data.
'------------------------------------------------------------------------------

Public Function CalcBridge_ToPMConstraintMessage(ByVal rawMessage As String) As String

    Dim msg As String

    msg = CStr(rawMessage)

    CalcBridge_ReplacePMMessage msg, _
        "Type de contrainte debut non reconnu", _
        "Type de contrainte d" & ChrW$(233) & "but invalide", _
        "choisir une contrainte d" & ChrW$(233) & "but support" & ChrW$(233) & "e"
    CalcBridge_ReplacePMMessage msg, _
        "Unknown start constraint type", _
        "Invalid start constraint type", _
        "choose a supported start constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Type de contrainte fin non reconnu", _
        "Type de contrainte fin invalide", _
        "choisir une contrainte fin support" & ChrW$(233) & "e"
    CalcBridge_ReplacePMMessage msg, _
        "Unknown finish constraint type", _
        "Invalid finish constraint type", _
        "choose a supported finish constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Actual Start avant contrainte debut", _
        "Actual Start incompatible avec Start No Earlier Than", _
        "repousser Actual Start ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Actual Start is before start constraint", _
        "Actual Start is incompatible with Start No Earlier Than", _
        "move Actual Start later or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Forecast Start avant contrainte debut", _
        "Forecast Start incompatible avec Start No Earlier Than", _
        "repousser Forecast Start ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Forecast Start is before start constraint", _
        "Forecast Start is incompatible with Start No Earlier Than", _
        "move Forecast Start later or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Actual Start apres contrainte debut max", _
        "Actual Start incompatible avec Start No Later Than", _
        "aligner Actual Start ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Actual Start is after latest start constraint", _
        "Actual Start is incompatible with Start No Later Than", _
        "align Actual Start or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Forecast Start apres contrainte debut max", _
        "Forecast Start incompatible avec Start No Later Than", _
        "aligner Forecast Start ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Forecast Start is after latest start constraint", _
        "Forecast Start is incompatible with Start No Later Than", _
        "align Forecast Start or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Calculated Start apres contrainte debut max", _
        "Start No Later Than impossible " & ChrW$(224) & " respecter", _
        "corriger la logique, la dur" & ChrW$(233) & "e ou la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Calculated Start is after latest start constraint", _
        "Start No Later Than cannot be met", _
        "fix logic, duration, or the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Actual Start different de contrainte Must Start On", _
        "Actual Start incompatible avec Must Start On", _
        "aligner Actual Start ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Actual Start differs from Must Start On constraint", _
        "Actual Start is incompatible with Must Start On", _
        "align Actual Start or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Forecast Start different de contrainte Must Start On", _
        "Forecast Start incompatible avec Must Start On", _
        "aligner Forecast Start ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Forecast Start differs from Must Start On constraint", _
        "Forecast Start is incompatible with Must Start On", _
        "align Forecast Start or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Calculated Start different de contrainte Must Start On", _
        "Must Start On impossible " & ChrW$(224) & " respecter", _
        "corriger la logique amont, la dur" & ChrW$(233) & "e ou la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Calculated Start differs from Must Start On constraint", _
        "Must Start On cannot be met", _
        "fix upstream logic, duration, or the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Actual Finish avant contrainte fin", _
        "Actual Finish incompatible avec Finish No Earlier Than", _
        "repousser Actual Finish ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Actual Finish is before finish constraint", _
        "Actual Finish is incompatible with Finish No Earlier Than", _
        "move Actual Finish later or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Forecast Finish avant contrainte fin amont", _
        "Forecast Finish incompatible avec les contraintes de fin amont", _
        "repousser Forecast Finish ou corriger la logique amont"
    CalcBridge_ReplacePMMessage msg, _
        "Forecast Finish is before upstream finish constraint", _
        "Forecast Finish is incompatible with upstream finish constraints", _
        "move Forecast Finish later or fix upstream logic"

    CalcBridge_ReplacePMMessage msg, _
        "Forecast Finish avant contrainte fin", _
        "Forecast Finish incompatible avec Finish No Earlier Than", _
        "repousser Forecast Finish ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Forecast Finish is before finish constraint", _
        "Forecast Finish is incompatible with Finish No Earlier Than", _
        "move Forecast Finish later or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Actual Finish apres contrainte fin max", _
        "Actual Finish incompatible avec Finish No Later Than", _
        "aligner Actual Finish ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Actual Finish is after latest finish constraint", _
        "Actual Finish is incompatible with Finish No Later Than", _
        "align Actual Finish or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Forecast Finish apres contrainte fin max", _
        "Forecast Finish incompatible avec Finish No Later Than", _
        "avancer Forecast Finish, r" & ChrW$(233) & "duire la dur" & ChrW$(233) & "e ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Forecast Finish is after latest finish constraint", _
        "Forecast Finish is incompatible with Finish No Later Than", _
        "move Forecast Finish earlier, reduce duration, or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Calculated Finish apres contrainte fin max", _
        "Finish No Later Than impossible " & ChrW$(224) & " respecter", _
        "corriger la logique, la dur" & ChrW$(233) & "e ou la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Calculated Finish is after latest finish constraint", _
        "Finish No Later Than cannot be met", _
        "fix logic, duration, or the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Actual Finish different de contrainte Must Finish On", _
        "Actual Finish incompatible avec Must Finish On", _
        "aligner Actual Finish ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Actual Finish differs from Must Finish On constraint", _
        "Actual Finish is incompatible with Must Finish On", _
        "align Actual Finish or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Forecast Finish different de contrainte Must Finish On", _
        "Forecast Finish incompatible avec Must Finish On", _
        "aligner Forecast Finish ou modifier la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Forecast Finish differs from Must Finish On constraint", _
        "Forecast Finish is incompatible with Must Finish On", _
        "align Forecast Finish or update the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Calculated Finish different de contrainte Must Finish On", _
        "Must Finish On impossible " & ChrW$(224) & " respecter", _
        "corriger la logique amont, la dur" & ChrW$(233) & "e ou la contrainte"
    CalcBridge_ReplacePMMessage msg, _
        "Calculated Finish differs from Must Finish On constraint", _
        "Must Finish On cannot be met", _
        "fix upstream logic, duration, or the constraint"

    CalcBridge_ReplacePMMessage msg, _
        "Calculated Start avant reseau impose par contrainte Must Finish On", _
        "Must Finish On incompatible avec la logique amont", _
        "le r" & ChrW$(233) & "seau impose un d" & ChrW$(233) & "marrage trop tardif pour respecter la fin impos" & ChrW$(233) & "e"
    CalcBridge_ReplacePMMessage msg, _
        "Calculated Start is before network allowed start due to Must Finish On constraint", _
        "Must Finish On is incompatible with upstream logic", _
        "the network forces a start too late to meet the imposed finish"

    CalcBridge_ReplacePMMessage msg, _
        "Actual Start different du debut impose par Must Finish On", _
        "Actual Start incompatible avec Must Finish On", _
        "la date saisie ne permet pas de respecter la fin impos" & ChrW$(233) & "e avec la dur" & ChrW$(233) & "e actuelle"
    CalcBridge_ReplacePMMessage msg, _
        "Actual Start differs from start implied by Must Finish On constraint", _
        "Actual Start is incompatible with Must Finish On", _
        "the entered date cannot meet the imposed finish with the current duration"

    CalcBridge_ReplacePMMessage msg, _
        "Forecast Start different du debut impose par Must Finish On", _
        "Forecast Start incompatible avec Must Finish On", _
        "la date saisie ne permet pas de respecter la fin impos" & ChrW$(233) & "e avec la dur" & ChrW$(233) & "e actuelle"
    CalcBridge_ReplacePMMessage msg, _
        "Forecast Start differs from start implied by Must Finish On constraint", _
        "Forecast Start is incompatible with Must Finish On", _
        "the entered date cannot meet the imposed finish with the current duration"

    CalcBridge_ReplacePMMessage msg, _
        "Duree incompatible avec contraintes Must Start On / Must Finish On", _
        "Dur" & ChrW$(233) & "e incompatible avec Must Start On et Must Finish On", _
        "corriger la dur" & ChrW$(233) & "e ou l'une des deux contraintes"
    CalcBridge_ReplacePMMessage msg, _
        "Duration is incompatible with Must Start On / Must Finish On constraints", _
        "Duration is incompatible with Must Start On and Must Finish On", _
        "fix duration or one of the two constraints"

    CalcBridge_ToPMConstraintMessage = msg

End Function


'------------------------------------------------------------------------------
' FR: Transforme la valeur Replace PM Message sans modifier la semantique du message source.
' EN: Transforms the Replace PM Message value without changing source-message semantics.
'------------------------------------------------------------------------------

Private Sub CalcBridge_ReplacePMMessage( _
    ByRef msg As String, _
    ByVal oldProblem As String, _
    ByVal newProblem As String, _
    ByVal actionText As String)

    msg = Replace( _
        msg, _
        oldProblem, _
        newProblem & vbCrLf & "-> " & actionText, _
        1, _
        -1, _
        vbTextCompare)

End Sub


'------------------------------------------------------------------------------
' FR: Construit la map Missing Baseline Rex Message a partir des donnees fournies par l'appelant.
' EN: Builds the Missing Baseline Rex Message map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildMissingBaselineRexMessage( _
    ByVal missingIds As Object, _
    ByVal idToWbs As Object) As String

    Dim idsLine As String
    Dim wbsLine As String

    idsLine = CalcBridge_BuildInlineList(missingIds, 20)
    wbsLine = CalcBridge_BuildInlineWBSList(missingIds, idToWbs, 20)

    CalcBridge_BuildMissingBaselineRexMessage = _
        "FR:" & vbCrLf & _
        "Analytics REX partiellement non calculées : Baseline Duration manquante sur au moins une tâche feuille." & vbCrLf & _
        "-> compléter Baseline Duration sur les lignes listées ci-dessous." & vbCrLf & vbCrLf & _
        "IDs : " & idsLine & vbCrLf & _
        "WBS : " & wbsLine & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        "REX analytics partially not calculated: Baseline Duration is missing on at least one leaf task." & vbCrLf & _
        "-> fill Baseline Duration on the lines listed below." & vbCrLf & vbCrLf & _
        "IDs: " & idsLine & vbCrLf & _
        "WBS: " & wbsLine

End Function


'------------------------------------------------------------------------------
' FR: Construit la collection LOE Explicit Predecessor Message a partir des donnees fournies par l'appelant.
' EN: Builds the LOE Explicit Predecessor Message collection from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildLOEExplicitPredecessorMessage(ByVal details As Collection) As String

    Dim detail As Object
    Dim blocksFR As String
    Dim blocksEN As String

    For Each detail In details
        CalcBridge_AppendTextBlock blocksFR, _
            "La tache " & CStr(detail("Succ WBS")) & " reference directement la LOE " & CStr(detail("LOE WBS")) & "." & vbCrLf & vbCrLf & _
            "Details :" & vbCrLf & vbCrLf & _
            "Successeur :" & vbCrLf & CStr(detail("Succ WBS")) & " (ID " & CStr(detail("Succ ID")) & ")" & vbCrLf & vbCrLf & _
            "LOE detectee :" & vbCrLf & CStr(detail("LOE WBS")) & " (ID " & CStr(detail("LOE ID")) & ")" & vbCrLf & vbCrLf & _
            "Lien :" & vbCrLf & CalcBridge_FormatLOELinkLabel(detail)

        CalcBridge_AppendTextBlock blocksEN, _
            "Task " & CStr(detail("Succ WBS")) & " directly references LOE " & CStr(detail("LOE WBS")) & "." & vbCrLf & vbCrLf & _
            "Details:" & vbCrLf & vbCrLf & _
            "Successor:" & vbCrLf & CStr(detail("Succ WBS")) & " (ID " & CStr(detail("Succ ID")) & ")" & vbCrLf & vbCrLf & _
            "Detected LOE:" & vbCrLf & CStr(detail("LOE WBS")) & " (ID " & CStr(detail("LOE ID")) & ")" & vbCrLf & vbCrLf & _
            "Link:" & vbCrLf & CalcBridge_FormatLOELinkLabel(detail)
    Next detail

    CalcBridge_BuildLOEExplicitPredecessorMessage = _
        "FR:" & vbCrLf & _
        "LOE utilisee comme predecesseur" & vbCrLf & vbCrLf & _
        blocksFR & vbCrLf & vbCrLf & _
        "Une LOE est pilotee par le reseau mais ne doit pas piloter d'autres taches." & vbCrLf & vbCrLf & _
        "-> remplacer la LOE par une vraie tache feuille ou une milestone." & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        "LOE used as predecessor" & vbCrLf & vbCrLf & _
        blocksEN & vbCrLf & vbCrLf & _
        "A LOE is driven by the network but must not drive other tasks." & vbCrLf & vbCrLf & _
        "-> replace the LOE with a real leaf task or milestone."

End Function


'------------------------------------------------------------------------------
' FR: Construit la collection LOE Parent Predecessor Message a partir des donnees fournies par l'appelant.
' EN: Builds the LOE Parent Predecessor Message collection from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildLOEParentPredecessorMessage(ByVal details As Collection) As String

    Dim detail As Object
    Dim blocksFR As String
    Dim blocksEN As String

    For Each detail In details
        CalcBridge_AppendTextBlock blocksFR, _
            "La tache " & CStr(detail("Succ WBS")) & " reference le parent " & CStr(detail("Parent WBS")) & "." & vbCrLf & _
            "Ce parent contient la LOE " & CStr(detail("LOE WBS")) & "." & vbCrLf & vbCrLf & _
            "Lors de l'expansion du lien parent, la LOE devient predecesseur indirect." & vbCrLf & vbCrLf & _
            "Details :" & vbCrLf & vbCrLf & _
            "Successeur :" & vbCrLf & CStr(detail("Succ WBS")) & " (ID " & CStr(detail("Succ ID")) & ")" & vbCrLf & vbCrLf & _
            "Predecesseur saisi :" & vbCrLf & CStr(detail("Parent WBS")) & " (ID " & CStr(detail("Parent ID")) & ")" & vbCrLf & vbCrLf & _
            "LOE detectee :" & vbCrLf & CStr(detail("LOE WBS")) & " (ID " & CStr(detail("LOE ID")) & ")" & vbCrLf & vbCrLf & _
            "Lien :" & vbCrLf & CalcBridge_FormatLOELinkLabel(detail)

        CalcBridge_AppendTextBlock blocksEN, _
            "Task " & CStr(detail("Succ WBS")) & " references parent " & CStr(detail("Parent WBS")) & "." & vbCrLf & _
            "This parent contains LOE " & CStr(detail("LOE WBS")) & "." & vbCrLf & vbCrLf & _
            "When the parent link is expanded, the LOE becomes an indirect predecessor." & vbCrLf & vbCrLf & _
            "Details:" & vbCrLf & vbCrLf & _
            "Successor:" & vbCrLf & CStr(detail("Succ WBS")) & " (ID " & CStr(detail("Succ ID")) & ")" & vbCrLf & vbCrLf & _
            "Entered predecessor:" & vbCrLf & CStr(detail("Parent WBS")) & " (ID " & CStr(detail("Parent ID")) & ")" & vbCrLf & vbCrLf & _
            "Detected LOE:" & vbCrLf & CStr(detail("LOE WBS")) & " (ID " & CStr(detail("LOE ID")) & ")" & vbCrLf & vbCrLf & _
            "Link:" & vbCrLf & CalcBridge_FormatLOELinkLabel(detail)
    Next detail

    CalcBridge_BuildLOEParentPredecessorMessage = _
        "FR:" & vbCrLf & _
        "LOE utilisee comme predecesseur via un lien parent" & vbCrLf & vbCrLf & _
        blocksFR & vbCrLf & vbCrLf & _
        "-> remplacer le parent par une tache feuille ou une milestone de fin." & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        "LOE used as predecessor through a parent link" & vbCrLf & vbCrLf & _
        blocksEN & vbCrLf & vbCrLf & _
        "-> replace the parent with a leaf task or finish milestone."

End Function


'------------------------------------------------------------------------------
' FR: Ajoute la valeur Text Block a la structure cible fournie par l'appelant.
' EN: Adds the Text Block value to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub CalcBridge_AppendTextBlock( _
    ByRef target As String, _
    ByVal blockText As String)

    If target <> "" Then target = target & vbCrLf & vbCrLf
    target = target & blockText

End Sub


'------------------------------------------------------------------------------
' FR: Retourne la map Wbs For ID sans exposer de mutateur sur l'etat source.
' EN: Returns the Wbs For ID map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function CalcBridge_GetWbsForId( _
    ByVal idToWbs As Object, _
    ByVal idVal As String) As String

    idVal = Trim$(CStr(idVal))

    If idVal <> "" Then
        If Not idToWbs Is Nothing Then
            If idToWbs.Exists(idVal) Then
                CalcBridge_GetWbsForId = CStr(idToWbs(idVal))
                Exit Function
            End If
        End If
    End If

    CalcBridge_GetWbsForId = idVal

End Function


'------------------------------------------------------------------------------
' FR: Normalise ou formate Format LOE Link Label selon le contrat canonique du composant.
' EN: Normalizes or formats Format LOE Link Label according to the component contract.
'------------------------------------------------------------------------------

Private Function CalcBridge_FormatLOELinkLabel(ByVal detail As Object) As String

    Dim lagVal As Double
    Dim signText As String

    lagVal = CDbl(detail("Lag"))

    If lagVal >= 0 Then
        signText = "+"
    Else
        signText = ""
    End If

    CalcBridge_FormatLOELinkLabel = CStr(detail("Link Type")) & signText & CStr(CLng(lagVal))

End Function


'------------------------------------------------------------------------------
' FR: Construit la collection Warning Ack Token List a partir des donnees fournies par l'appelant.
' EN: Builds the Warning Ack Token List collection from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildWarningAckTokenList( _
    ByVal idsDict As Object, _
    ByVal ackTokensById As Object) As String

    Dim key As Variant
    Dim tokenText As String

    If idsDict Is Nothing Then Exit Function
    If ackTokensById Is Nothing Then Exit Function

    For Each key In idsDict.Keys
        If ackTokensById.Exists(CStr(key)) Then
            If Trim$(CStr(ackTokensById(CStr(key)))) <> "" Then
                If tokenText <> "" Then tokenText = tokenText & ";"
                tokenText = tokenText & Trim$(CStr(ackTokensById(CStr(key))))
            End If
        End If
    Next key

    CalcBridge_BuildWarningAckTokenList = tokenText

End Function



'------------------------------------------------------------------------------
' FR: Construit la map Constraint Diagnostic Message a partir des donnees fournies par l'appelant.
' EN: Builds the Constraint Diagnostic Message map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildConstraintDiagnosticMessage( _
    ByVal diag As Object, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object, _
    ByVal cascadeDiagnostics As Object, _
    ByVal contextKey As String) As String

    Dim taskId As String
    Dim taskLabel As String
    Dim constraintType As String
    Dim checkedField As String
    Dim relationText As String
    Dim cascadeText As String

    If diag Is Nothing Then Exit Function
    If Not diag.Exists("TaskID") Then Exit Function

    taskId = CStr(diag("TaskID"))
    taskLabel = CalcBridge_DiagnosticTaskLabel(taskId, idToWbs, idToTaskName)
    constraintType = CalcBridge_DiagString(diag, "ConstraintType")
    checkedField = CalcBridge_DiagString(diag, "CheckedField")
    relationText = CalcBridge_DiagString(diag, "ExpectedOperator")
    cascadeText = CalcBridge_BuildCascadeRootCauseText(taskId, cascadeDiagnostics, idToWbs, idToTaskName)

    If constraintType = "" Then constraintType = CalcBridge_DiagString(diag, "ConstraintSide") & " constraint"
    If checkedField = "" Then checkedField = "Calculated date"
    If relationText = "" Then relationText = "="

    CalcBridge_BuildConstraintDiagnosticMessage = _
        "FR:" & vbCrLf & _
        "Contrainte impossible a respecter." & vbCrLf & vbCrLf & _
        "Tache : " & taskLabel & vbCrLf & _
        "Contrainte : " & constraintType & vbCrLf & _
        "Date contrainte : " & CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "ConstraintDate")) & vbCrLf & _
        checkedField & " : " & CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CheckedValue")) & vbCrLf & _
        "Valeur attendue : " & relationText & " " & CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "AllowedValue")) & vbCrLf & _
        "Calculated Start : " & CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CalculatedStart")) & vbCrLf & _
        "Calculated Finish : " & CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CalculatedFinish")) & _
        cascadeText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        "Constraint cannot be met." & vbCrLf & vbCrLf & _
        "Task: " & taskLabel & vbCrLf & _
        "Constraint: " & constraintType & vbCrLf & _
        "Constraint date: " & CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "ConstraintDate")) & vbCrLf & _
        checkedField & ": " & CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CheckedValue")) & vbCrLf & _
        "Expected value: " & relationText & " " & CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "AllowedValue")) & vbCrLf & _
        "Calculated Start: " & CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CalculatedStart")) & vbCrLf & _
        "Calculated Finish: " & CalcBridge_DiagnosticAnyValueText(CalcBridge_DiagValue(diag, "CalculatedFinish"))

End Function



'------------------------------------------------------------------------------
' FR: Construit la map Cascade Root Cause Text a partir des donnees fournies par l'appelant.
' EN: Builds the Cascade Root Cause Text map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CalcBridge_BuildCascadeRootCauseText( _
    ByVal taskId As String, _
    ByVal cascadeDiagnostics As Object, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object) As String

    Dim cascadeDiag As Object
    Dim parentId As String
    Dim rootId As String

    If cascadeDiagnostics Is Nothing Then Exit Function
    If Not cascadeDiagnostics.Exists(CStr(taskId)) Then Exit Function
    If Not IsObject(cascadeDiagnostics(CStr(taskId))) Then Exit Function

    Set cascadeDiag = cascadeDiagnostics(CStr(taskId))
    parentId = CalcBridge_DiagString(cascadeDiag, "ParentPropagatedFrom")
    rootId = CalcBridge_DiagString(cascadeDiag, "RootErrorID")

    If parentId = "" And rootId = "" Then Exit Function

    CalcBridge_BuildCascadeRootCauseText = _
        vbCrLf & vbCrLf & _
        "Propagation :" & vbCrLf & _
        "Bloque par : " & CalcBridge_DiagnosticTaskLabel(parentId, idToWbs, idToTaskName) & vbCrLf & _
        "Cause racine : " & CalcBridge_DiagnosticTaskLabel(rootId, idToWbs, idToTaskName)

End Function


'------------------------------------------------------------------------------
' FR: Construit la map Forecast Start Dependency Diagnostic Message a partir des donnees fournies par l'appelant.
' EN: Builds the Forecast Start Dependency Diagnostic Message map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function CalcBridge_BuildForecastStartDependencyDiagnosticMessage( _
    ByVal diag As Object, _
    ByVal idToWbs As Object, _
    ByVal idToTaskName As Object, _
    ByVal contextKey As String) As String

    Dim taskId As String
    Dim predId As String
    Dim taskLabel As String
    Dim predLabel As String
    Dim linkText As String
    Dim predDateKindFr As String
    Dim predDateKindEn As String
    Dim requestedLabelFr As String
    Dim requestedLabelEn As String

    If diag Is Nothing Then Exit Function
    If Not diag.Exists("TaskID") Then Exit Function
    If Not diag.Exists("BlockingPredecessorID") Then Exit Function

    taskId = CStr(diag("TaskID"))
    predId = CStr(diag("BlockingPredecessorID"))
    taskLabel = CalcBridge_DiagnosticTaskLabel(taskId, idToWbs, idToTaskName)
    predLabel = CalcBridge_DiagnosticTaskLabel(predId, idToWbs, idToTaskName)
    linkText = CStr(diag("BlockingLinkType")) & " " & CalcBridge_FormatDiagnosticLag(CDbl(diag("BlockingLag")))

    If UCase$(Trim$(CStr(diag("BlockingPredecessorDateKind")))) = "START" Then
        predDateKindFr = "Debut predecesseur"
        predDateKindEn = "Predecessor start"
    Else
        predDateKindFr = "Fin predecesseur"
        predDateKindEn = "Predecessor finish"
    End If

    If UCase$(Trim$(contextKey)) = "SCENARIO" Then
        requestedLabelFr = "Scenario Start demande"
        requestedLabelEn = "Requested Scenario Start"
    Else
        requestedLabelFr = "Test Start demande"
        requestedLabelEn = "Requested Test Start"
    End If

    CalcBridge_BuildForecastStartDependencyDiagnosticMessage = _
        "FR:" & vbCrLf & _
        "Test Start impossible." & vbCrLf & vbCrLf & _
        "Tache : " & taskLabel & vbCrLf & _
        "Dependance bloquante : " & predLabel & " (" & linkText & ")" & vbCrLf & _
        predDateKindFr & " : " & CalcBridge_DiagnosticDateText(diag("BlockingPredecessorDate")) & vbCrLf & _
        "Debut minimum autorise : " & CalcBridge_DiagnosticDateText(diag("MinimumAllowedStart")) & vbCrLf & _
        requestedLabelFr & " : " & CalcBridge_DiagnosticDateText(diag("RequestedStart")) & vbCrLf & _
        "-> avancer le predecesseur, modifier le lien/lag, ou choisir un Start >= " & _
            CalcBridge_DiagnosticDateText(diag("MinimumAllowedStart")) & "." & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        "Test Start impossible." & vbCrLf & vbCrLf & _
        "Task: " & taskLabel & vbCrLf & _
        "Blocking dependency: " & predLabel & " (" & linkText & ")" & vbCrLf & _
        predDateKindEn & ": " & CalcBridge_DiagnosticDateText(diag("BlockingPredecessorDate")) & vbCrLf & _
        "Earliest allowed start: " & CalcBridge_DiagnosticDateText(diag("MinimumAllowedStart")) & vbCrLf & _
        requestedLabelEn & ": " & CalcBridge_DiagnosticDateText(diag("RequestedStart")) & vbCrLf & _
        "-> move the predecessor earlier, update the link/lag, or choose a Start >= " & _
            CalcBridge_DiagnosticDateText(diag("MinimumAllowedStart")) & "."

End Function
