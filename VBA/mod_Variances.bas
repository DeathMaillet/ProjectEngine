Attribute VB_Name = "mod_Variances"
Option Explicit

'===============================================================================
' MODULE : mod_Variances
' DOMAINE / DOMAIN : Shared Infrastructure
'
' FR
' Calcule les references baseline, rollups parent et variances, puis ecrit leurs sorties WBS.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Calculates baseline references, parent rollups and variances, then writes their WBS outputs.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Compute_And_Push_Variances
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' ============================================================
' VARIANCE ENGINE – HYBRID BASELINE APPROACH
'
' Objectif :
' Calculer les variances Start / Finish / Duration directement
' en VBA, sans formule Excel, pour supporter :
' - les tâches sans baseline complète ;
' - les tâches dont la baseline peut être reconstruite par logique ;
' - les tâches summary / parents.
'
' Principe :
'
' 1. Référence baseline des tâches feuille
'
'    - Si Baseline Start + Baseline Finish existent :
'        on utilise ces dates saisies comme référence officielle PM.
'
'    - Sinon, si Baseline Duration existe et que les dépendances
'      permettent de positionner la tâche :
'        on reconstruit une baseline théorique à partir du réseau
'        logique déjà développé dans tbl_LOGIC_LINKS.
'
'    - Sinon :
'        la variance reste vide.
'
'
' 2. Référence baseline des tâches summary
'
'    Les parents ne doivent pas être saisis comme input planning.
'    Leur baseline de référence est donc calculée par roll-up :
'
'        Start    = min(Baseline Ref Start des enfants directs/indirects)
'        Finish   = max(Baseline Ref Finish des enfants directs/indirects)
'        Duration = Finish - Start + 1
'
'
' 3. Calcul des variances
'
'        Start Variance    = Calculated Start    - Baseline Ref Start
'        Finish Variance   = Calculated Finish   - Baseline Ref Finish
'        Duration Variance = Calculated Duration - Baseline Ref Duration
'
'
' 4. Sécurité WBS
'
'    Le WBS est considéré comme une table input PM.
'    Ce module NE RÉÉCRIT JAMAIS toute la table WBS.
'    Il écrit uniquement les trois colonnes autorisées :
'
'        - Start Variance
'        - Finish Variance
'        - Duration Variance
'
'    Le format de ces trois colonnes est forcé en numérique "0"
'    afin d'éviter qu'Excel transforme les jours de variance en dates
'    type 01/01/1900 ou 12:00:00 AM.
'
'    La colonne WBS est uniquement lue. Si nécessaire, elle est
'    normalisée en mémoire uniquement ("," -> ".") pour fiabiliser
'    les comparaisons parent/enfant.
'
' ============================================================


'------------------------------------------------------------------------------
' FR: Calcule And Push Variances pour le moteur variance calculation.
' EN: Computes And Push Variances for the variance calculation engine.
'------------------------------------------------------------------------------
Public Sub Compute_And_Push_Variances(Optional ByVal consoleMessages As Collection)

    Dim perfScope As clsPerfScope

    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim mapWBS As Object

    Dim rowById As Object
    Dim rowByWbs As Object
    Dim parentRows As Object
    Dim directChildrenRows As Object
    Dim leafRows As Object

    Dim refStartByRow As Object
    Dim refFinishByRow As Object
    Dim refDurByRow As Object

    Dim outSV() As Variant
    Dim outFV() As Variant
    Dim outDV() As Variant

    Dim requiredCols As Variant

    Set perfScope = Profiler_BeginScope("Compute_And_Push_Variances", "Variances")

    On Error GoTo ErrHandler

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)

    requiredCols = Array( _
        "ID", _
        "WBS", _
        "Predecessors WBS", _
        "Baseline Start", _
        "Baseline Duration", _
        "Baseline Finish", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration", _
        "Driving Logic", _
        "Start Variance", _
        "Finish Variance", _
        "Duration Variance" _
    )

    Core_RequireColumns mapWBS, requiredCols, "Compute_And_Push_Variances / tbl_WBS"

    Set rowById = CreateObject("Scripting.Dictionary")
    Set rowByWbs = CreateObject("Scripting.Dictionary")
    Set parentRows = CreateObject("Scripting.Dictionary")
    Set directChildrenRows = CreateObject("Scripting.Dictionary")
    Set leafRows = CreateObject("Scripting.Dictionary")

    Variance_BuildWbsIndexes _
        tblWBS, mapWBS, rowById, rowByWbs, parentRows, directChildrenRows, leafRows

    Set refStartByRow = CreateObject("Scripting.Dictionary")
    Set refFinishByRow = CreateObject("Scripting.Dictionary")
    Set refDurByRow = CreateObject("Scripting.Dictionary")

    Variance_BuildLeafBaselineReferences _
        tblWBS, mapWBS, rowByWbs, leafRows, _
        refStartByRow, refFinishByRow, refDurByRow

    Variance_RollupParentBaselineReferences _
        directChildrenRows, parentRows, _
        refStartByRow, refFinishByRow, refDurByRow

    Variance_BuildOutputArrays _
        tblWBS, mapWBS, _
        refStartByRow, refFinishByRow, refDurByRow, _
        outSV, outFV, outDV

    Variance_WriteOutputArraysToWBS tblWBS, outSV, outFV, outDV

    Exit Sub

ErrHandler:
    Variance_AddOrShowConsoleMessage consoleMessages, "STOP", _
        "Erreur dans Compute_And_Push_Variances" & vbCrLf & _
        "-> " & Err.Description, _
        "Error in Compute_And_Push_Variances" & vbCrLf & _
        "-> " & Err.Description

End Sub


'------------------------------------------------------------------------------
' FR: Construit la map Variance Build WBS Indexes a partir des donnees fournies par l'appelant.
' EN: Builds the Variance Build WBS Indexes map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Variance_BuildWbsIndexes( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal rowById As Object, _
    ByVal rowByWbs As Object, _
    ByVal parentRows As Object, _
    ByVal directChildrenRows As Object, _
    ByVal leafRows As Object)

    Dim r As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim parentWbs As String
    Dim parentRow As Long

    For r = 1 To tblWBS.ListRows.Count

        idVal = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).value))
        wbsVal = NormalizeWBSCode(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("WBS")).value))

        If idVal <> "" Then rowById(idVal) = r
        If wbsVal <> "" Then rowByWbs(wbsVal) = r

        If Not directChildrenRows.Exists(CStr(r)) Then
            Set directChildrenRows(CStr(r)) = New Collection
        End If

    Next r

    For r = 1 To tblWBS.ListRows.Count

        wbsVal = NormalizeWBSCode(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("WBS")).value))

        If wbsVal <> "" Then

            parentWbs = GetParentWBSCode(wbsVal)

            If parentWbs <> "" Then
                If rowByWbs.Exists(parentWbs) Then

                    parentRow = CLng(rowByWbs(parentWbs))

                    If Not directChildrenRows.Exists(CStr(parentRow)) Then
                        Set directChildrenRows(CStr(parentRow)) = New Collection
                    End If

                    directChildrenRows(CStr(parentRow)).Add r
                    parentRows(CStr(parentRow)) = True

                End If
            End If
        End If

    Next r

    For r = 1 To tblWBS.ListRows.Count
        If Not parentRows.Exists(CStr(r)) Then
            leafRows(CStr(r)) = True
        End If
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Construit la collection Variance Build Leaf Baseline References a partir des donnees fournies par l'appelant.
' EN: Builds the Variance Build Leaf Baseline References collection from data supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Variance_BuildLeafBaselineReferences( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal rowByWbs As Object, _
    ByVal leafRows As Object, _
    ByVal refStartByRow As Object, _
    ByVal refFinishByRow As Object, _
    ByVal refDurByRow As Object)

    Dim remaining As Object
    Dim progressed As Boolean
    Dim passCount As Long
    Dim maxPass As Long

    Dim key As Variant
    Dim r As Long

    Dim bs As Variant
    Dim bf As Variant
    Dim bd As Variant

    Dim predText As String
    Dim predTokens As Collection
    Dim token As Variant

    Dim predWbs As String
    Dim linkType As String
    Dim lagVal As Double
    Dim predRow As Long

    Dim hasValidPred As Boolean
    Dim hasUnresolvedPred As Boolean
    Dim bestStart As Variant
    Dim candidateStart As Variant

    Set remaining = CreateObject("Scripting.Dictionary")

    For Each key In leafRows.Keys
        remaining(CStr(key)) = True
    Next key

    maxPass = tblWBS.ListRows.Count + 5
    passCount = 0

    Do
        progressed = False
        passCount = passCount + 1

        For Each key In remaining.Keys

            r = CLng(key)

            bs = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS("Baseline Start")).value)
            bf = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS("Baseline Finish")).value)
            bd = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS("Baseline Duration")).value)

            If HasValue(bs) And HasValue(bf) Then

                refStartByRow(CStr(r)) = bs
                refFinishByRow(CStr(r)) = bf

                If HasValue(bd) And IsNumeric(bd) Then
                    refDurByRow(CStr(r)) = CLng(CDbl(bd))
                Else
                    refDurByRow(CStr(r)) = CLng(CDbl(bf) - CDbl(bs) + 1)
                End If

                remaining.Remove CStr(r)
                progressed = True
                GoTo NextRemainingTask

            End If

            If Not HasValue(bd) Then GoTo NextRemainingTask
            If Not IsNumeric(bd) Then GoTo NextRemainingTask
            If CDbl(bd) <= 0 Then GoTo NextRemainingTask

            predText = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS("Predecessors WBS")).value))

            If predText = "" Then GoTo NextRemainingTask

            Set predTokens = ParsePredecessorWbsTokens(predText)

            hasValidPred = False
            hasUnresolvedPred = False
            bestStart = Empty

            For Each token In predTokens

                ParsePredecessorToken CStr(token), predWbs, linkType, lagVal

                predWbs = NormalizeWBSCode(predWbs)

                If predWbs <> "" Then
                    If rowByWbs.Exists(predWbs) Then

                        predRow = CLng(rowByWbs(predWbs))

                        If leafRows.Exists(CStr(predRow)) Then

                            hasValidPred = True

                            If refStartByRow.Exists(CStr(predRow)) And refFinishByRow.Exists(CStr(predRow)) Then

                                Select Case UCase$(linkType)

                                    Case "SS"
                                        candidateStart = CDbl(refStartByRow(CStr(predRow))) + lagVal

                                    Case "FF"
                                        candidateStart = CDbl(refFinishByRow(CStr(predRow))) + lagVal - CDbl(bd) + 1

                                    Case Else
                                        candidateStart = CDbl(refFinishByRow(CStr(predRow))) + 1 + lagVal

                                End Select

                                If Not HasValue(bestStart) Then
                                    bestStart = candidateStart
                                ElseIf CDbl(candidateStart) > CDbl(bestStart) Then
                                    bestStart = candidateStart
                                End If

                            Else
                                hasUnresolvedPred = True
                            End If

                        End If
                    End If
                End If

            Next token

            If hasValidPred And Not hasUnresolvedPred And HasValue(bestStart) Then

                refStartByRow(CStr(r)) = bestStart
                refFinishByRow(CStr(r)) = CDbl(bestStart) + CDbl(bd) - 1
                refDurByRow(CStr(r)) = CLng(CDbl(bd))

                remaining.Remove CStr(r)
                progressed = True

            End If

NextRemainingTask:
        Next key

    Loop While remaining.Count > 0 And progressed And passCount <= maxPass

End Sub

'------------------------------------------------------------------------------
' FR: Propage ou agrege la map Variance Rollup Parent Baseline References selon les relations du reseau courant.
' EN: Propagates or rolls up the Variance Rollup Parent Baseline References map through current network relations.
'------------------------------------------------------------------------------

Private Sub Variance_RollupParentBaselineReferences( _
    ByVal directChildrenRows As Object, _
    ByVal parentRows As Object, _
    ByVal refStartByRow As Object, _
    ByVal refFinishByRow As Object, _
    ByVal refDurByRow As Object)

    Dim changed As Boolean
    Dim passCount As Long
    Dim maxPass As Long
    Dim parentKey As Variant
    Dim childRow As Variant

    Dim minStart As Variant
    Dim maxFinish As Variant
    Dim hasChildRef As Boolean

    maxPass = parentRows.Count + 5
    passCount = 0

    Do
        changed = False
        passCount = passCount + 1

        For Each parentKey In parentRows.Keys

            minStart = Empty
            maxFinish = Empty
            hasChildRef = False

            If directChildrenRows.Exists(CStr(parentKey)) Then

                For Each childRow In directChildrenRows(CStr(parentKey))

                    If refStartByRow.Exists(CStr(childRow)) And refFinishByRow.Exists(CStr(childRow)) Then

                        If Not hasChildRef Then
                            minStart = refStartByRow(CStr(childRow))
                            maxFinish = refFinishByRow(CStr(childRow))
                            hasChildRef = True
                        Else
                            If CDbl(refStartByRow(CStr(childRow))) < CDbl(minStart) Then
                                minStart = refStartByRow(CStr(childRow))
                            End If

                            If CDbl(refFinishByRow(CStr(childRow))) > CDbl(maxFinish) Then
                                maxFinish = refFinishByRow(CStr(childRow))
                            End If
                        End If

                    End If

                Next childRow

            End If

            If hasChildRef Then

                If Not refStartByRow.Exists(CStr(parentKey)) Then
                    refStartByRow(CStr(parentKey)) = minStart
                    refFinishByRow(CStr(parentKey)) = maxFinish
                    refDurByRow(CStr(parentKey)) = CLng(CDbl(maxFinish) - CDbl(minStart) + 1)
                    changed = True

                ElseIf CDbl(refStartByRow(CStr(parentKey))) <> CDbl(minStart) _
                    Or CDbl(refFinishByRow(CStr(parentKey))) <> CDbl(maxFinish) Then

                    refStartByRow(CStr(parentKey)) = minStart
                    refFinishByRow(CStr(parentKey)) = maxFinish
                    refDurByRow(CStr(parentKey)) = CLng(CDbl(maxFinish) - CDbl(minStart) + 1)
                    changed = True

                End If

            End If

        Next parentKey

    Loop While changed And passCount <= maxPass

End Sub

'------------------------------------------------------------------------------
' FR: Construit la map Variance Build Output Arrays a partir des donnees fournies par l'appelant.
' EN: Builds the Variance Build Output Arrays map from data supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Variance_BuildOutputArrays( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal refStartByRow As Object, _
    ByVal refFinishByRow As Object, _
    ByVal refDurByRow As Object, _
    ByRef outSV() As Variant, _
    ByRef outFV() As Variant, _
    ByRef outDV() As Variant)

    Dim r As Long
    Dim cs As Variant
    Dim cf As Variant
    Dim cd As Variant

    ReDim outSV(1 To tblWBS.ListRows.Count, 1 To 1)
    ReDim outFV(1 To tblWBS.ListRows.Count, 1 To 1)
    ReDim outDV(1 To tblWBS.ListRows.Count, 1 To 1)

    For r = 1 To tblWBS.ListRows.Count

        cs = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS("Calculated Start")).value)
        cf = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS("Calculated Finish")).value)
        cd = GetCellValue(tblWBS.DataBodyRange.Cells(r, mapWBS("Calculated Duration")).value)

        If HasValue(cs) And refStartByRow.Exists(CStr(r)) Then
            outSV(r, 1) = CLng(CDbl(cs) - CDbl(refStartByRow(CStr(r))))
        Else
            outSV(r, 1) = vbNullString
        End If

        If HasValue(cf) And refFinishByRow.Exists(CStr(r)) Then
            outFV(r, 1) = CLng(CDbl(cf) - CDbl(refFinishByRow(CStr(r))))
        Else
            outFV(r, 1) = vbNullString
        End If

        If HasValue(cd) And refDurByRow.Exists(CStr(r)) Then
            outDV(r, 1) = CLng(CDbl(cd) - CDbl(refDurByRow(CStr(r))))
        Else
            outDV(r, 1) = vbNullString
        End If

    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit ou synchronise Variance Write Output Arrays To WBS dans le stockage possede par le domaine.
' EN: Writes or synchronizes Variance Write Output Arrays To WBS in the store owned by the domain.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub Variance_WriteOutputArraysToWBS( _
    ByVal tblWBS As ListObject, _
    ByRef outSV() As Variant, _
    ByRef outFV() As Variant, _
    ByRef outDV() As Variant)

    Dim allowedFields As Variant
    Dim writeScopeToken As Long

    On Error GoTo ErrHandler

    allowedFields = Array( _
        "Start Variance", _
        "Finish Variance", _
        "Duration Variance" _
    )

    writeScopeToken = OpenAuthorizedWBSWriteScope( _
        "Variance_WriteOutputArraysToWBS", allowedFields)

    tblWBS.ListColumns("Start Variance").DataBodyRange.NumberFormat = "0"
    tblWBS.ListColumns("Finish Variance").DataBodyRange.NumberFormat = "0"
    tblWBS.ListColumns("Duration Variance").DataBodyRange.NumberFormat = "0"

    tblWBS.ListColumns("Start Variance").DataBodyRange.value = outSV
    tblWBS.ListColumns("Finish Variance").DataBodyRange.value = outFV
    tblWBS.ListColumns("Duration Variance").DataBodyRange.value = outDV

SafeExit:
    On Error Resume Next
    CloseAuthorizedWBSWriteScope writeScopeToken
    On Error GoTo 0
    Exit Sub

ErrHandler:
    Resume SafeExit

End Sub

'------------------------------------------------------------------------------
' FR: Normalise WBSCode dans un format exploitable.
' EN: Normalizes WBSCode into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizeWBSCode(ByVal value As String) As String

    value = Trim$(CStr(value))
    value = Replace(value, ",", ".")
    NormalizeWBSCode = value

End Function

'------------------------------------------------------------------------------
' FR: Retourne Parent WBSCode depuis le contexte variance calculation.
' EN: Returns Parent WBSCode from the variance calculation context.
'------------------------------------------------------------------------------
Private Function GetParentWBSCode(ByVal wbsCode As String) As String

    Dim p As Long

    wbsCode = NormalizeWBSCode(wbsCode)
    p = InStrRev(wbsCode, ".")

    If p > 0 Then
        GetParentWBSCode = Left$(wbsCode, p - 1)
    Else
        GetParentWBSCode = vbNullString
    End If

End Function

'------------------------------------------------------------------------------
' FR: Analyse Predecessor Wbs Tokens et extrait les informations utiles.
' EN: Parses Predecessor Wbs Tokens and extracts useful information.
'------------------------------------------------------------------------------
Private Function ParsePredecessorWbsTokens(ByVal predText As String) As Collection

    Dim result As Collection
    Dim parts() As String
    Dim i As Long
    Dim token As String

    Set result = New Collection

    predText = Trim$(CStr(predText))

    If predText <> "" Then
        parts = Split(predText, ";")

        For i = LBound(parts) To UBound(parts)
            token = Trim$(parts(i))
            If token <> "" Then result.Add token
        Next i
    End If

    Set ParsePredecessorWbsTokens = result

End Function

'------------------------------------------------------------------------------
' FR: Analyse Predecessor Token et extrait les informations utiles.
' EN: Parses Predecessor Token and extracts useful information.
'------------------------------------------------------------------------------
Private Sub ParsePredecessorToken( _
    ByVal token As String, _
    ByRef predWbs As String, _
    ByRef linkType As String, _
    ByRef lagVal As Double)

    Dim posSS As Long
    Dim posFF As Long
    Dim posSign As Long
    Dim posPlus As Long
    Dim posMinus As Long
    Dim logicPos As Long

    token = Trim$(CStr(token))
    token = Replace(token, ",", ".")

    linkType = "FS"
    lagVal = 0#

    posSS = InStr(1, UCase$(token), "SS", vbTextCompare)
    posFF = InStr(1, UCase$(token), "FF", vbTextCompare)

    If posSS > 0 Then
        logicPos = posSS
        linkType = "SS"
        predWbs = Left$(token, logicPos - 1)
        token = Mid$(token, logicPos + 2)
    ElseIf posFF > 0 Then
        logicPos = posFF
        linkType = "FF"
        predWbs = Left$(token, logicPos - 1)
        token = Mid$(token, logicPos + 2)
    Else
        posPlus = InStrRev(token, "+")
        posMinus = InStrRev(token, "-")

        posSign = 0
        If posPlus > 0 Then posSign = posPlus
        If posMinus > 0 Then
            If posMinus > posSign Then posSign = posMinus
        End If

        If posSign > 0 Then
            predWbs = Left$(token, posSign - 1)
            lagVal = CDbl(Mid$(token, posSign))
        Else
            predWbs = token
        End If

        predWbs = NormalizeWBSCode(predWbs)
        Exit Sub
    End If

    token = Trim$(token)

    If token <> "" Then
        If Left$(token, 1) = "+" Or Left$(token, 1) = "-" Then
            lagVal = CDbl(token)
        End If
    End If

    predWbs = NormalizeWBSCode(predWbs)

End Sub



'------------------------------------------------------------------------------
' FR: Ajoute la collection Variance Add Or Show Console Message a la structure cible fournie par l'appelant.
' EN: Adds the Variance Add Or Show Console Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub Variance_AddOrShowConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    CalcBridge_AddOrShowConsoleMessage consoleMessages, msgType, frText, enText

End Sub
