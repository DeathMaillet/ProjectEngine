Attribute VB_Name = "mod_CalcCoreNetwork"
Option Explicit

'===============================================================================
' MODULE : mod_CalcCoreNetwork
' DOMAINE / DOMAIN : Core Calculation
'
' FR
' Construit le graphe Core, l'ordre topologique, les relations parent/enfant et les rollups summary.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Builds the Core graph, topological order, parent/child relations and summary rollups.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Core_CreateLinksBySucc, Core_MakeLink, Core_AddLink, Core_GetLinkPredId, Core_GetLinkType, Core_GetLinkLag, Core_GetLinkSummarySourceId, Core_BuildRowById
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

'------------------------------------------------------------------------------
' FR: Construit la map ID successeur -> collection de records de lien a partir des donnees fournies par l'appelant.
' EN: Builds the successor-ID-to-link-records map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function Core_CreateLinksBySucc() As Object
    Set Core_CreateLinksBySucc = CreateObject("Scripting.Dictionary")
End Function

'------------------------------------------------------------------------------
' FR: Construit le record de lien PredID/Type/Lag/SummarySource a partir des donnees fournies par l'appelant.
' EN: Builds the PredID/Type/Lag/SummarySource link record from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function Core_MakeLink( _
    ByVal predId As String, _
    ByVal linkType As String, _
    ByVal lagVal As Double, _
    Optional ByVal summarySourceId As String = "") As Variant

    Dim linkArr(0 To 3) As Variant

    linkArr(0) = Trim$(CStr(predId))
    linkArr(1) = UCase$(Trim$(CStr(linkType)))
    linkArr(2) = CDbl(lagVal)
    linkArr(3) = Trim$(CStr(summarySourceId))

    Core_MakeLink = linkArr

End Function

'------------------------------------------------------------------------------
' FR: Ajoute le record de lien PredID/Type/Lag/SummarySource a la structure cible fournie par l'appelant.
' EN: Adds the PredID/Type/Lag/SummarySource link record to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub Core_AddLink( _
    ByVal linksBySuccId As Object, _
    ByVal succId As String, _
    ByVal predId As String, _
    ByVal linkType As String, _
    ByVal lagVal As Double, _
    Optional ByVal summarySourceId As String = "")

    Dim oneLink As Variant

    If linksBySuccId Is Nothing Then Exit Sub

    succId = Trim$(CStr(succId))
    predId = Trim$(CStr(predId))
    summarySourceId = Trim$(CStr(summarySourceId))

    If succId = "" Then Exit Sub
    If predId = "" Then Exit Sub

    If Not linksBySuccId.Exists(succId) Then
        Set linksBySuccId(succId) = New Collection
    End If

    oneLink = Core_MakeLink(predId, linkType, lagVal, summarySourceId)
    linksBySuccId(succId).Add oneLink

End Sub

'------------------------------------------------------------------------------
' FR: Retourne l'ID predecesseur stocke dans le record de lien sans exposer de mutateur sur l'etat source.
' EN: Returns the predecessor ID stored in the link record without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function Core_GetLinkPredId(ByRef oneLink As Variant) As String
    Core_GetLinkPredId = Trim$(CStr(oneLink(0)))
End Function

'------------------------------------------------------------------------------
' FR: Retourne le type FS/SS/FF du record de lien sans exposer de mutateur sur l'etat source.
' EN: Returns the link record FS/SS/FF type without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function Core_GetLinkType(ByRef oneLink As Variant) As String
    Core_GetLinkType = UCase$(Trim$(CStr(oneLink(1))))
End Function

'------------------------------------------------------------------------------
' FR: Retourne le lag stocke dans le record de lien sans exposer de mutateur sur l'etat source.
' EN: Returns the lag stored in the link record without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function Core_GetLinkLag(ByRef oneLink As Variant) As Double
    Core_GetLinkLag = CDbl(oneLink(2))
End Function

'------------------------------------------------------------------------------
' FR: Retourne l'ID summary source du lien expanse sans exposer de mutateur sur l'etat source.
' EN: Returns the expanded link's source summary ID without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function Core_GetLinkSummarySourceId(ByRef oneLink As Variant) As String

    On Error GoTo SafeExit

    If IsArray(oneLink) Then
        If UBound(oneLink) >= 3 Then
            Core_GetLinkSummarySourceId = Trim$(CStr(oneLink(3)))
            Exit Function
        End If
    End If

SafeExit:
    Core_GetLinkSummarySourceId = ""

End Function


'------------------------------------------------------------------------------
' FR: Construit l'index ID -> ligne du dataset a partir des donnees fournies par l'appelant.
' EN: Builds the dataset ID-to-row index from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function Core_BuildRowById( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object) As Object

    Dim perfScope As clsPerfScope

    Dim rowById As Object
    Dim r As Long
    Dim idVal As String

    Set perfScope = Profiler_BeginScope("Core_BuildRowById", "Dictionary")

    Set rowById = CreateObject("Scripting.Dictionary")

    For r = LBound(dataArr, 1) To UBound(dataArr, 1)
        idVal = Trim$(CStr(Core_GetVal(dataArr, r, mapCol, "ID")))
        If idVal <> "" Then
            rowById(idVal) = r
        End If
    Next r

    Set Core_BuildRowById = rowById

End Function


'------------------------------------------------------------------------------
' FR: Construit la map ID -> ParentID a partir des donnees fournies par l'appelant.
' EN: Builds the ID-to-ParentID map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function Core_BuildParentIds( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object) As Object

    Dim perfScope As clsPerfScope

    Dim parentIds As Object
    Dim r As Long
    Dim idVal As String

    Set perfScope = Profiler_BeginScope("Core_BuildParentIds", "Dictionary")

    Set parentIds = CreateObject("Scripting.Dictionary")

    For r = LBound(dataArr, 1) To UBound(dataArr, 1)
        idVal = Trim$(CStr(Core_GetVal(dataArr, r, mapCol, "ID")))
        If idVal <> "" Then
            If Core_IsSummaryRow(dataArr, r, mapCol) Then
                parentIds(idVal) = True
            End If
        End If
    Next r

    Set Core_BuildParentIds = parentIds

End Function


'------------------------------------------------------------------------------
' FR: Construit la map parent ID -> enfants directs a partir des donnees fournies par l'appelant.
' EN: Builds the parent-ID-to-direct-children map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function Core_BuildDirectChildrenById( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object) As Object

    Dim perfScope As clsPerfScope

    Dim directChildrenById As Object
    Dim r As Long
    Dim idVal As String
    Dim parentId As String
    Dim key As Variant

    Set perfScope = Profiler_BeginScope("Core_BuildDirectChildrenById", "Dictionary")

    Set directChildrenById = CreateObject("Scripting.Dictionary")

    For Each key In rowById.Keys
        Set directChildrenById(CStr(key)) = New Collection
    Next key

    For r = LBound(dataArr, 1) To UBound(dataArr, 1)
        idVal = Trim$(CStr(Core_GetVal(dataArr, r, mapCol, "ID")))
        parentId = Trim$(CStr(Core_GetVal(dataArr, r, mapCol, "ParentID")))

        If idVal <> "" And parentId <> "" Then
            If directChildrenById.Exists(parentId) Then
                directChildrenById(parentId).Add idVal
            End If
        End If
    Next r

    Set Core_BuildDirectChildrenById = directChildrenById

End Function


'------------------------------------------------------------------------------
' FR: Construit la map predecesseur ID -> successeurs a partir des donnees fournies par l'appelant.
' EN: Builds the predecessor-ID-to-successors map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function Core_BuildChildrenByPred( _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object) As Object

    Dim perfScope As clsPerfScope

    Dim childrenByPred As Object
    Dim anyId As Variant
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String

    Set perfScope = Profiler_BeginScope("Core_BuildChildrenByPred", "Dictionary")

    Set childrenByPred = CreateObject("Scripting.Dictionary")

    For Each anyId In rowById.Keys
        Set childrenByPred(CStr(anyId)) = New Collection
    Next anyId

    If linksBySuccId Is Nothing Then
        Set Core_BuildChildrenByPred = childrenByPred
        Exit Function
    End If

    For Each succId In linksBySuccId.Keys
        For Each oneLink In linksBySuccId(CStr(succId))
            predId = Core_GetLinkPredId(oneLink)
            If predId <> "" Then
                If childrenByPred.Exists(predId) Then
                    childrenByPred(predId).Add CStr(succId)
                End If
            End If
        Next oneLink
    Next succId

    Set Core_BuildChildrenByPred = childrenByPred

End Function


'------------------------------------------------------------------------------
' FR: Construit l'ensemble des IDs de taches feuilles valides a partir des donnees fournies par l'appelant.
' EN: Builds the set of valid leaf-task IDs from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function Core_BuildValidLeafIds( _
    ByVal rowById As Object, _
    ByVal parentIds As Object) As Object

    Dim perfScope As clsPerfScope

    Dim validIds As Object
    Dim idVal As Variant

    Set perfScope = Profiler_BeginScope("Core_BuildValidLeafIds", "Dictionary")

    Set validIds = CreateObject("Scripting.Dictionary")

    For Each idVal In rowById.Keys
        If parentIds Is Nothing Then
            validIds(CStr(idVal)) = True
        ElseIf Not parentIds.Exists(CStr(idVal)) Then
            validIds(CStr(idVal)) = True
        End If
    Next idVal

    Set Core_BuildValidLeafIds = validIds

End Function


'------------------------------------------------------------------------------
' FR: Construit la map des degres entrants du reseau a partir des donnees fournies par l'appelant.
' EN: Builds the network indegree map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function Core_BuildIndegree( _
    ByVal validIds As Object, _
    ByVal linksBySuccId As Object) As Object

    Dim perfScope As clsPerfScope

    Dim indegree As Object
    Dim idVal As Variant
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String

    Set perfScope = Profiler_BeginScope("Core_BuildIndegree", "Dictionary")

    Set indegree = CreateObject("Scripting.Dictionary")

    For Each idVal In validIds.Keys
        indegree(CStr(idVal)) = 0
    Next idVal

    If linksBySuccId Is Nothing Then
        Set Core_BuildIndegree = indegree
        Exit Function
    End If

    For Each succId In linksBySuccId.Keys
        If validIds.Exists(CStr(succId)) Then
            For Each oneLink In linksBySuccId(CStr(succId))
                predId = Core_GetLinkPredId(oneLink)
                If predId <> "" Then
                    If validIds.Exists(predId) Then
                        indegree(CStr(succId)) = CLng(indegree(CStr(succId))) + 1
                    End If
                End If
            Next oneLink
        End If
    Next succId

    Set Core_BuildIndegree = indegree

End Function


'------------------------------------------------------------------------------
' FR: Retourne la collection Topo Sort Leaf Network sans modifier les donnees d'entree.
' EN: Returns the Topo Sort Leaf Network collection without mutating input data.
'------------------------------------------------------------------------------

Public Function Core_TopoSortLeafNetwork( _
    ByVal validIds As Object, _
    ByVal childrenByPred As Object, _
    ByVal indegree As Object) As Collection

    Dim perfScope As clsPerfScope

    Dim q As Collection
    Dim topoOrder As Collection
    Dim currentId As Variant
    Dim childId As Variant
    Dim key As Variant

    Set perfScope = Profiler_BeginScope("Core_TopoSortLeafNetwork", "Network Sort")

    Set q = New Collection
    Set topoOrder = New Collection

    For Each key In validIds.Keys
        If CLng(indegree(CStr(key))) = 0 Then
            q.Add CStr(key)
        End If
    Next key

    Do While q.Count > 0
        currentId = q(1)
        q.Remove 1
        topoOrder.Add CStr(currentId)

        If childrenByPred.Exists(CStr(currentId)) Then
            For Each childId In childrenByPred(CStr(currentId))
                If validIds.Exists(CStr(childId)) Then
                    indegree(CStr(childId)) = CLng(indegree(CStr(childId))) - 1
                    If CLng(indegree(CStr(childId))) = 0 Then
                        q.Add CStr(childId)
                    End If
                End If
            Next childId
        End If
    Loop

    Set Core_TopoSortLeafNetwork = topoOrder

End Function

'------------------------------------------------------------------------------
' FR: Indique si la collection Topo Failure satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Topo Failure collection satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Public Function Core_HasTopoFailure( _
    ByVal topoOrder As Collection, _
    ByVal validIds As Object) As Boolean

    If topoOrder Is Nothing Then
        Core_HasTopoFailure = True
        Exit Function
    End If

    Core_HasTopoFailure = (topoOrder.Count <> validIds.Count)

End Function

'------------------------------------------------------------------------------
' FR: Propage ou agrege la collection Error To Children selon les relations du reseau courant.
' EN: Propagates or rolls up the Error To Children collection through current network relations.
'------------------------------------------------------------------------------

Public Sub Core_PropagateErrorToChildren( _
    ByVal startId As String, _
    ByVal childrenByPred As Object, _
    ByVal errorDict As Object)

    Dim q As Collection
    Dim currentId As Variant
    Dim childId As Variant

    If errorDict Is Nothing Then Exit Sub
    If childrenByPred Is Nothing Then Exit Sub

    Set q = New Collection
    q.Add CStr(startId)
    errorDict(CStr(startId)) = True

    Do While q.Count > 0
        currentId = q(1)
        q.Remove 1

        If childrenByPred.Exists(CStr(currentId)) Then
            For Each childId In childrenByPred(CStr(currentId))
                If Not errorDict.Exists(CStr(childId)) Then
                    errorDict(CStr(childId)) = True
                    q.Add CStr(childId)
                End If
            Next childId
        End If
    Loop

End Sub


'------------------------------------------------------------------------------
' FR: Propage ou agrege la map Summary Dates selon les relations du reseau courant.
' EN: Propagates or rolls up the Summary Dates map through current network relations.
'------------------------------------------------------------------------------

Public Sub Core_RollupSummaryDates( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowById As Object, _
    ByVal directChildrenById As Object, _
    ByVal parentIds As Object)

    Dim perfScope As clsPerfScope

    Dim changed As Boolean
    Dim passCount As Long
    Dim maxPass As Long

    Dim parentId As Variant
    Dim childId As Variant
    Dim rowIdx As Long
    Dim childRow As Long

    Dim minStart As Variant
    Dim maxFinish As Variant
    Dim childStart As Variant
    Dim childFinish As Variant
    Dim hasChildData As Boolean

    Set perfScope = Profiler_BeginScope("Core_RollupSummaryDates", "Core Rollup")

    maxPass = rowById.Count
    passCount = 0

    Do
        changed = False
        passCount = passCount + 1

        For Each parentId In parentIds.Keys

            If rowById.Exists(CStr(parentId)) Then
                rowIdx = CLng(rowById(CStr(parentId)))
                minStart = Empty
                maxFinish = Empty
                hasChildData = False

                If directChildrenById.Exists(CStr(parentId)) Then
                    For Each childId In directChildrenById(CStr(parentId))
                        If rowById.Exists(CStr(childId)) Then
                            childRow = CLng(rowById(CStr(childId)))
                            childStart = Core_GetVal(dataArr, childRow, mapCol, "Calculated Start")
                            childFinish = Core_GetVal(dataArr, childRow, mapCol, "Calculated Finish")

                            If HasValue(childStart) And HasValue(childFinish) Then
                                If Not hasChildData Then
                                    minStart = childStart
                                    maxFinish = childFinish
                                    hasChildData = True
                                Else
                                    minStart = Core_MinDateIfBoth(minStart, childStart)
                                    maxFinish = Core_MaxDateIfBoth(maxFinish, childFinish)
                                End If
                            End If
                        End If
                    Next childId
                End If

                If hasChildData Then
                    If CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Calculated Start")) <> CStr(minStart) _
                    Or CStr(Core_GetVal(dataArr, rowIdx, mapCol, "Calculated Finish")) <> CStr(maxFinish) Then
                        Core_SetCalcTriplet dataArr, rowIdx, mapCol, minStart, maxFinish, _
                            NormalizeCalendarType(Core_GetVal(dataArr, rowIdx, mapCol, "Cal"))
                        changed = True
                    End If
                End If
            End If

        Next parentId

    Loop While changed And passCount <= maxPass

End Sub


