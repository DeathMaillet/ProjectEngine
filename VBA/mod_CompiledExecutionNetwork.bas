Attribute VB_Name = "mod_CompiledExecutionNetwork"
Option Explicit

'===============================================================================
' MODULE : mod_CompiledExecutionNetwork
' DOMAINE / DOMAIN : Core Calculation / Analytics
'
' FR
' Compile une representation transactionnelle unique du reseau d'execution.
' Centralise les index, parentés, adjacences et ordres partages par les moteurs.
' Ne lit ni n'ecrit Excel et ne conserve aucun cache global.
'
' EN
' Compiles one transaction-owned execution-network representation.
' Centralizes indexes, parentage, adjacency and orders shared by the engines.
' Neither reads nor writes Excel and keeps no global cache.
'
' CONTRATS / CONTRACTS : CompileExecutionNetwork, CompiledNetwork_BuildCurrentFinishById, CompiledNetwork_BuildFinishByIdFromValues, CompiledExecutionNetworkHarness_*
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

' FR:
' Compile le dataset et ses liens deja parses/expanses en vues indexees partagees.
' Les diagnostics et politiques de lien restent ceux des proprietaires existants.
'
' EN:
' Compiles the dataset and its already parsed/expanded links into shared indexed
' views. Existing owners retain link diagnostics and policy.
Public Function CompileExecutionNetwork( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal linksBySuccId As Object) As clsCompiledExecutionNetwork

    Dim perfScope As clsPerfScope
    Dim result As clsCompiledExecutionNetwork
    Dim rowById As Object
    Dim parentIds As Object
    Dim directChildrenById As Object
    Dim childrenByPred As Object
    Dim validLeafIds As Object
    Dim loeIds As Object
    Dim indegree As Object
    Dim topoOrder As Collection
    Dim analyticsPredsById As Object
    Dim analyticsChildrenById As Object
    Dim predLagBySuccPred As Object
    Dim predTypeBySuccPred As Object
    Dim idToWbs As Object
    Dim nodeIndexById As Object

    Dim nodeIds() As String
    Dim rowsByNode() As Long
    Dim parentByNode() As Long
    Dim predOffsets() As Long
    Dim predNodes() As Long
    Dim predTypes() As String
    Dim predLags() As Double
    Dim corePredOffsets() As Long
    Dim corePredNodes() As Long
    Dim corePredIds() As String
    Dim corePredTypes() As String
    Dim corePredLags() As Double
    Dim corePredSummarySourceIds() As String
    Dim topoNodes() As Long
    Dim componentParents() As Long
    Dim componentByNode() As Long
    Dim componentDenseByRoot As Object

    Dim r As Long
    Dim nodeIndex As Long
    Dim edgeIndex As Long
    Dim coreEdgeIndex As Long
    Dim idVal As String
    Dim parentId As String
    Dim wbsVal As String
    Dim idKey As Variant
    Dim succId As Variant
    Dim predId As String
    Dim oneLink As Variant
    Dim linkType As String
    Dim linkLag As Double
    Dim linkKey As String
    Dim topoId As Variant
    Dim componentRoot As Long
    Dim componentCount As Long

    Set perfScope = Profiler_BeginScope("CompileExecutionNetwork", "Network Compile")

    Set rowById = Core_BuildRowById(dataArr, mapCol)
    Set parentIds = Core_BuildParentIds(dataArr, mapCol, rowById)
    Set directChildrenById = Core_BuildDirectChildrenById(dataArr, mapCol, rowById)
    Set childrenByPred = Core_BuildChildrenByPred(rowById, linksBySuccId)
    Set loeIds = CreateObject("Scripting.Dictionary")
    Set validLeafIds = Core_BuildValidLeafIds(rowById, parentIds)
    Set idToWbs = CreateObject("Scripting.Dictionary")
    Set nodeIndexById = CreateObject("Scripting.Dictionary")

    For Each idKey In rowById.Keys
        r = CLng(rowById(CStr(idKey)))
        If TaskTypeRules_IsLevelOfEffortRow(dataArr, mapCol, r) Then
            loeIds(CStr(idKey)) = True
            If validLeafIds.Exists(CStr(idKey)) Then validLeafIds.Remove CStr(idKey)
        End If
    Next idKey

    Set analyticsPredsById = CompiledNetwork_CreateEmptyCollections(validLeafIds)
    Set analyticsChildrenById = CompiledNetwork_CreateEmptyCollections(validLeafIds)
    Set predLagBySuccPred = CreateObject("Scripting.Dictionary")
    Set predTypeBySuccPred = CreateObject("Scripting.Dictionary")

    For Each succId In linksBySuccId.Keys
        If validLeafIds.Exists(CStr(succId)) Then
            For Each oneLink In linksBySuccId(CStr(succId))
                predId = Core_GetLinkPredId(oneLink)
                If validLeafIds.Exists(predId) Then
                    linkType = Core_GetLinkType(oneLink)
                    If linkType = "FS" Or linkType = "SS" Or linkType = "FF" Then
                        linkLag = Core_GetLinkLag(oneLink)
                        linkKey = CStr(succId) & "|" & predId
                        predLagBySuccPred(linkKey) = linkLag
                        predTypeBySuccPred(linkKey) = linkType
                        analyticsPredsById(CStr(succId)).Add predId
                        analyticsChildrenById(predId).Add CStr(succId)
                    End If
                End If
            Next oneLink
        End If
    Next succId

    Set indegree = Core_BuildIndegree(validLeafIds, linksBySuccId)
    Set topoOrder = Core_TopoSortLeafNetwork(validLeafIds, childrenByPred, indegree)

    ReDim nodeIds(1 To rowById.Count)
    ReDim rowsByNode(1 To rowById.Count)
    ReDim parentByNode(1 To rowById.Count)
    ReDim componentParents(1 To rowById.Count)
    ReDim componentByNode(1 To rowById.Count)

    nodeIndex = 0
    For r = LBound(dataArr, 1) To UBound(dataArr, 1)
        idVal = Trim$(CStr(dataArr(r, mapCol("ID"))))
        If idVal <> "" Then
            nodeIndex = nodeIndex + 1
            nodeIds(nodeIndex) = idVal
            rowsByNode(nodeIndex) = r
            nodeIndexById(idVal) = nodeIndex
            componentParents(nodeIndex) = nodeIndex
            If mapCol.Exists("WBS") Then
                wbsVal = NormalizeWBS(CStr(dataArr(r, mapCol("WBS"))))
                idToWbs(idVal) = wbsVal
            End If
        End If
    Next r

    For nodeIndex = 1 To rowById.Count
        r = rowsByNode(nodeIndex)
        parentId = Trim$(CStr(dataArr(r, mapCol("ParentID"))))
        If parentId <> "" Then
            If nodeIndexById.Exists(parentId) Then parentByNode(nodeIndex) = CLng(nodeIndexById(parentId))
        End If
    Next nodeIndex

    ' Core edges retain every expanded link and its diagnostic metadata.
    ReDim corePredOffsets(1 To rowById.Count + 1)
    coreEdgeIndex = 0
    For nodeIndex = 1 To rowById.Count
        corePredOffsets(nodeIndex) = coreEdgeIndex + 1
        idVal = nodeIds(nodeIndex)
        If linksBySuccId.Exists(idVal) Then
            coreEdgeIndex = coreEdgeIndex + linksBySuccId(idVal).Count
        End If
    Next nodeIndex
    corePredOffsets(rowById.Count + 1) = coreEdgeIndex + 1

    If coreEdgeIndex > 0 Then
        ReDim corePredNodes(1 To coreEdgeIndex)
        ReDim corePredIds(1 To coreEdgeIndex)
        ReDim corePredTypes(1 To coreEdgeIndex)
        ReDim corePredLags(1 To coreEdgeIndex)
        ReDim corePredSummarySourceIds(1 To coreEdgeIndex)
    Else
        ReDim corePredNodes(0 To 0)
        ReDim corePredIds(0 To 0)
        ReDim corePredTypes(0 To 0)
        ReDim corePredLags(0 To 0)
        ReDim corePredSummarySourceIds(0 To 0)
    End If

    coreEdgeIndex = 0
    For nodeIndex = 1 To rowById.Count
        idVal = nodeIds(nodeIndex)
        If linksBySuccId.Exists(idVal) Then
            For Each oneLink In linksBySuccId(idVal)
                coreEdgeIndex = coreEdgeIndex + 1
                predId = Core_GetLinkPredId(oneLink)
                corePredIds(coreEdgeIndex) = predId
                If nodeIndexById.Exists(predId) Then
                    corePredNodes(coreEdgeIndex) = CLng(nodeIndexById(predId))
                End If
                corePredTypes(coreEdgeIndex) = Core_GetLinkType(oneLink)
                corePredLags(coreEdgeIndex) = Core_GetLinkLag(oneLink)
                corePredSummarySourceIds(coreEdgeIndex) = Core_GetLinkSummarySourceId(oneLink)
            Next oneLink
        End If
    Next nodeIndex

    ReDim predOffsets(1 To rowById.Count + 1)
    edgeIndex = 0
    For nodeIndex = 1 To rowById.Count
        predOffsets(nodeIndex) = edgeIndex + 1
        idVal = nodeIds(nodeIndex)
        If analyticsPredsById.Exists(idVal) Then edgeIndex = edgeIndex + analyticsPredsById(idVal).Count
    Next nodeIndex
    predOffsets(rowById.Count + 1) = edgeIndex + 1

    If edgeIndex > 0 Then
        ReDim predNodes(1 To edgeIndex)
        ReDim predTypes(1 To edgeIndex)
        ReDim predLags(1 To edgeIndex)
    Else
        ReDim predNodes(0 To 0)
        ReDim predTypes(0 To 0)
        ReDim predLags(0 To 0)
    End If

    edgeIndex = 0
    For nodeIndex = 1 To rowById.Count
        idVal = nodeIds(nodeIndex)
        If analyticsPredsById.Exists(idVal) Then
            For Each topoId In analyticsPredsById(idVal)
                edgeIndex = edgeIndex + 1
                predNodes(edgeIndex) = CLng(nodeIndexById(CStr(topoId)))
                CompiledNetwork_Union componentParents, nodeIndex, predNodes(edgeIndex)
                linkKey = idVal & "|" & CStr(topoId)
                predTypes(edgeIndex) = CStr(predTypeBySuccPred(linkKey))
                predLags(edgeIndex) = CDbl(predLagBySuccPred(linkKey))
            Next topoId
        End If
    Next nodeIndex

    If topoOrder.Count > 0 Then
        ReDim topoNodes(1 To topoOrder.Count)
        nodeIndex = 0
        For Each topoId In topoOrder
            nodeIndex = nodeIndex + 1
            topoNodes(nodeIndex) = CLng(nodeIndexById(CStr(topoId)))
        Next topoId
    Else
        ReDim topoNodes(0 To 0)
    End If

    Set componentDenseByRoot = CreateObject("Scripting.Dictionary")
    componentCount = 0
    For nodeIndex = 1 To rowById.Count
        idVal = nodeIds(nodeIndex)
        If validLeafIds.Exists(idVal) Then
            componentRoot = CompiledNetwork_FindRoot(componentParents, nodeIndex)
            If Not componentDenseByRoot.Exists(CStr(componentRoot)) Then
                componentCount = componentCount + 1
                componentDenseByRoot(CStr(componentRoot)) = componentCount
            End If
            componentByNode(nodeIndex) = CLng(componentDenseByRoot(CStr(componentRoot)))
        End If
    Next nodeIndex

    Set result = New clsCompiledExecutionNetwork
    result.InitializeStructures rowById, parentIds, directChildrenById, childrenByPred, _
        validLeafIds, loeIds, topoOrder, analyticsPredsById, analyticsChildrenById, _
        predLagBySuccPred, predTypeBySuccPred, idToWbs, nodeIndexById
    result.InitializeIndexedArrays _
        nodeIds, rowsByNode, parentByNode, predOffsets, predNodes, predTypes, _
        predLags, topoNodes, componentByNode, edgeIndex, componentCount, _
        (topoOrder.Count = validLeafIds.Count)
    result.InitializeCoreIndexedArrays _
        corePredOffsets, corePredNodes, corePredIds, corePredTypes, corePredLags, _
        corePredSummarySourceIds

    Profiler_RecordCounter "CompiledNetworkCompilations", 1
    Profiler_RecordCounter "CompiledNetworkNodes", rowById.Count
    Profiler_RecordCounter "CompiledNetworkLeaves", validLeafIds.Count
    Profiler_RecordCounter "CompiledNetworkEdges", edgeIndex
    Profiler_RecordCounter "CompiledNetworkCoreEdges", coreEdgeIndex
    Profiler_RecordCounter "CompiledNetworkComponents", componentCount

    Set CompileExecutionNetwork = result

End Function

' FR: Construit en O(n) la date de fin Current de chaque composante compilee.
' EN: Builds each compiled component's Current finish in O(n).
' FR: Projette les aretes de summary vers les feuilles qui portent réellement
' la borne agrégée: Start minimum pour SS, Finish maximum pour FF. FS conserve
' son expansion historique complète.
Public Function CompiledNetwork_BuildSemanticAnalyticsViews( _
    ByVal network As clsCompiledExecutionNetwork, _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal startColumn As String, _
    ByVal finishColumn As String) As Object

    Dim startById As Object
    Dim finishById As Object
    Dim idKey As Variant
    Dim rowIndex As Long
    Dim value As Variant

    Set startById = CreateObject("Scripting.Dictionary")
    Set finishById = CreateObject("Scripting.Dictionary")
    For Each idKey In network.ValidLeafIds.Keys
        rowIndex = CLng(network.RowById(CStr(idKey)))
        value = GetCellValue(dataArr(rowIndex, mapCol(startColumn)))
        If HasValue(value) Then startById(CStr(idKey)) = value
        value = GetCellValue(dataArr(rowIndex, mapCol(finishColumn)))
        If HasValue(value) Then finishById(CStr(idKey)) = value
    Next idKey

    Set CompiledNetwork_BuildSemanticAnalyticsViews = _
        CompiledNetwork_BuildSemanticAnalyticsViewsFromValues(network, startById, finishById)

End Function

Public Function CompiledNetwork_BuildSemanticAnalyticsViewsFromValues( _
    ByVal network As clsCompiledExecutionNetwork, _
    ByVal startById As Object, _
    ByVal finishById As Object) As Object

    Dim result As Object
    Dim predsById As Object
    Dim childrenById As Object
    Dim lagByEdge As Object
    Dim typeByEdge As Object
    Dim edgeKeys As Object
    Dim groupMembers As Object
    Dim groupBestMembers As Object
    Dim groupBestValue As Object
    Dim groupTypes As Object
    Dim groupLags As Object
    Dim members As Collection
    Dim bestMembers As Collection
    Dim validIds As Object
    Dim nodeIds As Variant
    Dim offsets As Variant
    Dim predIds As Variant
    Dim predTypes As Variant
    Dim predLags As Variant
    Dim summarySources As Variant
    Dim succNode As Long
    Dim edgeIndex As Long
    Dim succId As String
    Dim predId As String
    Dim linkType As String
    Dim summarySource As String
    Dim groupKey As String
    Dim group As Variant
    Dim member As Variant
    Dim candidate As Variant
    Dim bestValue As Double
    Dim hasCandidate As Boolean
    Dim better As Boolean

    Set result = CreateObject("Scripting.Dictionary")
    Set validIds = network.ValidLeafIds
    Set predsById = CompiledNetwork_CreateEmptyCollections(validIds)
    Set childrenById = CompiledNetwork_CreateEmptyCollections(validIds)
    Set lagByEdge = CreateObject("Scripting.Dictionary")
    Set typeByEdge = CreateObject("Scripting.Dictionary")
    Set edgeKeys = CreateObject("Scripting.Dictionary")
    nodeIds = network.NodeIds
    offsets = network.CorePredOffsets
    predIds = network.CorePredIds
    predTypes = network.CorePredTypes
    predLags = network.CorePredLags
    summarySources = network.CorePredSummarySourceIds

    For succNode = 1 To network.NodeCount
        succId = CStr(nodeIds(succNode))
        If validIds.Exists(succId) Then
            Set groupMembers = CreateObject("Scripting.Dictionary")
            Set groupBestMembers = CreateObject("Scripting.Dictionary")
            Set groupBestValue = CreateObject("Scripting.Dictionary")
            Set groupTypes = CreateObject("Scripting.Dictionary")
            Set groupLags = CreateObject("Scripting.Dictionary")

            For edgeIndex = CLng(offsets(succNode)) To CLng(offsets(succNode + 1)) - 1
                predId = CStr(predIds(edgeIndex))
                linkType = UCase$(Trim$(CStr(predTypes(edgeIndex))))
                summarySource = Trim$(CStr(summarySources(edgeIndex)))
                If validIds.Exists(predId) Then
                    If Len(summarySource) > 0 And (linkType = "SS" Or linkType = "FF") Then
                        groupKey = summarySource & "|" & linkType & "|" & CStr(CDbl(predLags(edgeIndex)))
                        If Not groupMembers.Exists(groupKey) Then
                            Set members = New Collection
                            groupMembers.Add groupKey, members
                            groupTypes(groupKey) = linkType
                            groupLags(groupKey) = CDbl(predLags(edgeIndex))
                        Else
                            Set members = groupMembers(groupKey)
                        End If
                        members.Add predId

                        hasCandidate = False
                        If linkType = "SS" Then
                            If startById.Exists(predId) Then candidate = startById(predId): hasCandidate = True
                        Else
                            If finishById.Exists(predId) Then candidate = finishById(predId): hasCandidate = True
                        End If
                        If hasCandidate Then
                            better = False
                            If Not groupBestValue.Exists(groupKey) Then
                                better = True
                            Else
                                bestValue = CDbl(groupBestValue(groupKey))
                                If linkType = "SS" Then
                                    better = (CDbl(candidate) < bestValue)
                                Else
                                    better = (CDbl(candidate) > bestValue)
                                End If
                            End If
                            If better Then
                                groupBestValue(groupKey) = CDbl(candidate)
                                Set bestMembers = New Collection
                                bestMembers.Add predId
                                If groupBestMembers.Exists(groupKey) Then
                                    Set groupBestMembers(groupKey) = bestMembers
                                Else
                                    groupBestMembers.Add groupKey, bestMembers
                                End If
                            ElseIf Abs(CDbl(candidate) - CDbl(groupBestValue(groupKey))) < 0.000001 Then
                                Set bestMembers = groupBestMembers(groupKey)
                                bestMembers.Add predId
                            End If
                        End If
                    Else
                        CompiledNetwork_AddSemanticEdge predsById, childrenById, lagByEdge, typeByEdge, edgeKeys, _
                            succId, predId, linkType, CDbl(predLags(edgeIndex))
                    End If
                End If
            Next edgeIndex

            For Each group In groupMembers.Keys
                If groupBestMembers.Exists(CStr(group)) Then
                    Set members = groupBestMembers(CStr(group))
                Else
                    Set members = groupMembers(CStr(group))
                End If
                For Each member In members
                    CompiledNetwork_AddSemanticEdge predsById, childrenById, lagByEdge, typeByEdge, edgeKeys, _
                        succId, CStr(member), CStr(groupTypes(CStr(group))), CDbl(groupLags(CStr(group)))
                Next member
            Next group
        End If
    Next succNode

    result.Add "PredsById", predsById
    result.Add "ChildrenById", childrenById
    result.Add "PredLagBySuccPred", lagByEdge
    result.Add "PredTypeBySuccPred", typeByEdge
    Set CompiledNetwork_BuildSemanticAnalyticsViewsFromValues = result

End Function

Private Sub CompiledNetwork_AddSemanticEdge( _
    ByVal predsById As Object, _
    ByVal childrenById As Object, _
    ByVal lagByEdge As Object, _
    ByVal typeByEdge As Object, _
    ByVal edgeKeys As Object, _
    ByVal succId As String, _
    ByVal predId As String, _
    ByVal linkType As String, _
    ByVal lagValue As Double)

    Dim edgeKey As String
    Dim semanticKey As String

    edgeKey = succId & "|" & predId
    semanticKey = edgeKey & "|" & linkType & "|" & CStr(lagValue)
    If edgeKeys.Exists(semanticKey) Then Exit Sub
    edgeKeys(semanticKey) = True
    predsById(succId).Add predId
    childrenById(predId).Add succId
    lagByEdge(edgeKey) = lagValue
    typeByEdge(edgeKey) = linkType

End Sub

Public Function CompiledNetwork_BuildCurrentFinishById( _
    ByVal executionNetwork As clsCompiledExecutionNetwork, _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object) As Object

    Dim result As Object
    Dim componentFinish() As Variant
    Dim nodeIds As Variant
    Dim rowsByNode As Variant
    Dim componentByNode As Variant
    Dim nodeIndex As Long
    Dim componentIndex As Long
    Dim finishVal As Variant

    Set result = CreateObject("Scripting.Dictionary")
    If executionNetwork Is Nothing Then
        Set CompiledNetwork_BuildCurrentFinishById = result
        Exit Function
    End If
    If executionNetwork.ComponentCount = 0 Then
        Set CompiledNetwork_BuildCurrentFinishById = result
        Exit Function
    End If

    nodeIds = executionNetwork.NodeIds
    rowsByNode = executionNetwork.RowsByNode
    componentByNode = executionNetwork.ComponentByNode
    ReDim componentFinish(1 To executionNetwork.ComponentCount)

    For nodeIndex = 1 To executionNetwork.NodeCount
        componentIndex = componentByNode(nodeIndex)
        If componentIndex > 0 Then
            finishVal = GetCellValue(dataArr(rowsByNode(nodeIndex), mapCalc("Calculated Finish")))
            If HasValue(finishVal) Then
                If Not HasValue(componentFinish(componentIndex)) Then
                    componentFinish(componentIndex) = finishVal
                ElseIf CDbl(finishVal) > CDbl(componentFinish(componentIndex)) Then
                    componentFinish(componentIndex) = finishVal
                End If
            End If
        End If
    Next nodeIndex

    For nodeIndex = 1 To executionNetwork.NodeCount
        componentIndex = componentByNode(nodeIndex)
        If componentIndex > 0 Then
            If HasValue(componentFinish(componentIndex)) Then
                result(CStr(nodeIds(nodeIndex))) = componentFinish(componentIndex)
            End If
        End If
    Next nodeIndex

    Set CompiledNetwork_BuildCurrentFinishById = result

End Function

' FR:
' Projette en O(n) les fins fournies sur les composantes du reseau compile.
' Le dictionnaire source conserve la politique temporelle de son moteur.
'
' EN:
' Projects supplied finish values over compiled components in O(n).
' The source dictionary retains ownership of its engine's time policy.
Public Function CompiledNetwork_BuildFinishByIdFromValues( _
    ByVal executionNetwork As clsCompiledExecutionNetwork, _
    ByVal finishById As Object) As Object

    Dim result As Object
    Dim componentFinish() As Variant
    Dim nodeIds As Variant
    Dim componentByNode As Variant
    Dim nodeIndex As Long
    Dim componentIndex As Long
    Dim taskId As String
    Dim finishVal As Variant

    Set result = CreateObject("Scripting.Dictionary")
    If executionNetwork Is Nothing Then
        Set CompiledNetwork_BuildFinishByIdFromValues = result
        Exit Function
    End If
    If finishById Is Nothing Then
        Set CompiledNetwork_BuildFinishByIdFromValues = result
        Exit Function
    End If
    If executionNetwork.ComponentCount = 0 Then
        Set CompiledNetwork_BuildFinishByIdFromValues = result
        Exit Function
    End If

    nodeIds = executionNetwork.NodeIds
    componentByNode = executionNetwork.ComponentByNode
    ReDim componentFinish(1 To executionNetwork.ComponentCount)

    For nodeIndex = 1 To executionNetwork.NodeCount
        componentIndex = componentByNode(nodeIndex)
        If componentIndex > 0 Then
            taskId = CStr(nodeIds(nodeIndex))
            If finishById.Exists(taskId) Then
                finishVal = finishById(taskId)
                If HasValue(finishVal) Then
                    If Not HasValue(componentFinish(componentIndex)) Then
                        componentFinish(componentIndex) = finishVal
                    ElseIf CDbl(finishVal) > CDbl(componentFinish(componentIndex)) Then
                        componentFinish(componentIndex) = finishVal
                    End If
                End If
            End If
        End If
    Next nodeIndex

    For nodeIndex = 1 To executionNetwork.NodeCount
        componentIndex = componentByNode(nodeIndex)
        If componentIndex > 0 Then
            If HasValue(componentFinish(componentIndex)) Then
                result(CStr(nodeIds(nodeIndex))) = componentFinish(componentIndex)
            End If
        End If
    Next nodeIndex

    Set CompiledNetwork_BuildFinishByIdFromValues = result

End Function

' FR: Cree une collection vide par ID pour les adjacences compilees.
' EN: Creates one empty collection per ID for compiled adjacency.
Private Function CompiledNetwork_CreateEmptyCollections(ByVal ids As Object) As Object

    Dim result As Object
    Dim idKey As Variant

    Set result = CreateObject("Scripting.Dictionary")
    For Each idKey In ids.Keys
        Set result(CStr(idKey)) = New Collection
    Next idKey

    Set CompiledNetwork_CreateEmptyCollections = result

End Function

' FR: Retourne la racine union-find d'un noeud et compresse son chemin.
' EN: Returns a node's union-find root and compresses its path.
Private Function CompiledNetwork_FindRoot( _
    ByRef parents() As Long, _
    ByVal nodeIndex As Long) As Long

    Dim rootIndex As Long
    Dim nextIndex As Long

    rootIndex = nodeIndex
    Do While parents(rootIndex) <> rootIndex
        rootIndex = parents(rootIndex)
    Loop

    Do While parents(nodeIndex) <> nodeIndex
        nextIndex = parents(nodeIndex)
        parents(nodeIndex) = rootIndex
        nodeIndex = nextIndex
    Loop

    CompiledNetwork_FindRoot = rootIndex

End Function

' FR: Fusionne deux composantes du reseau compile.
' EN: Merges two compiled-network components.
Private Sub CompiledNetwork_Union( _
    ByRef parents() As Long, _
    ByVal leftNode As Long, _
    ByVal rightNode As Long)

    Dim leftRoot As Long
    Dim rightRoot As Long

    leftRoot = CompiledNetwork_FindRoot(parents, leftNode)
    rightRoot = CompiledNetwork_FindRoot(parents, rightNode)
    If leftRoot <> rightRoot Then parents(rightRoot) = leftRoot

End Sub

' FR:
' Compile le reseau CALC courant et verifie ses index, compteurs et topologie.
' Retourne PASS ou un diagnostic FAIL sans ecrire dans le classeur.
'
' EN:
' Compiles the current CALC network and verifies indexes, counts and topology.
' Returns PASS or a FAIL diagnostic without writing to the workbook.
Public Function CompiledExecutionNetworkHarness_Smoke() As String

    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim dataArr As Variant
    Dim linksBySuccId As Object
    Dim network As clsCompiledExecutionNetwork
    Dim nodeIds As Variant
    Dim rowsByNode As Variant
    Dim topoNodes As Variant
    Dim nodeIndex As Long
    Dim taskId As String

    On Error GoTo Fail

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    If tblCalc.DataBodyRange Is Nothing Then
        CompiledExecutionNetworkHarness_Smoke = "PASS: EMPTY"
        Exit Function
    End If

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    dataArr = tblCalc.DataBodyRange.Value2
    Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded(tblCalc)
    Set network = CompileExecutionNetwork(dataArr, mapCalc, linksBySuccId)

    If network Is Nothing Then Err.Raise vbObjectError + 4460, , "Compiled network is missing."
    If network.NodeCount <> network.RowById.Count Then Err.Raise vbObjectError + 4461, , "Node count mismatch."
    If network.LeafCount <> network.ValidLeafIds.Count Then Err.Raise vbObjectError + 4462, , "Leaf count mismatch."
    If network.TopoOrder.Count <> network.LeafCount Then Err.Raise vbObjectError + 4463, , "Topology does not cover every leaf."
    If Not network.IsValid Then Err.Raise vbObjectError + 4464, , "Compiled topology is invalid."

    nodeIds = network.NodeIds
    rowsByNode = network.RowsByNode
    For nodeIndex = 1 To network.NodeCount
        taskId = CStr(nodeIds(nodeIndex))
        If Not network.NodeIndexById.Exists(taskId) Then Err.Raise vbObjectError + 4465, , "Missing integer node index."
        If CLng(network.NodeIndexById(taskId)) <> nodeIndex Then Err.Raise vbObjectError + 4466, , "Unstable integer node index."
        If CLng(network.RowById(taskId)) <> CLng(rowsByNode(nodeIndex)) Then Err.Raise vbObjectError + 4467, , "CALC row index mismatch."
    Next nodeIndex

    topoNodes = network.TopoNodes
    If network.LeafCount > 0 Then
        If LBound(topoNodes) <> 1 Then Err.Raise vbObjectError + 4468, , "Unexpected topology lower bound."
        If UBound(topoNodes) <> network.LeafCount Then Err.Raise vbObjectError + 4469, , "Unexpected topology upper bound."
    End If

    CompiledExecutionNetworkHarness_Smoke = _
        "PASS|Nodes=" & CStr(network.NodeCount) & _
        "|Leaves=" & CStr(network.LeafCount) & _
        "|Edges=" & CStr(network.EdgeCount) & _
        "|Components=" & CStr(network.ComponentCount)
    Exit Function

Fail:
    CompiledExecutionNetworkHarness_Smoke = "FAIL|" & CStr(Err.Number) & "|" & Err.Description

End Function

' FR:
' Retourne le predicat TEST rapide sans initialiser le workflow.
' Cette fonction est strictement en lecture seule.
'
' EN:
' Returns the fast TEST predicate without initializing the workflow.
' This function is strictly read-only.
Public Function CompiledExecutionNetworkHarness_HasTestInputFast() As Boolean

    CompiledExecutionNetworkHarness_HasTestInputFast = _
        GanttLive_HasAnyGanttTestInputFast(ThisWorkbook.Worksheets("GANTT"))

End Function
' FR:
' Execute le contrat TEST silencieux sans journaliser le message de preflight.
' Le harnais doit vider les trois colonnes jaunes sur sa copie avant l'appel.
'
' EN:
' Runs the silent TEST contract without logging the preflight message.
' The harness must clear all three yellow columns on its copy before calling it.
Public Sub CompiledExecutionNetworkHarness_RunTestNoInputSilent()

    GanttTestService_RunTestEngine True, , , , False

End Sub

' FR:
' Execute le contrat SCENARIO silencieux sans journaliser le message de preflight.
' Le harnais doit vider les trois colonnes jaunes sur sa copie avant l'appel.
'
' EN:
' Runs the silent SCENARIO contract without logging the preflight message.
' The harness must clear all three yellow columns on its copy before calling it.
Public Sub CompiledExecutionNetworkHarness_RunScenarioNoInputSilent()

    GanttScenarioService_RunScenarioEngine True, , , , False

End Sub
' FR:
' Retourne le nombre de taches detectees par le preflight incremental canonique.
' Emet une erreur si le contrat demande un recalcul complet.
'
' EN:
' Returns the task count detected by the canonical incremental preflight.
' Raises an error when the contract requires a full recalculation.
Public Function CompiledExecutionNetworkHarness_GetChangedTaskCount() As Long

    Dim changedIds As Object
    Dim forceFullRecalc As Boolean

    Set changedIds = Get_Changed_TaskIds(forceFullRecalc)
    If forceFullRecalc Then Err.Raise vbObjectError + 4470, , "Incremental preflight requested a full recalculation."
    CompiledExecutionNetworkHarness_GetChangedTaskCount = changedIds.Count

End Function



