Attribute VB_Name = "mod_CoreBridgeIncrementalMap"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgeIncrementalMap
' DOMAINE / DOMAIN : Core Bridge
'
' FR
' Construit les maps de descendants et d'IDs impactes utilisees par le workflow incremental.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Builds descendant and impacted-ID maps used by the incremental workflow.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : BuildPredsBySucc_FromExpandedLinks, BuildParentByIdMap_FromCalc, Build_Successor_Map, Get_Impacted_Descendants
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Construit la map Preds By Succ From Expanded Links a partir des donnees fournies par l'appelant.
' EN: Builds the Preds By Succ From Expanded Links map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function BuildPredsBySucc_FromExpandedLinks( _
    ByVal rowById As Object, _
    ByVal linksBySuccId As Object) As Object

    Dim predsBySucc As Object
    Dim anyId As Variant
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String

    Set predsBySucc = CreateObject("Scripting.Dictionary")

    For Each anyId In rowById.Keys
        Set predsBySucc(CStr(anyId)) = New Collection
    Next anyId

    If linksBySuccId Is Nothing Then
        Set BuildPredsBySucc_FromExpandedLinks = predsBySucc
        Exit Function
    End If

    For Each succId In linksBySuccId.Keys
        If Not predsBySucc.Exists(CStr(succId)) Then
            Set predsBySucc(CStr(succId)) = New Collection
        End If

        For Each oneLink In linksBySuccId(CStr(succId))
            predId = Core_GetLinkPredId(oneLink)
            If predId <> "" Then
                predsBySucc(CStr(succId)).Add CStr(predId)
            End If
        Next oneLink
    Next succId

    Set BuildPredsBySucc_FromExpandedLinks = predsBySucc

End Function


'------------------------------------------------------------------------------
' FR: Construit l'index Parent By ID Map From CALC a partir des donnees fournies par l'appelant.
' EN: Builds the Parent By ID Map From CALC index from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function BuildParentByIdMap_FromCalc( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object) As Object

    Dim perfScope As clsPerfScope

    Dim parentById As Object
    Dim idKey As Variant
    Dim rowIdx As Long
    Dim parentId As String

    Set perfScope = Profiler_BeginScope("BuildParentByIdMap_FromCalc", "Network Build")

    Set parentById = CreateObject("Scripting.Dictionary")

    If Not mapCalc.Exists("ParentID") Then
        Set BuildParentByIdMap_FromCalc = parentById
        Exit Function
    End If

    For Each idKey In rowById.Keys
        rowIdx = CLng(rowById(CStr(idKey)))
        parentId = Trim$(CStr(dataArr(rowIdx, mapCalc("ParentID"))))

        If parentId <> "" Then
            parentById(CStr(idKey)) = parentId
        End If
    Next idKey

    Set BuildParentByIdMap_FromCalc = parentById

End Function


'------------------------------------------------------------------------------
' FR: Construit la map Successor Map a partir des donnees fournies par l'appelant.
' EN: Builds the Successor Map map from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function Build_Successor_Map(ByVal linksBySuccId As Object) As Object

    Dim perfScope As clsPerfScope

    Dim succByPred As Object
    Dim succId As Variant
    Dim oneLink As Variant
    Dim predId As String

    Set perfScope = Profiler_BeginScope("Build_Successor_Map", "Network Build")

    Set succByPred = CreateObject("Scripting.Dictionary")

    If linksBySuccId Is Nothing Then
        Set Build_Successor_Map = succByPred
        Exit Function
    End If

    For Each succId In linksBySuccId.Keys

        For Each oneLink In linksBySuccId(CStr(succId))

            predId = Core_GetLinkPredId(oneLink)

            If predId <> "" Then
                If Not succByPred.Exists(predId) Then
                    Set succByPred(predId) = New Collection
                End If

                succByPred(predId).Add CStr(succId)
            End If

        Next oneLink

    Next succId

    Set Build_Successor_Map = succByPred

End Function

'------------------------------------------------------------------------------
' FR: Retourne Impacted Descendants depuis le contexte core bridge.
' EN: Returns Impacted Descendants from the core bridge context.
'------------------------------------------------------------------------------
Public Function Get_Impacted_Descendants( _
    ByVal changedIds As Object, _
    ByVal succByPred As Object) As Object

    Dim impacted As Object
    Dim queue As Collection
    Dim idVal As Variant
    Dim currentId As String
    Dim succId As Variant

    Set impacted = CreateObject("Scripting.Dictionary")
    Set queue = New Collection

    If changedIds Is Nothing Then
        Set Get_Impacted_Descendants = impacted
        Exit Function
    End If

    For Each idVal In changedIds.Keys
        impacted(CStr(idVal)) = True
        queue.Add CStr(idVal)
    Next idVal

    Do While queue.Count > 0

        currentId = CStr(queue(1))
        queue.Remove 1

        If Not succByPred Is Nothing Then
            If succByPred.Exists(currentId) Then

                For Each succId In succByPred(currentId)

                    If Not impacted.Exists(CStr(succId)) Then
                        impacted(CStr(succId)) = True
                        queue.Add CStr(succId)
                    End If

                Next succId

            End If
        End If

    Loop

    Set Get_Impacted_Descendants = impacted

End Function


