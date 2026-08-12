Attribute VB_Name = "mod_CalcIncremental"
Option Explicit

'===============================================================================
' MODULE : mod_CalcIncremental
' DOMAINE / DOMAIN : Calculation / Data Sync
'
' FR
' Compare les signatures courantes a CALC_STATE et construit le scope de recalcul incremental.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Compares current signatures with CALC_STATE and builds the incremental recalculation scope.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Get_Changed_TaskIds, Build_Impacted_TaskIds
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const CALC_SHEET_NAME_INC As String = "CALC"
Private Const CALC_TABLE_NAME_INC As String = "tbl_CALC"


'------------------------------------------------------------------------------
' FR: Retourne Changed Task Ids depuis le contexte incremental calculation.
' EN: Returns Changed Task Ids from the incremental calculation context.
'------------------------------------------------------------------------------
Public Function Get_Changed_TaskIds(ByRef forceFullRecalc As Boolean) As Object

    Dim perfScope As clsPerfScope
    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim arrCalc As Variant
    Dim stateSigById As Object
    Dim stateRunStatus As String
    Dim changedIds As Object
    Dim r As Long
    Dim idVal As String
    Dim currentSig As String

    Set perfScope = Profiler_BeginScope("Get_Changed_TaskIds", "Incremental")
    Set changedIds = CreateObject("Scripting.Dictionary")
    forceFullRecalc = False

    On Error GoTo ForceFull
    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET_NAME_INC)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE_NAME_INC)

    If tblCalc Is Nothing Then GoTo ForceFull
    If tblCalc.DataBodyRange Is Nothing Then
        Debug.Print "INCREMENTAL CHECK - tbl_CALC empty: no changed task"
        Set Get_Changed_TaskIds = changedIds
        Exit Function
    End If

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    ValidateIncrementalCalcColumns mapCalc
    arrCalc = tblCalc.DataBodyRange.value
    Set stateSigById = CalcState_GetSignatureByIdMap()
    stateRunStatus = CalcState_GetRunStatus()

    If stateSigById.Count = 0 Then GoTo ForceFull
    If UCase$(Trim$(stateRunStatus)) <> "OK" Then GoTo ForceFull

    For r = 1 To UBound(arrCalc, 1)
        idVal = Trim$(CStr(arrCalc(r, mapCalc("ID"))))
        If idVal = "" Then GoTo NextRow
        currentSig = IncrementalSignature_BuildRow(arrCalc, r, mapCalc)
        If Not stateSigById.Exists(idVal) Then
            changedIds(idVal) = True
        ElseIf CStr(stateSigById(idVal)) <> currentSig Then
            changedIds(idVal) = True
        End If
NextRow:
    Next r

    Debug.Print "INCREMENTAL CHECK - changed tasks count: " & changedIds.Count
    Set Get_Changed_TaskIds = changedIds
    Exit Function

ForceFull:
    Debug.Print "INCREMENTAL CHECK - FULL FORCED: " & Err.Description
    forceFullRecalc = True
    Set changedIds = CreateObject("Scripting.Dictionary")
    Set Get_Changed_TaskIds = changedIds

End Function
'------------------------------------------------------------------------------
' FR: Valide Incremental Calc Columns et signale les incoherences detectees.
' EN: Validates Incremental Calc Columns and reports detected inconsistencies.
'------------------------------------------------------------------------------
Private Sub ValidateIncrementalCalcColumns(ByVal mapCalc As Object)

    Dim requiredCols As Variant
    Dim c As Variant

    requiredCols = IncrementalSignature_RequiredColumns()

    For Each c In requiredCols
        If Not mapCalc.Exists(CStr(c)) Then
            Err.Raise vbObjectError + 9201, _
                "ValidateIncrementalCalcColumns", _
                "Missing required column in tbl_CALC: " & CStr(c)
        End If
    Next c

End Sub

'------------------------------------------------------------------------------
' FR: Construit Calc Current Row Signature pour le traitement incremental calculation.
' EN: Builds Calc Current Row Signature for incremental calculation processing.

'------------------------------------------------------------------------------
' FR: Retourne Arr Val Incremental depuis le contexte incremental calculation.
' EN: Returns Arr Val Incremental from the incremental calculation context.
'------------------------------------------------------------------------------
Private Function GetArrVal_Incremental( _
    ByRef arr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal colName As String) As Variant

    If Not mapCol.Exists(colName) Then
        GetArrVal_Incremental = Empty
        Exit Function
    End If

    GetArrVal_Incremental = arr(rowIdx, mapCol(colName))

End Function

'------------------------------------------------------------------------------
' FR: Normalise Signature Token Incremental dans un format exploitable.
' EN: Normalizes Signature Token Incremental into a usable format.
'------------------------------------------------------------------------------
Private Function NormalizeSignatureToken_Incremental(ByVal v As Variant) As String

    If isError(v) Then
        NormalizeSignatureToken_Incremental = "#ERR"
        Exit Function
    End If

    If IsEmpty(v) Then
        NormalizeSignatureToken_Incremental = ""
        Exit Function
    End If

    If VarType(v) = vbNull Then
        NormalizeSignatureToken_Incremental = ""
        Exit Function
    End If

    If Trim$(CStr(v)) = "" Then
        NormalizeSignatureToken_Incremental = ""
        Exit Function
    End If

    If IsDate(v) Then
        NormalizeSignatureToken_Incremental = Format$(CDate(v), "yyyy-mm-dd")
    Else
        NormalizeSignatureToken_Incremental = Trim$(CStr(v))
    End If

End Function


'=====================================================
' Build_Impacted_TaskIds
'=====================================================
' Entrée :
' - changedIds (Dictionary of String -> True)
' - childrenByPred (predId -> children list)
' - parentById (id -> parentId)
'
' Sortie :
' - impactedIds (Dictionary)
'
' Règles :
' - inclut changed
' - inclut tous les descendants
' - inclut tous les parents (rollup)
'=====================================================


'------------------------------------------------------------------------------
' FR: Construit la collection Impacted Task IDs a partir des donnees fournies par l'appelant.
' EN: Builds the Impacted Task IDs collection from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function Build_Impacted_TaskIds( _
    ByVal changedIds As Object, _
    ByVal childrenByPred As Object, _
    ByVal parentById As Object) As Object

    Dim perfScope As clsPerfScope

    Dim impactedIds As Object
    Dim idVal As Variant
    Dim snapshotKeys As Variant
    Dim i As Long

    Set perfScope = Profiler_BeginScope("Build_Impacted_TaskIds", "Network Scan")

    Set impactedIds = CreateObject("Scripting.Dictionary")

    If changedIds Is Nothing Then
        Set Build_Impacted_TaskIds = impactedIds
        Exit Function
    End If

    ' 1. Seed = changed tasks
    For Each idVal In changedIds.Keys
        impactedIds(CStr(idVal)) = True
    Next idVal

    ' 2. Add all downstream successors / descendants in the dependency network
    For Each idVal In changedIds.Keys
        AddDescendants_Incremental CStr(idVal), childrenByPred, impactedIds
    Next idVal

    ' 3. Add parents for roll-up refresh
    If impactedIds.Count > 0 Then
        snapshotKeys = impactedIds.Keys

        For i = LBound(snapshotKeys) To UBound(snapshotKeys)
            AddParents_Incremental CStr(snapshotKeys(i)), parentById, impactedIds
        Next i
    End If

    Set Build_Impacted_TaskIds = impactedIds

End Function


'------------------------------------------------------------------------------
' FR: Ajoute Descendants a la structure cible.
' EN: Adds Descendants to the target structure.
'------------------------------------------------------------------------------
Private Sub AddDescendants( _
    ByVal rootId As String, _
    ByVal childrenByPred As Object, _
    ByVal impactedIds As Object)

    Dim stack As Object
    Dim currentId As String
    Dim child As Variant

    Set stack = CreateObject("System.Collections.ArrayList")
    stack.Add rootId

    Do While stack.Count > 0

        currentId = CStr(stack.item(stack.Count - 1))
        stack.RemoveAt stack.Count - 1

        If childrenByPred.Exists(currentId) Then
            For Each child In childrenByPred(currentId)

                If Not impactedIds.Exists(CStr(child)) Then
                    impactedIds(CStr(child)) = True
                    stack.Add CStr(child)
                End If

            Next child
        End If

    Loop

End Sub


'------------------------------------------------------------------------------
' FR: Ajoute Parents a la structure cible.
' EN: Adds Parents to the target structure.
'------------------------------------------------------------------------------
Private Sub AddParents( _
    ByVal startId As String, _
    ByVal parentById As Object, _
    ByVal impactedIds As Object)

    Dim currentId As String
    Dim parentId As String

    currentId = startId

    Do While parentById.Exists(currentId)

        parentId = CStr(parentById(currentId))
        If parentId = "" Then Exit Do

        If Not impactedIds.Exists(parentId) Then
            impactedIds(parentId) = True
        End If

        currentId = parentId

    Loop

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute Descendants Incremental a la structure cible.
' EN: Adds Descendants Incremental to the target structure.
'------------------------------------------------------------------------------
Private Sub AddDescendants_Incremental( _
    ByVal rootId As String, _
    ByVal childrenByPred As Object, _
    ByVal impactedIds As Object)

    Dim queue As Collection
    Dim currentId As String
    Dim child As Variant

    If childrenByPred Is Nothing Then Exit Sub

    Set queue = New Collection
    queue.Add rootId

    Do While queue.Count > 0

        currentId = CStr(queue(1))
        queue.Remove 1

        If childrenByPred.Exists(currentId) Then
            For Each child In childrenByPred(currentId)

                If Not impactedIds.Exists(CStr(child)) Then
                    impactedIds(CStr(child)) = True
                    queue.Add CStr(child)
                End If

            Next child
        End If

    Loop

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute Parents Incremental a la structure cible.
' EN: Adds Parents Incremental to the target structure.
'------------------------------------------------------------------------------
Private Sub AddParents_Incremental( _
    ByVal startId As String, _
    ByVal parentById As Object, _
    ByVal impactedIds As Object)

    Dim currentId As String
    Dim parentId As String

    If parentById Is Nothing Then Exit Sub

    currentId = startId

    Do While parentById.Exists(currentId)

        parentId = Trim$(CStr(parentById(currentId)))
        If parentId = "" Then Exit Do

        If Not impactedIds.Exists(parentId) Then
            impactedIds(parentId) = True
        End If

        currentId = parentId

    Loop

End Sub
