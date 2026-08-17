Attribute VB_Name = "mod_CanonicalIdentityIndex"
Option Explicit

'===============================================================================
' MODULE : mod_CanonicalIdentityIndex
' DOMAINE / DOMAIN : Canonical Read Models
'
' FR
' Possede les index immuables ID/WBS/lignes et la map Driving Logic partages par les lecteurs canoniques.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Owns immutable ID/WBS/row indexes and the shared Driving Logic map used by canonical readers.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : CanonicalIdentity_BuildColumnMap, CanonicalIdentity_BuildWbsToIdMap, CanonicalIdentity_GetWbsToIdMap, CanonicalIdentity_BuildCalcRowById, CanonicalIdentity_GetDrivingLogicByIdMap, CanonicalIdentity_GetWbsRowByIdMap, CanonicalIdentity_GetCalcRowByIdMap, CanonicalIdentity_GetWbsByIdMap
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' Canonical structural read model for tbl_WBS and tbl_CALC identities.
' It owns normalization and indexing only: no business calculation, UI or write.

Private Const CANONICAL_WBS_SHEET As String = "WBS"
Private Const CANONICAL_WBS_TABLE As String = "tbl_WBS"
Private Const CANONICAL_CALC_SHEET As String = "CALC"
Private Const CANONICAL_CALC_TABLE As String = "tbl_CALC"

'------------------------------------------------------------------------------
' FR: Construit une map immuable nom de colonne vers index pour une table.
' EN: Builds an immutable column-name-to-index map for a table.
'------------------------------------------------------------------------------
Public Function CanonicalIdentity_BuildColumnMap(ByVal tbl As ListObject) As Object

    Dim perfScope As clsPerfScope
    Dim values As Object
    Dim i As Long

    Set perfScope = Profiler_BeginScope("CanonicalIdentity_BuildColumnMap", "Excel Metadata")
    Set values = CreateObject("Scripting.Dictionary")

    For i = 1 To tbl.ListColumns.Count
        values(tbl.ListColumns(i).Name) = i
    Next i

    Set CanonicalIdentity_BuildColumnMap = CanonicalIdentity_Freeze(values)

End Function

'------------------------------------------------------------------------------
' FR: Construit l'index immuable WBS normalise vers ID avec politique last-row-wins.
' EN: Builds the immutable normalized-WBS-to-ID index with last-row-wins policy.
'------------------------------------------------------------------------------
Public Function CanonicalIdentity_BuildWbsToIdMap( _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object) As Object

    Dim perfScope As clsPerfScope
    Dim values As Object
    Dim arr As Variant
    Dim r As Long
    Dim wbsVal As String
    Dim idVal As String

    Set perfScope = Profiler_BeginScope("CanonicalIdentity_BuildWbsToIdMap", "Dictionary")
    Set values = CreateObject("Scripting.Dictionary")

    If Not tblWBS.DataBodyRange Is Nothing Then
        arr = tblWBS.DataBodyRange.value

        For r = 1 To UBound(arr, 1)
            wbsVal = NormalizeWBS(CStr(arr(r, mapWBS("WBS"))))
            idVal = Trim$(CStr(arr(r, mapWBS("ID"))))

            If wbsVal <> "" And idVal <> "" Then
                values(wbsVal) = idVal
            End If
        Next r
    End If

    Set CanonicalIdentity_BuildWbsToIdMap = CanonicalIdentity_Freeze(values)

End Function

'------------------------------------------------------------------------------
' FR: Lit tbl_WBS et retourne son index immuable WBS normalise vers ID.
' EN: Reads tbl_WBS and returns its immutable normalized-WBS-to-ID index.
'------------------------------------------------------------------------------
Public Function CanonicalIdentity_GetWbsToIdMap() As Object

    Dim tblWBS As ListObject
    Dim mapWBS As Object

    Set tblWBS = ThisWorkbook.Worksheets(CANONICAL_WBS_SHEET).ListObjects(CANONICAL_WBS_TABLE)
    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    Set CanonicalIdentity_GetWbsToIdMap = CanonicalIdentity_BuildWbsToIdMap(tblWBS, mapWBS)

End Function

'------------------------------------------------------------------------------
' FR: Construit l'index immuable ID vers numero de ligne d'un dataset CALC.
' EN: Builds the immutable ID-to-row-number index for a CALC dataset.
'------------------------------------------------------------------------------
Public Function CanonicalIdentity_BuildCalcRowById( _
    ByRef arrCalc As Variant, _
    ByVal mapCalc As Object) As Object

    Dim perfScope As clsPerfScope
    Dim values As Object
    Dim r As Long
    Dim idVal As String

    Set perfScope = Profiler_BeginScope("CanonicalIdentity_BuildCalcRowById", "Dictionary")
    Set values = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arrCalc, 1)
        idVal = Trim$(CStr(arrCalc(r, mapCalc("ID"))))
        If idVal <> "" Then values(idVal) = r
    Next r

    Set CanonicalIdentity_BuildCalcRowById = CanonicalIdentity_Freeze(values)

End Function

'------------------------------------------------------------------------------
' FR: Lit tbl_CALC et indexe sa Driving Logic normalisee par ID.
' EN: Reads tbl_CALC and indexes its normalized Driving Logic by ID.
'------------------------------------------------------------------------------
Public Function CanonicalIdentity_GetDrivingLogicByIdMap() As Object

    Dim perfScope As clsPerfScope
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim values As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set perfScope = Profiler_BeginScope("CanonicalIdentity_GetDrivingLogicByIdMap", "Excel Read")
    Set values = CreateObject("Scripting.Dictionary")
    Set tblCalc = ThisWorkbook.Worksheets(CANONICAL_CALC_SHEET).ListObjects(CANONICAL_CALC_TABLE)

    If Not tblCalc.DataBodyRange Is Nothing Then
        Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
        arr = tblCalc.DataBodyRange.value

        For r = 1 To UBound(arr, 1)
            idVal = Trim$(CStr(arr(r, mapCalc("ID"))))
            If idVal <> "" Then
                values(idVal) = UCase$(Trim$(CStr(arr(r, mapCalc("Driving Logic")))))
            End If
        Next r
    End If

    Set CanonicalIdentity_GetDrivingLogicByIdMap = CanonicalIdentity_Freeze(values)

End Function

'------------------------------------------------------------------------------
' FR: Lit tbl_WBS et indexe le numero de ligne de chaque ID non vide.
' EN: Reads tbl_WBS and indexes the row number of each nonblank ID.
'------------------------------------------------------------------------------
Public Function CanonicalIdentity_GetWbsRowByIdMap() As Object

    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim arr As Variant

    Set tblWBS = ThisWorkbook.Worksheets(CANONICAL_WBS_SHEET).ListObjects(CANONICAL_WBS_TABLE)
    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)

    If tblWBS.DataBodyRange Is Nothing Then
        Set CanonicalIdentity_GetWbsRowByIdMap = CanonicalIdentity_EmptyMap()
    Else
        arr = tblWBS.DataBodyRange.value
        Set CanonicalIdentity_GetWbsRowByIdMap = CanonicalIdentity_BuildRowById(arr, mapWBS)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Lit tbl_CALC et indexe le numero de ligne de chaque ID non vide.
' EN: Reads tbl_CALC and indexes the row number of each nonblank ID.
'------------------------------------------------------------------------------
Public Function CanonicalIdentity_GetCalcRowByIdMap() As Object

    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim arr As Variant

    Set tblCalc = ThisWorkbook.Worksheets(CANONICAL_CALC_SHEET).ListObjects(CANONICAL_CALC_TABLE)
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)

    If tblCalc.DataBodyRange Is Nothing Then
        Set CanonicalIdentity_GetCalcRowByIdMap = CanonicalIdentity_EmptyMap()
    Else
        arr = tblCalc.DataBodyRange.value
        Set CanonicalIdentity_GetCalcRowByIdMap = CanonicalIdentity_BuildRowById(arr, mapCalc)
    End If

End Function

'------------------------------------------------------------------------------
' FR: Lit tbl_WBS et indexe son WBS normalise par ID, y compris un WBS vide.
' EN: Reads tbl_WBS and indexes normalized WBS by ID, including a blank WBS.
'------------------------------------------------------------------------------
Public Function CanonicalIdentity_GetWbsByIdMap() As Object

    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim values As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set values = CreateObject("Scripting.Dictionary")
    Set tblWBS = ThisWorkbook.Worksheets(CANONICAL_WBS_SHEET).ListObjects(CANONICAL_WBS_TABLE)

    If Not tblWBS.DataBodyRange Is Nothing Then
        Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
        arr = tblWBS.DataBodyRange.value

        For r = 1 To UBound(arr, 1)
            idVal = Trim$(CStr(arr(r, mapWBS("ID"))))
            If idVal <> "" Then values(idVal) = NormalizeWBS(CStr(arr(r, mapWBS("WBS"))))
        Next r
    End If

    Set CanonicalIdentity_GetWbsByIdMap = CanonicalIdentity_Freeze(values)

End Function

'------------------------------------------------------------------------------
' FR: Indique si la condition Build Row By Id est satisfaite sans modifier les donnees source.
' EN: Returns whether the Build Row By Id condition is satisfied without mutating source data.
'------------------------------------------------------------------------------

Private Function CanonicalIdentity_BuildRowById( _
    ByRef sourceData As Variant, _
    ByVal columnMap As Object) As Object

    Set CanonicalIdentity_BuildRowById = CanonicalIdentity_BuildCalcRowById(sourceData, columnMap)

End Function

'------------------------------------------------------------------------------
' FR: Indique si la condition Empty Map est satisfaite sans modifier les donnees source.
' EN: Returns whether the Empty Map condition is satisfied without mutating source data.
'------------------------------------------------------------------------------

Private Function CanonicalIdentity_EmptyMap() As Object

    Dim values As Object

    Set values = CreateObject("Scripting.Dictionary")
    Set CanonicalIdentity_EmptyMap = CanonicalIdentity_Freeze(values)

End Function

'------------------------------------------------------------------------------
' FR: Indique si la condition Freeze est satisfaite sans modifier les donnees source.
' EN: Returns whether the Freeze condition is satisfied without mutating source data.
'------------------------------------------------------------------------------

Private Function CanonicalIdentity_Freeze(ByVal values As Object) As Object

    Dim readOnlyMap As clsCanonicalReadOnlyMap

    Set readOnlyMap = New clsCanonicalReadOnlyMap
    readOnlyMap.InitializeFrom values
    Set CanonicalIdentity_Freeze = readOnlyMap

End Function
