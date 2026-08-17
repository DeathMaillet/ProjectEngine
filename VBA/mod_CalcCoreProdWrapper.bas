Attribute VB_Name = "mod_CalcCoreProdWrapper"
Option Explicit

'===============================================================================
' MODULE : mod_CalcCoreProdWrapper
' DOMAINE / DOMAIN : Core Calculation
'
' FR
' Prepare le dataset CALC et le reseau canonique avant d'appeler le moteur Core en production.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Prepares the CALC dataset and canonical network before invoking the production Core engine.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Run_Calc_Core_PROD_Pilot, FillCalcParentAndSummaryFromWBS, BuildCoreLinksBySucc_FromLogicLinksTable_Expanded
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' Wrapper PROD pilote du nouveau cœur.
'
' Cette version corrige le point critique identifié :
' - expansion des liens summary -> leaf
' - production d'un réseau final leaf -> leaf
'
' Le core reste inchangé.
'=====================================================

'------------------------------------------------------------------------------
' FR: Lance le workflow Calc Core PROD Pilot.
' EN: Runs the Calc Core PROD Pilot workflow.
'------------------------------------------------------------------------------
Public Sub Run_Calc_Core_PROD_Pilot()

    Dim wsCalc As Worksheet
    Dim wsWBS As Worksheet
    Dim tblCalc As ListObject
    Dim tblWBS As ListObject

    Dim mapCalc As Object
    Dim dataArr As Variant
    Dim linksBySuccId As Object

    On Error GoTo ErrHandler

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")

    If tblCalc.DataBodyRange Is Nothing Then

        CalcCoreProd_ShowConsoleMessage _
            "WARNING", _
            "La table tbl_CALC est vide.", _
            "Table tbl_CALC is empty."

        Exit Sub

    End If

    If tblWBS.DataBodyRange Is Nothing Then

        CalcCoreProd_ShowConsoleMessage _
            "WARNING", _
            "La table tbl_WBS est vide.", _
            "Table tbl_WBS is empty."

        Exit Sub

    End If

    BeginMacroRun "Run_Calc_Core_PROD_Pilot"

    Ensure_Calc_Infrastructure
    FillCalcParentAndSummaryFromWBS tblCalc, tblWBS

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    dataArr = tblCalc.DataBodyRange.value

    Set linksBySuccId = BuildCoreLinksBySucc_FromLogicLinksTable_Expanded(tblCalc)

    Run_Calc_Core dataArr, mapCalc, linksBySuccId

    WriteCoreOutputsToCalc tblCalc, mapCalc, dataArr

    CalcCoreProd_ShowConsoleMessage _
        "INFO", _
        "Pilot core terminé." & vbCrLf & _
        "-> résultats écrits dans tbl_CALC", _
        "Pilot core completed." & vbCrLf & _
        "-> results written to tbl_CALC"

SafeExit:
    EndMacroRun
    Exit Sub

ErrHandler:

    CalcCoreProd_ShowConsoleMessage _
        "STOP", _
        "Erreur dans Run_Calc_Core_PROD_Pilot" & vbCrLf & _
        "-> " & Err.Description, _
        "Error in Run_Calc_Core_PROD_Pilot" & vbCrLf & _
        "-> " & Err.Description

    Resume SafeExit

End Sub




'------------------------------------------------------------------------------
' FR: Alimente la map Calc Parent And Summary From WBS dans la structure cible fournie par l'appelant.
' EN: Populates the Calc Parent And Summary From WBS map in the target structure supplied by the caller.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Public Sub FillCalcParentAndSummaryFromWBS( _
    ByVal tblCalc As ListObject, _
    ByVal tblWBS As ListObject)

    Dim perfScope As clsPerfScope

    Dim mapCalc As Object
    Dim mapWBS As Object

    Dim arrCalc As Variant
    Dim arrWBS As Variant

    Dim rowByCalcId As Object
    Dim idByWBS As Object
    Dim parentIdById As Object
    Dim childCountByParent As Object

    Dim r As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim parentWbs As String
    Dim parentId As String

    Set perfScope = Profiler_BeginScope("FillCalcParentAndSummaryFromWBS", "Excel Table Write")

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)

    RequireColumn mapCalc, "ID", "tbl_CALC"
    RequireColumn mapCalc, "ParentID", "tbl_CALC"
    RequireColumn mapCalc, "IsSummary", "tbl_CALC"

    RequireColumn mapWBS, "ID", "tbl_WBS"
    RequireColumn mapWBS, "WBS", "tbl_WBS"

    arrCalc = tblCalc.DataBodyRange.value
    arrWBS = tblWBS.DataBodyRange.value

    Set rowByCalcId = CreateObject("Scripting.Dictionary")
    Set idByWBS = CreateObject("Scripting.Dictionary")
    Set parentIdById = CreateObject("Scripting.Dictionary")
    Set childCountByParent = CreateObject("Scripting.Dictionary")

    For r = 1 To UBound(arrCalc, 1)
        idVal = Trim$(CStr(arrCalc(r, mapCalc("ID"))))
        If idVal <> "" Then rowByCalcId(idVal) = r
    Next r

    For r = 1 To UBound(arrWBS, 1)
        idVal = Trim$(CStr(arrWBS(r, mapWBS("ID"))))
        wbsVal = NormalizeWBS(arrWBS(r, mapWBS("WBS")))
        If idVal <> "" And wbsVal <> "" Then idByWBS(wbsVal) = idVal
    Next r

    For r = 1 To UBound(arrWBS, 1)
        idVal = Trim$(CStr(arrWBS(r, mapWBS("ID"))))
        wbsVal = NormalizeWBS(arrWBS(r, mapWBS("WBS")))

        If idVal <> "" And wbsVal <> "" Then
            parentWbs = GetParentWBS(wbsVal)
            parentId = ""

            If parentWbs <> "" Then
                If idByWBS.Exists(parentWbs) Then
                    parentId = CStr(idByWBS(parentWbs))
                    parentIdById(idVal) = parentId
                    childCountByParent(parentId) = CLng(GetDictValueAsLong(childCountByParent, parentId)) + 1
                End If
            End If
        End If
    Next r

    For r = 1 To UBound(arrCalc, 1)
        idVal = Trim$(CStr(arrCalc(r, mapCalc("ID"))))

        If idVal <> "" Then
            If parentIdById.Exists(idVal) Then
                arrCalc(r, mapCalc("ParentID")) = CStr(parentIdById(idVal))
            Else
                arrCalc(r, mapCalc("ParentID")) = ""
            End If

            If childCountByParent.Exists(idVal) Then
                arrCalc(r, mapCalc("IsSummary")) = True
            Else
                arrCalc(r, mapCalc("IsSummary")) = False
            End If
        Else
            arrCalc(r, mapCalc("ParentID")) = ""
            arrCalc(r, mapCalc("IsSummary")) = False
        End If
    Next r

    tblCalc.ListColumns("ParentID").DataBodyRange.value = GetSingleColumnArray(arrCalc, mapCalc("ParentID"))
    tblCalc.ListColumns("IsSummary").DataBodyRange.value = GetSingleColumnArray(arrCalc, mapCalc("IsSummary"))

End Sub


'------------------------------------------------------------------------------
' FR: Construit la collection Core Links By Succ From Logic Links Table Expanded a partir des donnees fournies par l'appelant.
' EN: Builds the Core Links By Succ From Logic Links Table Expanded collection from data supplied by the caller.
'------------------------------------------------------------------------------

Public Function BuildCoreLinksBySucc_FromLogicLinksTable_Expanded( _
    ByVal tblCalc As ListObject) As Object

    Dim perfScope As clsPerfScope
    Dim mapCalc As Object
    Dim arrCalc As Variant
    Dim network As clsParsedPlanningNetwork
    Dim link As clsParsedPlanningLink
    Dim linksBySuccId As Object
    Dim rowById As Object
    Dim directChildrenById As Object
    Dim leafDescCache As Object
    Dim dedupe As Object
    Dim r As Long
    Dim succIdRaw As String
    Dim predIdRaw As String
    Dim linkType As String
    Dim lagVal As Double
    Dim succLeafs As Collection
    Dim predLeafs As Collection
    Dim succLeaf As Variant
    Dim predLeaf As Variant
    Dim predRowIdx As Long
    Dim predIsSummary As Boolean
    Dim summarySourceId As String
    Dim key As String

    Set perfScope = Profiler_BeginScope("BuildCoreLinksBySucc_FromLogicLinksTable_Expanded", "Network Build")
    Set linksBySuccId = Core_CreateLinksBySucc()
    Set dedupe = CreateObject("Scripting.Dictionary")
    Set network = ParsedPlanningNetwork_LoadCanonical()
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)

    If Not network.HasColumn("Succ ID") Then Err.Raise vbObjectError + 901, "RequireColumn", "Missing column in tbl_LOGIC_LINKS: Succ ID"
    If Not network.HasColumn("Pred ID") Then Err.Raise vbObjectError + 901, "RequireColumn", "Missing column in tbl_LOGIC_LINKS: Pred ID"
    If Not network.HasColumn("Link Type") Then Err.Raise vbObjectError + 901, "RequireColumn", "Missing column in tbl_LOGIC_LINKS: Link Type"
    If Not network.HasColumn("Lag") Then Err.Raise vbObjectError + 901, "RequireColumn", "Missing column in tbl_LOGIC_LINKS: Lag"

    RequireColumn mapCalc, "ID", "tbl_CALC"
    RequireColumn mapCalc, "ParentID", "tbl_CALC"
    RequireColumn mapCalc, "IsSummary", "tbl_CALC"

    If network.Count = 0 Then
        Set BuildCoreLinksBySucc_FromLogicLinksTable_Expanded = linksBySuccId
        Exit Function
    End If

    arrCalc = tblCalc.DataBodyRange.value

    Set rowById = Core_BuildRowById(arrCalc, mapCalc)
    Set directChildrenById = Core_BuildDirectChildrenById(arrCalc, mapCalc, rowById)
    Set leafDescCache = CreateObject("Scripting.Dictionary")

    For r = 1 To network.Count

        Set link = network.Item(r)
        succIdRaw = link.SuccId
        predIdRaw = link.PredId
        linkType = link.LinkType
        lagVal = link.Lag

        If succIdRaw <> "" And predIdRaw <> "" Then

            predIsSummary = False
            summarySourceId = ""

            If rowById.Exists(predIdRaw) Then
                predRowIdx = CLng(rowById(predIdRaw))

                If Core_IsSummaryRow(arrCalc, predRowIdx, mapCalc) Then
                    predIsSummary = True
                    summarySourceId = predIdRaw
                End If
            End If

            Set succLeafs = GetLeafTargetsForNode(succIdRaw, arrCalc, mapCalc, rowById, directChildrenById, leafDescCache)
            Set predLeafs = GetLeafTargetsForNode(predIdRaw, arrCalc, mapCalc, rowById, directChildrenById, leafDescCache)

            For Each succLeaf In succLeafs
                For Each predLeaf In predLeafs

                    key = CStr(succLeaf) & "|" & _
                          CStr(predLeaf) & "|" & _
                          linkType & "|" & _
                          CStr(lagVal) & "|" & _
                          summarySourceId

                    If Not dedupe.Exists(key) Then
                        dedupe(key) = True

                        If predIsSummary Then
                            Core_AddLink linksBySuccId, CStr(succLeaf), CStr(predLeaf), linkType, lagVal, summarySourceId
                        Else
                            Core_AddLink linksBySuccId, CStr(succLeaf), CStr(predLeaf), linkType, lagVal
                        End If
                    End If

                Next predLeaf
            Next succLeaf

        End If
    Next r

    Set BuildCoreLinksBySucc_FromLogicLinksTable_Expanded = linksBySuccId

End Function


'------------------------------------------------------------------------------
' FR: Retourne Leaf Targets For Node depuis le contexte core bridge.
' EN: Returns Leaf Targets For Node from the core bridge context.
'------------------------------------------------------------------------------
Private Function GetLeafTargetsForNode( _
    ByVal nodeId As String, _
    ByRef arrCalc As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object, _
    ByVal directChildrenById As Object, _
    ByVal leafDescCache As Object) As Collection

    Dim perfScope As clsPerfScope

    Dim result As Collection
    Dim rowIdx As Long

    Set perfScope = Profiler_BeginScope("GetLeafTargetsForNode", "Network Build")

    Set result = New Collection

    If nodeId = "" Then
        Set GetLeafTargetsForNode = result
        Exit Function
    End If

    If Not rowById.Exists(nodeId) Then
        result.Add nodeId
        Set GetLeafTargetsForNode = result
        Exit Function
    End If

    rowIdx = CLng(rowById(nodeId))

    If Core_IsSummaryRow(arrCalc, rowIdx, mapCalc) Then
        Set result = CoreWrapper_GetLeafDescendantsFromCalc(nodeId, arrCalc, mapCalc, rowById, directChildrenById, leafDescCache)
    Else
        result.Add nodeId
    End If

    Set GetLeafTargetsForNode = result

End Function


'------------------------------------------------------------------------------
' FR: Retourne la collection Leaf Descendants From CALC sans exposer de mutateur sur l'etat source.
' EN: Returns the Leaf Descendants From CALC collection without exposing a mutator for source state.
'------------------------------------------------------------------------------

Private Function CoreWrapper_GetLeafDescendantsFromCalc( _
    ByVal startId As String, _
    ByRef arrCalc As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowById As Object, _
    ByVal directChildrenById As Object, _
    ByVal leafDescCache As Object) As Collection

    Dim perfScope As clsPerfScope

    Dim result As Collection
    Dim q As Collection
    Dim currentId As Variant
    Dim childId As Variant
    Dim rowIdx As Long

    Set perfScope = Profiler_BeginScope("CoreWrapper_GetLeafDescendantsFromCalc", "Network Scan")

    If leafDescCache.Exists(startId) Then
        Set CoreWrapper_GetLeafDescendantsFromCalc = CloneStringCollection(leafDescCache(startId))
        Exit Function
    End If

    Set result = New Collection
    Set q = New Collection
    q.Add startId

    Do While q.Count > 0
        currentId = q(1)
        q.Remove 1

        If rowById.Exists(CStr(currentId)) Then
            rowIdx = CLng(rowById(CStr(currentId)))

            If Core_IsSummaryRow(arrCalc, rowIdx, mapCalc) Then
                If directChildrenById.Exists(CStr(currentId)) Then
                    For Each childId In directChildrenById(CStr(currentId))
                        q.Add CStr(childId)
                    Next childId
                End If
            Else
                result.Add CStr(currentId)
            End If
        End If
    Loop

    Set leafDescCache(startId) = CloneStringCollection(result)
    Set CoreWrapper_GetLeafDescendantsFromCalc = result

End Function

'------------------------------------------------------------------------------
' FR: Construit la collection String Collection a partir des donnees fournies par l'appelant.
' EN: Builds the String Collection collection from data supplied by the caller.
'------------------------------------------------------------------------------

Private Function CloneStringCollection(ByVal sourceCol As Collection) As Collection

    Dim result As Collection
    Dim item As Variant

    Set result = New Collection

    For Each item In sourceCol
        result.Add CStr(item)
    Next item

    Set CloneStringCollection = result

End Function


'------------------------------------------------------------------------------
' FR: Valide Require Column et applique la politique d'erreur definie par le composant.
' EN: Validates Require Column and applies the component's defined failure policy.
'------------------------------------------------------------------------------

Private Sub RequireColumn( _
    ByVal mapCol As Object, _
    ByVal colName As String, _
    ByVal tableName As String)

    If Not mapCol.Exists(colName) Then
        Err.Raise vbObjectError + 901, "RequireColumn", _
            "Missing column in " & tableName & ": " & colName
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Retourne Single Column Array depuis le contexte core bridge.
' EN: Returns Single Column Array from the core bridge context.
'------------------------------------------------------------------------------
Private Function GetSingleColumnArray( _
    ByRef dataArr As Variant, _
    ByVal colIdx As Long) As Variant

    Dim outArr() As Variant
    Dim r As Long
    Dim rowCount As Long

    rowCount = UBound(dataArr, 1)
    ReDim outArr(1 To rowCount, 1 To 1)

    For r = 1 To rowCount
        outArr(r, 1) = dataArr(r, colIdx)
    Next r

    GetSingleColumnArray = outArr

End Function

'------------------------------------------------------------------------------
' FR: Retourne Dict Value As Long depuis le contexte core bridge.
' EN: Returns Dict Value As Long from the core bridge context.
'------------------------------------------------------------------------------
Private Function GetDictValueAsLong( _
    ByVal dictObj As Object, _
    ByVal key As String) As Long

    If dictObj.Exists(key) Then
        GetDictValueAsLong = CLng(dictObj(key))
    Else
        GetDictValueAsLong = 0
    End If

End Function

'------------------------------------------------------------------------------
' FR: Projette la collection Calc Core Prod Show Console Message vers l'interface autorisee par la politique runtime.
' EN: Projects the Calc Core Prod Show Console Message collection to the UI allowed by runtime policy.
'------------------------------------------------------------------------------

Private Sub CalcCoreProd_ShowConsoleMessage( _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    Dim consoleMessages As Collection

    Set consoleMessages = New Collection

    CalcBridge_AddConsoleMessage consoleMessages, _
        msgType, _
        "FR:" & vbCrLf & _
        frText & vbCrLf & vbCrLf & _
        "EN:" & vbCrLf & _
        enText

    CalcBridge_ShowPlanningConsole consoleMessages

End Sub
