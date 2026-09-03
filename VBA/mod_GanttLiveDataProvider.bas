Attribute VB_Name = "mod_GanttLiveDataProvider"
Option Explicit

'===============================================================================
' MODULE : mod_GanttLiveDataProvider
' DOMAINE / DOMAIN : Gantt
'
' FR
' Lit et projette les donnees necessaires au domaine sans porter sa politique metier.
' Ne rend aucune UI et ne lance aucun moteur.
'
' EN
' Reads and projects domain data without owning business policy.
' Does not render UI or run an engine.
'
' CONTRATS / CONTRACTS : GanttLive_HasAnyTestInput, GanttLive_HasAnyGanttTestInputFast, BuildGanttRowByWbsMap, ValidateGanttTestSourceColumns, FindGanttRowByWBS, GetParentIdFromWBS, GetLastGanttRow_Live, BuildCalcConstraintByIdMap_GanttLive, BuildTaskNameByIdFromWbs_Live, BuildCalcByIdMap_Live
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

Private Const GANTT_SHEET As String = "GANTT"
Private Const WBS_SHEET As String = "WBS"
Private Const WBS_TABLE As String = "tbl_WBS"
Private Const CALC_SHEET As String = "CALC"
Private Const CALC_TABLE As String = "tbl_CALC"

Private Const GANTT_FIRST_TASK_ROW As Long = 5
Private Const GANTT_COL_WBS As Long = 1
Private Const GANTT_COL_TEST_START As Long = 5
Private Const GANTT_COL_TEST_FINISH As Long = 6
Private Const GANTT_COL_TEST_PROGRESS As Long = 9


'------------------------------------------------------------------------------
' FR: Detecte si tbl_CALC_GANTT_TEST contient au moins une saisie TEST utilisateur.
' EN: Detects whether tbl_CALC_GANTT_TEST contains at least one user TEST input.
'------------------------------------------------------------------------------
Public Function GanttLive_HasAnyTestInput(ByVal tblTest As ListObject) As Boolean

    Dim arr As Variant
    Dim mapTest As Object
    Dim r As Long

    If tblTest Is Nothing Then Exit Function
    If tblTest.DataBodyRange Is Nothing Then Exit Function

    Set mapTest = CanonicalIdentity_BuildColumnMap(tblTest)
    arr = tblTest.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        If Trim$(CStr(arr(r, mapTest("Any Test Value")))) = "YES" Then
            GanttLive_HasAnyTestInput = True
            Exit Function
        End If
    Next r

End Function

'------------------------------------------------------------------------------
' FR:
' Detecte une hypothese TEST directement dans les trois colonnes editables du
' Gantt. Ce preflight ne construit aucune table de simulation et ne parcourt
' aucune shape.
'
' EN:
' Detects a TEST assumption directly in the three editable Gantt columns. This
' preflight builds no simulation table and scans no shape.
'------------------------------------------------------------------------------
Public Function GanttLive_HasAnyGanttTestInputFast(ByVal ws As Worksheet) As Boolean

    Dim lastRow As Long
    Dim startValues As Variant
    Dim finishValues As Variant
    Dim progressValues As Variant
    Dim r As Long

    If ws Is Nothing Then Exit Function

    lastRow = GetLastGanttRow_Live(ws)
    If lastRow < GANTT_FIRST_TASK_ROW Then Exit Function

    startValues = ws.Range( _
        ws.Cells(GANTT_FIRST_TASK_ROW, GANTT_COL_TEST_START), _
        ws.Cells(lastRow, GANTT_COL_TEST_START)).Value2
    finishValues = ws.Range( _
        ws.Cells(GANTT_FIRST_TASK_ROW, GANTT_COL_TEST_FINISH), _
        ws.Cells(lastRow, GANTT_COL_TEST_FINISH)).Value2
    progressValues = ws.Range( _
        ws.Cells(GANTT_FIRST_TASK_ROW, GANTT_COL_TEST_PROGRESS), _
        ws.Cells(lastRow, GANTT_COL_TEST_PROGRESS)).Value2

    For r = 1 To UBound(startValues, 1)
        If HasValue(startValues(r, 1)) Or _
           HasValue(finishValues(r, 1)) Or _
           HasValue(progressValues(r, 1)) Then
            GanttLive_HasAnyGanttTestInputFast = True
            Exit Function
        End If
    Next r

End Function


'------------------------------------------------------------------------------
' FR:
' Construit en une lecture l'index WBS normalise vers ligne Gantt utilise par
' les builders TEST et SCENARIO. L'index appartient au workflow appelant.
'
' EN:
' Builds the normalized-WBS-to-Gantt-row index used by TEST and SCENARIO
' builders in one read. The calling workflow owns the index.
'------------------------------------------------------------------------------
Public Function BuildGanttRowByWbsMap(ByVal ws As Worksheet) As Object

    Dim result As Object
    Dim lastRow As Long
    Dim values As Variant
    Dim wbsValues As Variant
    Dim tblWBS As ListObject
    Dim wbsColumnIndex As Long
    Dim r As Long
    Dim wbsVal As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    If ws Is Nothing Then
        Set BuildGanttRowByWbsMap = result
        Exit Function
    End If

    lastRow = GetLastGanttRow_Live(ws)
    If lastRow < GANTT_FIRST_TASK_ROW Then
        Set BuildGanttRowByWbsMap = result
        Exit Function
    End If

    values = ws.Range( _
        ws.Cells(GANTT_FIRST_TASK_ROW, GANTT_COL_WBS), _
        ws.Cells(lastRow, GANTT_COL_WBS)).Value2

    For r = 1 To UBound(values, 1)
        wbsVal = NormalizeWBS(CStr(values(r, 1)))
        If wbsVal <> "" Then
            If Not result.Exists(wbsVal) Then result(wbsVal) = GANTT_FIRST_TASK_ROW + r - 1
        End If
    Next r

    On Error Resume Next
    Set tblWBS = ThisWorkbook.Worksheets(WBS_SHEET).ListObjects(WBS_TABLE)
    On Error GoTo 0

    If Not tblWBS Is Nothing Then
        If Not tblWBS.DataBodyRange Is Nothing Then
            wbsColumnIndex = SchemaListColumn(tblWBS, VTS_TABLE_WBS, VTS_COL_WBS).Index
            wbsValues = tblWBS.DataBodyRange.Columns(wbsColumnIndex).Value2

            For r = 1 To UBound(wbsValues, 1)
                wbsVal = NormalizeWBS(CStr(wbsValues(r, 1)))
                If wbsVal <> "" Then
                    If Not result.Exists(wbsVal) Then result(wbsVal) = GANTT_FIRST_TASK_ROW + r - 1
                End If
            Next r
        End If
    End If

    Set BuildGanttRowByWbsMap = result

End Function

'------------------------------------------------------------------------------
' FR: Verifie les colonnes WBS/CALC indispensables a la construction d'un dataset TEST.
' EN: Validates the WBS/CALC columns required to build a TEST dataset.
'------------------------------------------------------------------------------
Public Sub ValidateGanttTestSourceColumns(ByVal mapWBS As Object, ByVal mapCalc As Object)

    Dim reqWBS As Variant
    Dim reqCalc As Variant
    Dim c As Variant

    reqWBS = Array( _
        VTS_COL_ID, VTS_COL_WBS, VTS_COL_TASK_TYPE, VTS_COL_CAL, _
        VTS_COL_CALCULATED_START, VTS_COL_CALCULATED_FINISH, VTS_COL_CALCULATED_DURATION, _
        VTS_COL_PROGRESS_PERCENT, VTS_COL_DRIVING_LOGIC, _
        VTS_COL_ACTUAL_START, VTS_COL_ACTUAL_FINISH, _
        VTS_COL_FORECAST_START, VTS_COL_FORECAST_FINISH, _
        VTS_COL_BASELINE_START, VTS_COL_BASELINE_DURATION)

    reqCalc = Array("ID", "Predecessors WBS", "Driving Logic")

    For Each c In reqWBS
        Call RequireMapColumn_GanttLive(mapWBS, "tbl_WBS", CStr(c), "ValidateGanttTestSourceColumns")
    Next c

    For Each c In reqCalc
        Call RequireMapColumn_GanttLive(mapCalc, "tbl_CALC", CStr(c), "ValidateGanttTestSourceColumns")
    Next c

End Sub



'------------------------------------------------------------------------------
' FR: Recherche dans la feuille GANTT la ligne correspondant a un WBS normalise.
' EN: Finds the GANTT sheet row matching a normalized WBS.
'------------------------------------------------------------------------------
Public Function FindGanttRowByWBS(ByVal ws As Worksheet, ByVal wbsVal As String) As Long

    Dim lastRow As Long
    Dim r As Long

    lastRow = GetLastGanttRow_Live(ws)

    For r = GANTT_FIRST_TASK_ROW To lastRow
        If NormalizeWBS(CStr(ws.Cells(r, GANTT_COL_WBS).value)) = wbsVal Then
            FindGanttRowByWBS = r
            Exit Function
        End If
    Next r

    FindGanttRowByWBS = 0

End Function



'------------------------------------------------------------------------------
' FR: Derive le WBS parent puis retourne l'ID parent via la map WBS vers ID.
' EN: Derives the parent WBS and returns the parent ID through the WBS-to-ID map.
'------------------------------------------------------------------------------
Public Function GetParentIdFromWBS(ByVal wbsVal As String, ByVal wbsToId As Object) As String

    Dim parentWbs As String

    parentWbs = GetParentWBS(wbsVal)

    If parentWbs <> "" Then
        If wbsToId.Exists(parentWbs) Then
            GetParentIdFromWBS = CStr(wbsToId(parentWbs))
            Exit Function
        End If
    End If

    GetParentIdFromWBS = ""

End Function



'------------------------------------------------------------------------------
' FR: Calcule la derniere ligne de tache visible dans GANTT a partir de la colonne WBS.
' EN: Computes the last visible task row in GANTT from the WBS column.
'------------------------------------------------------------------------------
Public Function GetLastGanttRow_Live(ByVal ws As Worksheet) As Long

    GetLastGanttRow_Live = ws.Cells(ws.rows.Count, GANTT_COL_WBS).End(xlUp).Row

    If GetLastGanttRow_Live < GANTT_FIRST_TASK_ROW Then
        GetLastGanttRow_Live = GANTT_FIRST_TASK_ROW - 1
    End If

End Function



'------------------------------------------------------------------------------
' FR: Lit CALC pour indexer par ID les contraintes actives et leurs dates utilisees par le Core live.
' EN: Reads CALC to index active constraints and dates by ID for the live Core run.
'------------------------------------------------------------------------------
Public Function BuildCalcConstraintByIdMap_GanttLive() As Object

    Dim d As Object
    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim mapCalc As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim req As Variant
    Dim c As Variant

    Set d = CreateObject("Scripting.Dictionary")
    Set wsCalc = ThisWorkbook.Worksheets(CALC_SHEET)
    Set tblCalc = wsCalc.ListObjects(CALC_TABLE)

    If tblCalc.DataBodyRange Is Nothing Then
        Set BuildCalcConstraintByIdMap_GanttLive = d
        Exit Function
    End If

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    req = Array( _
        "ID", _
        "Task Name", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date" _
    )

    For Each c In req
        Call RequireMapColumn_GanttLive(mapCalc, "tbl_CALC", CStr(c), "BuildCalcConstraintByIdMap_GanttLive")
    Next c

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapCalc("ID"))))

        If idVal <> "" Then
            d(idVal) = Array( _
                arr(r, mapCalc("Task Name")), _
                arr(r, mapCalc("Constraint Active")), _
                arr(r, mapCalc("Start Constraint Type")), _
                GetCellValue(arr(r, mapCalc("Start Constraint Date"))), _
                arr(r, mapCalc("Finish Constraint Type")), _
                GetCellValue(arr(r, mapCalc("Finish Constraint Date"))) _
            )
        End If
    Next r

    Set BuildCalcConstraintByIdMap_GanttLive = d

End Function



'------------------------------------------------------------------------------
' FR: Indexe les noms de taches WBS par ID pour enrichir les datasets et messages live.
' EN: Indexes WBS task names by ID to enrich live datasets and messages.
'------------------------------------------------------------------------------
Public Function BuildTaskNameByIdFromWbs_Live() As Object

    Dim d As Object
    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")
    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildTaskNameByIdFromWbs_Live = d
        Exit Function
    End If

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)

    Call RequireMapColumn_GanttLive(mapWBS, "tbl_WBS", VTS_COL_ID, "BuildTaskNameByIdFromWbs_Live")
    Call RequireMapColumn_GanttLive(mapWBS, "tbl_WBS", VTS_COL_TASK_NAME, "BuildTaskNameByIdFromWbs_Live")

    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapWBS(VTS_COL_ID))))
        If idVal <> "" Then d(idVal) = arr(r, mapWBS(VTS_COL_TASK_NAME))
    Next r

    Set BuildTaskNameByIdFromWbs_Live = d

End Function
'------------------------------------------------------------------------------
' FR: Indexe les predecesseurs WBS CALC par ID pour les afficher dans la table TEST.
' EN: Indexes CALC predecessor WBS text by ID for display in the TEST table.
'------------------------------------------------------------------------------
Public Function BuildCalcByIdMap_Live(ByVal tblCalc As ListObject, ByVal mapCalc As Object) As Object

    Dim d As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim colId As Long
    Dim colPred As Long

    Set d = CreateObject("Scripting.Dictionary")

    colId = RequireMapColumn_GanttLive(mapCalc, "tbl_CALC", "ID", "BuildCalcByIdMap_Live")
    colPred = RequireMapColumn_GanttLive(mapCalc, "tbl_CALC", "Predecessors WBS", "BuildCalcByIdMap_Live")

    If tblCalc.DataBodyRange Is Nothing Then
        Set BuildCalcByIdMap_Live = d
        Exit Function
    End If

    arr = tblCalc.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, colId)))
        If idVal <> "" Then
            d(idVal) = Trim$(CStr(arr(r, colPred)))
        End If
    Next r

    Set BuildCalcByIdMap_Live = d

End Function




'------------------------------------------------------------------------------
' FR: Construit la correspondance ID vers WBS normalise depuis tbl_WBS.
' EN: Builds the ID to normalized WBS lookup from tbl_WBS.
'------------------------------------------------------------------------------
Public Function BuildIdToWbsMapFromWBS_Live() As Object

    Dim d As Object
    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String
    Dim wbsVal As String

    Set d = CreateObject("Scripting.Dictionary")
    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildIdToWbsMapFromWBS_Live = d
        Exit Function
    End If

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapWBS(VTS_COL_ID))))
        wbsVal = NormalizeWBS(CStr(arr(r, mapWBS(VTS_COL_WBS))))

        If idVal <> "" And wbsVal <> "" Then
            d(idVal) = wbsVal
        End If
    Next r

    Set BuildIdToWbsMapFromWBS_Live = d

End Function



'------------------------------------------------------------------------------
' FR: Lit tbl_LOGIC_LINKS, valide FS/SS/FF et expanse les predecesseurs summary vers leurs feuilles.
' EN: Reads tbl_LOGIC_LINKS, validates FS/SS/FF, and expands summary predecessors to leaves.
'------------------------------------------------------------------------------
Private Function BuildLivePredsFromLogicLinks_FS_SS_FF_WithLagAndType( _
    ByVal validIds As Object, _
    ByVal predsById As Object, _
    ByVal childrenById As Object, _
    ByVal directChildrenById As Object, _
    ByVal predLagBySuccPred As Object, _
    ByVal predTypeBySuccPred As Object, _
    ByVal structuralErrors As Object, _
    ByVal errMissingPred As Object, _
    ByVal errUnsupportedLinkType As Object, _
    ByRef hasStructureError As Boolean) As Boolean

    Dim network As clsParsedPlanningNetwork
    Dim link As clsParsedPlanningLink
    Dim r As Long
    Dim succId As String
    Dim predId As String
    Dim linkType As String
    Dim linkKey As String
    Dim expandedLeafPreds As Collection
    Dim leafPredId As Variant

    BuildLivePredsFromLogicLinks_FS_SS_FF_WithLagAndType = False

    On Error GoTo SafeExit

    Set network = ParsedPlanningNetwork_LoadCanonical()

    If Not network.HasColumn("Succ ID") Then
        hasStructureError = True
        Exit Function
    End If

    If Not network.HasColumn("Pred ID") Then
        hasStructureError = True
        Exit Function
    End If

    If Not network.HasColumn("Link Type") Then
        hasStructureError = True
        Exit Function
    End If

    If Not network.HasColumn("Lag") Then
        hasStructureError = True
        Exit Function
    End If

    If network.Count = 0 Then
        BuildLivePredsFromLogicLinks_FS_SS_FF_WithLagAndType = True
        Exit Function
    End If

    For r = 1 To network.Count

        Set link = network.Item(r)
        succId = link.SuccId
        predId = link.PredId
        linkType = link.LinkType

        If succId = "" Then GoTo NextLinkRow
        If Not validIds.Exists(succId) Then GoTo NextLinkRow

        If predId = "" Then
            structuralErrors(succId) = True
            errMissingPred(succId) = True
            hasStructureError = True
            GoTo NextLinkRow
        End If

        If linkType <> "FS" And linkType <> "SS" And linkType <> "FF" Then
            structuralErrors(succId) = True
            errUnsupportedLinkType(succId) = True
            hasStructureError = True
            GoTo NextLinkRow
        End If

        If Not predsById.Exists(predId) Then
            structuralErrors(succId) = True
            errMissingPred(succId) = True
            hasStructureError = True
            GoTo NextLinkRow
        End If

        Set expandedLeafPreds = GetLiveLeafDescendantsFromHierarchy(predId, directChildrenById, validIds)

        If expandedLeafPreds Is Nothing Then
            structuralErrors(succId) = True
            errMissingPred(succId) = True
            hasStructureError = True
            GoTo NextLinkRow
        End If

        If expandedLeafPreds.Count = 0 Then
            structuralErrors(succId) = True
            errMissingPred(succId) = True
            hasStructureError = True
            GoTo NextLinkRow
        End If

        For Each leafPredId In expandedLeafPreds

            predsById(succId).Add CStr(leafPredId)
            childrenById(CStr(leafPredId)).Add succId

            linkKey = succId & "|" & CStr(leafPredId)
            predLagBySuccPred(linkKey) = link.Lag
            predTypeBySuccPred(linkKey) = linkType

        Next leafPredId

NextLinkRow:
    Next r

    BuildLivePredsFromLogicLinks_FS_SS_FF_WithLagAndType = True
    Exit Function

SafeExit:
    BuildLivePredsFromLogicLinks_FS_SS_FF_WithLagAndType = False

End Function



'------------------------------------------------------------------------------
' FR: Retourne les feuilles descendantes d'un ID summary ou l'ID lui-meme si c'est deja une feuille.
' EN: Returns leaf descendants of a summary ID, or the ID itself when already a leaf.
'------------------------------------------------------------------------------
Private Function GetLiveLeafDescendantsFromHierarchy( _
    ByVal startId As String, _
    ByVal directChildrenById As Object, _
    ByVal validIds As Object) As Collection

    Dim result As Collection
    Dim childId As Variant
    Dim childLeafs As Collection
    Dim leafId As Variant

    Set result = New Collection

    If validIds.Exists(startId) Then
        result.Add startId
        Set GetLiveLeafDescendantsFromHierarchy = result
        Exit Function
    End If

    If Not directChildrenById.Exists(startId) Then
        Set GetLiveLeafDescendantsFromHierarchy = result
        Exit Function
    End If

    If directChildrenById(startId).Count = 0 Then
        Set GetLiveLeafDescendantsFromHierarchy = result
        Exit Function
    End If

    For Each childId In directChildrenById(startId)

        Set childLeafs = GetLiveLeafDescendantsFromHierarchy(CStr(childId), directChildrenById, validIds)

        If Not childLeafs Is Nothing Then
            For Each leafId In childLeafs
                result.Add CStr(leafId)
            Next leafId
        End If

    Next childId

    Set GetLiveLeafDescendantsFromHierarchy = result

End Function




'------------------------------------------------------------------------------
' FR: Indexe les valeurs WBS brutes par ID pour reconstruire un dataset Core TEST fidele.
' EN: Indexes raw WBS values by ID to rebuild a faithful TEST Core dataset.
'------------------------------------------------------------------------------
Public Function BuildRawWbsSourceById_Live() As Object

    Dim d As Object
    Dim wsWBS As Worksheet
    Dim tblWBS As ListObject
    Dim mapWBS As Object
    Dim arr As Variant
    Dim r As Long
    Dim idVal As String

    Set d = CreateObject("Scripting.Dictionary")
    Set wsWBS = ThisWorkbook.Worksheets(WBS_SHEET)
    Set tblWBS = wsWBS.ListObjects(WBS_TABLE)

    If tblWBS.DataBodyRange Is Nothing Then
        Set BuildRawWbsSourceById_Live = d
        Exit Function
    End If

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    arr = tblWBS.DataBodyRange.value

    For r = 1 To UBound(arr, 1)
        idVal = Trim$(CStr(arr(r, mapWBS(VTS_COL_ID))))

        If idVal <> "" Then
            d(idVal) = Array( _
                GetCellValue(arr(r, mapWBS(VTS_COL_ACTUAL_START))), _
                GetCellValue(arr(r, mapWBS(VTS_COL_ACTUAL_FINISH))), _
                GetCellValue(arr(r, mapWBS(VTS_COL_FORECAST_START))), _
                GetCellValue(arr(r, mapWBS(VTS_COL_FORECAST_FINISH))), _
                GetCellValue(arr(r, mapWBS(VTS_COL_BASELINE_START))), _
                GetCellValue(arr(r, mapWBS(VTS_COL_BASELINE_DURATION))), _
                GetCellValue(arr(r, mapWBS(VTS_COL_PROGRESS_PERCENT))), _
                NormalizeCalendarType(arr(r, mapWBS(VTS_COL_CAL))) _
            )
        End If
    Next r

    Set BuildRawWbsSourceById_Live = d

End Function
