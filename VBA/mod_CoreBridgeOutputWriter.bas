Attribute VB_Name = "mod_CoreBridgeOutputWriter"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgeOutputWriter
' DOMAINE / DOMAIN : Core Bridge
'
' FR
' Ecrit les sorties Core Full/Partial dans CALC puis repousse les champs autorises vers WBS.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Writes Full/Partial Core outputs to CALC and then pushes authorized fields back to WBS.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : ApplyCalcDateFormats, WriteCoreOutputsToCalc_Partial, Push_Calculated_Back_To_WBS_Partial, WriteCoreOutputsToCalc, Push_Calculated_Back_To_WBS
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Actualise Apply Calc Date Formats sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply Calc Date Formats without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Public Sub ApplyCalcDateFormats(ByVal tblCalc As ListObject)

    On Error Resume Next

    tblCalc.ListColumns("Baseline Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Baseline Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Actual Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Actual Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Forecast Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Forecast Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Deadline").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Calculated Start").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Calculated Finish").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Deadline Float").DataBodyRange.NumberFormat = "0"

    On Error GoTo 0

End Sub


'------------------------------------------------------------------------------
' FR: Ecrit Core Outputs To Calc Partial vers le stockage cible.
' EN: Writes Core Outputs To Calc Partial to the target storage.
'------------------------------------------------------------------------------
Public Sub WriteCoreOutputsToCalc_Partial( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByRef dataArr As Variant, _
    ByVal impactedIds As Object)

    Dim perfScope As clsPerfScope

    Dim rowById As Object
    Dim idVal As Variant
    Dim rowIdx As Long

    Set perfScope = Profiler_BeginScope("WriteCoreOutputsToCalc_Partial", "Excel Cell Write")

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If impactedIds Is Nothing Then Exit Sub
    If impactedIds.Count = 0 Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Calculated Start") Then Exit Sub
    If Not mapCalc.Exists("Calculated Finish") Then Exit Sub
    If Not mapCalc.Exists("Calculated Duration") Then Exit Sub
    If Not mapCalc.Exists("Error flag") Then Exit Sub
    If Not mapCalc.Exists("ErrorMsg") Then Exit Sub

    Set rowById = Core_BuildRowById(dataArr, mapCalc)

    For Each idVal In impactedIds.Keys

        If rowById.Exists(CStr(idVal)) Then

            rowIdx = CLng(rowById(CStr(idVal)))

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Calculated Start")).value = _
                dataArr(rowIdx, mapCalc("Calculated Start"))

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Calculated Finish")).value = _
                dataArr(rowIdx, mapCalc("Calculated Finish"))

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Calculated Duration")).value = _
                dataArr(rowIdx, mapCalc("Calculated Duration"))

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Error flag")).value = _
                dataArr(rowIdx, mapCalc("Error flag"))

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("ErrorMsg")).value = _
                dataArr(rowIdx, mapCalc("ErrorMsg"))

        End If

    Next idVal

End Sub


'------------------------------------------------------------------------------
' FR: Pousse Calculated Back To WBS Partial vers sa table ou feuille cible.
' EN: Pushes Calculated Back To WBS Partial to its target table or sheet.
'------------------------------------------------------------------------------
Public Sub Push_Calculated_Back_To_WBS_Partial(ByVal impactedIds As Object)

    Dim perfScope As clsPerfScope

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim calcRowById As Object

    Dim allowedFields As Variant
    Dim authorizedFields As Variant
    Dim r As Long
    Dim i As Long
    Dim id As String
    Dim calcRow As Long
    Dim fieldKey As String
    Dim calcFieldName As String
    Dim writeScopeToken As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    Set perfScope = Profiler_BeginScope("Push_Calculated_Back_To_WBS_Partial", "Excel Cell Write")

    On Error GoTo SafeExit

    If impactedIds Is Nothing Then Exit Sub
    If impactedIds.Count = 0 Then Exit Sub

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")

    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set calcRowById = CreateObject("Scripting.Dictionary")

    If Not mapWBS.Exists(VTS_COL_ID) Then
        Err.Raise vbObjectError + 2101, "Push_Calculated_Back_To_WBS_Partial", _
            "Missing column in tbl_WBS: ID"
    End If

    If Not mapCalc.Exists("ID") Then
        Err.Raise vbObjectError + 2102, "Push_Calculated_Back_To_WBS_Partial", _
            "Missing column in tbl_CALC: ID"
    End If

    allowedFields = Array( _
        VTS_COL_CALCULATED_START, _
        VTS_COL_CALCULATED_FINISH, _
        VTS_COL_DRIVING_LOGIC, _
        VTS_COL_DEADLINE_FLOAT _
    )

    authorizedFields = Array( _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CALCULATED_START), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CALCULATED_FINISH), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_DRIVING_LOGIC), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_DEADLINE_FLOAT))

    For i = LBound(allowedFields) To UBound(allowedFields)

        fieldKey = CStr(allowedFields(i))
        calcFieldName = SchemaCanonicalEnglishColumnTitle(VTS_TABLE_WBS, fieldKey)

        If Not mapWBS.Exists(fieldKey) Then
            Err.Raise vbObjectError + 2110 + i, "Push_Calculated_Back_To_WBS_Partial", _
                "Missing output column in tbl_WBS: " & fieldKey
        End If

        If Not mapCalc.Exists(calcFieldName) Then
            Err.Raise vbObjectError + 2120 + i, "Push_Calculated_Back_To_WBS_Partial", _
                "Missing output column in tbl_CALC: " & calcFieldName
        End If

    Next i

    For r = 1 To tblCalc.ListRows.Count
        id = Trim$(CStr(tblCalc.DataBodyRange.Cells(r, mapCalc("ID")).value))
        If id <> "" Then
            calcRowById(id) = r
        End If
    Next r

    writeScopeToken = OpenAuthorizedWBSWriteScope( _
        "Push_Calculated_Back_To_WBS_Partial", authorizedFields)

    For r = 1 To tblWBS.ListRows.Count

        id = Trim$(CStr(tblWBS.DataBodyRange.Cells(r, mapWBS(VTS_COL_ID)).value))

        If id <> "" Then
            If impactedIds.Exists(id) Then

                If calcRowById.Exists(id) Then

                    calcRow = CLng(calcRowById(id))

                    For i = LBound(allowedFields) To UBound(allowedFields)
                        fieldKey = CStr(allowedFields(i))
                        calcFieldName = SchemaCanonicalEnglishColumnTitle(VTS_TABLE_WBS, fieldKey)
                        tblWBS.DataBodyRange.Cells(r, mapWBS(fieldKey)).value = _
                            tblCalc.DataBodyRange.Cells(calcRow, mapCalc(calcFieldName)).value
                    Next i

                Else

                    For i = LBound(allowedFields) To UBound(allowedFields)
                        fieldKey = CStr(allowedFields(i))
                        tblWBS.DataBodyRange.Cells(r, mapWBS(fieldKey)).ClearContents
                    Next i

                End If

            End If
        End If

    Next r

SafeExit:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    CloseAuthorizedWBSWriteScope writeScopeToken
    On Error GoTo 0

    If errorNumber <> 0 Then
        CalcBridge_ShowSingleConsoleMessage _
            "STOP", _
            "Erreur dans Push_Calculated_Back_To_WBS_Partial : " & errorDescription, _
            "Error in Push_Calculated_Back_To_WBS_Partial: " & errorDescription
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit en bloc les sorties du moteur Core dans tbl_CALC sans recalcul metier.
' EN: Bulk-writes Core engine outputs to tbl_CALC without business recalculation.
'------------------------------------------------------------------------------
Public Sub WriteCoreOutputsToCalc( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByRef dataArr As Variant)

    Dim perfScope As clsPerfScope

    Dim rowCount As Long
    Dim outStart() As Variant
    Dim outFinish() As Variant
    Dim outDur() As Variant
    Dim outErr() As Variant
    Dim outErrMsg() As Variant
    Dim r As Long

    Set perfScope = Profiler_BeginScope("WriteCoreOutputsToCalc", "Excel Table Write")

    rowCount = UBound(dataArr, 1)

    ReDim outStart(1 To rowCount, 1 To 1)
    ReDim outFinish(1 To rowCount, 1 To 1)
    ReDim outDur(1 To rowCount, 1 To 1)
    ReDim outErr(1 To rowCount, 1 To 1)
    ReDim outErrMsg(1 To rowCount, 1 To 1)

    For r = 1 To rowCount
        outStart(r, 1) = dataArr(r, mapCalc("Calculated Start"))
        outFinish(r, 1) = dataArr(r, mapCalc("Calculated Finish"))
        outDur(r, 1) = dataArr(r, mapCalc("Calculated Duration"))
        outErr(r, 1) = dataArr(r, mapCalc("Error flag"))
        outErrMsg(r, 1) = dataArr(r, mapCalc("ErrorMsg"))
    Next r

    tblCalc.ListColumns("Calculated Start").DataBodyRange.value = outStart
    tblCalc.ListColumns("Calculated Finish").DataBodyRange.value = outFinish
    tblCalc.ListColumns("Calculated Duration").DataBodyRange.value = outDur
    tblCalc.ListColumns("Error flag").DataBodyRange.value = outErr
    tblCalc.ListColumns("ErrorMsg").DataBodyRange.value = outErrMsg

End Sub

'------------------------------------------------------------------------------
' FR: Projette en bloc les sorties CALC autorisees vers WBS puis restaure les formules WBS gerees.
' EN: Bulk-projects authorized CALC outputs to WBS, then restores managed WBS formulas.
'------------------------------------------------------------------------------
Public Sub Push_Calculated_Back_To_WBS()

    Dim perfScope As clsPerfScope

    Dim wsWBS As Worksheet
    Dim wsCalc As Worksheet
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject

    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim calcRowById As Object

    Dim allowedFields As Variant
    Dim authorizedFields As Variant
    Dim writeScopeToken As Long
    Dim outCols As Object

    Dim arrWBS As Variant
    Dim arrCalc As Variant
    Dim outArr() As Variant

    Dim r As Long
    Dim c As Long
    Dim i As Long

    Dim wbsRows As Long
    Dim calcRows As Long

    Dim id As String
    Dim fieldKey As String
    Dim calcFieldName As String
    Dim calcRow As Long

    Dim consoleMessages As Collection
    Dim errorNumber As Long
    Dim errorDescription As String

    Set perfScope = Profiler_BeginScope("Push_Calculated_Back_To_WBS", "Excel Table Write")

    On Error GoTo SafeExit

    Set consoleMessages = New Collection

    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set wsCalc = ThisWorkbook.Worksheets("CALC")

    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")

    If tblWBS.DataBodyRange Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub

    Set mapWBS = SchemaBuildColumnKeyMap(tblWBS, VTS_TABLE_WBS)
    Set mapCalc = CreateObject("Scripting.Dictionary")
    Set calcRowById = CreateObject("Scripting.Dictionary")
    Set outCols = CreateObject("Scripting.Dictionary")

    allowedFields = Array( _
        VTS_COL_CALCULATED_START, _
        VTS_COL_CALCULATED_FINISH, _
        VTS_COL_DRIVING_LOGIC, _
        VTS_COL_CRITICAL_PATH, _
        VTS_COL_LONGEST_PATH, _
        VTS_COL_CRITICAL_PATH_REX, _
        VTS_COL_TOTAL_FLOAT, _
        VTS_COL_FREE_FLOAT, _
        VTS_COL_TOTAL_FLOAT_REX, _
        VTS_COL_FREE_FLOAT_REX, _
        VTS_COL_DEADLINE_FLOAT)

    authorizedFields = Array( _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CALCULATED_START), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CALCULATED_FINISH), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_DRIVING_LOGIC), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CRITICAL_PATH), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_LONGEST_PATH), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CRITICAL_PATH_REX), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_TOTAL_FLOAT), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_FREE_FLOAT), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_TOTAL_FLOAT_REX), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_FREE_FLOAT_REX), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_DEADLINE_FLOAT), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_BASELINE_FINISH), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_ACTUAL_DURATION), _
        SchemaCurrentColumnTitle(VTS_TABLE_WBS, VTS_COL_CALCULATED_DURATION))

    For c = 1 To tblCalc.ListColumns.Count
        mapCalc(tblCalc.ListColumns(c).Name) = c
    Next c

    If Not mapWBS.Exists(VTS_COL_ID) Then
        CoreBridgeOutputWriter_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_WBS.", _
            "Column ID was not found in tbl_WBS."
        GoTo SafeExit
    End If

    If Not mapCalc.Exists("ID") Then
        CoreBridgeOutputWriter_AddConsoleMessage consoleMessages, "STOP", _
            "La colonne ID est introuvable dans tbl_CALC.", _
            "Column ID was not found in tbl_CALC."
        GoTo SafeExit
    End If

    For i = LBound(allowedFields) To UBound(allowedFields)

        fieldKey = CStr(allowedFields(i))
        calcFieldName = SchemaCanonicalEnglishColumnTitle(VTS_TABLE_WBS, fieldKey)

        If Not mapWBS.Exists(fieldKey) Then
            CoreBridgeOutputWriter_AddConsoleMessage consoleMessages, "STOP", _
                "Colonne de sortie introuvable dans tbl_WBS : " & fieldKey, _
                "Output column not found in tbl_WBS: " & fieldKey
            GoTo SafeExit
        End If

        If Not mapCalc.Exists(calcFieldName) Then
            CoreBridgeOutputWriter_AddConsoleMessage consoleMessages, "STOP", _
                "Colonne de sortie introuvable dans tbl_CALC : " & calcFieldName, _
                "Output column not found in tbl_CALC: " & calcFieldName
            GoTo SafeExit
        End If

    Next i

    arrWBS = tblWBS.DataBodyRange.value
    arrCalc = tblCalc.DataBodyRange.value

    wbsRows = UBound(arrWBS, 1)
    calcRows = UBound(arrCalc, 1)

    For r = 1 To calcRows
        id = Trim$(CStr(arrCalc(r, mapCalc("ID"))))
        If id <> "" Then
            If Not calcRowById.Exists(id) Then
                calcRowById(id) = r
            End If
        End If
    Next r

    For i = LBound(allowedFields) To UBound(allowedFields)

        fieldKey = CStr(allowedFields(i))
        calcFieldName = SchemaCanonicalEnglishColumnTitle(VTS_TABLE_WBS, fieldKey)
        ReDim outArr(1 To wbsRows, 1 To 1)

        For r = 1 To wbsRows

            id = Trim$(CStr(arrWBS(r, mapWBS(VTS_COL_ID))))

            If id <> "" Then
                If calcRowById.Exists(id) Then
                    calcRow = CLng(calcRowById(id))
                    outArr(r, 1) = arrCalc(calcRow, mapCalc(calcFieldName))
                Else
                    outArr(r, 1) = Empty
                End If
            Else
                outArr(r, 1) = Empty
            End If

        Next r

        outCols(fieldKey) = outArr

    Next i

    writeScopeToken = OpenAuthorizedWBSWriteScope( _
        "Push_Calculated_Back_To_WBS", authorizedFields)

    For i = LBound(allowedFields) To UBound(allowedFields)
        fieldKey = CStr(allowedFields(i))
        SchemaListColumn(tblWBS, VTS_TABLE_WBS, fieldKey).DataBodyRange.value = outCols(fieldKey)
    Next i

    RestoreWBSFormulaColumns tblWBS

SafeExit:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    CloseAuthorizedWBSWriteScope writeScopeToken
    On Error GoTo 0

    If errorNumber <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        CoreBridgeOutputWriter_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur dans Push_Calculated_Back_To_WBS : " & errorDescription, _
            "Error in Push_Calculated_Back_To_WBS: " & errorDescription
    End If

    If Not consoleMessages Is Nothing Then
        CalcBridge_ShowPlanningConsole consoleMessages
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute une erreur du Full Output Writer a la collection de console.
' EN: Adds a Full Output Writer error to the console collection.
'------------------------------------------------------------------------------
Private Sub CoreBridgeOutputWriter_AddConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

End Sub
