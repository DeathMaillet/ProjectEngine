Attribute VB_Name = "mod_ArrayIndexedCoreHarness"
Option Explicit

'===============================================================================
' MODULE : mod_ArrayIndexedCoreHarness
' DOMAINE / DOMAIN : Core Calculation / Permanent Harness
'
' FR
' Fixtures deterministes en memoire protegeant la migration du Core vers les
' arrays indexes. Aucun classeur metier n'est modifie.
'
' EN
' Deterministic in-memory fixtures protecting the array-indexed Core migration.
' No business workbook is modified.
'
' CONTRATS / CONTRACTS : ArrayIndexedCoreSemanticHarness_Run
'===============================================================================

Private Const HARNESS_CAL As String = "7j/7"

Public Function ArrayIndexedCoreSemanticHarness_Run(ByVal outputPath As String) As String

    Dim fileNo As Integer
    Dim failures As Long
    Dim linkType As Variant
    Dim lagValue As Variant
    Dim caseName As String

    On Error GoTo Fail

    fileNo = FreeFile
    Open outputPath For Output As #fileNo
    Print #fileNo, "Case" & vbTab & "Domain" & vbTab & "Legacy" & vbTab & _
        "Indexed" & vbTab & "Expected" & vbTab & "Attributes" & vbTab & "Status" & vbTab & "Details"

    For Each linkType In Array("FS", "SS", "FF")
        For Each lagValue In Array(-2#, 0#, 2#)
            caseName = CStr(linkType) & IIf(CDbl(lagValue) > 0, "+", "") & CStr(lagValue)
            If Not ArrayIndexedCoreHarness_RunAtomicLinkCase( _
                fileNo, caseName, CStr(linkType), CDbl(lagValue)) Then
                failures = failures + 1
            End If
        Next lagValue
    Next linkType

    If Not ArrayIndexedCoreHarness_RunSourcePolicyCase(fileNo, "BASELINE_ONLY", 0) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunSourcePolicyCase(fileNo, "FORECAST_OVER_BASELINE", 1) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunSourcePolicyCase(fileNo, "ACTUAL_OVER_FORECAST", 2) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunMilestoneCase(fileNo) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunMultiplePredecessorCase(fileNo, "FS_SS_FF_MIX") Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunMilestoneLinkCase(fileNo, "MILESTONE_PRED_FS", True, False, "FS", 2#) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunMilestoneLinkCase(fileNo, "MILESTONE_SUCC_SS", False, True, "SS", -1#) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunMilestoneLinkCase(fileNo, "MILESTONE_TO_MILESTONE_FF", True, True, "FF", 1#) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunLOECase(fileNo, "LOE_VALID", 0) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunLOECase(fileNo, "LOE_MISSING_SS", 1) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunLOECase(fileNo, "LOE_MISSING_FF", 2) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunDiagnosticCase(fileNo, "MISSING_PREDECESSOR", 1) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunDiagnosticCase(fileNo, "SIMPLE_CYCLE", 2) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunDiagnosticCase(fileNo, "SELF_CYCLE", 3) Then failures = failures + 1
    If Not ArrayIndexedCoreHarness_RunParentExpansionMetadataCase(fileNo) Then failures = failures + 1

    Close #fileNo

    If failures = 0 Then
        ArrayIndexedCoreSemanticHarness_Run = "PASS"
    Else
        ArrayIndexedCoreSemanticHarness_Run = "FAIL|" & CStr(failures)
    End If
    Exit Function

Fail:
    On Error Resume Next
    If fileNo > 0 Then Close #fileNo
    ArrayIndexedCoreSemanticHarness_Run = "FAIL|" & CStr(Err.Number) & "|" & Err.Description

End Function

Private Function ArrayIndexedCoreHarness_RunAtomicLinkCase( _
    ByVal fileNo As Integer, _
    ByVal caseName As String, _
    ByVal linkType As String, _
    ByVal lagValue As Double) As Boolean

    Dim mapCol As Object
    Dim sourceData As Variant
    Dim legacyData As Variant
    Dim indexedData As Variant
    Dim linksBySuccId As Object
    Dim network As clsCompiledExecutionNetwork
    Dim expectedStart As Variant
    Dim expectedFinish As Variant
    Dim predStart As Double
    Dim predFinish As Double
    Dim durationValue As Double
    Dim parity As Boolean
    Dim expectedPass As Boolean
    Dim attributesPass As Boolean
    Dim details As String

    On Error GoTo Fail

    Set mapCol = ArrayIndexedCoreHarness_BuildColumnMap()
    sourceData = ArrayIndexedCoreHarness_BuildBaseData(mapCol, 2)
    predStart = CDbl(DateSerial(2026, 1, 5))
    durationValue = 2#

    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 1, "P", "1", "Task", predStart, 3#
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 2, "S", "2", "Task", CDbl(DateSerial(2026, 1, 1)), durationValue

    Set linksBySuccId = Core_CreateLinksBySucc()
    Core_AddLink linksBySuccId, "S", "P", linkType, lagValue

    legacyData = sourceData
    Run_Calc_Core legacyData, mapCol, linksBySuccId

    indexedData = sourceData
    Set network = CompileExecutionNetwork(indexedData, mapCol, linksBySuccId)
    Run_Calc_Core indexedData, mapCol, linksBySuccId, , , , , network

    predFinish = CDbl(legacyData(1, mapCol("Calculated Finish")))
    Select Case linkType
        Case "SS"
            expectedStart = ApplyLag(predStart, lagValue, HARNESS_CAL, "SS")
            expectedFinish = AddWorkingDays(expectedStart, durationValue, HARNESS_CAL)
        Case "FF"
            expectedFinish = ApplyLag(predFinish, lagValue, HARNESS_CAL, "FF")
            expectedStart = SubtractWorkingDays(expectedFinish, durationValue, HARNESS_CAL)
        Case Else
            expectedStart = ApplyLag(predFinish, lagValue, HARNESS_CAL, "FS")
            expectedFinish = AddWorkingDays(expectedStart, durationValue, HARNESS_CAL)
    End Select

    parity = ArrayIndexedCoreHarness_OutputRowsEqual(legacyData, indexedData, mapCol, 1, 2)
    expectedPass = ArrayIndexedCoreHarness_ValueEqual(indexedData(2, mapCol("Calculated Start")), expectedStart) And _
        ArrayIndexedCoreHarness_ValueEqual(indexedData(2, mapCol("Calculated Finish")), expectedFinish) And _
        ArrayIndexedCoreHarness_ValueEqual(indexedData(2, mapCol("Calculated Duration")), durationValue)
    attributesPass = ArrayIndexedCoreHarness_CompiledEdgeMatches(network, "S", "P", linkType, lagValue)
    details = "Edges=" & CStr(network.EdgeCount)

    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "ATOMIC_LINK", _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 2), _
        ArrayIndexedCoreHarness_RowResult(indexedData, mapCol, 2), _
        ArrayIndexedCoreHarness_FormatPair(expectedStart, expectedFinish), _
        linkType & "|" & CStr(lagValue), parity And expectedPass And attributesPass, details

    ArrayIndexedCoreHarness_RunAtomicLinkCase = parity And expectedPass And attributesPass
    Exit Function

Fail:
    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "ATOMIC_LINK", "", "", "", _
        linkType & "|" & CStr(lagValue), False, CStr(Err.Number) & "|" & Err.Description

End Function

Private Function ArrayIndexedCoreHarness_RunSourcePolicyCase( _
    ByVal fileNo As Integer, _
    ByVal caseName As String, _
    ByVal policyMode As Long) As Boolean

    Dim mapCol As Object
    Dim sourceData As Variant
    Dim legacyData As Variant
    Dim indexedData As Variant
    Dim linksBySuccId As Object
    Dim network As clsCompiledExecutionNetwork
    Dim expectedStart As Double
    Dim expectedFinish As Double
    Dim parity As Boolean
    Dim expectedPass As Boolean

    On Error GoTo Fail

    Set mapCol = ArrayIndexedCoreHarness_BuildColumnMap()
    sourceData = ArrayIndexedCoreHarness_BuildBaseData(mapCol, 1)
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 1, "A", "1", "Task", CDbl(DateSerial(2026, 1, 1)), 3#

    expectedStart = CDbl(DateSerial(2026, 1, 1))
    expectedFinish = CDbl(DateSerial(2026, 1, 3))

    If policyMode >= 1 Then
        sourceData(1, mapCol("Forecast Start")) = CDbl(DateSerial(2026, 2, 2))
        sourceData(1, mapCol("Forecast Finish")) = CDbl(DateSerial(2026, 2, 4))
        expectedStart = CDbl(DateSerial(2026, 2, 2))
        expectedFinish = CDbl(DateSerial(2026, 2, 4))
    End If
    If policyMode >= 2 Then
        sourceData(1, mapCol("Actual Start")) = CDbl(DateSerial(2026, 3, 3))
        sourceData(1, mapCol("Actual Finish")) = CDbl(DateSerial(2026, 3, 5))
        expectedStart = CDbl(DateSerial(2026, 3, 3))
        expectedFinish = CDbl(DateSerial(2026, 3, 5))
    End If

    Set linksBySuccId = Core_CreateLinksBySucc()
    legacyData = sourceData
    Run_Calc_Core legacyData, mapCol, linksBySuccId
    indexedData = sourceData
    Set network = CompileExecutionNetwork(indexedData, mapCol, linksBySuccId)
    Run_Calc_Core indexedData, mapCol, linksBySuccId, , , , , network

    parity = ArrayIndexedCoreHarness_OutputRowsEqual(legacyData, indexedData, mapCol, 1, 1)
    expectedPass = ArrayIndexedCoreHarness_ValueEqual(indexedData(1, mapCol("Calculated Start")), expectedStart) And _
        ArrayIndexedCoreHarness_ValueEqual(indexedData(1, mapCol("Calculated Finish")), expectedFinish)

    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "SOURCE_POLICY", _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 1), _
        ArrayIndexedCoreHarness_RowResult(indexedData, mapCol, 1), _
        ArrayIndexedCoreHarness_FormatPair(expectedStart, expectedFinish), "", _
        parity And expectedPass, ""
    ArrayIndexedCoreHarness_RunSourcePolicyCase = parity And expectedPass
    Exit Function

Fail:
    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "SOURCE_POLICY", "", "", "", "", False, _
        CStr(Err.Number) & "|" & Err.Description

End Function

Private Function ArrayIndexedCoreHarness_RunMilestoneCase(ByVal fileNo As Integer) As Boolean

    Dim mapCol As Object
    Dim sourceData As Variant
    Dim legacyData As Variant
    Dim indexedData As Variant
    Dim linksBySuccId As Object
    Dim network As clsCompiledExecutionNetwork
    Dim expectedDate As Double
    Dim passed As Boolean

    On Error GoTo Fail

    Set mapCol = ArrayIndexedCoreHarness_BuildColumnMap()
    sourceData = ArrayIndexedCoreHarness_BuildBaseData(mapCol, 1)
    expectedDate = CDbl(DateSerial(2026, 4, 6))
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 1, "M", "1", "Milestone", expectedDate, Empty
    Set linksBySuccId = Core_CreateLinksBySucc()

    legacyData = sourceData
    Run_Calc_Core legacyData, mapCol, linksBySuccId
    indexedData = sourceData
    Set network = CompileExecutionNetwork(indexedData, mapCol, linksBySuccId)
    Run_Calc_Core indexedData, mapCol, linksBySuccId, , , , , network

    passed = ArrayIndexedCoreHarness_OutputRowsEqual(legacyData, indexedData, mapCol, 1, 1) And _
        ArrayIndexedCoreHarness_ValueEqual(indexedData(1, mapCol("Calculated Start")), expectedDate) And _
        ArrayIndexedCoreHarness_ValueEqual(indexedData(1, mapCol("Calculated Finish")), expectedDate) And _
        ArrayIndexedCoreHarness_ValueEqual(indexedData(1, mapCol("Calculated Duration")), 1#)

    ArrayIndexedCoreHarness_WriteResult fileNo, "MILESTONE_BASELINE", "MILESTONE", _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 1), _
        ArrayIndexedCoreHarness_RowResult(indexedData, mapCol, 1), _
        ArrayIndexedCoreHarness_FormatPair(expectedDate, expectedDate), "", passed, ""
    ArrayIndexedCoreHarness_RunMilestoneCase = passed
    Exit Function

Fail:
    ArrayIndexedCoreHarness_WriteResult fileNo, "MILESTONE_BASELINE", "MILESTONE", "", "", "", "", False, _
        CStr(Err.Number) & "|" & Err.Description

End Function

Private Function ArrayIndexedCoreHarness_RunMultiplePredecessorCase( _
    ByVal fileNo As Integer, _
    ByVal caseName As String) As Boolean

    Dim mapCol As Object
    Dim sourceData As Variant
    Dim legacyData As Variant
    Dim indexedData As Variant
    Dim linksBySuccId As Object
    Dim network As clsCompiledExecutionNetwork
    Dim parity As Boolean

    On Error GoTo Fail

    Set mapCol = ArrayIndexedCoreHarness_BuildColumnMap()
    sourceData = ArrayIndexedCoreHarness_BuildBaseData(mapCol, 4)
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 1, "P1", "1", "Task", CDbl(DateSerial(2026, 1, 1)), 3#
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 2, "P2", "2", "Task", CDbl(DateSerial(2026, 1, 4)), 2#
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 3, "P3", "3", "Task", CDbl(DateSerial(2026, 1, 7)), 2#
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 4, "S", "4", "Task", CDbl(DateSerial(2026, 1, 1)), 3#

    Set linksBySuccId = Core_CreateLinksBySucc()
    Core_AddLink linksBySuccId, "S", "P1", "FS", -1
    Core_AddLink linksBySuccId, "S", "P2", "SS", 2
    Core_AddLink linksBySuccId, "S", "P3", "FF", 1

    legacyData = sourceData
    Run_Calc_Core legacyData, mapCol, linksBySuccId
    indexedData = sourceData
    Set network = CompileExecutionNetwork(indexedData, mapCol, linksBySuccId)
    Run_Calc_Core indexedData, mapCol, linksBySuccId, , , , , network

    parity = ArrayIndexedCoreHarness_OutputRowsEqual(legacyData, indexedData, mapCol, 1, 4)
    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "MULTI_PREDECESSOR", _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 4), _
        ArrayIndexedCoreHarness_RowResult(indexedData, mapCol, 4), _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 4), _
        "FS-1|SS+2|FF+1", parity, "Edges=" & CStr(network.EdgeCount)
    ArrayIndexedCoreHarness_RunMultiplePredecessorCase = parity
    Exit Function

Fail:
    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "MULTI_PREDECESSOR", "", "", "", _
        "FS-1|SS+2|FF+1", False, CStr(Err.Number) & "|" & Err.Description

End Function

Private Function ArrayIndexedCoreHarness_RunMilestoneLinkCase( _
    ByVal fileNo As Integer, _
    ByVal caseName As String, _
    ByVal predecessorMilestone As Boolean, _
    ByVal successorMilestone As Boolean, _
    ByVal linkType As String, _
    ByVal lagValue As Double) As Boolean

    Dim mapCol As Object
    Dim sourceData As Variant
    Dim legacyData As Variant
    Dim indexedData As Variant
    Dim linksBySuccId As Object
    Dim network As clsCompiledExecutionNetwork
    Dim passed As Boolean

    On Error GoTo Fail
    Set mapCol = ArrayIndexedCoreHarness_BuildColumnMap()
    sourceData = ArrayIndexedCoreHarness_BuildBaseData(mapCol, 2)
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 1, "MP", "1", _
        IIf(predecessorMilestone, "Milestone", "Task"), CDbl(DateSerial(2026, 5, 4)), _
        IIf(predecessorMilestone, Empty, 3#)
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 2, "MS", "2", _
        IIf(successorMilestone, "Milestone", "Task"), CDbl(DateSerial(2026, 5, 1)), _
        IIf(successorMilestone, Empty, 2#)

    Set linksBySuccId = Core_CreateLinksBySucc()
    Core_AddLink linksBySuccId, "MS", "MP", linkType, lagValue
    legacyData = sourceData
    Run_Calc_Core legacyData, mapCol, linksBySuccId
    indexedData = sourceData
    Set network = CompileExecutionNetwork(indexedData, mapCol, linksBySuccId)
    Run_Calc_Core indexedData, mapCol, linksBySuccId, , , , , network

    passed = ArrayIndexedCoreHarness_OutputRowsEqual(legacyData, indexedData, mapCol, 1, 2) And _
        ArrayIndexedCoreHarness_CompiledEdgeMatches(network, "MS", "MP", linkType, lagValue)
    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "MILESTONE_LINK", _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 2), _
        ArrayIndexedCoreHarness_RowResult(indexedData, mapCol, 2), _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 2), _
        linkType & "|" & CStr(lagValue), passed, ""
    ArrayIndexedCoreHarness_RunMilestoneLinkCase = passed
    Exit Function

Fail:
    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "MILESTONE_LINK", "", "", "", _
        linkType & "|" & CStr(lagValue), False, CStr(Err.Number) & "|" & Err.Description

End Function

Private Function ArrayIndexedCoreHarness_RunLOECase( _
    ByVal fileNo As Integer, _
    ByVal caseName As String, _
    ByVal invalidMode As Long) As Boolean

    Dim mapCol As Object
    Dim sourceData As Variant
    Dim legacyData As Variant
    Dim indexedData As Variant
    Dim linksBySuccId As Object
    Dim network As clsCompiledExecutionNetwork
    Dim passed As Boolean

    On Error GoTo Fail
    Set mapCol = ArrayIndexedCoreHarness_BuildColumnMap()
    sourceData = ArrayIndexedCoreHarness_BuildBaseData(mapCol, 3)
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 1, "LS", "1", "Task", CDbl(DateSerial(2026, 6, 1)), 3#
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 2, "LF", "2", "Task", CDbl(DateSerial(2026, 6, 8)), 4#
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 3, "LOE", "3", "LOE", Empty, Empty

    Set linksBySuccId = Core_CreateLinksBySucc()
    If invalidMode <> 1 Then Core_AddLink linksBySuccId, "LOE", "LS", "SS", 1#
    If invalidMode <> 2 Then Core_AddLink linksBySuccId, "LOE", "LF", "FF", -1#

    legacyData = sourceData
    Run_Calc_Core legacyData, mapCol, linksBySuccId
    indexedData = sourceData
    Set network = CompileExecutionNetwork(indexedData, mapCol, linksBySuccId)
    Run_Calc_Core indexedData, mapCol, linksBySuccId, , , , , network

    passed = ArrayIndexedCoreHarness_OutputRowsEqual(legacyData, indexedData, mapCol, 1, 3)
    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "LOE", _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 3), _
        ArrayIndexedCoreHarness_RowResult(indexedData, mapCol, 3), _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 3), _
        "Mode=" & CStr(invalidMode), passed, ""
    ArrayIndexedCoreHarness_RunLOECase = passed
    Exit Function

Fail:
    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "LOE", "", "", "", _
        "Mode=" & CStr(invalidMode), False, CStr(Err.Number) & "|" & Err.Description

End Function

Private Function ArrayIndexedCoreHarness_RunDiagnosticCase( _
    ByVal fileNo As Integer, _
    ByVal caseName As String, _
    ByVal diagnosticMode As Long) As Boolean

    Dim mapCol As Object
    Dim sourceData As Variant
    Dim legacyData As Variant
    Dim indexedData As Variant
    Dim linksBySuccId As Object
    Dim network As clsCompiledExecutionNetwork
    Dim passed As Boolean

    On Error GoTo Fail
    Set mapCol = ArrayIndexedCoreHarness_BuildColumnMap()
    sourceData = ArrayIndexedCoreHarness_BuildBaseData(mapCol, 2)
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 1, "DA", "1", "Task", CDbl(DateSerial(2026, 7, 1)), 2#
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 2, "DB", "2", "Task", CDbl(DateSerial(2026, 7, 1)), 2#
    Set linksBySuccId = Core_CreateLinksBySucc()

    Select Case diagnosticMode
        Case 1
            Core_AddLink linksBySuccId, "DB", "MISSING", "FS", 0#
        Case 2
            Core_AddLink linksBySuccId, "DB", "DA", "FS", 0#
            Core_AddLink linksBySuccId, "DA", "DB", "SS", 1#
        Case 3
            Core_AddLink linksBySuccId, "DA", "DA", "FF", 0#
    End Select

    legacyData = sourceData
    Run_Calc_Core legacyData, mapCol, linksBySuccId
    indexedData = sourceData
    Set network = CompileExecutionNetwork(indexedData, mapCol, linksBySuccId)
    Run_Calc_Core indexedData, mapCol, linksBySuccId, , , , , network

    passed = ArrayIndexedCoreHarness_OutputRowsEqual(legacyData, indexedData, mapCol, 1, 2)
    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "DIAGNOSTIC", _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 1) & ";" & ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 2), _
        ArrayIndexedCoreHarness_RowResult(indexedData, mapCol, 1) & ";" & ArrayIndexedCoreHarness_RowResult(indexedData, mapCol, 2), _
        "Legacy parity", "Mode=" & CStr(diagnosticMode), passed, ""
    ArrayIndexedCoreHarness_RunDiagnosticCase = passed
    Exit Function

Fail:
    ArrayIndexedCoreHarness_WriteResult fileNo, caseName, "DIAGNOSTIC", "", "", "", _
        "Mode=" & CStr(diagnosticMode), False, CStr(Err.Number) & "|" & Err.Description

End Function

Private Function ArrayIndexedCoreHarness_RunParentExpansionMetadataCase(ByVal fileNo As Integer) As Boolean

    Dim mapCol As Object
    Dim sourceData As Variant
    Dim legacyData As Variant
    Dim indexedData As Variant
    Dim linksBySuccId As Object
    Dim network As clsCompiledExecutionNetwork
    Dim summarySources As Variant
    Dim passed As Boolean

    On Error GoTo Fail
    Set mapCol = ArrayIndexedCoreHarness_BuildColumnMap()
    sourceData = ArrayIndexedCoreHarness_BuildBaseData(mapCol, 3)
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 1, "PA", "1.1", "Task", CDbl(DateSerial(2026, 8, 1)), 2#
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 2, "PB", "1.2", "Task", CDbl(DateSerial(2026, 8, 4)), 2#
    ArrayIndexedCoreHarness_SetTask sourceData, mapCol, 3, "PS", "2", "Task", CDbl(DateSerial(2026, 8, 1)), 2#
    Set linksBySuccId = Core_CreateLinksBySucc()
    Core_AddLink linksBySuccId, "PS", "PA", "SS", -1#, "PARENT-A"
    Core_AddLink linksBySuccId, "PS", "PB", "SS", -1#, "PARENT-A"

    legacyData = sourceData
    Run_Calc_Core legacyData, mapCol, linksBySuccId
    indexedData = sourceData
    Set network = CompileExecutionNetwork(indexedData, mapCol, linksBySuccId)
    Run_Calc_Core indexedData, mapCol, linksBySuccId, , , , , network
    summarySources = network.CorePredSummarySourceIds

    passed = ArrayIndexedCoreHarness_OutputRowsEqual(legacyData, indexedData, mapCol, 1, 3) And _
        CStr(summarySources(1)) = "PARENT-A" And CStr(summarySources(2)) = "PARENT-A"
    ArrayIndexedCoreHarness_WriteResult fileNo, "PARENT_EXPANSION_METADATA", "PARENT_EXPANSION", _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 3), _
        ArrayIndexedCoreHarness_RowResult(indexedData, mapCol, 3), _
        ArrayIndexedCoreHarness_RowResult(legacyData, mapCol, 3), _
        "SS-1|PARENT-A", passed, "CoreEdges=2"
    ArrayIndexedCoreHarness_RunParentExpansionMetadataCase = passed
    Exit Function

Fail:
    ArrayIndexedCoreHarness_WriteResult fileNo, "PARENT_EXPANSION_METADATA", "PARENT_EXPANSION", _
        "", "", "", "SS-1|PARENT-A", False, CStr(Err.Number) & "|" & Err.Description

End Function

Private Function ArrayIndexedCoreHarness_BuildColumnMap() As Object

    Dim result As Object
    Dim names As Variant
    Dim index As Long

    names = Array( _
        "ID", "WBS", "Task Name", "ParentID", "IsSummary", "Task Type", "Cal", _
        "Actual Start", "Actual Finish", "Forecast Start", "Forecast Finish", _
        "Baseline Start", "Baseline Duration", "Constraint Active", _
        "Start Constraint Type", "Start Constraint Date", _
        "Finish Constraint Type", "Finish Constraint Date", _
        "Calculated Start", "Calculated Finish", "Calculated Duration", _
        "Error flag", "ErrorMsg", "Driving Logic")

    Set result = CreateObject("Scripting.Dictionary")
    For index = LBound(names) To UBound(names)
        result(CStr(names(index))) = index + 1
    Next index
    Set ArrayIndexedCoreHarness_BuildColumnMap = result

End Function

Private Function ArrayIndexedCoreHarness_BuildBaseData( _
    ByVal mapCol As Object, _
    ByVal rowCount As Long) As Variant

    Dim result() As Variant
    ReDim result(1 To rowCount, 1 To mapCol.Count)
    ArrayIndexedCoreHarness_BuildBaseData = result

End Function

Private Sub ArrayIndexedCoreHarness_SetTask( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowIndex As Long, _
    ByVal taskId As String, _
    ByVal wbsCode As String, _
    ByVal taskType As String, _
    ByVal baselineStart As Variant, _
    ByVal baselineDuration As Variant)

    dataArr(rowIndex, mapCol("ID")) = taskId
    dataArr(rowIndex, mapCol("WBS")) = wbsCode
    dataArr(rowIndex, mapCol("Task Name")) = taskId
    dataArr(rowIndex, mapCol("ParentID")) = ""
    dataArr(rowIndex, mapCol("IsSummary")) = False
    dataArr(rowIndex, mapCol("Task Type")) = taskType
    dataArr(rowIndex, mapCol("Cal")) = HARNESS_CAL
    dataArr(rowIndex, mapCol("Baseline Start")) = baselineStart
    dataArr(rowIndex, mapCol("Baseline Duration")) = baselineDuration
    dataArr(rowIndex, mapCol("Constraint Active")) = "NO"

End Sub

Private Function ArrayIndexedCoreHarness_OutputRowsEqual( _
    ByRef leftData As Variant, _
    ByRef rightData As Variant, _
    ByVal mapCol As Object, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long) As Boolean

    Dim fields As Variant
    Dim fieldName As Variant
    Dim rowIndex As Long

    fields = Array("Calculated Start", "Calculated Finish", "Calculated Duration", _
        "Error flag", "ErrorMsg", "Driving Logic")

    For rowIndex = firstRow To lastRow
        For Each fieldName In fields
            If Not ArrayIndexedCoreHarness_ValueEqual( _
                leftData(rowIndex, mapCol(CStr(fieldName))), _
                rightData(rowIndex, mapCol(CStr(fieldName)))) Then Exit Function
        Next fieldName
    Next rowIndex

    ArrayIndexedCoreHarness_OutputRowsEqual = True

End Function

Private Function ArrayIndexedCoreHarness_CompiledEdgeMatches( _
    ByVal network As clsCompiledExecutionNetwork, _
    ByVal successorId As String, _
    ByVal predecessorId As String, _
    ByVal expectedType As String, _
    ByVal expectedLag As Double) As Boolean

    Dim successorIndex As Long
    Dim predecessorIndex As Long
    Dim offsets As Variant
    Dim nodes As Variant
    Dim types As Variant
    Dim lags As Variant
    Dim edgeIndex As Long

    If network Is Nothing Then Exit Function
    If Not network.NodeIndexById.Exists(successorId) Then Exit Function
    If Not network.NodeIndexById.Exists(predecessorId) Then Exit Function

    successorIndex = CLng(network.NodeIndexById(successorId))
    predecessorIndex = CLng(network.NodeIndexById(predecessorId))
    offsets = network.PredOffsets
    nodes = network.PredNodes
    types = network.PredTypes
    lags = network.PredLags

    For edgeIndex = CLng(offsets(successorIndex)) To CLng(offsets(successorIndex + 1)) - 1
        If CLng(nodes(edgeIndex)) = predecessorIndex Then
            ArrayIndexedCoreHarness_CompiledEdgeMatches = _
                (CStr(types(edgeIndex)) = expectedType) And _
                (Abs(CDbl(lags(edgeIndex)) - expectedLag) < 0.0000001)
            Exit Function
        End If
    Next edgeIndex

End Function

Private Function ArrayIndexedCoreHarness_ValueEqual(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean

    If IsEmpty(leftValue) And IsEmpty(rightValue) Then
        ArrayIndexedCoreHarness_ValueEqual = True
    ElseIf IsNumeric(leftValue) And IsNumeric(rightValue) Then
        ArrayIndexedCoreHarness_ValueEqual = (Abs(CDbl(leftValue) - CDbl(rightValue)) < 0.0000001)
    Else
        ArrayIndexedCoreHarness_ValueEqual = (CStr(leftValue) = CStr(rightValue))
    End If

End Function

Private Function ArrayIndexedCoreHarness_RowResult( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object, _
    ByVal rowIndex As Long) As String

    ArrayIndexedCoreHarness_RowResult = _
        ArrayIndexedCoreHarness_FormatValue(dataArr(rowIndex, mapCol("Calculated Start"))) & "|" & _
        ArrayIndexedCoreHarness_FormatValue(dataArr(rowIndex, mapCol("Calculated Finish"))) & "|" & _
        ArrayIndexedCoreHarness_FormatValue(dataArr(rowIndex, mapCol("Calculated Duration"))) & "|" & _
        Replace(CStr(dataArr(rowIndex, mapCol("ErrorMsg"))), vbCrLf, " ")

End Function

Private Function ArrayIndexedCoreHarness_FormatPair( _
    ByVal startValue As Variant, _
    ByVal finishValue As Variant) As String

    ArrayIndexedCoreHarness_FormatPair = _
        ArrayIndexedCoreHarness_FormatValue(startValue) & "|" & _
        ArrayIndexedCoreHarness_FormatValue(finishValue)

End Function

Private Function ArrayIndexedCoreHarness_FormatValue(ByVal value As Variant) As String
    If IsEmpty(value) Then
        ArrayIndexedCoreHarness_FormatValue = "<EMPTY>"
    ElseIf IsNumeric(value) Then
        ArrayIndexedCoreHarness_FormatValue = Format$(CDbl(value), "0.############")
    Else
        ArrayIndexedCoreHarness_FormatValue = CStr(value)
    End If
End Function

Private Sub ArrayIndexedCoreHarness_WriteResult( _
    ByVal fileNo As Integer, _
    ByVal caseName As String, _
    ByVal domainName As String, _
    ByVal legacyResult As String, _
    ByVal indexedResult As String, _
    ByVal expectedResult As String, _
    ByVal attributesResult As String, _
    ByVal passed As Boolean, _
    ByVal details As String)

    Print #fileNo, caseName & vbTab & domainName & vbTab & legacyResult & vbTab & _
        indexedResult & vbTab & expectedResult & vbTab & attributesResult & vbTab & _
        IIf(passed, "PASS", "FAIL") & vbTab & Replace(details, vbTab, " ")

End Sub
