Attribute VB_Name = "mod_IncrementalSignatureHarness"
Option Explicit

'===============================================================================
' MODULE : mod_IncrementalSignatureHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Incremental Signature sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Incremental Signature contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : IncrementalSignatureHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Verifie le contrat canonique et la detection incrementale sur une copie de test.
' EN: Verifies the canonical contract and incremental detection on a test copy.
'------------------------------------------------------------------------------
Public Function IncrementalSignatureHarness_Smoke() As String

    Dim columns As Variant
    Dim mapCalc As Object
    Dim arrCalc As Variant
    Dim expected As String
    Dim actual As String
    Dim i As Long
    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim calcColumns As Object
    Dim changedIds As Object
    Dim forceFullRecalc As Boolean
    Dim taskId As String
    Dim originalDeadline As Variant
    Dim deadlineChanged As Boolean
    Dim oldEvents As Boolean
    Dim eventsCaptured As Boolean
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo Fail

    columns = IncrementalSignature_RequiredColumns()
    IncrementalSignatureHarness_Assert UBound(columns) - LBound(columns) + 1 = 17, "required column count"

    Set mapCalc = CreateObject("Scripting.Dictionary")
    ReDim arrCalc(1 To 2, 1 To 17)
    For i = LBound(columns) To UBound(columns)
        mapCalc(CStr(columns(i))) = i - LBound(columns) + 1
    Next i

    arrCalc(1, mapCalc("ID")) = " 42 "
    arrCalc(1, mapCalc("ParentID")) = " 7 "
    arrCalc(1, mapCalc("IsSummary")) = "Oui"
    arrCalc(1, mapCalc("Predecessors WBS")) = " 1.1 ; 2.2 "
    arrCalc(1, mapCalc("Cal")) = " 7J/7 "
    arrCalc(1, mapCalc("Baseline Start")) = DateSerial(2026, 1, 2)
    arrCalc(1, mapCalc("Baseline Duration")) = 1.5
    arrCalc(1, mapCalc("Actual Start")) = DateSerial(2026, 1, 3)
    arrCalc(1, mapCalc("Actual Finish")) = DateSerial(2026, 1, 4)
    arrCalc(1, mapCalc("Forecast Start")) = DateSerial(2026, 1, 5)
    arrCalc(1, mapCalc("Forecast Finish")) = DateSerial(2026, 1, 6)
    arrCalc(1, mapCalc("Deadline")) = DateSerial(2026, 1, 7)
    arrCalc(1, mapCalc("Constraint Active")) = False
    arrCalc(1, mapCalc("Start Constraint Type")) = " Start Type "
    arrCalc(1, mapCalc("Start Constraint Date")) = DateSerial(2026, 1, 8)
    arrCalc(1, mapCalc("Finish Constraint Type")) = " Finish Type "
    arrCalc(1, mapCalc("Finish Constraint Date")) = DateSerial(2026, 1, 9)

    expected = "|ID=42|ParentID=7|IsSummary=TRUE|PredWBS=1.1;2.2|Cal=7j/7" & _
        "|BS=20260102|BD=" & Format$(1.5, "0.############") & _
        "|AS=20260103|AF=20260104|FS=20260105|FF=20260106|DL=20260107" & _
        "|CActive=FALSE|SCType=Start Type|SCDate=20260108|FCType=Finish Type|FCDate=20260109"
    actual = IncrementalSignature_BuildRow(arrCalc, 1, mapCalc)
    IncrementalSignatureHarness_Assert actual = expected, "normalized signature"

    expected = "|ID=|ParentID=|IsSummary=|PredWBS=|Cal=7j/7" & _
        "|BS=|BD=|AS=|AF=|FS=|FF=|DL=|CActive=|SCType=|SCDate=|FCType=|FCDate="
    actual = IncrementalSignature_BuildRow(arrCalc, 2, mapCalc)
    IncrementalSignatureHarness_Assert actual = expected, "empty signature"

    Write_CalcState_Snapshot "OK"
    Set changedIds = Get_Changed_TaskIds(forceFullRecalc)
    IncrementalSignatureHarness_Assert Not forceFullRecalc, "unchanged snapshot does not force full recalculation"
    IncrementalSignatureHarness_Assert changedIds.Count = 0, "unchanged snapshot has no changed task"

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set calcColumns = CanonicalIdentity_BuildColumnMap(tblCalc)
    IncrementalSignatureHarness_Assert Not tblCalc.DataBodyRange Is Nothing, "CALC fixture has rows"
    taskId = Trim$(CStr(tblCalc.DataBodyRange.Cells(1, calcColumns("ID")).value))
    originalDeadline = tblCalc.DataBodyRange.Cells(1, calcColumns("Deadline")).value

    oldEvents = Application.EnableEvents
    eventsCaptured = True
    Application.EnableEvents = False
    If IsDate(originalDeadline) Then
        tblCalc.DataBodyRange.Cells(1, calcColumns("Deadline")).value = CDate(originalDeadline) + 1
    Else
        tblCalc.DataBodyRange.Cells(1, calcColumns("Deadline")).value = DateSerial(2099, 12, 31)
    End If
    deadlineChanged = True

    Set changedIds = Get_Changed_TaskIds(forceFullRecalc)
    IncrementalSignatureHarness_Assert Not forceFullRecalc, "controlled mutation does not force full recalculation"
    IncrementalSignatureHarness_Assert changedIds.Exists(taskId), "controlled mutation identifies changed task"

    tblCalc.DataBodyRange.Cells(1, calcColumns("Deadline")).value = originalDeadline
    deadlineChanged = False
    Application.EnableEvents = oldEvents

    IncrementalSignatureHarness_Smoke = "PASS"
    Exit Function

Fail:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If deadlineChanged Then tblCalc.DataBodyRange.Cells(1, calcColumns("Deadline")).value = originalDeadline
    If eventsCaptured Then Application.EnableEvents = oldEvents
    On Error GoTo 0
    IncrementalSignatureHarness_Smoke = "FAIL: " & errNum & " - " & errDesc

End Function

'------------------------------------------------------------------------------
' FR: Interrompt le harnais lorsqu'un invariant de signature n'est pas respecte.
' EN: Stops the harness when a signature invariant is not satisfied.
'------------------------------------------------------------------------------
Private Sub IncrementalSignatureHarness_Assert(ByVal condition As Boolean, ByVal contractName As String)
    If Not condition Then Err.Raise vbObjectError + 3403, "IncrementalSignatureHarness", contractName
End Sub
