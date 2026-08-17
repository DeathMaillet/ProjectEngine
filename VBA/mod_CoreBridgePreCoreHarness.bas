Attribute VB_Name = "mod_CoreBridgePreCoreHarness"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgePreCoreHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Core Bridge Pre Core sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Core Bridge Pre Core contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : CoreBridgePreCoreHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private gPreCoreHarnessTracePath As String

'------------------------------------------------------------------------------
' FR: Execute un smoke contractuel des validations Pre-Core, mappings incrementaux
'     et writes partiels CoreBridge sur une copie temporaire.
' EN: Runs a contractual smoke for Pre-Core validations, incremental maps, and
'     partial CoreBridge writes on a temporary workbook copy.
'------------------------------------------------------------------------------
Public Function CoreBridgePreCoreHarness_Smoke( _
    ByVal matrixPath As String) As String

    Dim wsCalc As Worksheet
    Dim wsWBS As Worksheet
    Dim tblCalc As ListObject
    Dim tblWBS As ListObject
    Dim tblLinks As ListObject
    Dim mapCalc As Object
    Dim mapWBS As Object
    Dim mapLinks As Object
    Dim messages As Collection

    On Error GoTo Fail

    gPreCoreHarnessTracePath = matrixPath & ".trace.txt"
    CoreBridgePreCoreHarness_DeleteFile gPreCoreHarnessTracePath
    CoreBridgePreCoreHarness_DeleteFile matrixPath
    CoreBridgePreCoreHarness_Trace "01 enter"

    PlanningConsolePolicy_DisableNonInteractive
    Set messages = New Collection

    Set wsCalc = ThisWorkbook.Worksheets("CALC")
    Set wsWBS = ThisWorkbook.Worksheets("WBS")
    Set tblCalc = wsCalc.ListObjects("tbl_CALC")
    Set tblWBS = wsWBS.ListObjects("tbl_WBS")
    Set tblLinks = wsCalc.ListObjects("tbl_LOGIC_LINKS")

    CoreBridgePreCoreHarness_Assert Not tblCalc.DataBodyRange Is Nothing, "tbl_CALC has rows"
    CoreBridgePreCoreHarness_Assert Not tblWBS.DataBodyRange Is Nothing, "tbl_WBS has rows"
    CoreBridgePreCoreHarness_Assert tblCalc.ListRows.Count >= 6, "tbl_CALC has at least six rows"
    CoreBridgePreCoreHarness_Assert tblWBS.ListRows.Count >= 2, "tbl_WBS has at least two rows"

    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    Set mapLinks = CanonicalIdentity_BuildColumnMap(tblLinks)

    CoreBridgePreCoreHarness_PrepareRows tblCalc, mapCalc, tblWBS, mapWBS, tblLinks, mapLinks
    CoreBridgePreCoreHarness_Trace "02 rows prepared"

    CoreBridgePreCoreHarness_TestValidateCalc tblCalc
    CoreBridgePreCoreHarness_Trace "03 validate calc ok"

    CoreBridgePreCoreHarness_TestIncrementalMaps tblCalc, mapCalc
    CoreBridgePreCoreHarness_Trace "04 incremental maps ok"

    CoreBridgePreCoreHarness_TestPreCoreStops tblCalc, mapCalc, tblLinks, mapLinks, messages
    CoreBridgePreCoreHarness_Trace "05 pre-core stops ok"

    CoreBridgePreCoreHarness_TestPartialOutputWrites tblCalc, mapCalc, tblWBS, mapWBS
    CoreBridgePreCoreHarness_Trace "06 partial output writes ok"

    CoreBridgePreCoreHarness_TestFullOutputWrites tblCalc, mapCalc, tblWBS, mapWBS
    CoreBridgePreCoreHarness_Trace "07 full output writes ok"

    CoreBridgePreCoreHarness_WriteMatrix matrixPath
    CoreBridgePreCoreHarness_Trace "08 matrix written"

    CoreBridgePreCoreHarness_Smoke = "PASS"
    Exit Function

Fail:
    On Error Resume Next
    CoreBridgePreCoreHarness_Trace "FAIL " & Err.Number & " " & Err.Description
    PlanningConsolePolicy_DisableNonInteractive
    CoreBridgePreCoreHarness_Smoke = "FAIL: " & Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Prepare Rows et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Prepare Rows contract and reports any divergence to the harness.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: writes to an Excel table owned by the workflow.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_PrepareRows( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object, _
    ByVal tblLinks As ListObject, _
    ByVal mapLinks As Object)

    Dim r As Long
    Dim requiredCalc As Variant
    Dim requiredWBS As Variant
    Dim requiredLinks As Variant

    requiredCalc = Array("ID", "WBS", "Task Name", "Task Type", "ParentID", "IsSummary", _
        "Calculated Start", "Calculated Finish", "Calculated Duration", "Driving Logic", _
        "Deadline Float", "Error flag", "ErrorMsg")
    requiredWBS = Array("ID", "WBS", "Calculated Start", "Calculated Finish", "Driving Logic", "Deadline Float")
    requiredLinks = Array("Succ ID", "Pred ID", "Link Type", "Lag")

    CoreBridgePreCoreHarness_RequireColumns mapCalc, requiredCalc, "tbl_CALC"
    CoreBridgePreCoreHarness_RequireColumns mapWBS, requiredWBS, "tbl_WBS"
    CoreBridgePreCoreHarness_RequireColumns mapLinks, requiredLinks, "tbl_LOGIC_LINKS"

    For r = 1 To 6
        CoreBridgePreCoreHarness_ClearCalcRow tblCalc, mapCalc, r
    Next r

    For r = 1 To 2
        tblWBS.DataBodyRange.Cells(r, mapWBS("ID")).Value = "CBH-OUT-" & CStr(r)
        tblWBS.DataBodyRange.Cells(r, mapWBS("WBS")).Value = "31.13." & CStr(r)
        tblWBS.DataBodyRange.Cells(r, mapWBS("Calculated Start")).Value = DateSerial(2026, 1, r)
        tblWBS.DataBodyRange.Cells(r, mapWBS("Calculated Finish")).Value = DateSerial(2026, 1, r + 1)
        tblWBS.DataBodyRange.Cells(r, mapWBS("Driving Logic")).Value = "before-" & CStr(r)
        tblWBS.DataBodyRange.Cells(r, mapWBS("Deadline Float")).Value = 99 + r
    Next r

    For r = 1 To tblLinks.ListRows.Count
        tblLinks.DataBodyRange.Cells(r, mapLinks("Succ ID")).ClearContents
        tblLinks.DataBodyRange.Cells(r, mapLinks("Pred ID")).ClearContents
        tblLinks.DataBodyRange.Cells(r, mapLinks("Link Type")).ClearContents
        tblLinks.DataBodyRange.Cells(r, mapLinks("Lag")).ClearContents
    Next r

    If tblLinks.ListRows.Count = 0 Then tblLinks.ListRows.Add

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Clear Calc Row et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Clear Calc Row contract and reports any divergence to the harness.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: writes to an Excel table owned by the workflow.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_ClearCalcRow( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal rowIdx As Long)

    Dim key As Variant

    For Each key In mapCalc.Keys
        tblCalc.DataBodyRange.Cells(rowIdx, CLng(mapCalc(CStr(key)))).ClearContents
    Next key

    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("ID")).Value = "CBH-ROW-" & CStr(rowIdx)
    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("WBS")).Value = "31.13." & CStr(rowIdx)
    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Task Name")).Value = "Harness row " & CStr(rowIdx)
    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Task Type")).Value = "Task"
    tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("IsSummary")).Value = False

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Test Validate Calc et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Test Validate Calc contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_TestValidateCalc(ByVal tblCalc As ListObject)

    CoreBridgePreCoreHarness_Assert ValidateCalcAfterSync(tblCalc), "ValidateCalcAfterSync returns True"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Test Incremental Maps et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Test Incremental Maps contract and reports any divergence to the harness.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_TestIncrementalMaps( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object)

    Dim arrCalc As Variant
    Dim rowById As Object
    Dim linksBySucc As Object
    Dim predsBySucc As Object
    Dim succByPred As Object
    Dim changedIds As Object
    Dim impacted As Object
    Dim parentById As Object

    tblCalc.DataBodyRange.Cells(1, mapCalc("ID")).Value = "CBH-A"
    tblCalc.DataBodyRange.Cells(2, mapCalc("ID")).Value = "CBH-B"
    tblCalc.DataBodyRange.Cells(3, mapCalc("ID")).Value = "CBH-C"
    tblCalc.DataBodyRange.Cells(2, mapCalc("ParentID")).Value = "CBH-A"

    arrCalc = tblCalc.DataBodyRange.Value
    Set rowById = Core_BuildRowById(arrCalc, mapCalc)
    Set linksBySucc = Core_CreateLinksBySucc()
    Core_AddLink linksBySucc, "CBH-B", "CBH-A", "FS", 0#
    Core_AddLink linksBySucc, "CBH-C", "CBH-B", "FS", 0#

    Set predsBySucc = BuildPredsBySucc_FromExpandedLinks(rowById, linksBySucc)
    CoreBridgePreCoreHarness_Assert predsBySucc.Exists("CBH-C"), "predsBySucc contains CBH-C"
    CoreBridgePreCoreHarness_Assert predsBySucc("CBH-C").Count = 1, "CBH-C has one predecessor"
    CoreBridgePreCoreHarness_Assert CStr(predsBySucc("CBH-C")(1)) = "CBH-B", "CBH-C predecessor is CBH-B"

    Set parentById = BuildParentByIdMap_FromCalc(arrCalc, mapCalc, rowById)
    CoreBridgePreCoreHarness_Assert parentById.Exists("CBH-B"), "parent map contains CBH-B"
    CoreBridgePreCoreHarness_Assert CStr(parentById("CBH-B")) = "CBH-A", "CBH-B parent is CBH-A"

    Set succByPred = Build_Successor_Map(linksBySucc)
    Set changedIds = CreateObject("Scripting.Dictionary")
    changedIds("CBH-A") = True
    Set impacted = Get_Impacted_Descendants(changedIds, succByPred)

    CoreBridgePreCoreHarness_Assert impacted.Exists("CBH-A"), "impacted includes root"
    CoreBridgePreCoreHarness_Assert impacted.Exists("CBH-B"), "impacted includes first successor"
    CoreBridgePreCoreHarness_Assert impacted.Exists("CBH-C"), "impacted includes transitive successor"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Test Pre Core Stops et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Test Pre Core Stops contract and reports any divergence to the harness.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_TestPreCoreStops( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal tblLinks As ListObject, _
    ByVal mapLinks As Object, _
    ByVal messages As Collection)

    Dim arrCalc As Variant
    Dim rowById As Object
    Dim validIds As Object
    Dim linksBySucc As Object

    tblCalc.DataBodyRange.Cells(1, mapCalc("ID")).Value = "CBH-MISSING-SUCC"
    tblCalc.DataBodyRange.Cells(1, mapCalc("WBS")).Value = "31.13.M"
    tblLinks.DataBodyRange.Cells(1, mapLinks("Succ ID")).Value = "CBH-MISSING-SUCC"
    tblLinks.DataBodyRange.Cells(1, mapLinks("Pred ID")).Value = "CBH-NOPE"
    tblLinks.DataBodyRange.Cells(1, mapLinks("Link Type")).Value = "FS"
    tblLinks.DataBodyRange.Cells(1, mapLinks("Lag")).Value = 0

    arrCalc = tblCalc.DataBodyRange.Value
    Set rowById = Core_BuildRowById(arrCalc, mapCalc)
    CoreBridgePreCoreHarness_Assert CalcBridge_PreCore_CheckMissingPredecessors(tblCalc, mapCalc, rowById, messages), "missing predecessor detected"
    CoreBridgePreCoreHarness_Assert CStr(tblCalc.DataBodyRange.Cells(1, mapCalc("ErrorMsg")).Value) = "Missing predecessor", "missing predecessor marked"

    tblCalc.DataBodyRange.Cells(3, mapCalc("ID")).Value = "CBH-CYCLE-A"
    tblCalc.DataBodyRange.Cells(4, mapCalc("ID")).Value = "CBH-CYCLE-B"
    Set validIds = CreateObject("Scripting.Dictionary")
    validIds("CBH-CYCLE-A") = True
    validIds("CBH-CYCLE-B") = True
    Set linksBySucc = Core_CreateLinksBySucc()
    Core_AddLink linksBySucc, "CBH-CYCLE-A", "CBH-CYCLE-B", "FS", 0#
    Core_AddLink linksBySucc, "CBH-CYCLE-B", "CBH-CYCLE-A", "FS", 0#

    CoreBridgePreCoreHarness_Assert CalcBridge_PreCore_CheckCycles(tblCalc, mapCalc, validIds, linksBySucc, messages), "cycle detected"
    CoreBridgePreCoreHarness_Assert CStr(tblCalc.DataBodyRange.Cells(3, mapCalc("ErrorMsg")).Value) = "Cycle detected", "cycle row A marked"
    CoreBridgePreCoreHarness_Assert CStr(tblCalc.DataBodyRange.Cells(4, mapCalc("ErrorMsg")).Value) = "Cycle detected", "cycle row B marked"

    tblCalc.DataBodyRange.Cells(5, mapCalc("ID")).Value = "CBH-LOE"
    tblCalc.DataBodyRange.Cells(5, mapCalc("WBS")).Value = "31.13.L"
    tblCalc.DataBodyRange.Cells(5, mapCalc("Task Type")).Value = "Level of Effort"
    tblCalc.DataBodyRange.Cells(6, mapCalc("ID")).Value = "CBH-LOE-SUCC"
    tblCalc.DataBodyRange.Cells(6, mapCalc("WBS")).Value = "31.13.S"
    tblLinks.DataBodyRange.Cells(1, mapLinks("Succ ID")).Value = "CBH-LOE-SUCC"
    tblLinks.DataBodyRange.Cells(1, mapLinks("Pred ID")).Value = "CBH-LOE"
    tblLinks.DataBodyRange.Cells(1, mapLinks("Link Type")).Value = "FS"
    tblLinks.DataBodyRange.Cells(1, mapLinks("Lag")).Value = 0

    CoreBridgePreCoreHarness_Assert CalcBridge_PreCore_CheckLOEAsPredecessor(tblCalc, mapCalc, messages), "LOE predecessor detected"
    CoreBridgePreCoreHarness_Assert CStr(tblCalc.DataBodyRange.Cells(5, mapCalc("ErrorMsg")).Value) = "LOE cannot be used as predecessor", "LOE predecessor row marked"
    CoreBridgePreCoreHarness_Assert messages.Count >= 3, "pre-core messages captured"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Test Partial Output Writes et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Test Partial Output Writes contract and reports any divergence to the harness.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' EN - Side effect: writes to an Excel table owned by the workflow.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_TestPartialOutputWrites( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object)

    Dim arrCalc As Variant
    Dim impactedIds As Object
    Dim row2Before As Variant
    Dim wbsRow2Before As Variant

    tblCalc.DataBodyRange.Cells(1, mapCalc("ID")).Value = "CBH-OUT-1"
    tblCalc.DataBodyRange.Cells(2, mapCalc("ID")).Value = "CBH-OUT-2"
    tblCalc.DataBodyRange.Cells(1, mapCalc("Calculated Start")).Value = DateSerial(2026, 4, 1)
    tblCalc.DataBodyRange.Cells(1, mapCalc("Calculated Finish")).Value = DateSerial(2026, 4, 5)
    tblCalc.DataBodyRange.Cells(1, mapCalc("Calculated Duration")).Value = 5
    tblCalc.DataBodyRange.Cells(1, mapCalc("Driving Logic")).Value = "after-1"
    tblCalc.DataBodyRange.Cells(1, mapCalc("Deadline Float")).Value = -2
    tblCalc.DataBodyRange.Cells(2, mapCalc("Calculated Start")).Value = DateSerial(2026, 5, 1)
    row2Before = tblCalc.DataBodyRange.Cells(2, mapCalc("Calculated Start")).Value
    wbsRow2Before = tblWBS.DataBodyRange.Cells(2, mapWBS("Calculated Start")).Value

    arrCalc = tblCalc.DataBodyRange.Value
    arrCalc(1, mapCalc("Calculated Start")) = DateSerial(2026, 6, 1)
    arrCalc(1, mapCalc("Calculated Finish")) = DateSerial(2026, 6, 4)
    arrCalc(1, mapCalc("Calculated Duration")) = 4
    arrCalc(1, mapCalc("Error flag")) = vbNullString
    arrCalc(1, mapCalc("ErrorMsg")) = vbNullString
    arrCalc(2, mapCalc("Calculated Start")) = DateSerial(2026, 7, 1)

    Set impactedIds = CreateObject("Scripting.Dictionary")
    impactedIds("CBH-OUT-1") = True

    WriteCoreOutputsToCalc_Partial tblCalc, mapCalc, arrCalc, impactedIds

    CoreBridgePreCoreHarness_Assert CLng(tblCalc.DataBodyRange.Cells(1, mapCalc("Calculated Duration")).Value) = 4, "partial CALC write updated impacted row"
    CoreBridgePreCoreHarness_Assert CLng(CDate(tblCalc.DataBodyRange.Cells(2, mapCalc("Calculated Start")).Value)) = CLng(CDate(row2Before)), "partial CALC write left non-impacted row"

    Push_Calculated_Back_To_WBS_Partial impactedIds

    CoreBridgePreCoreHarness_Assert CStr(tblWBS.DataBodyRange.Cells(1, mapWBS("Driving Logic")).Value) = "after-1", "partial WBS push copied driving logic"
    CoreBridgePreCoreHarness_Assert CDbl(tblWBS.DataBodyRange.Cells(1, mapWBS("Deadline Float")).Value) = -2#, "partial WBS push copied deadline float"
    CoreBridgePreCoreHarness_Assert CLng(CDate(tblWBS.DataBodyRange.Cells(2, mapWBS("Calculated Start")).Value)) = CLng(CDate(wbsRow2Before)), "partial WBS push left non-impacted row"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Test Full Output Writes et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Test Full Output Writes contract and reports any divergence to the harness.
' FR - Effet de bord : ecrit dans une table Excel detenue par le workflow.
' FR - Effet de bord : efface uniquement les donnees ou objets cibles du contrat.
' EN - Side effect: writes to an Excel table owned by the workflow.
' EN - Side effect: clears only data or objects targeted by the contract.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_TestFullOutputWrites( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByVal tblWBS As ListObject, _
    ByVal mapWBS As Object)

    Dim arrCalc As Variant
    Dim fullWBSFields As Variant
    Dim sharedFields As Variant
    Dim impactedIds As Object
    Dim expectedShared As Object
    Dim i As Long
    Dim fieldName As String
    Dim wbsSentinel As Date

    fullWBSFields = Array( _
        "Calculated Start", "Calculated Finish", "Driving Logic", _
        "Critical Path", "Longest Path", "Critical Path REX", _
        "Total Float", "Free Float", "Total Float REX", _
        "Free Float REX", "Deadline Float")
    sharedFields = Array("Calculated Start", "Calculated Finish", "Driving Logic", "Deadline Float")

    CoreBridgePreCoreHarness_RequireColumns mapWBS, fullWBSFields, "tbl_WBS"
    CoreBridgePreCoreHarness_RequireColumns mapCalc, fullWBSFields, "tbl_CALC"

    tblCalc.DataBodyRange.Cells(1, mapCalc("ID")).Value = "CBH-OUT-1"
    tblCalc.DataBodyRange.Cells(2, mapCalc("ID")).Value = "CBH-OUT-2"
    tblWBS.DataBodyRange.Cells(1, mapWBS("ID")).Value = "CBH-OUT-1"
    tblWBS.DataBodyRange.Cells(2, mapWBS("ID")).Value = "CBH-OUT-2"

    wbsSentinel = DateSerial(2025, 12, 31)
    tblWBS.DataBodyRange.Cells(1, mapWBS("Calculated Start")).Value = wbsSentinel

    arrCalc = tblCalc.DataBodyRange.Value
    arrCalc(1, mapCalc("Calculated Start")) = DateSerial(2026, 8, 3)
    arrCalc(1, mapCalc("Calculated Finish")) = DateSerial(2026, 8, 7)
    arrCalc(1, mapCalc("Calculated Duration")) = 5
    arrCalc(1, mapCalc("Error flag")) = "FULL-ERR"
    arrCalc(1, mapCalc("ErrorMsg")) = "Full writer harness"
    arrCalc(2, mapCalc("Calculated Start")) = DateSerial(2026, 9, 1)
    arrCalc(2, mapCalc("Calculated Finish")) = DateSerial(2026, 9, 2)
    arrCalc(2, mapCalc("Calculated Duration")) = 2
    arrCalc(2, mapCalc("Error flag")) = vbNullString
    arrCalc(2, mapCalc("ErrorMsg")) = vbNullString

    WriteCoreOutputsToCalc tblCalc, mapCalc, arrCalc

    CoreBridgePreCoreHarness_Assert CLng(CDate(tblCalc.DataBodyRange.Cells(1, mapCalc("Calculated Start")).Value)) = CLng(DateSerial(2026, 8, 3)), "full CALC write copied start"
    CoreBridgePreCoreHarness_Assert CLng(CDate(tblCalc.DataBodyRange.Cells(1, mapCalc("Calculated Finish")).Value)) = CLng(DateSerial(2026, 8, 7)), "full CALC write copied finish"
    CoreBridgePreCoreHarness_Assert CLng(tblCalc.DataBodyRange.Cells(1, mapCalc("Calculated Duration")).Value) = 5, "full CALC write copied duration"
    CoreBridgePreCoreHarness_Assert CStr(tblCalc.DataBodyRange.Cells(1, mapCalc("Error flag")).Value) = "FULL-ERR", "full CALC write copied error flag"
    CoreBridgePreCoreHarness_Assert CStr(tblCalc.DataBodyRange.Cells(1, mapCalc("ErrorMsg")).Value) = "Full writer harness", "full CALC write copied error message"
    CoreBridgePreCoreHarness_Assert CLng(CDate(tblWBS.DataBodyRange.Cells(1, mapWBS("Calculated Start")).Value)) = CLng(wbsSentinel), "CALC write precedes WBS push"

    tblCalc.DataBodyRange.Cells(1, mapCalc("Driving Logic")).Value = "FULL-DRIVE"
    tblCalc.DataBodyRange.Cells(1, mapCalc("Critical Path")).Value = "Yes"
    tblCalc.DataBodyRange.Cells(1, mapCalc("Longest Path")).Value = "No"
    tblCalc.DataBodyRange.Cells(1, mapCalc("Critical Path REX")).Value = "CP-REX"
    tblCalc.DataBodyRange.Cells(1, mapCalc("Total Float")).Value = 11
    tblCalc.DataBodyRange.Cells(1, mapCalc("Free Float")).Value = 7
    tblCalc.DataBodyRange.Cells(1, mapCalc("Total Float REX")).Value = 5
    tblCalc.DataBodyRange.Cells(1, mapCalc("Free Float REX")).Value = 3
    tblCalc.DataBodyRange.Cells(1, mapCalc("Deadline Float")).Value = -4

    Push_Calculated_Back_To_WBS

    For i = LBound(fullWBSFields) To UBound(fullWBSFields)
        fieldName = CStr(fullWBSFields(i))
        CoreBridgePreCoreHarness_Assert _
            CStr(tblWBS.DataBodyRange.Cells(1, mapWBS(fieldName)).Value) = _
            CStr(tblCalc.DataBodyRange.Cells(1, mapCalc(fieldName)).Value), _
            "full WBS push copied " & fieldName
    Next i

    CoreBridgePreCoreHarness_Assert Left$(CStr(tblWBS.ListColumns("Calculated Duration").DataBodyRange.Cells(1, 1).FormulaLocal), 1) = "=", "full WBS push restored calculated duration formula"

    Set expectedShared = CreateObject("Scripting.Dictionary")
    For i = LBound(sharedFields) To UBound(sharedFields)
        fieldName = CStr(sharedFields(i))
        expectedShared(fieldName) = tblWBS.DataBodyRange.Cells(1, mapWBS(fieldName)).Value
        tblWBS.DataBodyRange.Cells(1, mapWBS(fieldName)).ClearContents
    Next i

    Set impactedIds = CreateObject("Scripting.Dictionary")
    impactedIds("CBH-OUT-1") = True
    Push_Calculated_Back_To_WBS_Partial impactedIds

    For i = LBound(sharedFields) To UBound(sharedFields)
        fieldName = CStr(sharedFields(i))
        CoreBridgePreCoreHarness_Assert _
            CStr(tblWBS.DataBodyRange.Cells(1, mapWBS(fieldName)).Value) = CStr(expectedShared(fieldName)), _
            "full and partial WBS outputs agree for " & fieldName
    Next i

End Sub
'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Write Matrix et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Write Matrix contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_WriteMatrix(ByVal matrixPath As String)

    Dim fileNo As Integer

    fileNo = FreeFile
    Open matrixPath For Output As #fileNo
    Print #fileNo, "Check" & vbTab & "Domain" & vbTab & "Result"
    Print #fileNo, "ValidateCalcAfterSync" & vbTab & "Pre-Core Validation" & vbTab & "PASS"
    Print #fileNo, "Missing predecessor" & vbTab & "Pre-Core Validation" & vbTab & "PASS"
    Print #fileNo, "Cycle detection" & vbTab & "Pre-Core Validation" & vbTab & "PASS"
    Print #fileNo, "LOE predecessor" & vbTab & "Pre-Core Validation" & vbTab & "PASS"
    Print #fileNo, "Incremental descendants" & vbTab & "Incremental Mapping" & vbTab & "PASS"
    Print #fileNo, "Partial CALC write" & vbTab & "Output Writer" & vbTab & "PASS"
    Print #fileNo, "Partial WBS push" & vbTab & "Output Writer" & vbTab & "PASS"
    Print #fileNo, "Full CALC write" & vbTab & "Output Writer" & vbTab & "PASS"
    Print #fileNo, "Full WBS push" & vbTab & "Output Writer" & vbTab & "PASS"
    Print #fileNo, "Full vs partial coherence" & vbTab & "Output Writer" & vbTab & "PASS"
    Print #fileNo, "CALC then WBS write order" & vbTab & "Output Writer" & vbTab & "PASS"
    Print #fileNo, "WBS formula restoration" & vbTab & "Output Writer" & vbTab & "PASS"
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Require Columns et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Require Columns contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_RequireColumns( _
    ByVal mapObj As Object, _
    ByVal columnNames As Variant, _
    ByVal tableName As String)

    Dim i As Long

    For i = LBound(columnNames) To UBound(columnNames)
        If Not mapObj.Exists(CStr(columnNames(i))) Then
            Err.Raise vbObjectError + 9510, "CoreBridgePreCoreHarness", _
                "Missing column in " & tableName & ": " & CStr(columnNames(i))
        End If
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Assert et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Assert contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_Assert(ByVal condition As Boolean, ByVal messageText As String)

    If Not condition Then
        CoreBridgePreCoreHarness_Trace "ASSERT FAIL " & messageText
        Err.Raise vbObjectError + 9511, "CoreBridgePreCoreHarness", messageText
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Trace et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Trace contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_Trace(ByVal messageText As String)

    Dim fileNo As Integer

    If Trim$(gPreCoreHarnessTracePath) = "" Then Exit Sub

    fileNo = FreeFile
    Open gPreCoreHarnessTracePath For Append As #fileNo
    Print #fileNo, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & messageText
    Close #fileNo

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Pre Core Harness Delete File et signale toute divergence au harnais.
' EN: Verifies the Pre Core Harness Delete File contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CoreBridgePreCoreHarness_DeleteFile(ByVal filePath As String)

    On Error Resume Next
    If Len(Dir$(filePath, vbNormal)) > 0 Then Kill filePath
    On Error GoTo 0

End Sub
