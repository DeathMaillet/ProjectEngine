Attribute VB_Name = "mod_ParsedNetworkHarness"
Option Explicit

'===============================================================================
' MODULE : mod_ParsedNetworkHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Parsed Network sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Parsed Network contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : ParsedPlanningNetworkHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Verifie le contrat Parsed Planning Network Harness Smoke et signale toute divergence au harnais.
' EN: Verifies the Parsed Planning Network Harness Smoke contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Public Function ParsedPlanningNetworkHarness_Smoke() As String

    Dim tblLinks As ListObject
    Dim columns As Object
    Dim network As clsParsedPlanningNetwork
    Dim link As clsParsedPlanningLink
    Dim arr As Variant
    Dim r As Long
    Dim expectedType As String
    Dim expectedLag As Double
    Dim countBefore As Long
    Dim mutationError As Long

    On Error GoTo Fail

    Set tblLinks = ThisWorkbook.Worksheets("CALC").ListObjects("tbl_LOGIC_LINKS")
    Set columns = CanonicalIdentity_BuildColumnMap(tblLinks)
    Set network = ParsedPlanningNetwork_ParseTable(tblLinks)

    If tblLinks.DataBodyRange Is Nothing Then
        ParsedPlanningNetworkHarness_Assert network.Count = 0, "empty table count"
    Else
        arr = tblLinks.DataBodyRange.value
        ParsedPlanningNetworkHarness_Assert network.Count = UBound(arr, 1), "row count parity"

        For r = 1 To network.Count
            Set link = network.Item(r)
            expectedType = UCase$(Trim$(CStr(arr(r, columns("Link Type")))))
            If expectedType = "" Then expectedType = "FS"
            expectedLag = 0#
            If IsNumeric(arr(r, columns("Lag"))) Then expectedLag = CDbl(arr(r, columns("Lag")))

            ParsedPlanningNetworkHarness_Assert link.RowIndex = r, "row order"
            ParsedPlanningNetworkHarness_Assert link.SuccId = Trim$(CStr(arr(r, columns("Succ ID")))), "Succ ID parity"
            ParsedPlanningNetworkHarness_Assert link.PredId = Trim$(CStr(arr(r, columns("Pred ID")))), "Pred ID parity"
            ParsedPlanningNetworkHarness_Assert link.LinkType = expectedType, "Link Type parity"
            ParsedPlanningNetworkHarness_Assert Abs(link.Lag - expectedLag) < 0.0000001, "Lag parity"
            If columns.Exists("Raw Token") Then
                ParsedPlanningNetworkHarness_Assert link.RawToken = Trim$(CStr(arr(r, columns("Raw Token")))), "Raw Token parity"
            End If
        Next r
    End If

    ParsedPlanningNetworkHarness_Assert network.HasColumn("Succ ID"), "Succ ID schema"
    ParsedPlanningNetworkHarness_Assert network.HasColumn("Pred ID"), "Pred ID schema"

    countBefore = network.Count
    On Error Resume Next
    CallByName network, "Add", VbMethod, Nothing
    mutationError = Err.Number
    Err.Clear
    On Error GoTo Fail
    ParsedPlanningNetworkHarness_Assert mutationError <> 0, "network rejects Add"
    ParsedPlanningNetworkHarness_Assert network.Count = countBefore, "network remains unchanged"

    ParsedPlanningNetworkHarness_Smoke = "PASS"
    Exit Function

Fail:
    ParsedPlanningNetworkHarness_Smoke = "FAIL: " & Err.Number & " - " & Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Parsed Planning Network Harness Assert et signale toute divergence au harnais.
' EN: Verifies the Parsed Planning Network Harness Assert contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub ParsedPlanningNetworkHarness_Assert(ByVal condition As Boolean, ByVal contractName As String)
    If Not condition Then Err.Raise vbObjectError + 3290, "ParsedPlanningNetworkHarness", contractName
End Sub
