Attribute VB_Name = "mod_CanonicalIdentityHarness"
Option Explicit

'===============================================================================
' MODULE : mod_CanonicalIdentityHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Canonical Identity sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Canonical Identity contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : CanonicalIdentityHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' Test-only contract harness. PowerShell runs it on an isolated workbook copy.

Private Const HARNESS_WBS_SHEET As String = "WBS"
Private Const HARNESS_WBS_TABLE As String = "tbl_WBS"
Private Const HARNESS_CALC_SHEET As String = "CALC"
Private Const HARNESS_CALC_TABLE As String = "tbl_CALC"

'------------------------------------------------------------------------------
' FR: Prouve la parite, l'immuabilite et les politiques du Canonical Identity Index.
' EN: Proves Canonical Identity Index parity, immutability, and indexing policies.
'------------------------------------------------------------------------------
Public Function CanonicalIdentityHarness_Smoke() As String

    Dim tblWBS As ListObject
    Dim tblCalc As ListObject
    Dim newColumns As Object
    Dim newWbsToId As Object
    Dim newDriving As Object
    Dim expected As Object
    Dim calcColumns As Object
    Dim calcRows As Object
    Dim wbsRows As Object
    Dim wbsById As Object
    Dim arrCalc As Variant
    Dim arrWBS As Variant
    Dim r As Long
    Dim idVal As String
    Dim wbsVal As String
    Dim key As Variant

    On Error GoTo Fail

    CanonicalIdentityHarness_Trace "01 enter"
    Set tblWBS = ThisWorkbook.Worksheets(HARNESS_WBS_SHEET).ListObjects(HARNESS_WBS_TABLE)
    Set tblCalc = ThisWorkbook.Worksheets(HARNESS_CALC_SHEET).ListObjects(HARNESS_CALC_TABLE)

    CanonicalIdentityHarness_Trace "02 tables resolved"
    Set newColumns = CanonicalIdentity_BuildColumnMap(tblWBS)
    CanonicalIdentityHarness_Assert newColumns.Count = tblWBS.ListColumns.Count, "column map count"
    For r = 1 To tblWBS.ListColumns.Count
        CanonicalIdentityHarness_Assert newColumns.Exists(tblWBS.ListColumns(r).Name), "column map key"
        CanonicalIdentityHarness_Assert CLng(newColumns(tblWBS.ListColumns(r).Name)) = r, "column map index"
    Next r

    CanonicalIdentityHarness_Trace "03 columns parity"
    Set newWbsToId = CanonicalIdentity_BuildWbsToIdMap(tblWBS, newColumns)
    Set expected = CreateObject("Scripting.Dictionary")
    If Not tblWBS.DataBodyRange Is Nothing Then
        arrWBS = tblWBS.DataBodyRange.value
        For r = 1 To UBound(arrWBS, 1)
            wbsVal = NormalizeWBS(CStr(arrWBS(r, newColumns("WBS"))))
            idVal = Trim$(CStr(arrWBS(r, newColumns("ID"))))
            If wbsVal <> "" And idVal <> "" Then expected(wbsVal) = idVal
        Next r
    End If
    CanonicalIdentityHarness_AssertMapsEqual expected, newWbsToId, "WBS to ID policy"

    CanonicalIdentityHarness_Trace "04 WBS parity"
    Set newDriving = CanonicalIdentity_GetDrivingLogicByIdMap()
    Set expected = CreateObject("Scripting.Dictionary")
    If Not tblCalc.DataBodyRange Is Nothing Then
        Set calcColumns = CanonicalIdentity_BuildColumnMap(tblCalc)
        arrCalc = tblCalc.DataBodyRange.value
        For r = 1 To UBound(arrCalc, 1)
            idVal = Trim$(CStr(arrCalc(r, calcColumns("ID"))))
            If idVal <> "" Then expected(idVal) = UCase$(Trim$(CStr(arrCalc(r, calcColumns("Driving Logic")))))
        Next r
    End If
    CanonicalIdentityHarness_AssertMapsEqual expected, newDriving, "Driving Logic policy"

    CanonicalIdentityHarness_Trace "05 driving parity"
    Set calcColumns = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set calcRows = CanonicalIdentity_GetCalcRowByIdMap()

    If tblCalc.DataBodyRange Is Nothing Then
        CanonicalIdentityHarness_Assert calcRows.Count = 0, "empty CALC row index"
    Else
        arrCalc = tblCalc.DataBodyRange.value
        Set calcRows = CanonicalIdentity_BuildCalcRowById(arrCalc, calcColumns)
        CanonicalIdentityHarness_Assert calcRows.Count <= UBound(arrCalc, 1), "CALC row index count"

        For Each key In calcRows.Keys
            CanonicalIdentityHarness_Assert _
                Trim$(CStr(arrCalc(CLng(calcRows(key)), calcColumns("ID")))) = CStr(key), _
                "CALC row index value"
        Next key
    End If

    Set wbsRows = CanonicalIdentity_GetWbsRowByIdMap()
    Set wbsById = CanonicalIdentity_GetWbsByIdMap()
    CanonicalIdentityHarness_Assert wbsRows.Count = wbsById.Count, "WBS ID projections share identity policy"

    CanonicalIdentityHarness_Trace "06 row projections"
    CanonicalIdentityHarness_AssertReadOnly newColumns, "column map immutable"
    CanonicalIdentityHarness_AssertReadOnly newWbsToId, "WBS to ID map immutable"
    CanonicalIdentityHarness_AssertReadOnly newDriving, "Driving Logic map immutable"
    CanonicalIdentityHarness_AssertMissingKeyRead newColumns, "missing-key read contract"

    CanonicalIdentityHarness_Trace "07 immutable"
    CanonicalIdentityHarness_Smoke = "PASS"
    Exit Function

Fail:
    CanonicalIdentityHarness_Trace "FAIL " & CStr(Err.Number) & " " & Err.Description
    CanonicalIdentityHarness_Smoke = "FAIL: " & Err.Number & " - " & Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Verifie le contrat Canonical Identity Harness Assert Maps Equal et signale toute divergence au harnais.
' EN: Verifies the Canonical Identity Harness Assert Maps Equal contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CanonicalIdentityHarness_AssertMapsEqual( _
    ByVal expected As Object, _
    ByVal actual As Object, _
    ByVal contractName As String)

    Dim key As Variant

    CanonicalIdentityHarness_Assert expected.Count = actual.Count, contractName & " count"

    For Each key In expected.Keys
        CanonicalIdentityHarness_Assert actual.Exists(key), contractName & " key " & CStr(key)
        CanonicalIdentityHarness_Assert CStr(expected(key)) = CStr(actual(key)), contractName & " value " & CStr(key)
    Next key

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Canonical Identity Harness Assert Missing Key Read et signale toute divergence au harnais.
' EN: Verifies the Canonical Identity Harness Assert Missing Key Read contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CanonicalIdentityHarness_AssertMissingKeyRead( _
    ByVal candidate As Object, _
    ByVal contractName As String)

    Dim missingValue As Variant
    Dim countBefore As Long

    countBefore = candidate.Count
    missingValue = candidate("__CANONICAL_MISSING_KEY__")

    CanonicalIdentityHarness_Assert IsEmpty(missingValue), contractName & " returns Empty"
    CanonicalIdentityHarness_Assert candidate.Count = countBefore, contractName & " preserves count"
    CanonicalIdentityHarness_Assert Not candidate.Exists("__CANONICAL_MISSING_KEY__"), contractName & " does not add key"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Canonical Identity Harness Trace et signale toute divergence au harnais.
' EN: Verifies the Canonical Identity Harness Trace contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CanonicalIdentityHarness_Trace(ByVal message As String)

    Dim fileNumber As Integer
    Dim tracePath As String

    On Error Resume Next
    tracePath = Environ$("TEMP") & "\canonical_identity_trace.txt"
    fileNumber = FreeFile
    Open tracePath For Append As #fileNumber
    Print #fileNumber, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & message
    Close #fileNumber
    On Error GoTo 0

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Canonical Identity Harness Assert Read Only et signale toute divergence au harnais.
' EN: Verifies the Canonical Identity Harness Assert Read Only contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CanonicalIdentityHarness_AssertReadOnly( _
    ByVal candidate As Object, _
    ByVal contractName As String)

    Dim mutationError As Long

    On Error Resume Next
    CallByName candidate, "Add", VbMethod, "__CANONICAL_MUTATION__", 1
    mutationError = Err.Number
    Err.Clear
    On Error GoTo 0

    CanonicalIdentityHarness_Assert mutationError <> 0, contractName & " rejects Add"
    CanonicalIdentityHarness_Assert Not candidate.Exists("__CANONICAL_MUTATION__"), contractName & " remains unchanged"

End Sub

'------------------------------------------------------------------------------
' FR: Verifie le contrat Canonical Identity Harness Assert et signale toute divergence au harnais.
' EN: Verifies the Canonical Identity Harness Assert contract and reports any divergence to the harness.
'------------------------------------------------------------------------------

Private Sub CanonicalIdentityHarness_Assert(ByVal condition As Boolean, ByVal contractName As String)

    If Not condition Then Err.Raise vbObjectError + 3280, "CanonicalIdentityHarness", contractName

End Sub
