Attribute VB_Name = "mod_OwnerDataAccessHarness"
Option Explicit

'===============================================================================
' MODULE : mod_OwnerDataAccessHarness
' DOMAINE / DOMAIN : Validation Harnesses
'
' FR
' Harnais de preuve du contrat Owner Data Access sur des copies de test.
' N'appartient a aucun workflow produit et ne doit pas etre appele en usage normal.
'
' EN
' Proof harness for the Owner Data Access contract on test copies.
' Is not production workflow code and must not run during normal use.
'
' CONTRATS / CONTRACTS : OwnerDataAccessHarness_Smoke
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


' Test-only harness. It seeds owner tables on a temporary workbook copy and
' validates the business projections without exposing Excel objects to callers.

'------------------------------------------------------------------------------
' FR: Valide les contrats proprietaires CalcState, Gantt Simulation et projection Dashboard S-Curve.
' EN: Validates CalcState, Gantt Simulation, and S-Curve Dashboard owner contracts.
'------------------------------------------------------------------------------
Public Function OwnerDataAccessHarness_Smoke() As String

    Dim wsState As Worksheet
    Dim tblState As ListObject
    Dim stateRow As ListRow
    Dim signatures As Object
    Dim wsSimulation As Worksheet
    Dim tblSimulation As ListObject
    Dim simulationRow As ListRow
    Dim resultById As Object
    Dim legacyResultById As Object
    Dim scenarioById As Object
    Dim resultValues As Variant
    Dim scenarioValues As Variant
    Dim wsSCurve As Worksheet
    Dim tblSCurve As ListObject
    Dim wsCalcSCurve As Worksheet
    Dim tblCalcSCurve As ListObject
    Dim scurveProjection As Object

    On Error GoTo Fail

    Ensure_CalcState_Table
    Set wsState = ThisWorkbook.Worksheets("CALC_STATE")
    Set tblState = wsState.ListObjects("tbl_CALC_STATE")
    OwnerDataAccessHarness_ClearRows tblState
    Set stateRow = tblState.ListRows.Add
    stateRow.Range.Cells(1, tblState.ListColumns("ID").Index).value = "OWNER-DATA-STATE-1"
    stateRow.Range.Cells(1, tblState.ListColumns("Row Signature").Index).value = "SIG-OWNER-1"
    stateRow.Range.Cells(1, tblState.ListColumns("Run Status").Index).value = "OK"

    Set signatures = CalcState_GetSignatureByIdMap()
    OwnerDataAccessHarness_Assert signatures.Count = 1, "CalcState signature count"
    OwnerDataAccessHarness_Assert signatures.Exists("OWNER-DATA-STATE-1"), "CalcState signature ID"
    OwnerDataAccessHarness_Assert CStr(signatures("OWNER-DATA-STATE-1")) = "SIG-OWNER-1", "CalcState signature value"
    OwnerDataAccessHarness_Assert CalcState_GetRunStatus() = "OK", "CalcState run status"
    CalcState_MarkDirty
    OwnerDataAccessHarness_Assert CStr(tblState.ListColumns("Run Status").DataBodyRange.Cells(1, 1).value) = "DIRTY", "CalcState dirty mutation"

    Set wsSimulation = Ensure_CalcGanttTest_Sheet()
    Set tblSimulation = Ensure_CalcGanttTest_Table(wsSimulation)
    ResizeTableToRowCount_Generic tblSimulation, 0
    Set simulationRow = tblSimulation.ListRows.Add
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "ID", "OWNER-DATA-SIM-1"
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "Task Type", "LEVEL_OF_EFFORT"
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "Is Summary", "No"
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "Calc Test Start", 46000
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "Calc Test Finish", 46004
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "Calc Test Duration", 5
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "Calc Test Progress", 0.4
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "Input Progress", 0.35
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "Error Flag", ""
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "Any Test Value", "Yes"

    Set resultById = GanttSimulation_BuildResultByIdMap()
    OwnerDataAccessHarness_Assert resultById.Exists("OWNER-DATA-SIM-1"), "Simulation result ID"
    resultValues = resultById("OWNER-DATA-SIM-1")
    OwnerDataAccessHarness_Assert CLng(resultValues(2)) = 5, "Simulation result duration"
    OwnerDataAccessHarness_Assert CStr(resultValues(6)) = "Yes", "Simulation Any Test Value"

    Set legacyResultById = GanttLive_BuildTestByIdMap()
    OwnerDataAccessHarness_Assert legacyResultById.Count = resultById.Count, "GanttLive facade count"
    resultValues = legacyResultById("OWNER-DATA-SIM-1")
    OwnerDataAccessHarness_Assert CLng(resultValues(2)) = 5, "GanttLive facade payload"

    Set scenarioById = GanttSimulation_BuildScenarioBaselineById()
    OwnerDataAccessHarness_Assert scenarioById.Exists("OWNER-DATA-SIM-1"), "Scenario baseline ID"
    scenarioValues = scenarioById("OWNER-DATA-SIM-1")
    OwnerDataAccessHarness_Assert CLng(scenarioValues(1)) = 5, "Scenario baseline duration"
    OwnerDataAccessHarness_Assert CBool(scenarioValues(4)), "Scenario baseline LOE alias"

    OwnerDataAccessHarness_Assert Not GanttSimulation_HasErrors(), "Simulation error state false"
    OwnerDataAccessHarness_SetValue tblSimulation, 1, "Error Flag", "ERROR"
    OwnerDataAccessHarness_Assert GanttSimulation_HasErrors(), "Simulation error state true"

    Set wsSCurve = ThisWorkbook.Worksheets("SCURVE")
    Set tblSCurve = wsSCurve.ListObjects("tbl_SCURVE")
    ResizeTableToRowCount_Generic tblSCurve, 3
    OwnerDataAccessHarness_SetValue tblSCurve, 1, "Date", CLng(Date) - 1
    OwnerDataAccessHarness_SetValue tblSCurve, 2, "Date", CLng(Date)
    OwnerDataAccessHarness_SetValue tblSCurve, 3, "Date", CLng(Date) + 1
    OwnerDataAccessHarness_SetValue tblSCurve, 1, "Cumulative Baseline", 0.1
    OwnerDataAccessHarness_SetValue tblSCurve, 2, "Cumulative Baseline", 0.4
    OwnerDataAccessHarness_SetValue tblSCurve, 3, "Cumulative Baseline", 0.8
    OwnerDataAccessHarness_SetValue tblSCurve, 1, "Cumulative Actual", 0.05
    OwnerDataAccessHarness_SetValue tblSCurve, 2, "Cumulative Actual", 0.25
    OwnerDataAccessHarness_SetValue tblSCurve, 3, "Cumulative Actual", 0.25
    OwnerDataAccessHarness_SetValue tblSCurve, 1, "Calculated Curve Dashed", 0.05
    OwnerDataAccessHarness_SetValue tblSCurve, 2, "Calculated Curve Dashed", 0.25
    OwnerDataAccessHarness_SetValue tblSCurve, 3, "Calculated Curve Dashed", 0.75

    Set wsCalcSCurve = ThisWorkbook.Worksheets("CALC_SCURVE")
    Set tblCalcSCurve = wsCalcSCurve.ListObjects("tbl_CALC_SCURVE")
    ResizeTableToRowCount_Generic tblCalcSCurve, 2
    OwnerDataAccessHarness_SetValue tblCalcSCurve, 1, "SCurve Actualized Weight", 0.2
    OwnerDataAccessHarness_SetValue tblCalcSCurve, 2, "SCurve Actualized Weight", 0.3

    Set scurveProjection = SCurve_BuildDashboardProjection()
    OwnerDataAccessHarness_Assert scurveProjection.Exists("DateRange"), "S-Curve date range"
    OwnerDataAccessHarness_Assert scurveProjection.Exists("BaselineRange"), "S-Curve baseline range"
    OwnerDataAccessHarness_Assert scurveProjection.Exists("ActualRange"), "S-Curve actual range"
    OwnerDataAccessHarness_Assert scurveProjection.Exists("ForecastRange"), "S-Curve forecast range"
    OwnerDataAccessHarness_Assert scurveProjection("DateRange").Rows.Count = 3, "S-Curve projection row count"
    OwnerDataAccessHarness_Assert Abs(CDbl(scurveProjection("PlannedProgress")) - 0.4) < 0.0000001, "S-Curve planned progress"
    OwnerDataAccessHarness_Assert Abs(CDbl(scurveProjection("ActualProgress")) - 0.5) < 0.0000001, "S-Curve actual progress"

    OwnerDataAccessHarness_Smoke = "PASS"
    Exit Function

Fail:
    OwnerDataAccessHarness_Smoke = "FAIL: " & Err.Number & " - " & Err.Description

End Function

'------------------------------------------------------------------------------
' FR: Supprime les lignes d'une table de fixture du harnais.
' EN: Deletes rows from a harness fixture table.
'------------------------------------------------------------------------------
Private Sub OwnerDataAccessHarness_ClearRows(ByVal tbl As ListObject)

    Do While tbl.ListRows.Count > 0
        tbl.ListRows(tbl.ListRows.Count).Delete
    Loop

End Sub

'------------------------------------------------------------------------------
' FR: Ecrit une valeur de fixture dans une colonne nommee.
' EN: Writes a fixture value into a named column.
'------------------------------------------------------------------------------
Private Sub OwnerDataAccessHarness_SetValue( _
    ByVal tbl As ListObject, _
    ByVal rowIndex As Long, _
    ByVal columnName As String, _
    ByVal value As Variant)

    tbl.ListColumns(columnName).DataBodyRange.Cells(rowIndex, 1).value = value

End Sub

'------------------------------------------------------------------------------
' FR: Arrete le harnais avec un diagnostic precis si un contrat echoue.
' EN: Stops the harness with a precise diagnostic when a contract fails.
'------------------------------------------------------------------------------
Private Sub OwnerDataAccessHarness_Assert(ByVal condition As Boolean, ByVal contractName As String)

    If Not condition Then Err.Raise vbObjectError + 3260, "OwnerDataAccessHarness", contractName

End Sub
