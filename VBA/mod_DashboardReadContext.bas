Attribute VB_Name = "mod_DashboardReadContext"
Option Explicit

'===============================================================================
' MODULE : mod_DashboardReadContext
' DOMAINE / DOMAIN : Dashboard
'
' FR
' Charge le contexte de lecture WBS, CALC et S-Curve commun aux trois modes Dashboard.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Loads the WBS, CALC and S-Curve read context shared by the three Dashboard modes.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : DashboardReadContext_Load
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const DASHBOARD_WBS_SHEET As String = "WBS"
Private Const DASHBOARD_WBS_TABLE As String = "tbl_WBS"
Private Const DASHBOARD_CALC_SHEET As String = "CALC"
Private Const DASHBOARD_CALC_TABLE As String = "tbl_CALC"

'------------------------------------------------------------------------------
' FR: Construit le contexte de lecture commun aux trois workflows Dashboard.
' EN: Builds the read context shared by the three Dashboard workflows.
'------------------------------------------------------------------------------
Public Function DashboardReadContext_Load() As clsDashboardReadContext

    Dim context As clsDashboardReadContext
    Dim tblWBS As ListObject
    Dim tblCalc As ListObject
    Dim mapWBS As Object
    Dim mapCalc As Object
    Dim scurveProjection As Object

    Set tblWBS = ThisWorkbook.Worksheets(DASHBOARD_WBS_SHEET).ListObjects(DASHBOARD_WBS_TABLE)
    Set tblCalc = ThisWorkbook.Worksheets(DASHBOARD_CALC_SHEET).ListObjects(DASHBOARD_CALC_TABLE)
    Set mapWBS = CanonicalIdentity_BuildColumnMap(tblWBS)
    Set mapCalc = CanonicalIdentity_BuildColumnMap(tblCalc)
    Set scurveProjection = SCurve_BuildDashboardProjection()

    Set context = New clsDashboardReadContext
    context.InitializeDashboardReadContext tblWBS, tblCalc, mapWBS, mapCalc, scurveProjection
    Set DashboardReadContext_Load = context

End Function
