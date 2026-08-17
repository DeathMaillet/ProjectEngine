Attribute VB_Name = "mod_ParsedPlanningNetwork"
Option Explicit

'===============================================================================
' MODULE : mod_ParsedPlanningNetwork
' DOMAINE / DOMAIN : Parsed Planning Network
'
' FR
' Lit LOGIC_LINKS et construit le reseau canonique de records immuables Succ/Pred/Type/Lag.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Reads LOGIC_LINKS and builds the canonical network of immutable Succ/Pred/Type/Lag records.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : ParsedPlanningNetwork_LoadCanonical, ParsedPlanningNetwork_ParseTable
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const NETWORK_CALC_SHEET As String = "CALC"
Private Const NETWORK_LINKS_TABLE As String = "tbl_LOGIC_LINKS"

'------------------------------------------------------------------------------
' FR: Lit tbl_LOGIC_LINKS et retourne son snapshot structurel immuable et ordonne.
' EN: Reads tbl_LOGIC_LINKS and returns its immutable ordered structural snapshot.
'------------------------------------------------------------------------------
Public Function ParsedPlanningNetwork_LoadCanonical() As clsParsedPlanningNetwork

    Dim tblLinks As ListObject

    Set tblLinks = ThisWorkbook.Worksheets(NETWORK_CALC_SHEET).ListObjects(NETWORK_LINKS_TABLE)
    Set ParsedPlanningNetwork_LoadCanonical = ParsedPlanningNetwork_ParseTable(tblLinks)

End Function

'------------------------------------------------------------------------------
' FR: Parse une table de liens sans filtrage, validation ou projection metier.
' EN: Parses a link table without business filtering, validation, or projection.
'------------------------------------------------------------------------------
Public Function ParsedPlanningNetwork_ParseTable( _
    ByVal tblLinks As ListObject) As clsParsedPlanningNetwork

    Dim perfScope As clsPerfScope
    Dim columns As Object
    Dim links As Collection
    Dim network As clsParsedPlanningNetwork
    Dim link As clsParsedPlanningLink
    Dim sourceData As Variant
    Dim r As Long
    Dim succId As String
    Dim succWbs As String
    Dim predId As String
    Dim predWbs As String
    Dim rawLinkType As String
    Dim linkType As String
    Dim lag As Double
    Dim rawToken As String
    Dim expandedFrom As String

    Set perfScope = Profiler_BeginScope("ParsedPlanningNetwork_ParseTable", "Excel Read")
    Set links = New Collection
    Set columns = CanonicalIdentity_BuildColumnMap(tblLinks)

    If Not tblLinks.DataBodyRange Is Nothing Then
        sourceData = tblLinks.DataBodyRange.value

        For r = 1 To UBound(sourceData, 1)
            succId = ParsedPlanningNetwork_StringValue(sourceData, r, columns, "Succ ID", False)
            succWbs = ParsedPlanningNetwork_StringValue(sourceData, r, columns, "Succ WBS", True)
            predId = ParsedPlanningNetwork_StringValue(sourceData, r, columns, "Pred ID", False)
            predWbs = ParsedPlanningNetwork_StringValue(sourceData, r, columns, "Pred WBS", True)
            rawLinkType = UCase$(ParsedPlanningNetwork_StringValue(sourceData, r, columns, "Link Type", False))
            linkType = rawLinkType
            If columns.Exists("Link Type") And linkType = "" Then linkType = "FS"
            lag = ParsedPlanningNetwork_LagValue(sourceData, r, columns)
            rawToken = ParsedPlanningNetwork_StringValue(sourceData, r, columns, "Raw Token", False)
            expandedFrom = ParsedPlanningNetwork_StringValue(sourceData, r, columns, "Expanded From", False)

            Set link = New clsParsedPlanningLink
            link.Initialize r, succId, succWbs, predId, predWbs, rawLinkType, linkType, lag, rawToken, expandedFrom
            links.Add link
        Next r
    End If

    Set network = New clsParsedPlanningNetwork
    network.InitializeNetworkFrom links, columns
    Set ParsedPlanningNetwork_ParseTable = network

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate String Value selon le contrat canonique du composant.
' EN: Normalizes or formats String Value according to the component contract.
'------------------------------------------------------------------------------

Private Function ParsedPlanningNetwork_StringValue( _
    ByRef sourceData As Variant, _
    ByVal rowIndex As Long, _
    ByVal columns As Object, _
    ByVal columnName As String, _
    ByVal shouldNormalizeWbs As Boolean) As String

    Dim value As String

    If Not columns.Exists(columnName) Then Exit Function
    value = Trim$(CStr(sourceData(rowIndex, columns(columnName))))
    If shouldNormalizeWbs Then value = NormalizeWBS(value)
    ParsedPlanningNetwork_StringValue = value

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Lag Value selon le contrat canonique du composant.
' EN: Normalizes or formats Lag Value according to the component contract.
'------------------------------------------------------------------------------

Private Function ParsedPlanningNetwork_LagValue( _
    ByRef sourceData As Variant, _
    ByVal rowIndex As Long, _
    ByVal columns As Object) As Double

    If Not columns.Exists("Lag") Then Exit Function
    If IsNumeric(sourceData(rowIndex, columns("Lag"))) Then
        ParsedPlanningNetwork_LagValue = CDbl(sourceData(rowIndex, columns("Lag")))
    End If

End Function
