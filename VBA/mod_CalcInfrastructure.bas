Attribute VB_Name = "mod_CalcInfrastructure"
Option Explicit

'===============================================================================
' MODULE : mod_CalcInfrastructure
' DOMAINE / DOMAIN : Calculation / Data Sync
'
' FR
' Cree et valide les feuilles, tables et colonnes techniques requises par le calcul.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Creates and validates the technical sheets, tables and columns required by calculation.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Ensure_Calc_Infrastructure
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


Private Const CALC_SHEET_NAME As String = "CALC"
Private Const CALC_TABLE_NAME As String = "tbl_CALC"
Private Const LOGIC_TABLE_NAME As String = "tbl_LOGIC_LINKS"

Private Const CALC_TABLE_TOP_LEFT As String = "A1"
Private Const TABLE_GAP_COLS As Long = 3


'------------------------------------------------------------------------------
' FR: Verifie ou cree Calc Infrastructure si necessaire.
' EN: Ensures or creates Calc Infrastructure when needed.
'------------------------------------------------------------------------------
Public Sub Ensure_Calc_Infrastructure(Optional ByVal consoleMessages As Collection)

    Dim perfScope As clsPerfScope

    Dim wsCalc As Worksheet
    Dim tblCalc As ListObject
    Dim tblLogic As ListObject
    Dim logicTopLeft As Range

    Dim calcSheetWasCreated As Boolean
    Dim calcTableWasCreated As Boolean
    Dim logicTableWasCreated As Boolean

    Set perfScope = Profiler_BeginScope("Ensure_Calc_Infrastructure", "Excel Infrastructure")

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    Set wsCalc = EnsureWorksheetExists(CALC_SHEET_NAME, calcSheetWasCreated)

    Set tblCalc = EnsureTableWithHeaders( _
        wsCalc, _
        CALC_TABLE_NAME, _
        wsCalc.Range(CALC_TABLE_TOP_LEFT), _
        GetCalcHeaders(), _
        calcTableWasCreated _
    )

    Set logicTopLeft = GetLogicTableTopLeft(wsCalc)

    Set tblLogic = EnsureTableWithHeaders( _
        wsCalc, _
        LOGIC_TABLE_NAME, _
        logicTopLeft, _
        GetLogicLinksHeaders(), _
        logicTableWasCreated _
    )

    If calcSheetWasCreated Or calcTableWasCreated Then
        Apply_tbl_CALC_ColumnFormats tblCalc
    End If

SafeExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True

    CalcInfrastructure_AddOrShowConsoleMessage consoleMessages, "STOP", _
        "Erreur Ensure_Calc_Infrastructure" & vbCrLf & _
        "-> " & Err.Description, _
        "Error Ensure_Calc_Infrastructure" & vbCrLf & _
        "-> " & Err.Description

End Sub



'------------------------------------------------------------------------------
' FR: Verifie ou cree Worksheet Exists si necessaire.
' EN: Ensures or creates Worksheet Exists when needed.
'------------------------------------------------------------------------------
Private Function EnsureWorksheetExists(ByVal sheetName As String, Optional ByRef wasCreated As Boolean = False) As Worksheet

    Dim ws As Worksheet
    Dim wsStart As Worksheet

    wasCreated = False
    Set wsStart = ActiveSheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        wasCreated = True

        If Not wsStart Is Nothing Then
            wsStart.Activate
        End If
    End If

    Set EnsureWorksheetExists = ws

End Function

'------------------------------------------------------------------------------
' FR: Retourne Logic Table Top Left depuis le contexte calculation infrastructure.
' EN: Returns Logic Table Top Left from the calculation infrastructure context.
'------------------------------------------------------------------------------
Private Function GetLogicTableTopLeft(ByVal ws As Worksheet) As Range

    Dim calcHeaders As Variant
    Dim calcHeaderCount As Long
    Dim startCol As Long

    calcHeaders = GetCalcHeaders()
    calcHeaderCount = UBound(calcHeaders) - LBound(calcHeaders) + 1

    startCol = ws.Range(CALC_TABLE_TOP_LEFT).Column + calcHeaderCount + TABLE_GAP_COLS

    Set GetLogicTableTopLeft = ws.Cells(1, startCol)

End Function

'------------------------------------------------------------------------------
' FR: Verifie ou cree Table With Headers si necessaire.
' EN: Ensures or creates Table With Headers when needed.
'------------------------------------------------------------------------------
Private Function EnsureTableWithHeaders( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal headers As Variant, _
    Optional ByRef wasCreated As Boolean = False) As ListObject

    Dim tbl As ListObject
    Dim i As Long
    Dim headerCount As Long
    Dim rng As Range
    Dim lc As ListColumn
    Dim headerName As String

    wasCreated = False

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0

    headerCount = UBound(headers) - LBound(headers) + 1

    If tbl Is Nothing Then

        Set rng = ws.Range(topLeft, topLeft.Offset(1, headerCount - 1))

        rng.Clear

        For i = 0 To headerCount - 1
            rng.Cells(1, i + 1).value = headers(LBound(headers) + i)
        Next i

        Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        tbl.Name = tableName
        wasCreated = True

        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.ClearContents
        End If

    Else

        For i = LBound(headers) To UBound(headers)
            headerName = CStr(headers(i))

            Set lc = Nothing
            On Error Resume Next
            Set lc = tbl.ListColumns(headerName)
            On Error GoTo 0

            If lc Is Nothing Then
                Set lc = tbl.ListColumns.Add
                lc.Name = headerName
            End If
        Next i

    End If

    Set EnsureTableWithHeaders = tbl

End Function

'------------------------------------------------------------------------------
' FR: Retourne Calc Headers depuis le contexte calculation infrastructure.
' EN: Returns Calc Headers from the calculation infrastructure context.
'------------------------------------------------------------------------------
Private Function GetCalcHeaders() As Variant

    Dim arr
    arr = Array( _
        "ID", "WBS", "Task Name", "ParentID", "IsSummary", "S", "Predecessors WBS", _
        "Cal", _
        "Baseline Start", "Baseline Duration", "Baseline Finish", "Actual Start", "Actual Finish", "Actual Duration", _
        "Forecast Start", "Forecast Finish", "Deadline", "Calculated Start", "Calculated Finish", "Calculated Duration", _
        "Driving Logic", "Constraint Active", "Start Constraint Type", "Start Constraint Date", _
        "Finish Constraint Type", "Finish Constraint Date", "Error flag", "ErrorMsg", "Critical Path", "Longest Path", "Total Float", "Free Float", "Deadline Float", _
        "Critical Path REX", "Total Float REX", "Free Float REX" _
    )

    GetCalcHeaders = arr

End Function

'------------------------------------------------------------------------------
' FR: Retourne Logic Links Headers depuis le contexte calculation infrastructure.
' EN: Returns Logic Links Headers from the calculation infrastructure context.
'------------------------------------------------------------------------------
Private Function GetLogicLinksHeaders() As Variant

    Dim arr
    arr = Array("Succ ID", "Succ WBS", "Pred ID", "Pred WBS", "Link Type", "Lag", "Raw Token", "Expanded From")

    GetLogicLinksHeaders = arr

End Function

'------------------------------------------------------------------------------
' FR: Cree tbl CALC pour le domaine calculation infrastructure.
' EN: Creates tbl CALC for the calculation infrastructure domain.
'------------------------------------------------------------------------------
Private Function Create_tbl_CALC(ByVal wsCalc As Worksheet) As ListObject

    Dim headers1 As Variant
    Dim headers2 As Variant
    Dim headers3 As Variant
    Dim headers As Variant

    Dim lastCol As Long
    Dim rng As Range
    Dim tbl As ListObject

    headers1 = Array( _
        "ID", _
        "Predecessors WBS", _
        "Baseline Start", _
        "Baseline Duration", _
        "Baseline Finish", _
        "Actual Start", _
        "Actual Finish", _
        "Actual Duration" _
    )

    headers2 = Array( _
        "Forecast Start", _
        "Forecast Finish", _
        "WBS", _
        "ParentID", _
        "IsSummary", _
        "S", _
        "Calculated Start", _
        "Calculated Finish", _
        "Calculated Duration" _
    )

    headers3 = Array( _
        "Driving Logic", _
        "Error flag", _
        "ErrorMsg", _
        "Critical Path", _
        "Longest Path", _
        "Total Float", _
        "Free Float", _
        "Deadline Float", _
        "Critical Path REX", _
        "Total Float REX", _
        "Free Float REX" _
    )

    headers = JoinArrays(headers1, headers2, headers3)

    wsCalc.Cells.Clear

    lastCol = UBound(headers) - LBound(headers) + 1
    Set rng = wsCalc.Range("A1").Resize(2, lastCol)

    WriteHeaderArrayToRange rng.rows(1), headers
    rng.rows(2).ClearContents

    Set tbl = wsCalc.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "tbl_CALC"

    Set Create_tbl_CALC = tbl

End Function


'------------------------------------------------------------------------------
' FR: Retourne la valeur Join Arrays sans modifier les donnees d'entree.
' EN: Returns the Join Arrays value without mutating input data.
'------------------------------------------------------------------------------

Private Function JoinArrays(ParamArray arrs() As Variant) As Variant

    Dim totalCount As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim arr As Variant
    Dim result() As Variant

    totalCount = 0
    For i = LBound(arrs) To UBound(arrs)
        arr = arrs(i)
        totalCount = totalCount + (UBound(arr) - LBound(arr) + 1)
    Next i

    ReDim result(0 To totalCount - 1)
    k = 0

    For i = LBound(arrs) To UBound(arrs)
        arr = arrs(i)
        For j = LBound(arr) To UBound(arr)
            result(k) = arr(j)
            k = k + 1
        Next j
    Next i

    JoinArrays = result

End Function

'------------------------------------------------------------------------------
' FR: Ecrit Header Array To Range vers le stockage cible.
' EN: Writes Header Array To Range to the target storage.
'------------------------------------------------------------------------------
Private Sub WriteHeaderArrayToRange(ByVal targetRow As Range, ByVal headers As Variant)

    Dim i As Long

    For i = LBound(headers) To UBound(headers)
        targetRow.Cells(1, i - LBound(headers) + 1).value = headers(i)
    Next i

End Sub

'------------------------------------------------------------------------------
' FR: Actualise Apply tbl CALC Column Formats sans modifier les regles metier qui produisent les donnees.
' EN: Refreshes Apply tbl CALC Column Formats without changing the business rules that produce the data.
'------------------------------------------------------------------------------

Private Sub Apply_tbl_CALC_ColumnFormats(ByVal tblCalc As ListObject)

    On Error Resume Next

    ' Dates
    tblCalc.ListColumns("Baseline Start").Range.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Baseline Finish").Range.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Actual Start").Range.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Actual Finish").Range.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Forecast Start").Range.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Forecast Finish").Range.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Calculated Start").Range.NumberFormat = "dd/mm/yyyy"
    tblCalc.ListColumns("Calculated Finish").Range.NumberFormat = "dd/mm/yyyy"

    ' Text
    tblCalc.ListColumns("Predecessors WBS").Range.NumberFormat = "@"
    tblCalc.ListColumns("WBS").Range.NumberFormat = "@"
    tblCalc.ListColumns("Task Name").Range.NumberFormat = "@"
    tblCalc.ListColumns("Cal").Range.NumberFormat = "@"
    tblCalc.ListColumns("Driving Logic").Range.NumberFormat = "@"
    tblCalc.ListColumns("Error flag").Range.NumberFormat = "@"
    tblCalc.ListColumns("ErrorMsg").Range.NumberFormat = "@"
    tblCalc.ListColumns("Critical Path").Range.NumberFormat = "@"
    tblCalc.ListColumns("Longest Path").Range.NumberFormat = "@"
    tblCalc.ListColumns("Critical Path REX").Range.NumberFormat = "@"
    tblCalc.ListColumns("IsSummary").Range.NumberFormat = "@"
    tblCalc.ListColumns("S").Range.NumberFormat = "@"

    ' Numeric / general
    tblCalc.ListColumns("ID").Range.NumberFormat = "0"
    tblCalc.ListColumns("ParentID").Range.NumberFormat = "0"
    tblCalc.ListColumns("Baseline Duration").Range.NumberFormat = "0"
    tblCalc.ListColumns("Actual Duration").Range.NumberFormat = "0"
    tblCalc.ListColumns("Calculated Duration").Range.NumberFormat = "0"
    tblCalc.ListColumns("Total Float").Range.NumberFormat = "0"
    tblCalc.ListColumns("Free Float").Range.NumberFormat = "0"
    tblCalc.ListColumns("Deadline Float").Range.NumberFormat = "0"
    tblCalc.ListColumns("Total Float REX").Range.NumberFormat = "0"
    tblCalc.ListColumns("Free Float REX").Range.NumberFormat = "0"

    On Error GoTo 0

End Sub



'------------------------------------------------------------------------------
' FR: Ajoute la collection Calc Infrastructure Add Or Show Console Message a la structure cible fournie par l'appelant.
' EN: Adds the Calc Infrastructure Add Or Show Console Message collection to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Private Sub CalcInfrastructure_AddOrShowConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    CalcBridge_AddOrShowConsoleMessage consoleMessages, msgType, frText, enText

End Sub


