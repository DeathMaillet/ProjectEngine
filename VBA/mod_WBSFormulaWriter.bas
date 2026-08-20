Attribute VB_Name = "mod_WBSFormulaWriter"
Option Explicit

'===============================================================================
' MODULE : mod_WBSFormulaWriter
' DOMAINE / DOMAIN : WBS
'
' FR
' Restaure les formules WBS gerees sous un scope d'ecriture explicitement possede.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Restores managed WBS formulas under an explicitly owned write scope.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : RestoreWBSFormulaColumns
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Restaure les formules et formats geres des colonnes calculees de tbl_WBS.
' EN: Restores managed formulas and formats in calculated tbl_WBS columns.
'------------------------------------------------------------------------------
Public Sub RestoreWBSFormulaColumns(ByVal tblWBS As ListObject)

    Dim perfScope As clsPerfScope
    Dim consoleMessages As Collection
    Dim authorizedFields As Variant
    Dim baselineFinishCol As ListColumn
    Dim actualDurationCol As ListColumn
    Dim calculatedDurationCol As ListColumn
    Dim writeScopeToken As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    Set perfScope = Profiler_BeginScope("RestoreWBSFormulaColumns", "Excel Formula Restore")

    On Error GoTo SafeExit

    Set consoleMessages = New Collection

    If tblWBS Is Nothing Then Exit Sub
    If tblWBS.DataBodyRange Is Nothing Then Exit Sub

    On Error Resume Next
    Set baselineFinishCol = tblWBS.ListColumns("Baseline Finish")
    Set actualDurationCol = tblWBS.ListColumns("Actual Duration")
    Set calculatedDurationCol = tblWBS.ListColumns("Calculated Duration")
    Err.Clear
    On Error GoTo SafeExit

    authorizedFields = Array("Baseline Finish", "Actual Duration", "Calculated Duration")
    writeScopeToken = OpenAuthorizedWBSWriteScope( _
        "RestoreWBSFormulaColumns", authorizedFields)

    If Not baselineFinishCol Is Nothing Then
        RestoreWBSFormulaColumnIfNeeded baselineFinishCol, _
            "=IF(OR([@[Baseline Start]]="""",[@[Baseline Duration]]=""""),"""",[@[Baseline Start]]+[@[Baseline Duration]]-1)", _
            "dd/mm/yyyy"
    End If

    If Not actualDurationCol Is Nothing Then
        RestoreWBSFormulaColumnIfNeeded actualDurationCol, _
            "=IF(OR([@[Actual Start]]="""",[@[Actual Finish]]=""""),"""",[@[Actual Finish]]-[@[Actual Start]]+1)", _
            "0"
    End If

    If Not calculatedDurationCol Is Nothing Then
        RestoreWBSFormulaColumnIfNeeded calculatedDurationCol, _
            "=IF(OR([@[Calculated Start]]="""",[@[Calculated Finish]]=""""),"""",[@[Calculated Finish]]-[@[Calculated Start]]+1)", _
            "0"
    End If

SafeExit:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    CloseAuthorizedWBSWriteScope writeScopeToken
    On Error GoTo 0

    If errorNumber <> 0 Then
        If consoleMessages Is Nothing Then Set consoleMessages = New Collection
        WBSFormulaWriter_AddConsoleMessage consoleMessages, "STOP", _
            "Erreur dans RestoreWBSFormulaColumns : " & errorDescription, _
            "Error in RestoreWBSFormulaColumns: " & errorDescription
        CalcBridge_ShowPlanningConsole consoleMessages
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Reecrit une colonne de formule WBS uniquement si sa formule ou son format diverge.
' EN: Rewrites a WBS formula column only when its formula or format differs.
'------------------------------------------------------------------------------
Private Sub RestoreWBSFormulaColumnIfNeeded( _
    ByVal targetColumn As ListColumn, _
    ByVal expectedFormula As String, _
    ByVal expectedNumberFormat As String)

    Dim targetRange As Range
    Dim currentFormulas As Variant
    Dim currentValue As Variant
    Dim currentFormat As Variant
    Dim rowCount As Long
    Dim r As Long
    Dim needsWrite As Boolean

    If targetColumn Is Nothing Then Exit Sub
    Set targetRange = targetColumn.DataBodyRange
    If targetRange Is Nothing Then Exit Sub

    rowCount = targetRange.Rows.Count
    currentFormulas = targetRange.Formula

    For r = 1 To rowCount
        If rowCount = 1 And Not IsArray(currentFormulas) Then
            currentValue = currentFormulas
        Else
            currentValue = currentFormulas(r, 1)
        End If

        If IsError(currentValue) Or IsNull(currentValue) Then
            needsWrite = True
        ElseIf StrComp(CStr(currentValue), expectedFormula, vbTextCompare) <> 0 Then
            needsWrite = True
        End If
    Next r

    If needsWrite Then targetRange.Formula = expectedFormula

    currentFormat = targetRange.NumberFormat
    If IsNull(currentFormat) Then
        targetRange.NumberFormat = expectedNumberFormat
    ElseIf StrComp(CStr(currentFormat), expectedNumberFormat, vbTextCompare) <> 0 Then
        targetRange.NumberFormat = expectedNumberFormat
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute une erreur de restauration des formules WBS a la console.
' EN: Adds a WBS formula restoration error to the console.
'------------------------------------------------------------------------------
Private Sub WBSFormulaWriter_AddConsoleMessage( _
    ByVal consoleMessages As Collection, _
    ByVal msgType As String, _
    ByVal frText As String, _
    ByVal enText As String)

    If consoleMessages Is Nothing Then Exit Sub

    CalcBridge_AddConsoleMessage consoleMessages, msgType, _
        BiMsg(frText, enText)

End Sub
