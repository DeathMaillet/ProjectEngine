Attribute VB_Name = "mod_CalcCoreShared"
Option Explicit

'===============================================================================
' MODULE : mod_CalcCoreShared
' DOMAINE / DOMAIN : Core Calculation
'
' FR
' Fournit les primitives partagees du Core pour dates, valeurs, task types et acces aux colonnes.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Provides shared Core primitives for dates, values, task types and column access.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : Core_GetVal, Core_SetVal, Core_HasVal, Core_ClearCalcOutputs_Row, Core_ClearAllCalcOutputs, Core_IsSummaryRow, Core_CalcInclusiveDuration, Core_SetCalcTriplet
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================

'------------------------------------------------------------------------------
' FR: Retourne la map Val sans exposer de mutateur sur l'etat source.
' EN: Returns the Val map without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function Core_GetVal( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal colName As String) As Variant

    If mapCol Is Nothing Then
        Core_GetVal = Empty
        Exit Function
    End If

    If Not mapCol.Exists(colName) Then
        Core_GetVal = Empty
        Exit Function
    End If

    Core_GetVal = dataArr(rowIdx, mapCol(colName))

End Function

'------------------------------------------------------------------------------
' FR: Active ou initialise Core Set Val dans l'etat runtime du composant.
' EN: Activates or initializes Core Set Val in the component runtime state.
'------------------------------------------------------------------------------

Public Sub Core_SetVal( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal colName As String, _
    ByVal newValue As Variant)

    If mapCol Is Nothing Then Exit Sub
    If Not mapCol.Exists(colName) Then Exit Sub

    dataArr(rowIdx, mapCol(colName)) = newValue

End Sub

'------------------------------------------------------------------------------
' FR: Indique si la map Val satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Val map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Public Function Core_HasVal( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal colName As String) As Boolean

    Dim v As Variant

    If mapCol Is Nothing Then Exit Function
    If Not mapCol.Exists(colName) Then Exit Function

    v = dataArr(rowIdx, mapCol(colName))
    Core_HasVal = HasValue(v)

End Function

'------------------------------------------------------------------------------
' FR: Reinitialise Core Clear Calc Outputs Row dans le perimetre possede par le composant.
' EN: Resets Core Clear Calc Outputs Row within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub Core_ClearCalcOutputs_Row( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object)

    Core_SetVal dataArr, rowIdx, mapCol, "Calculated Start", Empty
    Core_SetVal dataArr, rowIdx, mapCol, "Calculated Finish", Empty
    Core_SetVal dataArr, rowIdx, mapCol, "Calculated Duration", Empty
    Core_SetVal dataArr, rowIdx, mapCol, "Error flag", ""
    Core_SetVal dataArr, rowIdx, mapCol, "ErrorMsg", ""

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Core Clear All Calc Outputs dans le perimetre possede par le composant.
' EN: Resets Core Clear All Calc Outputs within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub Core_ClearAllCalcOutputs( _
    ByRef dataArr As Variant, _
    ByVal mapCol As Object)

    Dim r As Long

    If IsEmpty(dataArr) Then Exit Sub

    For r = LBound(dataArr, 1) To UBound(dataArr, 1)
        Core_ClearCalcOutputs_Row dataArr, r, mapCol
    Next r

End Sub

'------------------------------------------------------------------------------
' FR: Indique si la map Summary Row satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Summary Row map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Public Function Core_IsSummaryRow( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object) As Boolean

    Dim v As Variant
    Dim txt As String

    If mapCol Is Nothing Then Exit Function
    If Not mapCol.Exists("IsSummary") Then Exit Function

    v = Core_GetVal(dataArr, rowIdx, mapCol, "IsSummary")

    Select Case VarType(v)
        Case vbBoolean
            Core_IsSummaryRow = CBool(v)
            Exit Function

        Case vbString
            txt = Trim$(LCase$(CStr(v)))
            Core_IsSummaryRow = (txt = "true" Or txt = "1" Or txt = "yes" Or txt = "summary")
            Exit Function

        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
            Core_IsSummaryRow = (CDbl(v) <> 0)
            Exit Function
    End Select

    Core_IsSummaryRow = False

End Function


'------------------------------------------------------------------------------
' FR: Retourne la valeur Calc Inclusive Duration sans modifier les donnees d'entree.
' EN: Returns the Calc Inclusive Duration value without mutating input data.
'------------------------------------------------------------------------------

Public Function Core_CalcInclusiveDuration( _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    Optional ByVal calendarType As String = "") As Variant

    Dim perfScope As clsPerfScope

    Set perfScope = Profiler_BeginScope("Core_CalcInclusiveDuration", "Calendar Calculation")

    If Not HasValue(startVal) Then
        Core_CalcInclusiveDuration = Empty
        Exit Function
    End If

    If Not HasValue(finishVal) Then
        Core_CalcInclusiveDuration = Empty
        Exit Function
    End If

    Core_CalcInclusiveDuration = DateDiffWorkingDays(startVal, finishVal, calendarType)

End Function

'------------------------------------------------------------------------------
' FR: Active ou initialise Core Set Calc Triplet dans l'etat runtime du composant.
' EN: Activates or initializes Core Set Calc Triplet in the component runtime state.
'------------------------------------------------------------------------------

Public Sub Core_SetCalcTriplet( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal startVal As Variant, _
    ByVal finishVal As Variant, _
    Optional ByVal calendarType As String = "")

    Dim durVal As Variant

    Core_SetVal dataArr, rowIdx, mapCol, "Calculated Start", startVal
    Core_SetVal dataArr, rowIdx, mapCol, "Calculated Finish", finishVal

    durVal = Core_CalcInclusiveDuration(startVal, finishVal, calendarType)
    Core_SetVal dataArr, rowIdx, mapCol, "Calculated Duration", durVal

End Sub

'------------------------------------------------------------------------------
' FR: Active ou initialise Core Set Error Flag Row dans l'etat runtime du composant.
' EN: Activates or initializes Core Set Error Flag Row in the component runtime state.
'------------------------------------------------------------------------------

Public Sub Core_SetErrorFlag_Row( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal isError As Boolean)

    If isError Then
        Core_SetVal dataArr, rowIdx, mapCol, "Error flag", "ERROR"
    Else
        Core_SetVal dataArr, rowIdx, mapCol, "Error flag", ""
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Ajoute la map Error Message Row a la structure cible fournie par l'appelant.
' EN: Adds the Error Message Row map to the target structure supplied by the caller.
'------------------------------------------------------------------------------

Public Sub Core_AddErrorMessage_Row( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object, _
    ByVal msgText As String, _
    Optional ByVal markError As Boolean = True)

    Dim currentMsg As String

    If Trim$(msgText) = "" Then Exit Sub

    currentMsg = Trim$(CStr(Core_GetVal(dataArr, rowIdx, mapCol, "ErrorMsg")))

    If currentMsg = "" Then
        Core_SetVal dataArr, rowIdx, mapCol, "ErrorMsg", msgText
    ElseIf InStr(1, currentMsg, msgText, vbTextCompare) = 0 Then
        Core_SetVal dataArr, rowIdx, mapCol, "ErrorMsg", currentMsg & " / " & msgText
    End If

    If markError Then
        Core_SetErrorFlag_Row dataArr, rowIdx, mapCol, True
    End If

End Sub

'------------------------------------------------------------------------------
' FR: Reinitialise Core Clear Error State Row dans le perimetre possede par le composant.
' EN: Resets Core Clear Error State Row within the state owned by the component.
'------------------------------------------------------------------------------

Public Sub Core_ClearErrorState_Row( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object)

    Core_SetErrorFlag_Row dataArr, rowIdx, mapCol, False
    Core_SetVal dataArr, rowIdx, mapCol, "ErrorMsg", ""

End Sub

'------------------------------------------------------------------------------
' FR: Retourne la valeur Max Date If Both sans modifier les donnees d'entree.
' EN: Returns the Max Date If Both value without mutating input data.
'------------------------------------------------------------------------------

Public Function Core_MaxDateIfBoth( _
    ByVal leftVal As Variant, _
    ByVal rightVal As Variant) As Variant

    If HasValue(leftVal) And HasValue(rightVal) Then
        Core_MaxDateIfBoth = maxDate(leftVal, rightVal)
    ElseIf HasValue(leftVal) Then
        Core_MaxDateIfBoth = leftVal
    ElseIf HasValue(rightVal) Then
        Core_MaxDateIfBoth = rightVal
    Else
        Core_MaxDateIfBoth = Empty
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Min Date If Both sans modifier les donnees d'entree.
' EN: Returns the Min Date If Both value without mutating input data.
'------------------------------------------------------------------------------

Public Function Core_MinDateIfBoth( _
    ByVal leftVal As Variant, _
    ByVal rightVal As Variant) As Variant

    If HasValue(leftVal) And HasValue(rightVal) Then
        If CDbl(leftVal) <= CDbl(rightVal) Then
            Core_MinDateIfBoth = leftVal
        Else
            Core_MinDateIfBoth = rightVal
        End If
    ElseIf HasValue(leftVal) Then
        Core_MinDateIfBoth = leftVal
    ElseIf HasValue(rightVal) Then
        Core_MinDateIfBoth = rightVal
    Else
        Core_MinDateIfBoth = Empty
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la date de debut source applicable a la tache sans exposer de mutateur sur l'etat source.
' EN: Returns the source start date applicable to the task without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function Core_GetSourceStart( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object) As Variant

    If Core_HasVal(dataArr, rowIdx, mapCol, "Actual Start") Then
        Core_GetSourceStart = Core_GetVal(dataArr, rowIdx, mapCol, "Actual Start")
    ElseIf Core_HasVal(dataArr, rowIdx, mapCol, "Forecast Start") Then
        Core_GetSourceStart = Core_GetVal(dataArr, rowIdx, mapCol, "Forecast Start")
    ElseIf Core_HasVal(dataArr, rowIdx, mapCol, "Baseline Start") Then
        Core_GetSourceStart = Core_GetVal(dataArr, rowIdx, mapCol, "Baseline Start")
    Else
        Core_GetSourceStart = Empty
    End If

End Function

'------------------------------------------------------------------------------
' FR: Retourne la date de fin source applicable a la tache sans exposer de mutateur sur l'etat source.
' EN: Returns the source finish date applicable to the task without exposing a mutator for source state.
'------------------------------------------------------------------------------

Public Function Core_GetSourceFinish( _
    ByRef dataArr As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCol As Object) As Variant

    If Core_HasVal(dataArr, rowIdx, mapCol, "Actual Finish") Then
        Core_GetSourceFinish = Core_GetVal(dataArr, rowIdx, mapCol, "Actual Finish")
    ElseIf Core_HasVal(dataArr, rowIdx, mapCol, "Forecast Finish") Then
        Core_GetSourceFinish = Core_GetVal(dataArr, rowIdx, mapCol, "Forecast Finish")
    Else
        Core_GetSourceFinish = Empty
    End If

End Function
'------------------------------------------------------------------------------
' FR: Valide Core Require Columns et applique la politique d'erreur definie par le composant.
' EN: Validates Core Require Columns and applies the component's defined failure policy.
'------------------------------------------------------------------------------

Public Sub Core_RequireColumns( _
    ByVal mapCol As Object, _
    ByVal requiredCols As Variant, _
    ByVal contextName As String)

    Dim c As Variant
    Dim missing As String

    missing = ""

    For Each c In requiredCols
        If Not mapCol.Exists(CStr(c)) Then
            If missing <> "" Then missing = missing & vbCrLf
            missing = missing & "- " & CStr(c)
        End If
    Next c

    If missing <> "" Then
        Err.Raise vbObjectError + 701, "Core_RequireColumns", _
            "Missing required columns in " & contextName & ":" & vbCrLf & missing
    End If

End Sub

