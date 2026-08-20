Attribute VB_Name = "mod_IncrementalSignature"
Option Explicit

'===============================================================================
' MODULE : mod_IncrementalSignature
' DOMAINE / DOMAIN : Shared Infrastructure
'
' FR
' Possede les 17 champs, leur ordre, normalisation et serialisation de la signature incrementale.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Owns the 17 fields, ordering, normalization and serialization of the incremental signature.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : IncrementalSignature_RequiredColumns, IncrementalSignature_BuildRow
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Retourne les colonnes CALC du contrat de signature incrementale, dans leur ordre canonique.
' EN: Returns the CALC columns in the incremental signature contract, in canonical order.
'------------------------------------------------------------------------------
Public Function IncrementalSignature_RequiredColumns() As Variant

    IncrementalSignature_RequiredColumns = Array( _
        "ID", _
        "ParentID", _
        "IsSummary", _
        "Predecessors WBS", _
        "Cal", _
        "Baseline Start", _
        "Baseline Duration", _
        "Actual Start", _
        "Actual Finish", _
        "Forecast Start", _
        "Forecast Finish", _
        "Deadline", _
        "Constraint Active", _
        "Start Constraint Type", _
        "Start Constraint Date", _
        "Finish Constraint Type", _
        "Finish Constraint Date" _
    )

End Function

'------------------------------------------------------------------------------
' FR: Construit la signature incrementale canonique d'une ligne CALC sans lire ni modifier Excel.
' EN: Builds the canonical incremental signature for a CALC row without reading or writing Excel.
'------------------------------------------------------------------------------
Public Function IncrementalSignature_BuildRow( _
    ByRef arrCalc As Variant, _
    ByVal rowIdx As Long, _
    ByVal mapCalc As Object) As String

    Dim s As String

    s = ""
    s = s & "|ID=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("ID")), "TEXT")
    s = s & "|ParentID=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("ParentID")), "TEXT")
    s = s & "|IsSummary=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("IsSummary")), "BOOLEAN")
    s = s & "|PredWBS=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Predecessors WBS")), "PREDWBS")
    s = s & "|Cal=" & IncrementalSignature_NormalizeValue(NormalizeCalendarType(arrCalc(rowIdx, mapCalc("Cal"))), "TEXT")

    s = s & "|BS=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Baseline Start")), "DATE")
    s = s & "|BD=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Baseline Duration")), "NUMBER")
    s = s & "|AS=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Actual Start")), "DATE")
    s = s & "|AF=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Actual Finish")), "DATE")
    s = s & "|FS=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Forecast Start")), "DATE")
    s = s & "|FF=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Forecast Finish")), "DATE")
    s = s & "|DL=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Deadline")), "DATE")
    s = s & "|CActive=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Constraint Active")), "BOOLEAN")
    s = s & "|SCType=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Start Constraint Type")), "TEXT")
    s = s & "|SCDate=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Start Constraint Date")), "DATE")
    s = s & "|FCType=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Finish Constraint Type")), "TEXT")
    s = s & "|FCDate=" & IncrementalSignature_NormalizeValue(arrCalc(rowIdx, mapCalc("Finish Constraint Date")), "DATE")

    IncrementalSignature_BuildRow = s

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Normalize Value selon le contrat canonique du composant.
' EN: Normalizes or formats Normalize Value according to the component contract.
'------------------------------------------------------------------------------

Private Function IncrementalSignature_NormalizeValue( _
    ByVal v As Variant, _
    Optional ByVal valueKind As String = "TEXT") As String

    Select Case UCase$(Trim$(valueKind))
        Case "BOOLEAN"
            IncrementalSignature_NormalizeValue = IncrementalSignature_NormalizeBoolean(v)
        Case "DATE"
            IncrementalSignature_NormalizeValue = IncrementalSignature_NormalizeDate(v)
        Case "NUMBER"
            IncrementalSignature_NormalizeValue = IncrementalSignature_NormalizeNumber(v)
        Case "PREDWBS"
            IncrementalSignature_NormalizeValue = IncrementalSignature_NormalizeText(v, True)
        Case Else
            IncrementalSignature_NormalizeValue = IncrementalSignature_NormalizeText(v, False)
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Normalize Boolean selon le contrat canonique du composant.
' EN: Normalizes or formats Normalize Boolean according to the component contract.
'------------------------------------------------------------------------------

Private Function IncrementalSignature_NormalizeBoolean(ByVal v As Variant) As String

    Dim s As String

    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function

    If VarType(v) = vbBoolean Then
        If CBool(v) Then
            IncrementalSignature_NormalizeBoolean = "TRUE"
        Else
            IncrementalSignature_NormalizeBoolean = "FALSE"
        End If
        Exit Function
    End If

    s = UCase$(Trim$(CStr(v)))
    Select Case s
        Case "TRUE", "VRAI", "-1", "1", "YES", "OUI"
            IncrementalSignature_NormalizeBoolean = "TRUE"
        Case "FALSE", "FAUX", "0", "NO", "NON", ""
            IncrementalSignature_NormalizeBoolean = "FALSE"
        Case Else
            IncrementalSignature_NormalizeBoolean = s
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Normalize Date selon le contrat canonique du composant.
' EN: Normalizes or formats Normalize Date according to the component contract.
'------------------------------------------------------------------------------

Private Function IncrementalSignature_NormalizeDate(ByVal v As Variant) As String

    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Trim$(CStr(v)) = "" Then Exit Function

    On Error GoTo Fallback
    If IsDate(v) Then
        IncrementalSignature_NormalizeDate = Format$(CDate(v), "yyyymmdd")
        Exit Function
    End If
    If IsNumeric(v) Then
        If CDbl(v) > 0 Then
            IncrementalSignature_NormalizeDate = Format$(CDate(CDbl(v)), "yyyymmdd")
            Exit Function
        End If
    End If

Fallback:
    IncrementalSignature_NormalizeDate = Trim$(CStr(v))

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Normalize Number selon le contrat canonique du composant.
' EN: Normalizes or formats Normalize Number according to the component contract.
'------------------------------------------------------------------------------

Private Function IncrementalSignature_NormalizeNumber(ByVal v As Variant) As String

    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If Trim$(CStr(v)) = "" Then Exit Function

    If IsNumeric(v) Then
        IncrementalSignature_NormalizeNumber = Format$(CDbl(v), "0.############")
    Else
        IncrementalSignature_NormalizeNumber = Trim$(CStr(v))
    End If

End Function

'------------------------------------------------------------------------------
' FR: Normalise ou formate Normalize Text selon le contrat canonique du composant.
' EN: Normalizes or formats Normalize Text according to the component contract.
'------------------------------------------------------------------------------

Private Function IncrementalSignature_NormalizeText( _
    ByVal v As Variant, _
    Optional ByVal removeSpaces As Boolean = False) As String

    Dim s As String

    If IsEmpty(v) Then Exit Function
    If IsNull(v) Then Exit Function

    s = Trim$(CStr(v))
    If removeSpaces Then s = Replace$(s, " ", "")
    IncrementalSignature_NormalizeText = s

End Function
