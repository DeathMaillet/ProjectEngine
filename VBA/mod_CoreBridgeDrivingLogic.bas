Attribute VB_Name = "mod_CoreBridgeDrivingLogic"
Option Explicit

'===============================================================================
' MODULE : mod_CoreBridgeDrivingLogic
' DOMAINE / DOMAIN : Core Bridge
'
' FR
' Construit et pousse la Driving Logic issue des dependances calculees.
' Ne doit pas contourner les contrats publics des autres domaines.
'
' EN
' Builds and pushes Driving Logic derived from calculated dependencies.
' Must not bypass public contracts owned by other domains.
'
' CONTRATS / CONTRACTS : WriteCoreDrivingLogicToCalc, WriteCoreDrivingLogicToCalc_Partial
' CALLBACKS EXTERNES / EXTERNAL CALLBACKS : Aucun / None
'===============================================================================


'------------------------------------------------------------------------------
' FR: Ecrit Core Driving Logic To Calc vers le stockage cible.
' EN: Writes Core Driving Logic To Calc to the target storage.
'------------------------------------------------------------------------------
Public Sub WriteCoreDrivingLogicToCalc( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByRef dataArr As Variant)

    Dim rowById As Object
    Dim parentIds As Object
    Dim rowCount As Long
    Dim outDriving() As Variant
    Dim r As Long
    Dim idVal As String
    Dim baselineStart As Variant
    Dim actualStart As Variant
    Dim actualFinish As Variant
    Dim forecastStart As Variant
    Dim forecastFinish As Variant
    Dim calcStart As Variant

    Set rowById = Core_BuildRowById(dataArr, mapCalc)
    Set parentIds = Core_BuildParentIds(dataArr, mapCalc, rowById)

    rowCount = UBound(dataArr, 1)
    ReDim outDriving(1 To rowCount, 1 To 1)

    For r = 1 To rowCount

        idVal = Trim$(CStr(dataArr(r, mapCalc("ID"))))
        outDriving(r, 1) = ""

        If idVal <> "" Then

            If parentIds.Exists(idVal) Then

                outDriving(r, 1) = "SUMMARY"

            ElseIf TaskTypeRules_IsLevelOfEffortRow(dataArr, mapCalc, r) Then

                outDriving(r, 1) = "LOE"

            Else

                baselineStart = dataArr(r, mapCalc("Baseline Start"))
                actualStart = dataArr(r, mapCalc("Actual Start"))
                actualFinish = dataArr(r, mapCalc("Actual Finish"))
                forecastStart = dataArr(r, mapCalc("Forecast Start"))
                forecastFinish = dataArr(r, mapCalc("Forecast Finish"))
                calcStart = dataArr(r, mapCalc("Calculated Start"))

                If CalcBridge_IsDrivenByActiveConstraint(dataArr, mapCalc, r) Then
                    outDriving(r, 1) = "CONSTRAINT"
                ElseIf HasValue(actualStart) Or HasValue(actualFinish) Then
                    outDriving(r, 1) = "ACTUAL"
                ElseIf HasValue(forecastStart) Or HasValue(forecastFinish) Then
                    outDriving(r, 1) = "FORECAST"
                ElseIf HasValue(calcStart) And HasValue(baselineStart) Then
                    If CDbl(calcStart) > CDbl(baselineStart) Then
                        outDriving(r, 1) = "DEPENDENCY"
                    Else
                        outDriving(r, 1) = "BASELINE"
                    End If
                Else
                    outDriving(r, 1) = "BASELINE"
                End If

            End If

        End If

    Next r

    tblCalc.ListColumns("Driving Logic").DataBodyRange.value = outDriving

End Sub

'------------------------------------------------------------------------------
' FR: Indique si la map Driven By Active Constraint satisfait la condition attendue, sans modifier les donnees source.
' EN: Returns whether the Driven By Active Constraint map satisfies the expected condition without mutating source data.
'------------------------------------------------------------------------------

Private Function CalcBridge_IsDrivenByActiveConstraint( _
    ByRef dataArr As Variant, _
    ByVal mapCalc As Object, _
    ByVal rowIdx As Long) As Boolean

    Dim activeVal As String
    Dim startType As String
    Dim finishType As String
    Dim startDate As Variant
    Dim finishDate As Variant
    Dim calcStart As Variant
    Dim calcFinish As Variant

    CalcBridge_IsDrivenByActiveConstraint = False

    If mapCalc Is Nothing Then Exit Function
    If Not mapCalc.Exists("Constraint Active") Then Exit Function
    If Not mapCalc.Exists("Start Constraint Type") Then Exit Function
    If Not mapCalc.Exists("Start Constraint Date") Then Exit Function
    If Not mapCalc.Exists("Finish Constraint Type") Then Exit Function
    If Not mapCalc.Exists("Finish Constraint Date") Then Exit Function
    If Not mapCalc.Exists("Calculated Start") Then Exit Function
    If Not mapCalc.Exists("Calculated Finish") Then Exit Function

    activeVal = UCase$(Trim$(CStr(dataArr(rowIdx, mapCalc("Constraint Active")))))
    If activeVal <> "YES" Then Exit Function

    startType = UCase$(Trim$(CStr(dataArr(rowIdx, mapCalc("Start Constraint Type")))))
    finishType = UCase$(Trim$(CStr(dataArr(rowIdx, mapCalc("Finish Constraint Type")))))
    startDate = dataArr(rowIdx, mapCalc("Start Constraint Date"))
    finishDate = dataArr(rowIdx, mapCalc("Finish Constraint Date"))
    calcStart = dataArr(rowIdx, mapCalc("Calculated Start"))
    calcFinish = dataArr(rowIdx, mapCalc("Calculated Finish"))

    If startType = "MUST START ON" Then
        CalcBridge_IsDrivenByActiveConstraint = True
        Exit Function
    End If

    If finishType = "MUST FINISH ON" Then
        CalcBridge_IsDrivenByActiveConstraint = True
        Exit Function
    End If

    Select Case startType
        Case "START NO EARLIER THAN", "START NO LATER THAN"
            If CalcBridge_DatesMatch(calcStart, startDate) Then
                CalcBridge_IsDrivenByActiveConstraint = True
                Exit Function
            End If
    End Select

    Select Case finishType
        Case "FINISH NO EARLIER THAN", "FINISH NO LATER THAN"
            If CalcBridge_DatesMatch(calcFinish, finishDate) Then
                CalcBridge_IsDrivenByActiveConstraint = True
                Exit Function
            End If
    End Select

End Function

'------------------------------------------------------------------------------
' FR: Retourne la valeur Dates Match sans modifier les donnees d'entree.
' EN: Returns the Dates Match value without mutating input data.
'------------------------------------------------------------------------------

Private Function CalcBridge_DatesMatch(ByVal leftVal As Variant, ByVal rightVal As Variant) As Boolean

    CalcBridge_DatesMatch = False

    If Not HasValue(leftVal) Then Exit Function
    If Not HasValue(rightVal) Then Exit Function
    If Not IsDate(leftVal) Then Exit Function
    If Not IsDate(rightVal) Then Exit Function

    CalcBridge_DatesMatch = (CLng(CDbl(CDate(leftVal))) = CLng(CDbl(CDate(rightVal))))

End Function


'------------------------------------------------------------------------------
' FR: Ecrit Core Driving Logic To Calc Partial vers le stockage cible.
' EN: Writes Core Driving Logic To Calc Partial to the target storage.
'------------------------------------------------------------------------------
Public Sub WriteCoreDrivingLogicToCalc_Partial( _
    ByVal tblCalc As ListObject, _
    ByVal mapCalc As Object, _
    ByRef dataArr As Variant, _
    ByVal impactedIds As Object)

    Dim perfScope As clsPerfScope

    Dim rowById As Object
    Dim parentIds As Object
    Dim idVal As Variant
    Dim rowIdx As Long

    Dim baselineStart As Variant
    Dim actualStart As Variant
    Dim actualFinish As Variant
    Dim forecastStart As Variant
    Dim forecastFinish As Variant
    Dim calcStart As Variant
    Dim drivingLogic As String

    Set perfScope = Profiler_BeginScope("WriteCoreDrivingLogicToCalc_Partial", "Excel Cell Write")

    If tblCalc Is Nothing Then Exit Sub
    If tblCalc.DataBodyRange Is Nothing Then Exit Sub
    If impactedIds Is Nothing Then Exit Sub
    If impactedIds.Count = 0 Then Exit Sub

    If Not mapCalc.Exists("ID") Then Exit Sub
    If Not mapCalc.Exists("Driving Logic") Then Exit Sub
    If Not mapCalc.Exists("Baseline Start") Then Exit Sub
    If Not mapCalc.Exists("Actual Start") Then Exit Sub
    If Not mapCalc.Exists("Actual Finish") Then Exit Sub
    If Not mapCalc.Exists("Forecast Start") Then Exit Sub
    If Not mapCalc.Exists("Forecast Finish") Then Exit Sub
    If Not mapCalc.Exists("Calculated Start") Then Exit Sub

    Set rowById = Core_BuildRowById(dataArr, mapCalc)
    Set parentIds = Core_BuildParentIds(dataArr, mapCalc, rowById)

    For Each idVal In impactedIds.Keys

        If rowById.Exists(CStr(idVal)) Then

            rowIdx = CLng(rowById(CStr(idVal)))
            drivingLogic = ""

            If parentIds.Exists(CStr(idVal)) Then

                drivingLogic = "SUMMARY"

            ElseIf TaskTypeRules_IsLevelOfEffortRow(dataArr, mapCalc, rowIdx) Then

                drivingLogic = "LOE"

            Else

                baselineStart = dataArr(rowIdx, mapCalc("Baseline Start"))
                actualStart = dataArr(rowIdx, mapCalc("Actual Start"))
                actualFinish = dataArr(rowIdx, mapCalc("Actual Finish"))
                forecastStart = dataArr(rowIdx, mapCalc("Forecast Start"))
                forecastFinish = dataArr(rowIdx, mapCalc("Forecast Finish"))
                calcStart = dataArr(rowIdx, mapCalc("Calculated Start"))

                If CalcBridge_IsDrivenByActiveConstraint(dataArr, mapCalc, rowIdx) Then
                    drivingLogic = "CONSTRAINT"
                ElseIf HasValue(actualStart) Or HasValue(actualFinish) Then
                    drivingLogic = "ACTUAL"
                ElseIf HasValue(forecastStart) Or HasValue(forecastFinish) Then
                    drivingLogic = "FORECAST"
                ElseIf HasValue(calcStart) And HasValue(baselineStart) Then
                    If CDbl(calcStart) > CDbl(baselineStart) Then
                        drivingLogic = "DEPENDENCY"
                    Else
                        drivingLogic = "BASELINE"
                    End If
                Else
                    drivingLogic = "BASELINE"
                End If

            End If

            tblCalc.DataBodyRange.Cells(rowIdx, mapCalc("Driving Logic")).value = drivingLogic

        End If

    Next idVal

End Sub


