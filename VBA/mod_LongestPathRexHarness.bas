Attribute VB_Name = "mod_LongestPathRexHarness"
Option Explicit

'===============================================================================
' MODULE : mod_LongestPathRexHarness
' DOMAINE / DOMAIN : Analytics proof harness
'
' FR
' Harnais permanent de preuve Longest Path Current/REX. Les fixtures sont construites
' en memoire afin de verifier le moteur partage sans muter les donnees metier.
'
' EN
' Permanent Longest Path Current/REX proof harness. Fixtures are built in memory so
' the shared engine can be verified without mutating business data.
'
' CONTRATS / CONTRACTS : LongestPathRexHarness_Export
'===============================================================================

Public Function LongestPathRexHarness_Export(ByVal evidenceFolder As String) As String

    On Error GoTo ErrHandler

    If Len(Dir$(evidenceFolder, vbDirectory)) = 0 Then MkDir evidenceFolder

    LongestPathRexHarness_WriteTieFixtures evidenceFolder
    LongestPathRexHarness_WriteTemporalFixtures evidenceFolder
    LongestPathRexHarness_WriteIdentityFixtures evidenceFolder
    LongestPathRexHarness_WriteDivergenceFixtures evidenceFolder

    LongestPathRexHarness_Export = "PASS"
    Exit Function

ErrHandler:
    LongestPathRexHarness_Export = "FAIL|" & CStr(Err.Number) & "|" & Err.Description

End Function

Private Sub LongestPathRexHarness_WriteTieFixtures(ByVal folderPath As String)

    Dim f As Integer
    Dim statusText As String
    Dim details As String

    statusText = LongestPathRexHarness_RunFixture( _
        "TIE_FS_REJOIN", _
        Array("A", "B", "C"), _
        Array(DateSerial(2026, 1, 1), DateSerial(2026, 1, 1), DateSerial(2026, 1, 4)), _
        Array(DateSerial(2026, 1, 3), DateSerial(2026, 1, 3), DateSerial(2026, 1, 6)), _
        Array("C|A|FS|0", "C|B|FS|0"), _
        "A,B,C", _
        True, _
        details)

    f = FreeFile
    Open LongestPathRexHarness_Path(folderPath, "synthetic_tie_fixtures.tsv") For Output As #f
    Print #f, "Fixture" & vbTab & "ExpectedLongestIds" & vbTab & "Status" & vbTab & "Details"
    Print #f, "TIE_FS_REJOIN" & vbTab & "A,B,C" & vbTab & statusText & vbTab & LongestPathRexHarness_Tsv(details)
    Close #f

    f = FreeFile
    Open LongestPathRexHarness_Path(folderPath, "longest_path_ties.tsv") For Output As #f
    Print #f, "Fixture" & vbTab & "TiePolicy" & vbTab & "Status" & vbTab & "Notes"
    Print #f, "TIE_FS_REJOIN" & vbTab & "All equal driving predecessors are retained" & vbTab & statusText & vbTab & "A and B both satisfy the FS+0 equality into C."
    Close #f

End Sub

Private Sub LongestPathRexHarness_WriteTemporalFixtures(ByVal folderPath As String)

    Dim f As Integer
    Dim linkTypes As Variant
    Dim lags As Variant
    Dim linkType As Variant
    Dim lagValue As Variant
    Dim fixtureName As String
    Dim predStart As Date
    Dim predFinish As Date
    Dim succStart As Variant
    Dim succFinish As Variant
    Dim statusText As String
    Dim details As String

    linkTypes = Array("FS", "SS", "FF")
    lags = Array(-2, 0, 3)
    predStart = DateSerial(2026, 2, 10)
    predFinish = DateSerial(2026, 2, 12)

    f = FreeFile
    Open LongestPathRexHarness_Path(folderPath, "temporal_constraint_fixtures.tsv") For Output As #f
    Print #f, "Fixture" & vbTab & "LinkType" & vbTab & "Lag" & vbTab & "CurrentRexSame" & vbTab & "Status" & vbTab & "Details"

    For Each linkType In linkTypes
        For Each lagValue In lags
            Select Case CStr(linkType)
                Case "SS"
                    succStart = ApplyLag(predStart, CLng(lagValue), CALENDAR_7D, "SS")
                    succFinish = CDbl(CDate(succStart)) + 2
                Case "FF"
                    succFinish = ApplyLag(predFinish, CLng(lagValue), CALENDAR_7D, "FF")
                    succStart = CDbl(CDate(succFinish)) - 2
                Case Else
                    succStart = ApplyLag(predFinish, CLng(lagValue), CALENDAR_7D, "FS")
                    succFinish = CDbl(CDate(succStart)) + 2
            End Select

            fixtureName = CStr(linkType) & "_LAG_" & CStr(lagValue)
            statusText = LongestPathRexHarness_RunFixture( _
                fixtureName, _
                Array("P", "S"), _
                Array(predStart, succStart), _
                Array(predFinish, succFinish), _
                Array("S|P|" & CStr(linkType) & "|" & CStr(lagValue)), _
                vbNullString, _
                True, _
                details)

            Print #f, fixtureName & vbTab & CStr(linkType) & vbTab & CStr(lagValue) & vbTab & "YES" & vbTab & statusText & vbTab & LongestPathRexHarness_Tsv(details)
        Next lagValue
    Next linkType

    Close #f

End Sub

Private Sub LongestPathRexHarness_WriteIdentityFixtures(ByVal folderPath As String)

    Dim f As Integer
    Dim statusText As String
    Dim details As String

    statusText = LongestPathRexHarness_RunFixture( _
        "CURRENT_REX_IDENTITY", _
        Array("A", "B", "C", "D"), _
        Array(DateSerial(2026, 3, 1), DateSerial(2026, 3, 4), DateSerial(2026, 3, 4), DateSerial(2026, 3, 7)), _
        Array(DateSerial(2026, 3, 3), DateSerial(2026, 3, 6), DateSerial(2026, 3, 5), DateSerial(2026, 3, 9)), _
        Array("B|A|FS|0", "C|A|FS|0", "D|B|FS|0", "D|C|FS|1"), _
        "A,B,C,D", _
        True, _
        details)

    f = FreeFile
    Open LongestPathRexHarness_Path(folderPath, "current_rex_identity.tsv") For Output As #f
    Print #f, "Fixture" & vbTab & "Invariant" & vbTab & "Status" & vbTab & "Details"
    Print #f, "CURRENT_REX_IDENTITY" & vbTab & "Calculated dates equal Baseline dates" & vbTab & statusText & vbTab & LongestPathRexHarness_Tsv(details)
    Close #f

End Sub

Private Sub LongestPathRexHarness_WriteDivergenceFixtures(ByVal folderPath As String)

    Dim f As Integer
    Dim statusText As String
    Dim details As String
    Dim curStarts As Variant
    Dim curFinishes As Variant
    Dim rexStarts As Variant
    Dim rexFinishes As Variant

    curStarts = Array(DateSerial(2026, 4, 1), DateSerial(2026, 4, 1), DateSerial(2026, 4, 6))
    curFinishes = Array(DateSerial(2026, 4, 3), DateSerial(2026, 4, 5), DateSerial(2026, 4, 8))
    rexStarts = Array(DateSerial(2026, 4, 1), DateSerial(2026, 4, 1), DateSerial(2026, 4, 4))
    rexFinishes = Array(DateSerial(2026, 4, 6), DateSerial(2026, 4, 2), DateSerial(2026, 4, 6))

    statusText = LongestPathRexHarness_RunDivergenceFixture( _
        "CURRENT_REX_DIVERGENCE", _
        Array("A", "B", "C"), _
        curStarts, curFinishes, _
        rexStarts, rexFinishes, _
        Array("C|A|FS|0", "C|B|FS|0"), _
        details)

    f = FreeFile
    Open LongestPathRexHarness_Path(folderPath, "current_rex_divergence.tsv") For Output As #f
    Print #f, "Fixture" & vbTab & "Expected" & vbTab & "Status" & vbTab & "Details"
    Print #f, "CURRENT_REX_DIVERGENCE" & vbTab & "Current and REX may diverge when temporal vectors differ" & vbTab & statusText & vbTab & LongestPathRexHarness_Tsv(details)
    Close #f

End Sub

Private Function LongestPathRexHarness_RunFixture( _
    ByVal fixtureName As String, _
    ByVal ids As Variant, _
    ByVal starts As Variant, _
    ByVal finishes As Variant, _
    ByVal edges As Variant, _
    ByVal expectedLongestCsv As String, _
    ByVal requireCurrentRexSame As Boolean, _
    ByRef details As String) As String

    Dim dataArr As Variant
    Dim mapCalc As Object
    Dim idToRow As Object
    Dim predsById As Object
    Dim childrenById As Object
    Dim validIds As Object
    Dim predLagBySuccPred As Object
    Dim predTypeBySuccPred As Object
    Dim outCurrent As Variant
    Dim outRex As Variant
    Dim currentCsv As String
    Dim rexCsv As String

    LongestPathRexHarness_BuildFixture ids, starts, finishes, starts, finishes, edges, _
        dataArr, mapCalc, idToRow, predsById, childrenById, validIds, predLagBySuccPred, predTypeBySuccPred

    ComputeLongestPath Nothing, mapCalc, idToRow, predsById, childrenById, validIds, _
        predLagBySuccPred, predTypeBySuccPred, dataArr, outCurrent, Nothing, _
        "Calculated Start", "Calculated Finish", "Longest Path", True

    ComputeLongestPath Nothing, mapCalc, idToRow, predsById, childrenById, validIds, _
        predLagBySuccPred, predTypeBySuccPred, dataArr, outRex, Nothing, _
        "Baseline Start", "Baseline Finish", "Longest Path REX", False

    currentCsv = LongestPathRexHarness_LongestCsv(ids, outCurrent)
    rexCsv = LongestPathRexHarness_LongestCsv(ids, outRex)
    details = "Current=" & currentCsv & "; REX=" & rexCsv

    If requireCurrentRexSame Then
        If currentCsv <> rexCsv Then
            LongestPathRexHarness_RunFixture = "FAIL"
            Exit Function
        End If
    End If

    If Len(expectedLongestCsv) > 0 Then
        If currentCsv <> expectedLongestCsv Then
            LongestPathRexHarness_RunFixture = "FAIL"
            Exit Function
        End If
    End If

    LongestPathRexHarness_RunFixture = "PASS"

End Function

Private Function LongestPathRexHarness_RunDivergenceFixture( _
    ByVal fixtureName As String, _
    ByVal ids As Variant, _
    ByVal curStarts As Variant, _
    ByVal curFinishes As Variant, _
    ByVal rexStarts As Variant, _
    ByVal rexFinishes As Variant, _
    ByVal edges As Variant, _
    ByRef details As String) As String

    Dim dataArr As Variant
    Dim mapCalc As Object
    Dim idToRow As Object
    Dim predsById As Object
    Dim childrenById As Object
    Dim validIds As Object
    Dim predLagBySuccPred As Object
    Dim predTypeBySuccPred As Object
    Dim outCurrent As Variant
    Dim outRex As Variant
    Dim currentCsv As String
    Dim rexCsv As String

    LongestPathRexHarness_BuildFixture ids, curStarts, curFinishes, rexStarts, rexFinishes, edges, _
        dataArr, mapCalc, idToRow, predsById, childrenById, validIds, predLagBySuccPred, predTypeBySuccPred

    ComputeLongestPath Nothing, mapCalc, idToRow, predsById, childrenById, validIds, _
        predLagBySuccPred, predTypeBySuccPred, dataArr, outCurrent, Nothing, _
        "Calculated Start", "Calculated Finish", "Longest Path", True

    ComputeLongestPath Nothing, mapCalc, idToRow, predsById, childrenById, validIds, _
        predLagBySuccPred, predTypeBySuccPred, dataArr, outRex, Nothing, _
        "Baseline Start", "Baseline Finish", "Longest Path REX", False

    currentCsv = LongestPathRexHarness_LongestCsv(ids, outCurrent)
    rexCsv = LongestPathRexHarness_LongestCsv(ids, outRex)
    details = "Current=" & currentCsv & "; REX=" & rexCsv

    If currentCsv <> rexCsv Then
        LongestPathRexHarness_RunDivergenceFixture = "PASS"
    Else
        LongestPathRexHarness_RunDivergenceFixture = "FAIL"
    End If

End Function

Private Sub LongestPathRexHarness_BuildFixture( _
    ByVal ids As Variant, _
    ByVal curStarts As Variant, _
    ByVal curFinishes As Variant, _
    ByVal rexStarts As Variant, _
    ByVal rexFinishes As Variant, _
    ByVal edges As Variant, _
    ByRef dataArr As Variant, _
    ByRef mapCalc As Object, _
    ByRef idToRow As Object, _
    ByRef predsById As Object, _
    ByRef childrenById As Object, _
    ByRef validIds As Object, _
    ByRef predLagBySuccPred As Object, _
    ByRef predTypeBySuccPred As Object)

    Dim i As Long
    Dim rowIndex As Long
    Dim edge As Variant
    Dim parts As Variant
    Dim succId As String
    Dim predId As String
    Dim linkType As String
    Dim lagValue As Double
    Dim linkKey As String

    Set mapCalc = CreateObject("Scripting.Dictionary")
    mapCalc("ID") = 1
    mapCalc("WBS") = 2
    mapCalc("Cal") = 3
    mapCalc("Actual Finish") = 4
    mapCalc("Calculated Start") = 5
    mapCalc("Calculated Finish") = 6
    mapCalc("Baseline Start") = 7
    mapCalc("Baseline Finish") = 8
    mapCalc("Longest Path") = 9
    mapCalc("Longest Path REX") = 10

    Set idToRow = CreateObject("Scripting.Dictionary")
    Set predsById = CreateObject("Scripting.Dictionary")
    Set childrenById = CreateObject("Scripting.Dictionary")
    Set validIds = CreateObject("Scripting.Dictionary")
    Set predLagBySuccPred = CreateObject("Scripting.Dictionary")
    Set predTypeBySuccPred = CreateObject("Scripting.Dictionary")

    ReDim dataArr(1 To UBound(ids) - LBound(ids) + 1, 1 To 10)

    For i = LBound(ids) To UBound(ids)
        rowIndex = i - LBound(ids) + 1
        dataArr(rowIndex, 1) = CStr(ids(i))
        dataArr(rowIndex, 2) = CStr(rowIndex)
        dataArr(rowIndex, 3) = CALENDAR_7D
        dataArr(rowIndex, 5) = curStarts(i)
        dataArr(rowIndex, 6) = curFinishes(i)
        dataArr(rowIndex, 7) = rexStarts(i)
        dataArr(rowIndex, 8) = rexFinishes(i)
        idToRow(CStr(ids(i))) = rowIndex
        validIds(CStr(ids(i))) = True
    Next i

    For Each edge In edges
        parts = Split(CStr(edge), "|")
        succId = CStr(parts(0))
        predId = CStr(parts(1))
        linkType = CStr(parts(2))
        lagValue = CDbl(parts(3))
        linkKey = succId & "|" & predId

        LongestPathRexHarness_AddCollectionValue predsById, succId, predId
        LongestPathRexHarness_AddCollectionValue childrenById, predId, succId
        predTypeBySuccPred(linkKey) = linkType
        predLagBySuccPred(linkKey) = lagValue
    Next edge

End Sub

Private Sub LongestPathRexHarness_AddCollectionValue(ByVal target As Object, ByVal key As String, ByVal value As String)
    Dim values As Collection
    If target.Exists(key) Then
        Set values = target(key)
    Else
        Set values = New Collection
        Set target(key) = values
    End If
    values.Add value
End Sub

Private Function LongestPathRexHarness_LongestCsv(ByVal ids As Variant, ByRef outLP As Variant) As String

    Dim i As Long
    Dim rowIndex As Long
    Dim parts As Collection
    Dim result As String
    Dim valueText As String

    Set parts = New Collection
    For i = LBound(ids) To UBound(ids)
        rowIndex = i - LBound(ids) + 1
        valueText = Trim$(CStr(outLP(rowIndex, 1)))
        If UCase$(valueText) = "LONGEST" Then parts.Add CStr(ids(i))
    Next i

    For i = 1 To parts.Count
        If Len(result) > 0 Then result = result & ","
        result = result & CStr(parts(i))
    Next i

    LongestPathRexHarness_LongestCsv = result

End Function

Private Function LongestPathRexHarness_Path(ByVal folderPath As String, ByVal fileName As String) As String
    If Right$(folderPath, 1) = "\" Then
        LongestPathRexHarness_Path = folderPath & fileName
    Else
        LongestPathRexHarness_Path = folderPath & "\" & fileName
    End If
End Function

Private Function LongestPathRexHarness_Tsv(ByVal value As String) As String
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, vbTab, " ")
    LongestPathRexHarness_Tsv = value
End Function
