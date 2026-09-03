Attribute VB_Name = "mod_VisibleTableHeaderLanguage"
Option Explicit

' Owns physical FR/EN ListObject header transactions for user-visible tables.
' It uses the visible-table schema as the only title catalogue.

Private Const VHL_ERROR_BASE As Long = vbObjectError + 5450
Private Const TMP_PREFIX As String = "__VTS_TMP_"

Public Function VisibleHeaders_CaptureOwnerSnapshot(ByVal ownerKey As String) As Object
    Set VisibleHeaders_CaptureOwnerSnapshot = VisibleHeaders_CaptureTablesSnapshot( _
        VisibleHeaders_TablesForOwner(ownerKey))
End Function

Public Function VisibleHeaders_CaptureAllSnapshot() As Object
    Set VisibleHeaders_CaptureAllSnapshot = VisibleHeaders_CaptureTablesSnapshot( _
        VisibleHeaders_AllTables())
End Function

Public Sub VisibleHeaders_PreflightOwner( _
    ByVal ownerKey As String, _
    ByVal targetLanguage As String)

    VisibleHeaders_PreflightTables VisibleHeaders_TablesForOwner(ownerKey), targetLanguage, True
End Sub

Public Sub VisibleHeaders_PreflightAll(ByVal targetLanguage As String)
    VisibleHeaders_PreflightTables VisibleHeaders_AllTables(), targetLanguage, True
End Sub

Public Sub VisibleHeaders_ApplyOwnerFromSnapshot( _
    ByVal snapshot As Object, _
    ByVal ownerKey As String, _
    ByVal targetLanguage As String)

    Dim tables As Variant
    tables = VisibleHeaders_TablesForOwner(ownerKey)
    VisibleHeaders_ApplyTablesFromSnapshot snapshot, tables, targetLanguage
End Sub

Public Sub VisibleHeaders_ApplyAllFromSnapshot( _
    ByVal snapshot As Object, _
    ByVal targetLanguage As String)

    VisibleHeaders_ApplyTablesFromSnapshot snapshot, VisibleHeaders_AllTables(), targetLanguage
End Sub

Public Sub VisibleHeaders_RestoreSnapshot(ByVal snapshot As Object)
    Dim tables As Variant
    Dim tableKey As Variant
    Dim tableSnapshot As Object
    Dim tableObject As ListObject
    Dim keys As Variant
    Dim i As Long

    If snapshot Is Nothing Then Exit Sub
    tables = snapshot("Tables")

    On Error Resume Next
    For Each tableKey In tables
        Set tableSnapshot = snapshot(CStr(tableKey))
        Set tableObject = VisibleHeaders_FindTable(CStr(tableKey))
        keys = tableSnapshot("ColumnKeys")
        For i = LBound(keys) To UBound(keys)
            If tableObject.ListColumns(i + 1).Name <> CStr(tableSnapshot("Header:" & CStr(i + 1))) Then
                tableObject.ListColumns(i + 1).Name = CStr(tableSnapshot("Header:" & CStr(i + 1)))
            End If
        Next i
        Schema_SetPhysicalHeaderLanguageOverride CStr(tableKey), CStr(tableSnapshot("PhysicalLanguage"))
    Next tableKey
    On Error GoTo 0
End Sub

Public Function VisibleHeaders_CurrentTableLanguage(ByVal tableKey As String) As String
    Dim tableObject As ListObject
    Set tableObject = VisibleHeaders_FindTable(tableKey)
    VisibleHeaders_CurrentTableLanguage = VisibleHeaders_DetectTableLanguage(tableObject, tableKey)
End Function

Public Sub VisibleHeaders_ValidateTable( _
    ByVal tableKey As String, _
    ByVal expectedLanguage As String)

    Dim tableObject As ListObject
    Dim detectedLanguage As String

    Set tableObject = VisibleHeaders_FindTable(tableKey)
    detectedLanguage = VisibleHeaders_DetectTableLanguage(tableObject, tableKey)
    If detectedLanguage <> VisibleHeaders_NormalizeLanguage(expectedLanguage) Then
        Err.Raise VHL_ERROR_BASE + 1, "VisibleHeaders_ValidateTable", _
            "Table '" & tableKey & "' is physically '" & detectedLanguage & _
            "' but expected '" & VisibleHeaders_NormalizeLanguage(expectedLanguage) & "'."
    End If
End Sub

Public Function VisibleHeaders_HeaderHash(ByVal tableKey As String) As String
    Dim tableObject As ListObject
    Dim i As Long
    Dim textValue As String

    Set tableObject = VisibleHeaders_FindTable(tableKey)
    For i = 1 To tableObject.ListColumns.Count
        textValue = textValue & "|" & tableObject.ListColumns(i).Name
    Next i
    VisibleHeaders_HeaderHash = CStr(Len(textValue)) & ":" & CStr(VisibleHeaders_TextChecksum(textValue))
End Function

Public Sub VisibleHeaders_MaterializePersistedLanguages()
    Dim tables As Variant
    Dim tableKey As Variant
    Dim targetLanguage As String
    Dim snapshot As Object

    tables = VisibleHeaders_AllTables()
    For Each tableKey In tables
        targetLanguage = SchemaPhysicalHeaderLanguageForTable(CStr(tableKey))
        Set snapshot = VisibleHeaders_CaptureTablesSnapshot(Array(CStr(tableKey)))
        VisibleHeaders_ApplyTablesFromSnapshot snapshot, Array(CStr(tableKey)), targetLanguage
        Set snapshot = Nothing
    Next tableKey
    Schema_ClearPhysicalHeaderLanguageOverrides
End Sub

Private Sub VisibleHeaders_PreflightTables( _
    ByVal tables As Variant, _
    ByVal targetLanguage As String, _
    ByVal requirePersistedMatch As Boolean)

    Dim tableKey As Variant
    Dim tableObject As ListObject
    Dim detectedLanguage As String
    Dim persistedLanguage As String

    Call VisibleHeaders_NormalizeLanguage(targetLanguage)
    For Each tableKey In tables
        Set tableObject = VisibleHeaders_FindTable(CStr(tableKey))
        If tableObject.Parent.ProtectContents Then
            Err.Raise VHL_ERROR_BASE + 15, "VisibleHeaders_PreflightTables", _
                "Worksheet '" & tableObject.Parent.Name & "' is protected; table '" & _
                CStr(tableKey) & "' headers cannot be renamed."
        End If
        detectedLanguage = VisibleHeaders_DetectTableLanguage(tableObject, CStr(tableKey))
        If requirePersistedMatch Then
            persistedLanguage = SchemaPhysicalHeaderLanguageForTable(CStr(tableKey))
            If detectedLanguage <> persistedLanguage Then
                Err.Raise VHL_ERROR_BASE + 2, "VisibleHeaders_PreflightTables", _
                    "Physical header language mismatch for table '" & CStr(tableKey) & _
                    "'. Persisted owner language is '" & persistedLanguage & _
                    "' but headers are '" & detectedLanguage & "'."
            End If
        End If
    Next tableKey
End Sub

Private Sub VisibleHeaders_ApplyTablesFromSnapshot( _
    ByVal snapshot As Object, _
    ByVal tables As Variant, _
    ByVal targetLanguage As String)

    Dim transactionId As String
    Dim tableKey As Variant
    Dim tableObject As ListObject
    Dim tableSnapshot As Object
    Dim keys As Variant
    Dim i As Long
    Dim normalizedLanguage As String

    On Error GoTo Failed

    normalizedLanguage = VisibleHeaders_NormalizeLanguage(targetLanguage)
    If VisibleHeaders_SnapshotAlreadyInLanguage(snapshot, tables, normalizedLanguage) Then
        For Each tableKey In tables
            Schema_SetPhysicalHeaderLanguageOverride CStr(tableKey), normalizedLanguage
        Next tableKey
        Exit Sub
    End If

    transactionId = Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Rnd() * 1000000#))

    For Each tableKey In tables
        Set tableObject = VisibleHeaders_FindTable(CStr(tableKey))
        Set tableSnapshot = snapshot(CStr(tableKey))
        keys = tableSnapshot("ColumnKeys")
        For i = LBound(keys) To UBound(keys)
            tableObject.ListColumns(i + 1).Name = VisibleHeaders_TemporaryTitle( _
                CStr(tableKey), CStr(tableSnapshot("Key:" & CStr(i + 1))), transactionId)
        Next i
    Next tableKey

    For Each tableKey In tables
        Set tableObject = VisibleHeaders_FindTable(CStr(tableKey))
        Set tableSnapshot = snapshot(CStr(tableKey))
        keys = tableSnapshot("ColumnKeys")
        For i = LBound(keys) To UBound(keys)
            tableObject.ListColumns(i + 1).Name = SchemaColumnTitle( _
                CStr(tableKey), CStr(tableSnapshot("Key:" & CStr(i + 1))), normalizedLanguage)
        Next i
        Schema_SetPhysicalHeaderLanguageOverride CStr(tableKey), normalizedLanguage
        VisibleHeaders_ValidateTable CStr(tableKey), normalizedLanguage
    Next tableKey

    VisibleHeaders_AssertNoTemporaryResiduals
    Exit Sub

Failed:
    VisibleHeaders_RestoreSnapshot snapshot
    Err.Raise Err.Number, "VisibleHeaders_ApplyTablesFromSnapshot", Err.Description
End Sub

Private Function VisibleHeaders_SnapshotAlreadyInLanguage( _
    ByVal snapshot As Object, _
    ByVal tables As Variant, _
    ByVal languageCode As String) As Boolean

    Dim tableKey As Variant
    Dim tableSnapshot As Object
    Dim keys As Variant
    Dim i As Long

    VisibleHeaders_SnapshotAlreadyInLanguage = True
    For Each tableKey In tables
        Set tableSnapshot = snapshot(CStr(tableKey))
        If CStr(tableSnapshot("PhysicalLanguage")) <> languageCode Then
            VisibleHeaders_SnapshotAlreadyInLanguage = False
            Exit Function
        End If
        keys = tableSnapshot("ColumnKeys")
        For i = LBound(keys) To UBound(keys)
            If CStr(tableSnapshot("Header:" & CStr(i + 1))) <> SchemaColumnTitle( _
                    CStr(tableKey), CStr(tableSnapshot("Key:" & CStr(i + 1))), languageCode) Then
                VisibleHeaders_SnapshotAlreadyInLanguage = False
                Exit Function
            End If
        Next i
    Next tableKey
End Function

Private Function VisibleHeaders_CaptureTablesSnapshot(ByVal tables As Variant) As Object
    Dim result As Object
    Dim tableSnapshot As Object
    Dim tableKey As Variant
    Dim tableObject As ListObject
    Dim keys As Variant
    Dim physicalKeys As Variant
    Dim i As Long

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbBinaryCompare
    result.Add "Tables", tables

    For Each tableKey In tables
        Set tableObject = VisibleHeaders_FindTable(CStr(tableKey))
        keys = SchemaColumnKeys(CStr(tableKey))
        If tableObject.ListColumns.Count <> UBound(keys) - LBound(keys) + 1 Then
            Err.Raise VHL_ERROR_BASE + 3, "VisibleHeaders_CaptureTablesSnapshot", _
                "Table '" & CStr(tableKey) & "' has " & CStr(tableObject.ListColumns.Count) & _
                " columns; expected " & CStr(UBound(keys) - LBound(keys) + 1) & "."
        End If

        Set tableSnapshot = CreateObject("Scripting.Dictionary")
        tableSnapshot.CompareMode = vbBinaryCompare
        tableSnapshot.Add "PhysicalLanguage", VisibleHeaders_DetectTableLanguage(tableObject, CStr(tableKey))
        physicalKeys = VisibleHeaders_PhysicalColumnKeys(tableObject, CStr(tableKey), CStr(tableSnapshot("PhysicalLanguage")))
        tableSnapshot.Add "ColumnKeys", physicalKeys
        For i = LBound(keys) To UBound(keys)
            tableSnapshot.Add "Key:" & CStr(i + 1), CStr(physicalKeys(i))
            tableSnapshot.Add "Header:" & CStr(i + 1), tableObject.ListColumns(i + 1).Name
        Next i
        result.Add CStr(tableKey), tableSnapshot
    Next tableKey

    Set VisibleHeaders_CaptureTablesSnapshot = result
End Function

Private Function VisibleHeaders_DetectTableLanguage( _
    ByVal tableObject As ListObject, _
    ByVal tableKey As String) As String

    Dim keys As Variant
    Dim i As Long
    Dim enTitle As String
    Dim frTitle As String
    Dim physicalTitle As String
    Dim recoveredTitle As String
    Dim detectedLanguage As String
    Dim seenTitles As Object
    Dim seenKeys As Object
    Dim matchedKey As String
    Dim matchedLanguage As String

    keys = SchemaColumnKeys(tableKey)
    If tableObject.ListColumns.Count <> UBound(keys) - LBound(keys) + 1 Then
        Err.Raise VHL_ERROR_BASE + 4, "VisibleHeaders_DetectTableLanguage", _
            "Table '" & tableKey & "' has " & CStr(tableObject.ListColumns.Count) & _
            " columns; expected " & CStr(UBound(keys) - LBound(keys) + 1) & "."
    End If

    Set seenTitles = CreateObject("Scripting.Dictionary")
    seenTitles.CompareMode = vbBinaryCompare
    Set seenKeys = CreateObject("Scripting.Dictionary")
    seenKeys.CompareMode = vbBinaryCompare

    For i = LBound(keys) To UBound(keys)
        physicalTitle = tableObject.ListColumns(i + 1).Name
        If seenTitles.Exists(physicalTitle) Then
            Err.Raise VHL_ERROR_BASE + 5, "VisibleHeaders_DetectTableLanguage", _
                "Duplicate physical header '" & physicalTitle & "' in table '" & tableKey & "'."
        End If
        seenTitles.Add physicalTitle, True

        matchedKey = VisibleHeaders_MatchColumnKey(tableKey, physicalTitle, matchedLanguage)
        If Len(matchedKey) = 0 Then
            recoveredTitle = VisibleHeaders_RecoverImportedMojibake(physicalTitle)
            If StrComp(recoveredTitle, physicalTitle, vbBinaryCompare) <> 0 Then
                matchedKey = VisibleHeaders_MatchColumnKey(tableKey, recoveredTitle, matchedLanguage)
            End If
        End If
        If Len(matchedKey) = 0 Then
            enTitle = VisibleHeaders_ExpectedTitleList(tableKey, VTS_LANG_EN)
            frTitle = VisibleHeaders_ExpectedTitleList(tableKey, VTS_LANG_FR)
            Err.Raise VHL_ERROR_BASE + 6, "VisibleHeaders_DetectTableLanguage", _
                "Unknown physical header in table '" & tableKey & "' at column " & _
                CStr(i + 1) & ". Observed '" & physicalTitle & "'. Expected one of EN {" & _
                enTitle & "} or FR {" & frTitle & "}."
        End If
        If seenKeys.Exists(matchedKey) Then
            Err.Raise VHL_ERROR_BASE + 12, "VisibleHeaders_DetectTableLanguage", _
                "Duplicate schema column key '" & matchedKey & "' in table '" & tableKey & "'."
        End If
        seenKeys.Add matchedKey, True

        If matchedLanguage = "BOTH" Then
            ' Neutral headers such as ID, WBS or Date do not decide table language.
        ElseIf matchedLanguage = VTS_LANG_EN Then
            If detectedLanguage = "" Then detectedLanguage = VTS_LANG_EN
            If detectedLanguage <> VTS_LANG_EN Then GoTo MixedState
        ElseIf matchedLanguage = VTS_LANG_FR Then
            If detectedLanguage = "" Then detectedLanguage = VTS_LANG_FR
            If detectedLanguage <> VTS_LANG_FR Then GoTo MixedState
        End If
    Next i

    For i = LBound(keys) To UBound(keys)
        If Not seenKeys.Exists(CStr(keys(i))) Then
            Err.Raise VHL_ERROR_BASE + 13, "VisibleHeaders_DetectTableLanguage", _
                "Missing schema column key '" & CStr(keys(i)) & "' in table '" & tableKey & "'."
        End If
    Next i

    If detectedLanguage = "" Then detectedLanguage = SchemaPhysicalHeaderLanguageForTable(tableKey)
    VisibleHeaders_DetectTableLanguage = detectedLanguage
    Exit Function

MixedState:
    Err.Raise VHL_ERROR_BASE + 7, "VisibleHeaders_DetectTableLanguage", _
        "Mixed EN/FR physical headers in table '" & tableKey & "'."
End Function

Private Function VisibleHeaders_PhysicalColumnKeys( _
    ByVal tableObject As ListObject, _
    ByVal tableKey As String, _
    ByVal languageCode As String) As Variant

    Dim keys As Variant
    Dim physicalKeys() As String
    Dim i As Long
    Dim matchedLanguage As String
    Dim physicalTitle As String
    Dim recoveredTitle As String

    keys = SchemaColumnKeys(tableKey)
    ReDim physicalKeys(LBound(keys) To UBound(keys))
    For i = LBound(keys) To UBound(keys)
        physicalTitle = tableObject.ListColumns(i + 1).Name
        physicalKeys(i) = VisibleHeaders_MatchColumnKey( _
            tableKey, physicalTitle, matchedLanguage)
        If Len(physicalKeys(i)) = 0 Then
            recoveredTitle = VisibleHeaders_RecoverImportedMojibake(physicalTitle)
            If StrComp(recoveredTitle, physicalTitle, vbBinaryCompare) <> 0 Then
                physicalKeys(i) = VisibleHeaders_MatchColumnKey( _
                    tableKey, recoveredTitle, matchedLanguage)
            End If
        End If
        If (matchedLanguage <> languageCode And matchedLanguage <> "BOTH") Or Len(physicalKeys(i)) = 0 Then
            Err.Raise VHL_ERROR_BASE + 14, "VisibleHeaders_PhysicalColumnKeys", _
                "Unable to map physical header '" & physicalTitle & _
                "' in table '" & tableKey & "' to language '" & languageCode & "'."
        End If
    Next i
    VisibleHeaders_PhysicalColumnKeys = physicalKeys
End Function

Private Function VisibleHeaders_RecoverImportedMojibake(ByVal textValue As String) As String
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HA9), "é")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HA8), "è")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HAA), "ê")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HAB), "ë")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HA0), "à")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HA2), "â")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HB9), "ù")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HBB), "û")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HB4), "ô")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HAE), "î")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HAF), "ï")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&HA7), "ç")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&H2030), "É")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&H2C6), "È")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&H160), "Ê")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&H2039), "Ë")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&H20AC), "À")
    textValue = Replace(textValue, ChrW$(&HC3) & ChrW$(&H2021), "Ç")
    textValue = Replace(textValue, ChrW$(&HE2) & ChrW$(&H20AC) & ChrW$(&H2122), "'")
    textValue = Replace(textValue, ChrW$(&HE2) & ChrW$(&H20AC) & ChrW$(&H201C), "-")
    textValue = Replace(textValue, ChrW$(&HE2) & ChrW$(&H20AC) & ChrW$(&H201D), "-")
    textValue = Replace(textValue, ChrW$(&HC2), vbNullString)
    VisibleHeaders_RecoverImportedMojibake = textValue
End Function

Private Function VisibleHeaders_MatchColumnKey( _
    ByVal tableKey As String, _
    ByVal physicalTitle As String, _
    ByRef matchedLanguage As String) As String

    Dim keys As Variant
    Dim key As Variant

    keys = SchemaColumnKeys(tableKey)
    For Each key In keys
        If StrComp(SchemaColumnTitle(tableKey, CStr(key), VTS_LANG_EN), _
                   SchemaColumnTitle(tableKey, CStr(key), VTS_LANG_FR), vbBinaryCompare) = 0 And _
           StrComp(physicalTitle, SchemaColumnTitle(tableKey, CStr(key), VTS_LANG_EN), vbBinaryCompare) = 0 Then
            matchedLanguage = "BOTH"
            VisibleHeaders_MatchColumnKey = CStr(key)
            Exit Function
        End If
        If StrComp(physicalTitle, SchemaColumnTitle(tableKey, CStr(key), VTS_LANG_EN), vbBinaryCompare) = 0 Then
            matchedLanguage = VTS_LANG_EN
            VisibleHeaders_MatchColumnKey = CStr(key)
            Exit Function
        End If
        If StrComp(physicalTitle, SchemaColumnTitle(tableKey, CStr(key), VTS_LANG_FR), vbBinaryCompare) = 0 Then
            matchedLanguage = VTS_LANG_FR
            VisibleHeaders_MatchColumnKey = CStr(key)
            Exit Function
        End If
    Next key
    matchedLanguage = ""
End Function

Private Function VisibleHeaders_ExpectedTitleList( _
    ByVal tableKey As String, _
    ByVal languageCode As String) As String

    Dim keys As Variant
    Dim key As Variant
    Dim result As String

    keys = SchemaColumnKeys(tableKey)
    For Each key In keys
        If Len(result) > 0 Then result = result & " / "
        result = result & SchemaColumnTitle(tableKey, CStr(key), languageCode)
    Next key
    VisibleHeaders_ExpectedTitleList = result
End Function

Private Function VisibleHeaders_FindTable(ByVal tableKey As String) As ListObject
    Dim worksheetObject As Worksheet
    Dim tableObject As ListObject

    For Each worksheetObject In ThisWorkbook.Worksheets
        For Each tableObject In worksheetObject.ListObjects
            If StrComp(tableObject.Name, tableKey, vbBinaryCompare) = 0 Then
                Set VisibleHeaders_FindTable = tableObject
                Exit Function
            End If
        Next tableObject
    Next worksheetObject

    Err.Raise VHL_ERROR_BASE + 8, "VisibleHeaders_FindTable", _
        "Visible table '" & tableKey & "' was not found."
End Function

Private Function VisibleHeaders_TablesForOwner(ByVal ownerKey As String) As Variant
    Select Case UCase$(Trim$(ownerKey))
        Case "WBS"
            VisibleHeaders_TablesForOwner = Array(VTS_TABLE_WBS)
        Case "SCURVE", "S-CURVE"
            VisibleHeaders_TablesForOwner = Array(VTS_TABLE_SCURVE)
        Case "CONSTRAINTS"
            VisibleHeaders_TablesForOwner = Array(VTS_TABLE_CONSTRAINTS)
        Case "EVENT", "EVENT_HISTORY"
            VisibleHeaders_TablesForOwner = Array(VTS_TABLE_EVENT_HISTORY, VTS_TABLE_EVENT_ACK)
        Case "DASHBOARD", "GANTT"
            VisibleHeaders_TablesForOwner = Array()
        Case Else
            Err.Raise VHL_ERROR_BASE + 9, "VisibleHeaders_TablesForOwner", _
                "Unknown language owner '" & ownerKey & "'."
    End Select
End Function

Private Function VisibleHeaders_AllTables() As Variant
    VisibleHeaders_AllTables = Array( _
        VTS_TABLE_WBS, VTS_TABLE_SCURVE, VTS_TABLE_CONSTRAINTS, _
        VTS_TABLE_EVENT_HISTORY, VTS_TABLE_EVENT_ACK)
End Function

Private Function VisibleHeaders_NormalizeLanguage(ByVal languageCode As String) As String
    Select Case UCase$(Trim$(languageCode))
        Case VTS_LANG_FR
            VisibleHeaders_NormalizeLanguage = VTS_LANG_FR
        Case VTS_LANG_EN
            VisibleHeaders_NormalizeLanguage = VTS_LANG_EN
        Case Else
            Err.Raise VHL_ERROR_BASE + 10, "VisibleHeaders_NormalizeLanguage", _
                "Unsupported target language '" & languageCode & "'. Expected EN or FR."
    End Select
End Function

Private Function VisibleHeaders_TemporaryTitle( _
    ByVal tableKey As String, _
    ByVal columnKey As String, _
    ByVal transactionId As String) As String

    VisibleHeaders_TemporaryTitle = TMP_PREFIX & Replace(tableKey, "tbl_", "") & "_" & _
        columnKey & "_" & transactionId
End Function

Private Sub VisibleHeaders_AssertNoTemporaryResiduals()
    Dim worksheetObject As Worksheet
    Dim tableObject As ListObject
    Dim columnObject As ListColumn

    For Each worksheetObject In ThisWorkbook.Worksheets
        For Each tableObject In worksheetObject.ListObjects
            If SchemaIsVisibleTable(tableObject.Name) Then
                For Each columnObject In tableObject.ListColumns
                    If Left$(columnObject.Name, Len(TMP_PREFIX)) = TMP_PREFIX Then
                        Err.Raise VHL_ERROR_BASE + 11, "VisibleHeaders_AssertNoTemporaryResiduals", _
                            "Temporary header residual '" & columnObject.Name & _
                            "' remains in table '" & tableObject.Name & "'."
                    End If
                Next columnObject
            End If
        Next tableObject
    Next worksheetObject
End Sub

Public Sub VisibleHeaders_MigrateLegacyLocalizedTitleRows()
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldCalculation As XlCalculation
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    On Error GoTo Failed

    ' Preflight every sheet before the first destructive edit. This routine is
    ' intentionally not called by Open, Reset, Update or language workflows.
    VisibleHeaders_PreflightLegacyRow "WBS", VTS_TABLE_WBS, 4, 5, 4
    VisibleHeaders_PreflightLegacyRow "SCURVE", VTS_TABLE_SCURVE, 1, 2, 1
    VisibleHeaders_PreflightLegacyRow "CONSTRAINTS", VTS_TABLE_CONSTRAINTS, 1, 2, 1
    VisibleHeaders_PreflightLegacyRow "EVENT_HISTORY", VTS_TABLE_EVENT_HISTORY, 3, 4, 3
    VisibleHeaders_PreflightLegacyRow "EVENT_ACK", VTS_TABLE_EVENT_ACK, 2, 3, 2

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    VisibleHeaders_DeleteLegacyRowIfPresent "WBS", VTS_TABLE_WBS, 4, 5, 4
    VisibleHeaders_DeleteLegacyRowIfPresent "SCURVE", VTS_TABLE_SCURVE, 1, 2, 1
    VisibleHeaders_DeleteLegacyRowIfPresent "CONSTRAINTS", VTS_TABLE_CONSTRAINTS, 1, 2, 1
    VisibleHeaders_DeleteLegacyRowIfPresent "EVENT_HISTORY", VTS_TABLE_EVENT_HISTORY, 3, 4, 3
    VisibleHeaders_DeleteLegacyRowIfPresent "EVENT_ACK", VTS_TABLE_EVENT_ACK, 2, 3, 2

CleanExit:
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Sub VisibleHeaders_PreflightLegacyRow( _
    ByVal worksheetName As String, _
    ByVal tableKey As String, _
    ByVal legacyRow As Long, _
    ByVal legacyHeaderRow As Long, _
    ByVal migratedHeaderRow As Long)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim candidate As Range
    Dim usedLastColumn As Long
    Dim i As Long
    Dim shapeObject As Shape
    Dim cellText As String
    Dim keys As Variant

    Set ws = ThisWorkbook.Worksheets(worksheetName)
    Set tbl = ws.ListObjects(tableKey)

    If tbl.HeaderRowRange.Row = migratedHeaderRow Then Exit Sub
    If tbl.HeaderRowRange.Row <> legacyHeaderRow Then
        Err.Raise VHL_ERROR_BASE + 20, "VisibleHeaders_PreflightLegacyRow", _
            "Unexpected physical header row for table '" & tableKey & "'. Found " & _
            CStr(tbl.HeaderRowRange.Row) & ", expected " & CStr(legacyHeaderRow) & _
            " before migration or " & CStr(migratedHeaderRow) & " after migration."
    End If
    If legacyRow <> tbl.HeaderRowRange.Row - 1 Then
        Err.Raise VHL_ERROR_BASE + 21, "VisibleHeaders_PreflightLegacyRow", _
            "The candidate legacy row is not immediately above table '" & tableKey & "'."
    End If
    keys = SchemaColumnKeys(tableKey)
    If tbl.ListColumns.Count <> UBound(keys) - LBound(keys) + 1 Then
        Err.Raise VHL_ERROR_BASE + 22, "VisibleHeaders_PreflightLegacyRow", _
            "Unexpected column count in table '" & tableKey & "'."
    End If

    Set candidate = ws.Cells(legacyRow, tbl.Range.Column).Resize(1, tbl.ListColumns.Count)
    If CBool(candidate.MergeCells) Then
        Err.Raise VHL_ERROR_BASE + 23, "VisibleHeaders_PreflightLegacyRow", _
            "Merged cells exist in the candidate legacy row for table '" & tableKey & "'."
    End If
    If Application.WorksheetFunction.CountA(candidate) <> tbl.ListColumns.Count Then
        Err.Raise VHL_ERROR_BASE + 24, "VisibleHeaders_PreflightLegacyRow", _
            "The candidate legacy row for table '" & tableKey & _
            "' does not contain exactly one label per visible column."
    End If

    usedLastColumn = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If tbl.Range.Column > 1 Then
        If Application.WorksheetFunction.CountA( _
                ws.Cells(legacyRow, 1).Resize(1, tbl.Range.Column - 1)) > 0 Then
            Err.Raise VHL_ERROR_BASE + 25, "VisibleHeaders_PreflightLegacyRow", _
                "Unexpected data exists before the legacy labels on sheet '" & worksheetName & "'."
        End If
    End If
    If usedLastColumn > tbl.Range.Column + tbl.ListColumns.Count - 1 Then
        If Application.WorksheetFunction.CountA(ws.Cells(legacyRow, _
                tbl.Range.Column + tbl.ListColumns.Count).Resize(1, _
                usedLastColumn - tbl.Range.Column - tbl.ListColumns.Count + 1)) > 0 Then
            Err.Raise VHL_ERROR_BASE + 26, "VisibleHeaders_PreflightLegacyRow", _
                "Unexpected data exists after the legacy labels on sheet '" & worksheetName & "'."
        End If
    End If

    For i = 1 To tbl.ListColumns.Count
        cellText = Trim$(CStr(candidate.Cells(1, i).Value2))
        If Not VisibleHeaders_IsKnownLocalizedTitle(tableKey, cellText) Then
            Err.Raise VHL_ERROR_BASE + 27, "VisibleHeaders_PreflightLegacyRow", _
                "Unknown content '" & cellText & "' in the candidate legacy row for table '" & _
                tableKey & "' at column " & CStr(i) & "."
        End If
    Next i

    For Each shapeObject In ws.Shapes
        If shapeObject.TopLeftCell.Row <= legacyRow And _
           shapeObject.BottomRightCell.Row >= legacyRow Then
            Err.Raise VHL_ERROR_BASE + 28, "VisibleHeaders_PreflightLegacyRow", _
                "Shape '" & shapeObject.Name & "' intersects the candidate legacy row on sheet '" & _
                worksheetName & "'."
        End If
    Next shapeObject
End Sub

Private Sub VisibleHeaders_DeleteLegacyRowIfPresent( _
    ByVal worksheetName As String, _
    ByVal tableKey As String, _
    ByVal legacyRow As Long, _
    ByVal legacyHeaderRow As Long, _
    ByVal migratedHeaderRow As Long)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim columnCount As Long
    Dim rowCount As Long
    Dim headerText As String

    Set ws = ThisWorkbook.Worksheets(worksheetName)
    Set tbl = ws.ListObjects(tableKey)
    If tbl.HeaderRowRange.Row = migratedHeaderRow Then Exit Sub

    columnCount = tbl.ListColumns.Count
    rowCount = tbl.ListRows.Count
    headerText = VisibleHeaders_HeaderText(tbl)
    ws.Rows(legacyRow).Delete Shift:=xlUp

    Set tbl = ws.ListObjects(tableKey)
    If tbl.HeaderRowRange.Row <> migratedHeaderRow Or _
       tbl.ListColumns.Count <> columnCount Or _
       tbl.ListRows.Count <> rowCount Or _
       VisibleHeaders_HeaderText(tbl) <> headerText Then
        Err.Raise VHL_ERROR_BASE + 29, "VisibleHeaders_DeleteLegacyRowIfPresent", _
            "Post-delete structure validation failed for table '" & tableKey & "'."
    End If
End Sub

Private Function VisibleHeaders_HeaderText(ByVal tableObject As ListObject) As String
    Dim i As Long

    For i = 1 To tableObject.ListColumns.Count
        VisibleHeaders_HeaderText = VisibleHeaders_HeaderText & ChrW$(30) & _
            tableObject.ListColumns(i).Name
    Next i
End Function

Private Function VisibleHeaders_IsKnownLocalizedTitle( _
    ByVal tableKey As String, _
    ByVal candidateTitle As String) As Boolean

    Dim keys As Variant
    Dim key As Variant

    keys = SchemaColumnKeys(tableKey)
    For Each key In keys
        If StrComp(candidateTitle, SchemaColumnTitle(tableKey, CStr(key), VTS_LANG_FR), _
                vbBinaryCompare) = 0 Or _
           StrComp(candidateTitle, SchemaColumnTitle(tableKey, CStr(key), VTS_LANG_EN), _
                vbBinaryCompare) = 0 Then
            VisibleHeaders_IsKnownLocalizedTitle = True
            Exit Function
        End If
    Next key
End Function

Private Function VisibleHeaders_TextChecksum(ByVal textValue As String) As Long
    Dim i As Long
    Dim checksum As Long

    For i = 1 To Len(textValue)
        checksum = ((checksum * 33) Xor AscW(Mid$(textValue, i, 1))) And &H7FFFFFFF
    Next i
    VisibleHeaders_TextChecksum = checksum
End Function
